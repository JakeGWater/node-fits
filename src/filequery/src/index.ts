import { Workbook, Cell } from 'exceljs'

function CHECK(any: any, message = 'CHECK FAILED'): asserts any {
    if (!any) {
        throw new Error(message)
    }
}

const args = process.argv.slice(2)

class NArray<T> {
    constructor(
        public Height: number,
        public Width: number,
        default_value: T
    ) {
        this.data = []
        for (let i = 0; i < Height; i++) {
            this.data[i] = []
            for (let j = 0; j < Width; j++) {
                this.data[i][j] = default_value
            }
        }
    }
    data: T[][] = []

    Size() {
        return [this.Height, this.Width]
    }

    Get(x: number, y: number): T {
        // console.log({ x, y }, this)
        return this.data[y][x]
    }

    Set(x: number, y: number, value: T) {
        // console.log({ x, y },)
        this.data[y][x] = value
    }

    toString() {
        return this.data.map(row => row.join('')).join('\n')
    }
}

type Frame = string[][]

abstract class Process {
    protected abstract process_frame(frame: Frame): Frame
    Process(frames: Frame[]): Frame[] {
        let out_frames: Frame[] = []
        for (let frame of frames) {
            out_frames.push(this.process_frame(frame))
        }
        return out_frames
    }
}

function DEEP_COPY(frame: Frame) {
    let out_frame: Frame = []
    for (let [key, val] of frame) {
        out_frame.push([key, val])
    }
    return out_frame
}

class ProcessHeader extends Process {
    constructor() { super() }
    process_frame(frame: Frame): Frame {
        let out_frame: Frame = DEEP_COPY(frame)
        if (out_frame.length > 0) {
            out_frame[0][0] = 'Title'
            out_frame[0][1] = frame[0][0]
        }
        return out_frame
    }
}

class ProcessColumns {
    constructor() { }
    Process(frames: Frame[]): Frame[] {
        let out_frames: Frame[] = []

        for (let frame of frames) {
            let frame_first_row = frame[0]
            if (frame_first_row.length === 2) {
                out_frames.push(frame)
            } else {
                for (let i = 1; i < frame_first_row.length; i++) {
                    let out_frame: Frame = []
                    for (let row of frame) {
                        out_frame.push([row[0], row[i]])
                    }
                    out_frames.push(out_frame)
                }
            }
        }

        return out_frames
    }
}

type Optional<T extends {}> = { [P in keyof T]?: NotAFunction<T[P]> }
type NotAFunction<T> = T extends Function ? never : T

class ProcessUnion {
    constructor(
        options: Optional<ProcessUnion> = {}
    ) {
        Object.assign(this, options)
    }
    public DefaultValue = ''
    Process(frames: Frame[]): Frame[] {
        let column_names_set = new Set()
        for (let row of frames) {
            for (let [key, _] of row) {
                column_names_set.add(key)
            }
        }

        let column_names = Array.from(column_names_set).sort()
        let out_table: Frame[] = []

        row: for (let row of frames) {
            let out_row = new Array(column_names.length)
            col_name: for (let col_name of column_names) {
                for (let [key, val] of row) {
                    if (key === col_name) {
                        out_row[column_names.indexOf(col_name)] = [key, val]
                        continue col_name
                    }
                }
                out_row[column_names.indexOf(col_name)] = [col_name, '']
            }
            out_table.push(out_row)
        }

        return out_table
    }
}

function SIZE(frame: Frame) {
    return [frame.length, frame[0].length]
}

class ProcessTabels {
    Process(frame: Frame): Frame[] {
        let [height, width] = SIZE(frame)
        let na = new NArray(height, width, false)
        let ranges: any[] = []
        for (let j = 0; j < width; j++) {
            for (let i = 0; i < height; i++) {
                if (na.Get(j, i)) {
                    // console.log(`Skipping ${i}, ${j}`)
                    continue
                } else {
                    // console.log(`Checking ${i}, ${j}`)
                }

                let cell = frame[i][j]
                if (cell) {
                    // console.log(`Section Start ${i}, ${j} - ${cell.value}`)
                    // this is the start of a section!!!
                    let ti = i // row
                    let tj = j // col
                    let tj_max = j
                    let row_empty = false

                    search: while (true) {
                        // stop if we run out of rows
                        if (ti >= height) {
                            break search
                        }

                        // next line if we run out of columns
                        if (tj >= width) {
                            // stop tho if this row is empty
                            if (row_empty) {
                                break search
                            }

                            tj = j // reset back to start
                            ti++ // move to next row

                            row_empty = true
                            continue
                        }

                        let tcell = frame[ti][tj]
                        na.Set(tj, ti, true)

                        // if the cell has something, advance to the right and increase the max
                        if (tcell) {
                            row_empty = false
                            tj++
                            tj_max = Math.max(tj, tj_max)
                        }
                        // if the cell is empty, but we're less than the max, keep going right
                        else if (tj < tj_max) {
                            tj++
                        }
                        // otherwise go to the next row
                        else {
                            // stop tho if this row is empty
                            if (row_empty) {
                                break search
                            }

                            tj = j
                            ti++

                            row_empty = true
                        }
                    }

                    ranges.push({ head: [i, j], tail: [ti, tj_max] })
                }
            }
        }

        let frames: Frame[] = []
        for (let range of ranges) {
            let [hi, hj] = range.head
            let [ti, tj] = range.tail
            let out: any[] = []
            for (let i = hi; i < ti; i++) {
                let key = frame[i][hj]
                let vals: any = []
                for (let j = hj; j < tj - 1; j++) {
                    vals.push(frame[i][j + 1])
                }
                out.push([key, ...vals])
            }
            frames.push(out)
        }

        return frames
    }
}


class FormatCsv {
    Format(frames: Frame[]): string {
        let out = [
            frames[0].map(([key, _]) => key).join(','),
            frames.map(row => row.map(([_, val]) => val).join(',')).join('\n'),
        ]
        return out.join('\n')
    }
}

async function main(filename: string) {
    let workbook = new Workbook()
    let file = await workbook.xlsx.readFile(filename)

    let sheet = file.getWorksheet(1)

    CHECK(sheet)

    let na = new NArray(sheet.rowCount, sheet.columnCount, false)

    let frame_starts = []

    console.log(`Seaching for frames across ${sheet.rowCount} rows and ${sheet.columnCount} columns`)

    let fframe: Frame = []
    for (let i = 0; i < sheet.rowCount; i++) {
        let frow: string[] = []
        for (let j = 0; j < sheet.columnCount; j++) {
            let cell = sheet.getCell(i + 1, j + 1)
            frow.push(cell.text)
        }
        fframe.push(frow)
    }

    let frames = new ProcessTabels().Process(fframe)

    // Process Columns
    // We assume the leftmost column is the key, and the rest are values
    // We create a frame for each set of values
    frames = new ProcessColumns().Process(frames)

    // We assume the first row is the header, and give that a Title: key
    frames = new ProcessHeader().Process(frames)

    // We unify all the frames to have the same columns in the same order
    // Empty columns will be the empty string
    frames = new ProcessUnion().Process(frames)

    // Format Frames to CSV
    // There is no safeguard to ensure the frames have the same columns in the same order
    // You should send this through a ProcessUnion first
    console.log(new FormatCsv().Format(frames))
}

main(args[0])
