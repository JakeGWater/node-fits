import { Workbook, Cell } from 'exceljs'

import yargs from 'yargs'

const argv = yargs
    .option('file', {
        alias: 'f',
        description: 'The file to query',
        type: 'string',
    })
    .help()
    .alias('help', 'h').argv

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

type Frame = [string, string][]

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

class ProcessUnion {
    constructor() { }
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

    let ranges: any[] = []
    for (let j = 0; j < sheet.columnCount; j++) {
        for (let i = 0; i < sheet.rowCount; i++) {
            if (na.Get(j, i)) {
                // console.log(`Skipping ${i}, ${j}`)
                continue
            } else {
                // console.log(`Checking ${i}, ${j}`)
            }

            let cell = sheet.getCell(i + 1, j + 1)
            if (cell.value) {
                // console.log(`Section Start ${i}, ${j} - ${cell.value}`)
                // this is the start of a section!!!
                let ti = i // row
                let tj = j // col
                let tj_max = j
                let row_empty = false

                search: while (true) {
                    // stop if we run out of rows
                    if (ti >= sheet.rowCount) {
                        break search
                    }

                    // next line if we run out of columns
                    if (tj >= sheet.columnCount) {
                        // stop tho if this row is empty
                        if (row_empty) {
                            break search
                        }

                        tj = j // reset back to start
                        ti++ // move to next row

                        row_empty = true
                        continue
                    }

                    // console.log(`Checking ${ti}, ${tj}`)
                    let tcell = sheet.getCell(ti + 1, tj + 1)
                    na.Set(tj, ti, true)

                    // if the cell has something, advance to the right and increase the max
                    // console.log(`Found ${ti}, ${tj} - ${tcell.text}`)
                    if (tcell.text) {
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

                ranges.push({ head: [i, j], tail: [ti, tj] })
            }
        }
    }

    // Extract Frames from Sheet
    let frames: Frame[] = []
    for (let frame of ranges) {
        let [hi, hj] = frame.head
        let [ti, _] = frame.tail
        let out: any[] = []
        for (let i = hi; i < ti; i++) {
            let key = sheet.getCell(i + 1, hj + 1).text
            let val = sheet.getCell(i + 1, hj + 2).text
            out.push([key, val])
        }
        frames.push(out)
    }

    // Process Frames
    frames = new ProcessHeader().Process(frames)
    frames = new ProcessUnion().Process(frames)

    // Format Frames
    console.log(new FormatCsv().Format(frames))
}

enum CellBorders {
    Top = 1,
    Left = 2,
    Bottom = 4,
    Right = 8,
}

function cellHasBorder(cell: Cell): CellBorders {
    let border = CellBorders.Top | CellBorders.Left | CellBorders.Bottom | CellBorders.Right
    if (cell.border.top) {
        border &= ~CellBorders.Top
    }
    if (cell.border.left) {
        border &= ~CellBorders.Left
    }
    if (cell.border.bottom) {
        border &= ~CellBorders.Bottom
    }
    if (cell.border.right) {
        border &= ~CellBorders.Right
    }
    return border
}

main(args[0])

