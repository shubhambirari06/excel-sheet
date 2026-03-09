import ExcelJS from 'exceljs'
import type { CellFormulaValue, CellRichTextValue } from 'exceljs'

export interface ParsedCell {
    address: string
    value: ExcelJS.CellValue
    formula?: string
    numFmt?: string
    style: Partial<ExcelJS.Style>
    isMerged: boolean
    masterAddress?: string
}

export interface ParsedSheet {
    name: string
    rowCount: number
    columnCount: number
    merges: string[]
    rows: Map<number, Map<number, ParsedCell>>
    columnWidths: Map<number, number>
    rowHeights: Map<number, number>
}

export interface ParsedWorkbook {
    sheets: ParsedSheet[]
    workbook: ExcelJS.Workbook
}

export async function parseExcelFile(file: File): Promise<ParsedWorkbook> {
    const buffer = await file.arrayBuffer()
    const workbook = new ExcelJS.Workbook()
    await workbook.xlsx.load(buffer)

    const sheets: ParsedSheet[] = []

    workbook.eachSheet((worksheet) => {
        const merges: string[] = []
        const rows = new Map<number, Map<number, ParsedCell>>()
        const columnWidths = new Map<number, number>()
        const rowHeights = new Map<number, number>()

        worksheet.columns.forEach((col, idx) => {
            if (col.width) {
                columnWidths.set(idx + 1, col.width)
            }
        })

        const mergeMap = new Map<string, string>()
        const model = worksheet.model as unknown as Record<string, unknown> | undefined
        if (model?.merges) {
            const rawMerges = model.merges as string[]
            rawMerges.forEach((m) => {
                merges.push(m)
                const [start, end] = m.split(':')
                const startCell = worksheet.getCell(start)
                const endCell = worksheet.getCell(end)
                const startRow = Number(startCell.row)
                const endRow = Number(endCell.row)
                const startCol = Number(startCell.col)
                const endCol = Number(endCell.col)
                for (let r = startRow; r <= endRow; r++) {
                    for (let c = startCol; c <= endCol; c++) {
                        const addr = worksheet.getCell(r, c).address
                        if (addr !== start) {
                            mergeMap.set(addr, start)
                        }
                    }
                }
            })
        }

        worksheet.eachRow({ includeEmpty: true }, (row, rowNumber) => {
            if (row.height) {
                rowHeights.set(rowNumber, row.height)
            }

            const cellMap = new Map<number, ParsedCell>()

            row.eachCell({ includeEmpty: true }, (cell, colNumber) => {
                const parsed: ParsedCell = {
                    address: cell.address,
                    value: cell.value,
                    style: cell.style ?? {},
                    isMerged: mergeMap.has(cell.address),
                    masterAddress: mergeMap.get(cell.address),
                }

                if (
                    cell.value &&
                    typeof cell.value === 'object' &&
                    'formula' in cell.value
                ) {
                    parsed.formula = (cell.value as CellFormulaValue).formula
                }

                if (cell.numFmt) {
                    parsed.numFmt = cell.numFmt
                }

                cellMap.set(colNumber, parsed)
            })

            rows.set(rowNumber, cellMap)
        })

        sheets.push({
            name: worksheet.name,
            rowCount: worksheet.rowCount,
            columnCount: worksheet.columnCount,
            merges,
            rows,
            columnWidths,
            rowHeights,
        })
    })

    return { sheets, workbook }
}

export function buildWorkbookJson(parsed: ParsedWorkbook): object {
    const sheets = parsed.sheets.map((sheet) => {
        const rowsArr: Record<string, unknown>[] = []

        sheet.rows.forEach((cellMap, rowIndex) => {
            const cells: Record<string, unknown>[] = []

            cellMap.forEach((cell, colIndex) => {
                const cellObj: Record<string, unknown> = {
                    index: colIndex - 1,
                }

                if (cell.formula) {
                    cellObj.formula = cell.formula
                    if (
                        cell.value &&
                        typeof cell.value === 'object' &&
                        'result' in (cell.value as unknown as Record<string, unknown>)
                    ) {
                        const result = (cell.value as CellFormulaValue).result
                        cellObj.value = convertValue(result)
                    }
                } else if (cell.value !== null && cell.value !== undefined) {
                    cellObj.value = convertCellValue(cell.value)
                }

                if (cell.numFmt) {
                    cellObj.format = cell.numFmt
                }

                cells.push(cellObj)
            })

            rowsArr.push({
                index: rowIndex - 1,
                cells,
            })
        })

        const columns = Array.from(sheet.columnWidths.entries()).map(
            ([idx, width]) => ({
                index: idx - 1,
                width: Math.round(width * 7.5),
            }),
        )

        return {
            name: sheet.name,
            rows: rowsArr,
            columns,
        }
    })

    return { Workbook: { sheets } }
}

function convertCellValue(val: ExcelJS.CellValue): unknown {
    if (val === null || val === undefined) return ''
    if (typeof val === 'string' || typeof val === 'number' || typeof val === 'boolean') {
        return val
    }
    if (val instanceof Date) {
        return val.toLocaleDateString()
    }
    if (typeof val === 'object') {
        if ('richText' in val) {
            const rt = (val as CellRichTextValue).richText
            return rt.map((seg) => seg.text).join('')
        }
        if ('formula' in val) {
            return convertValue((val as CellFormulaValue).result)
        }
        if ('error' in val) {
            return (val as unknown as Record<string, unknown>).error
        }
        if ('hyperlink' in val) {
            const hv = val as unknown as Record<string, unknown>
            return hv.text ?? hv.hyperlink ?? ''
        }
    }
    return String(val)
}

function convertValue(val: unknown): unknown {
    if (val === null || val === undefined) return ''
    if (val instanceof Date) return val.toLocaleDateString()
    if (typeof val === 'object') return String(val)
    return val
}
