import ExcelJS from 'exceljs'

export interface ParsedSheet {
    name: string
    rows: CellObject[][]
}

interface CellObject {
    value: string | number | boolean | null
    formula?: string
}

export async function parseExcelFile(file: File): Promise<ParsedSheet[]> {
    const ext = file.name.substring(file.name.lastIndexOf('.')).toLowerCase()

    if (ext === '.csv') {
        return parseCsvFile(file)
    }

    return parseXlsxFile(file)
}

async function parseXlsxFile(file: File): Promise<ParsedSheet[]> {
    const arrayBuffer = await file.arrayBuffer()
    const wb = new ExcelJS.Workbook()
    await wb.xlsx.load(arrayBuffer)

    const sheets: ParsedSheet[] = []

    wb.eachSheet((worksheet) => {
        const rows: CellObject[][] = []
        const maxRow = worksheet.rowCount
        const maxCol = worksheet.columnCount

        for (let r = 1; r <= maxRow; r++) {
            const row = worksheet.getRow(r)
            const cells: CellObject[] = []

            for (let c = 1; c <= maxCol; c++) {
                const cell = row.getCell(c)
                const cellObj: CellObject = { value: null }

                const cellFormula = (cell as unknown as Record<string, unknown>).formula
                if (typeof cellFormula === 'string' && cellFormula.length > 0) {
                    cellObj.formula = cellFormula
                    cellObj.value = (cell.result as string | number | boolean) ?? null
                } else if (cell.value !== null && cell.value !== undefined) {
                    if (cell.value instanceof Date) {
                        cellObj.value = cell.value.toISOString().split('T')[0]
                    }
                    else if (
                        typeof cell.value === 'object' &&
                        'richText' in (cell.value as object)
                    ) {
                        const rt = (cell.value as { richText: { text: string }[] }).richText
                        cellObj.value = rt.map((seg) => seg.text).join('')
                    }
                    else if (
                        typeof cell.value === 'object' &&
                        'result' in (cell.value as object)
                    ) {
                        cellObj.value = (cell.value as { result: unknown }).result as
                            | string
                            | number
                            | boolean
                            | null
                    }
                    else if (typeof cell.value !== 'object') {
                        cellObj.value = cell.value as string | number | boolean
                    }
                }

                cells.push(cellObj)
            }

            rows.push(cells)
        }

        sheets.push({
            name: worksheet.name || 'Sheet',
            rows,
        })
    })

    if (sheets.length === 0) {
        sheets.push({ name: 'Sheet1', rows: [] })
    }

    return sheets
}

async function parseCsvFile(file: File): Promise<ParsedSheet[]> {
    const text = await file.text()
    const lines = text.split('\n').filter((l) => l.trim().length > 0)
    const rows: CellObject[][] = []

    for (const line of lines) {
        const cells = parseCsvLine(line).map((val) => {
            const num = Number(val)
            const value: string | number | null =
                val === '' ? null : isNaN(num) ? val : num
            return { value } as CellObject
        })
        rows.push(cells)
    }

    const nameWithoutExt = file.name.replace(/\.[^/.]+$/, '')
    return [{ name: nameWithoutExt || 'Sheet1', rows }]
}

function parseCsvLine(line: string): string[] {
    const result: string[] = []
    let current = ''
    let inQuotes = false

    for (let i = 0; i < line.length; i++) {
        const ch = line[i]
        if (inQuotes) {
            if (ch === '"' && line[i + 1] === '"') {
                current += '"'
                i++
            } else if (ch === '"') {
                inQuotes = false
            } else {
                current += ch
            }
        } else {
            if (ch === '"') {
                inQuotes = true
            } else if (ch === ',') {
                result.push(current.trim())
                current = ''
            } else {
                current += ch
            }
        }
    }
    result.push(current.trim())
    return result
}

export function buildWorkbookJson(sheets: ParsedSheet[]): { Workbook: unknown } {
    const sheetObjects = sheets.map((sheet) => {
        const rows = sheet.rows.map((row, rowIndex) => {
            const cells = row.map((cellData) => {
                const cell: Record<string, unknown> = {}
                if (cellData.formula) {
                    cell.formula = cellData.formula
                } else if (cellData.value !== undefined && cellData.value !== null) {
                    cell.value = cellData.value
                }
                return cell
            })
            return { index: rowIndex, cells }
        })

        return {
            name: sheet.name || 'Sheet1',
            rows,
            columns: [],
            selectedRange: 'A1:A1',
        }
    })

    return {
        Workbook: {
            sheets: sheetObjects,
            definedNames: [],
        },
    }
}
