import ExcelJS from 'exceljs'
import { saveAs } from 'file-saver'

export async function exportSpreadsheetToXlsx(
    jsonResponse: unknown,
    fileName: string,
): Promise<void> {
    const workbookData = extractWorkbookData(jsonResponse)
    const wb = new ExcelJS.Workbook()
    wb.creator = 'Excel Sheet App'
    wb.created = new Date()

    for (const sheet of workbookData.sheets) {
        const ws = wb.addWorksheet(sheet.name || 'Sheet')

        for (const row of sheet.rows) {
            const rowValues: (string | number | boolean | null)[] = []
            for (const cell of row.cells) {
                rowValues.push(cell.value ?? null)
            }
            if (rowValues.some((v) => v !== null)) {
                ws.addRow(rowValues)
            }
        }

        ws.columns.forEach((column) => {
            let maxLength = 10
            column.eachCell?.({ includeEmpty: false }, (cell) => {
                const cellLen = cell.value ? String(cell.value).length : 0
                if (cellLen > maxLength) maxLength = cellLen
            })
            column.width = Math.min(maxLength + 2, 50)
        })
    }

    const buffer = await wb.xlsx.writeBuffer()
    const blob = new Blob([buffer], {
        type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
    })
    saveAs(blob, `${fileName}.xlsx`)
}

interface CellData {
    value: string | number | boolean | null
    formula?: string
}

interface RowData {
    index: number
    cells: CellData[]
}

interface SheetData {
    name: string
    rows: RowData[]
}

interface WorkbookExtract {
    sheets: SheetData[]
}

function extractWorkbookData(jsonResponse: unknown): WorkbookExtract {
    const sheets: SheetData[] = []

    const root = jsonResponse as Record<string, unknown>
    const jsonObject = (root?.jsonObject ?? root) as Record<string, unknown>
    const workbook = (jsonObject?.Workbook ?? jsonObject) as Record<string, unknown>
    const rawSheets = (workbook?.sheets ?? []) as Record<string, unknown>[]

    for (const rawSheet of rawSheets) {
        const sheetName = (rawSheet.name as string) || 'Sheet'
        const rawRows = (rawSheet.rows ?? []) as Record<string, unknown>[]
        const rows: RowData[] = []

        for (let rowIdx = 0; rowIdx < rawRows.length; rowIdx++) {
            const rawRow = rawRows[rowIdx]
            if (!rawRow) continue

            const rawCells = (rawRow.cells ?? []) as Record<string, unknown>[]
            const cells: CellData[] = []

            for (const rawCell of rawCells) {
                if (!rawCell) {
                    cells.push({ value: null })
                    continue
                }

                const formula = rawCell.formula as string | undefined
                let value: string | number | boolean | null = null

                if (rawCell.value !== undefined && rawCell.value !== null) {
                    value = rawCell.value as string | number | boolean
                }

                cells.push({ value, formula })
            }

            rows.push({ index: rowIdx, cells })
        }

        sheets.push({ name: sheetName, rows })
    }

    if (sheets.length === 0) {
        sheets.push({ name: 'Sheet1', rows: [] })
    }

    return { sheets }
}
