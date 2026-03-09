import ExcelJS from 'exceljs'
import type { CellFormulaValue } from 'exceljs'
import { saveAs } from 'file-saver'

export async function exportSpreadsheetToXlsx(
    ssJsonResponse: unknown,
    fileName: string,
    sourceWorkbook?: ExcelJS.Workbook,
): Promise<void> {
    const payload = ssJsonResponse as Record<string, unknown>
    const jsonData = (payload.jsonObject ?? payload) as Record<string, unknown>
    const wbData =
        (jsonData.Workbook as Record<string, unknown> | undefined) ??
        (jsonData.workbook as Record<string, unknown> | undefined)

    if (!wbData?.sheets) {
        throw new Error('No sheet data found in Spreadsheet JSON.')
    }

    const sheetsArr = wbData.sheets as Array<Record<string, unknown> | null | undefined>
    const workbook = sourceWorkbook ?? new ExcelJS.Workbook()

    for (const sheetJson of sheetsArr) {
        if (!sheetJson) continue

        const sheetName = (sheetJson.name as string) || 'Sheet1'
        const rowsArr =
            (sheetJson.rows as Array<Record<string, unknown> | null | undefined>) ?? []

        let worksheet = workbook.getWorksheet(sheetName)
        if (!worksheet) {
            worksheet = workbook.addWorksheet(sheetName)
        }

        const colsArr =
            (sheetJson.columns as Array<Record<string, unknown> | null | undefined>) ?? []
        for (let colPos = 0; colPos < colsArr.length; colPos++) {
            const colDef = colsArr[colPos]
            if (!colDef) continue
            const idx =
                typeof colDef.index === 'number' ? (colDef.index as number) : colPos
            if (idx == null || isNaN(idx)) continue
            const colNum = idx + 1
            const w = colDef.width as number | undefined
            if (w && colNum > 0) {
                try {
                    worksheet.getColumn(colNum).width = Math.round(w / 7.5)
                } catch {
                    // non-critical
                }
            }
        }

        for (let rowPos = 0; rowPos < rowsArr.length; rowPos++) {
            const rowDef = rowsArr[rowPos]
            if (!rowDef) continue

            const rowIdx =
                typeof rowDef.index === 'number' ? (rowDef.index as number) : rowPos
            if (rowIdx == null || isNaN(rowIdx)) continue
            const rowNum = rowIdx + 1
            if (rowNum < 1) continue

            const wsRow = worksheet.getRow(rowNum)

            const rowHeight = rowDef.height as number | undefined
            if (rowHeight && rowHeight > 0) {
                wsRow.height = rowHeight
            }

            const cells =
                (rowDef.cells as Array<Record<string, unknown> | null | undefined>) ?? []

            for (let cellPos = 0; cellPos < cells.length; cellPos++) {
                const cellDef = cells[cellPos]
                if (!cellDef) continue

                const colIdx =
                    typeof cellDef.index === 'number'
                        ? (cellDef.index as number)
                        : cellPos
                if (colIdx == null || isNaN(colIdx)) continue
                const colNum = colIdx + 1
                if (colNum < 1) continue

                const cell = wsRow.getCell(colNum)

                if (typeof cellDef.formula === 'string' && cellDef.formula.trim()) {
                    cell.value = {
                        formula: cellDef.formula as string,
                    } as CellFormulaValue
                } else if (cellDef.value !== undefined) {
                    try {
                        if (!cell.isMerged || cell.address === cell.master?.address) {
                            cell.value = toPrimitive(cellDef.value)
                        }
                    } catch {
                        cell.value = toPrimitive(cellDef.value)
                    }
                }

                if (cellDef.format && !cell.numFmt) {
                    cell.numFmt = cellDef.format as string
                }

                if (!sourceWorkbook) {
                    applyStyleFromJson(cell, cellDef)
                }
            }

            wsRow.commit()
        }
    }

    const buffer = await workbook.xlsx.writeBuffer()
    const blob = new Blob([buffer], {
        type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
    })
    saveAs(blob, `${fileName}.xlsx`)
}

function toPrimitive(val: unknown): ExcelJS.CellValue {
    if (val === null || val === undefined) return null
    if (typeof val === 'string' || typeof val === 'number' || typeof val === 'boolean') {
        return val
    }
    if (val instanceof Date) return val
    if (typeof val === 'object') {
        const obj = val as Record<string, unknown>
        if ('formula' in obj) {
            return { formula: obj.formula as string } as CellFormulaValue
        }
        return String(val)
    }
    return String(val)
}

function applyStyleFromJson(
    cell: ExcelJS.Cell,
    cellDef: Record<string, unknown>,
): void {
    const style = cellDef.style as Record<string, unknown> | undefined
    if (!style) return

    if (style.fontWeight === 'bold' || style.fontFamily || style.fontSize || style.color) {
        cell.font = {
            ...cell.font,
            bold: style.fontWeight === 'bold' || cell.font?.bold,
            name: (style.fontFamily as string) ?? cell.font?.name ?? 'Calibri',
            size: (style.fontSize as number) ?? cell.font?.size ?? 11,
            color: style.color
                ? { argb: (style.color as string).replace('#', 'FF') }
                : cell.font?.color,
        }
    }

    if (style.textAlign || style.verticalAlign) {
        cell.alignment = {
            ...cell.alignment,
            horizontal: mapAlignment(style.textAlign as string),
            vertical: mapVerticalAlignment(style.verticalAlign as string),
            wrapText: (style.whiteSpace as string) === 'normal' || cell.alignment?.wrapText,
        }
    }

    if (style.backgroundColor) {
        const bg = (style.backgroundColor as string).replace('#', '')
        cell.fill = {
            type: 'pattern',
            pattern: 'solid',
            fgColor: { argb: `FF${bg}` },
        }
    }

    if (style.borderTop || style.borderBottom || style.borderLeft || style.borderRight) {
        cell.border = {
            top: parseBorder(style.borderTop as string) ?? cell.border?.top,
            bottom: parseBorder(style.borderBottom as string) ?? cell.border?.bottom,
            left: parseBorder(style.borderLeft as string) ?? cell.border?.left,
            right: parseBorder(style.borderRight as string) ?? cell.border?.right,
        }
    }
}

function mapAlignment(align?: string): ExcelJS.Alignment['horizontal'] | undefined {
    if (!align) return undefined
    const map: Record<string, ExcelJS.Alignment['horizontal']> = {
        left: 'left',
        center: 'center',
        right: 'right',
    }
    return map[align]
}

function mapVerticalAlignment(align?: string): ExcelJS.Alignment['vertical'] | undefined {
    if (!align) return undefined
    const map: Record<string, ExcelJS.Alignment['vertical']> = {
        top: 'top',
        middle: 'middle',
        bottom: 'bottom',
    }
    return map[align]
}

function parseBorder(borderStr?: string): ExcelJS.Border | undefined {
    if (!borderStr) return undefined
    return { style: 'thin', color: { argb: 'FF000000' } }
}
