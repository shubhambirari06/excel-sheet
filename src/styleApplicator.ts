import type { SpreadsheetComponent } from '@syncfusion/ej2-react-spreadsheet'
import type { ParsedWorkbook } from './excelImport'
import type { FillPattern } from 'exceljs'

export function applyParsedStyles(
    spreadsheet: SpreadsheetComponent,
    parsed: ParsedWorkbook,
    sheetIndex: number = 0,
): void {
    const sheet = parsed.sheets[sheetIndex]
    if (!sheet) return

    const sheetName = sheet.name

    for (const merge of sheet.merges) {
        try {
            spreadsheet.merge(`${sheetName}!${merge}`)
        } catch {
            // merge may already exist
        }
    }

    sheet.rows.forEach((cellMap) => {
        cellMap.forEach((cell) => {
            if (cell.isMerged && cell.masterAddress) return

            const addr = `${sheetName}!${cell.address}`
            const style = cell.style
            if (!style) return

            const format: Record<string, string> = {}

            if (style.font) {
                if (style.font.bold) format.fontWeight = 'bold'
                if (style.font.italic) format.fontStyle = 'italic'
                if (style.font.size) format.fontSize = `${style.font.size}pt`
                if (style.font.name) format.fontFamily = style.font.name
                if (style.font.color?.argb) {
                    const argb = style.font.color.argb
                    format.color = `#${argb.length >= 8 ? argb.slice(2) : argb}`
                }
                if (style.font.underline) format.textDecoration = 'underline'
            }

            if (style.alignment) {
                if (style.alignment.horizontal) {
                    format.textAlign = style.alignment.horizontal
                }
                if (style.alignment.vertical) {
                    const vMap: Record<string, string> = {
                        top: 'top',
                        middle: 'middle',
                        bottom: 'bottom',
                    }
                    format.verticalAlign = vMap[style.alignment.vertical] ?? 'bottom'
                }
                if (style.alignment.wrapText) {
                    format.whiteSpace = 'normal'
                }
            }

            if (style.fill && style.fill.type === 'pattern') {
                const patternFill = style.fill as FillPattern
                if (patternFill.fgColor?.argb) {
                    const argb = patternFill.fgColor.argb
                    format.backgroundColor = `#${argb.length >= 8 ? argb.slice(2) : argb}`
                }
            }

            if (Object.keys(format).length > 0) {
                try {
                    spreadsheet.cellFormat(format, addr)
                } catch {
                    // non-critical
                }
            }

            if (cell.numFmt) {
                try {
                    spreadsheet.numberFormat(cell.numFmt, addr)
                } catch {
                    // non-critical
                }
            }

            if (style.border) {
                const b = style.border
                if (b.top || b.bottom || b.left || b.right) {
                    try {
                        spreadsheet.setBorder(
                            { border: '1px solid #000000' },
                            addr,
                            'Outer',
                        )
                    } catch {
                        // non-critical
                    }
                }
            }
        })
    })

    sheet.rowHeights.forEach((height, rowIdx) => {
        try {
            spreadsheet.setRowHeight(rowIdx - 1, height, sheetIndex)
        } catch {
            // non-critical
        }
    })
}
