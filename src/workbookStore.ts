import ExcelJS from 'exceljs'

let _sourceWorkbook: ExcelJS.Workbook | null = null
let _sourceBuffer: ArrayBuffer | null = null

export function setSourceWorkbook(wb: ExcelJS.Workbook | null): void {
    _sourceWorkbook = wb
    if (wb) {
        wb.xlsx.writeBuffer().then((buf) => {
            _sourceBuffer = buf as ArrayBuffer
        }).catch(() => {
            _sourceBuffer = null
        })
    } else {
        _sourceBuffer = null
    }
}

export function getSourceWorkbook(): ExcelJS.Workbook | null {
    return _sourceWorkbook
}

export async function cloneSourceWorkbook(): Promise<ExcelJS.Workbook | null> {
    if (!_sourceBuffer && _sourceWorkbook) {
        try {
            _sourceBuffer = (await _sourceWorkbook.xlsx.writeBuffer()) as ArrayBuffer
        } catch {
            _sourceBuffer = null
        }
    }

    if (!_sourceBuffer) return null
    const wb = new ExcelJS.Workbook()
    await wb.xlsx.load(_sourceBuffer)
    return wb
}

export function clearSourceWorkbook(): void {
    _sourceWorkbook = null
    _sourceBuffer = null
}
