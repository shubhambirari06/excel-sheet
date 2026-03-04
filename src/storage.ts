const STORAGE_KEY = 'excel_sheet_workbook_v1'
const STORAGE_META_KEY = 'excel_sheet_meta_v1'

export interface WorkbookMeta {
    name: string
    timestamp: string
}

export function saveWorkbookToLocalStorage(
    jsonResponse: unknown,
    name: string,
): void {
    const serialized = JSON.stringify(jsonResponse)

    const meta: WorkbookMeta = {
        name,
        timestamp: new Date().toISOString(),
    }

    localStorage.setItem(STORAGE_KEY, serialized)
    localStorage.setItem(STORAGE_META_KEY, JSON.stringify(meta))
}

export function loadWorkbookFromLocalStorage(): {
    jsonData: unknown
    meta: WorkbookMeta
} | null {
    const raw = localStorage.getItem(STORAGE_KEY)
    const metaRaw = localStorage.getItem(STORAGE_META_KEY)

    if (!raw || !metaRaw) return null

    try {
        const jsonData = JSON.parse(raw)
        const meta: WorkbookMeta = JSON.parse(metaRaw)
        return { jsonData, meta }
    } catch {
        console.error('Failed to parse stored workbook data.')
        return null
    }
}

export function hasSavedWorkbook(): boolean {
    return !!localStorage.getItem(STORAGE_KEY)
}
