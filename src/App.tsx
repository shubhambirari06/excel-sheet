import { useCallback, useEffect, useRef, useState } from 'react'
import './App.css'
import {
    SpreadsheetComponent,
    SheetsDirective,
    SheetDirective,
    RangesDirective,
    RangeDirective,
    ColumnsDirective,
    ColumnDirective,
} from '@syncfusion/ej2-react-spreadsheet'
import {data } from './datasource'
import {
    saveWorkbookToLocalStorage,
    loadWorkbookFromLocalStorage,
    hasSavedWorkbook,
} from './storage'
import { exportSpreadsheetToXlsx } from './excelExport'
import { parseExcelFile, buildWorkbookJson } from './excelImport'

const ACCEPTED_EXTENSIONS = '.xlsx,.xls,.csv'

function App() {
    const spreadsheetRef = useRef<SpreadsheetComponent | null>(null)
    const fileInputRef = useRef<HTMLInputElement | null>(null)
    const autoSaveTimerRef = useRef<number | null>(null)
    const isLoadingRef = useRef(false)

    const [status, setStatus] = useState('')
    const [fileName, setFileName] = useState('spreadsheet')
    const [showNameDialog, setShowNameDialog] = useState(false)
    const [tempFileName, setTempFileName] = useState('spreadsheet')

    const getSpreadsheet = useCallback((): SpreadsheetComponent | null => {
        const ss = spreadsheetRef.current
        if (!ss) setStatus(' Spreadsheet is not ready yet.')
        return ss
    }, [])

    const saveToLocal = useCallback(
        (nameOverride?: string) => {
            const spreadsheet = getSpreadsheet()
            if (!spreadsheet || isLoadingRef.current) return

            const saveName = nameOverride ?? fileName

            spreadsheet
                .saveAsJson()
                .then((response) => {
                    try {
                        saveWorkbookToLocalStorage(response, saveName)
                        setStatus(
                            ` Saved "${saveName}" at ${new Date().toLocaleTimeString()}`,
                        )
                    } catch (error) {
                        console.error('Save error:', error)
                        setStatus(' Unable to save to local storage (might be full).')
                    }
                })
                .catch((error: unknown) => {
                    console.error('saveAsJson error:', error)
                    setStatus(' Unable to serialize workbook.')
                })
        },
        [fileName, getSpreadsheet],
    )

    const restoreFromLocalStorage = useCallback(() => {
        const spreadsheet = getSpreadsheet()
        if (!spreadsheet) return

        const result = loadWorkbookFromLocalStorage()
        if (!result) {
            setStatus('ℹ️ No saved workbook found in local storage.')
            return
        }

        const { jsonData, meta } = result
        isLoadingRef.current = true

        try {
            spreadsheet.openFromJson({ file: jsonData as object })
            setFileName(meta.name)
            setStatus(
                ` Restored "${meta.name}" (saved ${new Date(meta.timestamp).toLocaleString()})`,
            )
        } catch (error) {
            console.error('openFromJson error:', error)
            setStatus(' Unable to restore saved workbook.')
        } finally {
            setTimeout(() => {
                isLoadingRef.current = false
            }, 500)
        }
    }, [getSpreadsheet])

    const handleLoadFromFile = useCallback(() => {
        fileInputRef.current?.click()
    }, [])

    const handleFileSelected = useCallback(
        async (event: React.ChangeEvent<HTMLInputElement>) => {
            const spreadsheet = getSpreadsheet()
            if (!spreadsheet) return

            const file = event.target.files?.[0]
            if (!file) return

            if (fileInputRef.current) fileInputRef.current.value = ''

            const ext = file.name.substring(file.name.lastIndexOf('.')).toLowerCase()
            if (!['.xlsx', '.xls', '.csv'].includes(ext)) {
                setStatus(` Invalid file type "${ext}". Use .xlsx, .xls, or .csv.`)
                return
            }

            setStatus(` Loading "${file.name}"...`)
            isLoadingRef.current = true

            try {
                const sheetData = await parseExcelFile(file)
                const workbookJson = buildWorkbookJson(sheetData)
                spreadsheet.openFromJson({ file: workbookJson })

                const nameWithoutExt = file.name.replace(/\.[^/.]+$/, '')
                setFileName(nameWithoutExt)
                setStatus(` Loaded "${file.name}" successfully.`)
            } catch (error) {
                console.error('File load error:', error)
                setStatus(
                    ` Failed to load "${file.name}". ${error instanceof Error ? error.message : ''}`,
                )
            } finally {
                setTimeout(() => {
                    isLoadingRef.current = false
                }, 500)
            }
        },
        [getSpreadsheet],
    )

    const downloadAsExcel = useCallback(() => {
        const spreadsheet = getSpreadsheet()
        if (!spreadsheet) return

        setStatus(` Preparing "${fileName}.xlsx"...`)

        spreadsheet
            .saveAsJson()
            .then(async (response) => {
                try {
                    await exportSpreadsheetToXlsx(response, fileName)
                    setStatus(` Downloaded "${fileName}.xlsx"`)
                } catch (error) {
                    console.error('Export error:', error)
                    setStatus(' Failed to export as .xlsx.')
                }
            })
            .catch((error: unknown) => {
                console.error('saveAsJson error:', error)
                setStatus(' Unable to read workbook data for export.')
            })
    }, [fileName, getSpreadsheet])

    const handleSaveClick = () => {
        setTempFileName(fileName)
        setShowNameDialog(true)
    }

    const handleConfirmSave = () => {
        const name = tempFileName.trim() || 'spreadsheet'
        setFileName(name)
        setShowNameDialog(false)
        setTimeout(() => saveToLocal(name), 50)
    }

    const handleDialogKeyDown = (e: React.KeyboardEvent<HTMLInputElement>) => {
        if (e.key === 'Enter') handleConfirmSave()
        if (e.key === 'Escape') setShowNameDialog(false)
    }

    const scheduleAutoSave = useCallback(() => {
        if (isLoadingRef.current) return
        if (autoSaveTimerRef.current !== null) {
            window.clearTimeout(autoSaveTimerRef.current)
        }
        autoSaveTimerRef.current = window.setTimeout(() => {
            saveToLocal()
        }, 3000)
    }, [saveToLocal])

    const handleCreated = useCallback(() => {
        if (hasSavedWorkbook()) {
            restoreFromLocalStorage()
        } else {
            setStatus('ℹ️ Using default data. Click "Save to Local" to persist.')
        }
    }, [restoreFromLocalStorage])

    useEffect(() => {
        const onUnload = () => saveToLocal()
        window.addEventListener('beforeunload', onUnload)
        return () => {
            window.removeEventListener('beforeunload', onUnload)
            if (autoSaveTimerRef.current !== null) {
                window.clearTimeout(autoSaveTimerRef.current)
            }
        }
    }, [saveToLocal])

    return (
        <div style={{ height: '100%', display: 'flex', flexDirection: 'column' }}>
            <div className="toolbar">
                <button type="button" className="toolbar-btn" onClick={handleSaveClick}>
                     Save to Local
                </button>
                <button type="button" className="toolbar-btn" onClick={handleLoadFromFile}>
                     Load Excel File
                </button>
                <button type="button" className="toolbar-btn primary" onClick={downloadAsExcel}>
                     Download .xlsx
                </button>

                <span className="toolbar-filename">
                    File: <strong>{fileName}</strong>
                </span>
                {status && <span className="toolbar-status">{status}</span>}

                <input
                    ref={fileInputRef}
                    type="file"
                    accept={ACCEPTED_EXTENSIONS}
                    style={{ display: 'none' }}
                    onChange={handleFileSelected}
                />
            </div>

            {showNameDialog && (
                <div className="dialog-overlay" onClick={() => setShowNameDialog(false)}>
                    <div className="dialog-box" onClick={(e) => e.stopPropagation()}>
                        <h3 style={{ marginTop: 0, marginBottom: 16, fontSize: 18 }}>
                            Save Workbook
                        </h3>
                        <label htmlFor="filename-input" className="dialog-label">
                            File name
                        </label>
                        <input
                            id="filename-input"
                            type="text"
                            value={tempFileName}
                            onChange={(e) => setTempFileName(e.target.value)}
                            onKeyDown={handleDialogKeyDown}
                            placeholder="Enter file name"
                            autoFocus
                            className="dialog-input"
                        />
                        <div
                            style={{
                                display: 'flex',
                                gap: 8,
                                justifyContent: 'flex-end',
                                marginTop: 20,
                            }}
                        >
                            <button
                                type="button"
                                className="dialog-btn cancel"
                                onClick={() => setShowNameDialog(false)}
                            >
                                Cancel
                            </button>
                            <button
                                type="button"
                                className="dialog-btn save"
                                onClick={handleConfirmSave}
                            >
                                Save
                            </button>
                        </div>
                    </div>
                </div>
            )}

            <div style={{ flex: 1, minHeight: 0 }}>
                <SpreadsheetComponent
                    ref={spreadsheetRef}
                    created={handleCreated}
                    cellSave={scheduleAutoSave}
                    actionComplete={scheduleAutoSave}
                    showRibbon={true}
                    showFormulaBar={true}
                    height="100%"
                    width="100%"
                    allowEditing={true}
                    allowSorting={true}
                    allowFiltering={true}
                    allowInsert={true}
                    allowDelete={true}
                    allowDataValidation={true}
                    allowConditionalFormat={true}
                    allowHyperlink={true}
                    allowNumberFormatting={true}
                    allowCellFormatting={true}
                    allowChart={true}
                    allowImage={true}
                    allowMerge={true}
                    allowFreezePane={true}
                    allowResizing={true}
                    allowUndoRedo={true}
                    allowWrap={true}
                    allowFindAndReplace={true}
                    allowOpen={true}
                    allowSave={true}
                >
                    <SheetsDirective>
                        <SheetDirective name="Inventory">
                            <RangesDirective>
                                <RangeDirective dataSource={data} startCell="A1" />
                            </RangesDirective>
                            <ColumnsDirective>
                                <ColumnDirective width={180} />
                                <ColumnDirective width={160} />
                                <ColumnDirective width={90} />
                                <ColumnDirective width={120} />
                                <ColumnDirective width={90} />
                                <ColumnDirective width={140} />
                            </ColumnsDirective>
                        </SheetDirective>
                    </SheetsDirective>
                </SpreadsheetComponent>
            </div>
        </div>
    )
}

export default App