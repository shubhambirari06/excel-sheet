import { useCallback, useEffect, useRef, useState } from "react";
import type { SpreadsheetComponent } from "@syncfusion/ej2-react-spreadsheet";
import { saveWorkbookToLocalStorage } from "../storage";
import { exportSpreadsheetToXlsx } from "../excelExport";
import { parseExcelFile, buildWorkbookJson } from "../excelImport";
import { applyParsedStyles } from "../styleApplicator";
import {
  setSourceWorkbook,
  cloneSourceWorkbook,
  clearSourceWorkbook,
} from "../workbookStore";
import {
  applyEcm1DetailLayout,
  applyEcm2DetailLayout,
  applyProjectInputLayout,
} from "../features/spreadsheet/layouts";
import {
  PROJECT_INPUT_SHEET_ALIASES,
  PROJECT_INPUT_SHEET_NAME,
  normalizeSheetName,
} from "../features/spreadsheet/sheetConfig";

export function useSpreadsheetController() {
  const spreadsheetRef = useRef<SpreadsheetComponent | null>(null);
  const fileInputRef = useRef<HTMLInputElement | null>(null);
  const autoSaveTimerRef = useRef<number | null>(null);
  const isLoadingRef = useRef(false);
  const applyLayoutOnDataBoundRef = useRef(true);
  const hasAppliedInitialLayoutRef = useRef(false);

  const [status, setStatus] = useState("");
  const [fileName, setFileName] = useState("spreadsheet");
  const [showNameDialog, setShowNameDialog] = useState(false);
  const [tempFileName, setTempFileName] = useState("spreadsheet");

  const focusProjectInputSheet = useCallback((spreadsheet: SpreadsheetComponent) => {
    const sheets = spreadsheet.sheets ?? [];
    const normalizedPrimary = normalizeSheetName(PROJECT_INPUT_SHEET_NAME);
    const exact = sheets.find(
      (sheet) => normalizeSheetName(sheet?.name) === normalizedPrimary,
    );
    const alias = sheets.find((sheet) =>
      PROJECT_INPUT_SHEET_ALIASES.some(
        (name) => normalizeSheetName(sheet?.name) === normalizeSheetName(name),
      ),
    );
    const targetSheetName = exact?.name ?? alias?.name;
    if (!targetSheetName) return;

    spreadsheet.goTo(`${targetSheetName}!A1`);
    spreadsheet.selectRange("A1");
  }, []);

  const applyAllLayouts = useCallback((spreadsheet: SpreadsheetComponent) => {
    applyProjectInputLayout({ spreadsheet, setStatus });
    applyEcm1DetailLayout({ spreadsheet, setStatus });
    applyEcm2DetailLayout({ spreadsheet });
    focusProjectInputSheet(spreadsheet);
    window.setTimeout(() => {
      applyEcm1DetailLayout({ spreadsheet, setStatus });
      applyEcm2DetailLayout({ spreadsheet });
      focusProjectInputSheet(spreadsheet);
    }, 150);
  }, [focusProjectInputSheet]);

  const getSpreadsheet = useCallback((): SpreadsheetComponent | null => {
    const ss = spreadsheetRef.current;
    if (!ss) setStatus(" Spreadsheet is not ready yet.");
    return ss;
  }, []);

  const resetViewportToTop = useCallback(() => {
    const spreadsheet = spreadsheetRef.current;
    if (!spreadsheet) return;
    spreadsheet.goTo("A1");
    spreadsheet.selectRange("A1");
  }, []);

  const saveToLocal = useCallback(
    (nameOverride?: string) => {
      const spreadsheet = getSpreadsheet();
      if (!spreadsheet || isLoadingRef.current) return;

      const saveName = nameOverride ?? fileName;

      spreadsheet
        .saveAsJson()
        .then((response) => {
          try {
            saveWorkbookToLocalStorage(response, saveName);
            setStatus(` Saved \"${saveName}\" at ${new Date().toLocaleTimeString()}`);
          } catch (error) {
            console.error("Save error:", error);
            setStatus(" Unable to save to local storage (might be full).");
          }
        })
        .catch((error: unknown) => {
          console.error("saveAsJson error:", error);
          setStatus(" Unable to serialize workbook.");
        });
    },
    [fileName, getSpreadsheet],
  );

  const handleLoadFromFile = useCallback(() => {
    fileInputRef.current?.click();
  }, []);

  const handleFileSelected = useCallback(
    async (event: React.ChangeEvent<HTMLInputElement>) => {
      const file = event.target.files?.[0];
      if (!file) return;

      if (fileInputRef.current) fileInputRef.current.value = "";

      const ext = file.name.substring(file.name.lastIndexOf(".")).toLowerCase();
      if (![".xlsx", ".xls", ".csv"].includes(ext)) {
        setStatus(` Invalid file type \"${ext}\". Use .xlsx, .xls, or .csv.`);
        return;
      }

      isLoadingRef.current = true;
      setStatus(`Loading \"${file.name}\"...`);

      try {
        const parsed = await parseExcelFile(file);
        const spreadsheet = getSpreadsheet();
        if (!spreadsheet) return;

        setSourceWorkbook(parsed.workbook);
        // Re-run layout on imported workbooks so header cells are restored per template.
        applyLayoutOnDataBoundRef.current = true;
        hasAppliedInitialLayoutRef.current = false;

        const workbookJson = buildWorkbookJson(parsed);
        spreadsheet.openFromJson({ file: workbookJson });

        setTimeout(() => {
          try {
            for (let i = 0; i < parsed.sheets.length; i++) {
              applyParsedStyles(spreadsheet, parsed, i);
            }
            // Enforce template layout after import regardless of dataBound timing.
            applyAllLayouts(spreadsheet);
          } catch (styleError) {
            console.warn("Style application warning:", styleError);
          }
        }, 300);

        const nameWithoutExt = file.name.replace(/\.[^/.]+$/, "");
        setFileName(nameWithoutExt);
        setStatus(`Loaded \"${file.name}\" successfully.`);
      } catch (error) {
        console.error("File load error:", error);
        setStatus(
          `Failed to load \"${file.name}\". ${error instanceof Error ? error.message : ""}`,
        );
        clearSourceWorkbook();
      } finally {
        setTimeout(() => {
          isLoadingRef.current = false;
        }, 800);
      }
    },
    [applyAllLayouts, getSpreadsheet],
  );

  const downloadAsExcel = useCallback(() => {
    const spreadsheet = getSpreadsheet();
    if (!spreadsheet) return;

    setStatus(`Preparing \"${fileName}.xlsx\"...`);

    spreadsheet
      .saveAsJson()
      .then(async (response) => {
        try {
          const source = (await cloneSourceWorkbook()) ?? undefined;
          await exportSpreadsheetToXlsx(response, fileName, source);
          setStatus(`Downloaded \"${fileName}.xlsx\"`);
        } catch (error) {
          console.error("Export error:", error);
          setStatus(
            `Failed to export as .xlsx. ${error instanceof Error ? error.message : ""}`,
          );
        }
      })
      .catch((error: unknown) => {
        console.error("saveAsJson error:", error);
        setStatus(" Unable to read workbook data for export.");
      });
  }, [fileName, getSpreadsheet]);

  const handleSaveClick = useCallback(() => {
    setTempFileName(fileName);
    setShowNameDialog(true);
  }, [fileName]);

  const handleConfirmSave = useCallback(() => {
    const name = tempFileName.trim() || "spreadsheet";
    setFileName(name);
    setShowNameDialog(false);
    setTimeout(() => saveToLocal(name), 50);
  }, [saveToLocal, tempFileName]);

  const handleDialogKeyDown = useCallback(
    (e: React.KeyboardEvent<HTMLInputElement>) => {
      if (e.key === "Enter") handleConfirmSave();
      if (e.key === "Escape") setShowNameDialog(false);
    },
    [handleConfirmSave],
  );

  const scheduleAutoSave = useCallback(() => {
    if (isLoadingRef.current) return;
    if (autoSaveTimerRef.current !== null) {
      window.clearTimeout(autoSaveTimerRef.current);
    }
    autoSaveTimerRef.current = window.setTimeout(() => {
      saveToLocal();
    }, 3000);
  }, [saveToLocal]);

  const handleCreated = useCallback(() => {
    clearSourceWorkbook();
    applyLayoutOnDataBoundRef.current = true;
    hasAppliedInitialLayoutRef.current = false;
    setStatus(" Default workbook loaded from in-app datasource.");
  }, []);

  const handleDataBound = useCallback(() => {
    if (!applyLayoutOnDataBoundRef.current) return;
    if (hasAppliedInitialLayoutRef.current) return;

    const spreadsheet = getSpreadsheet();
    if (!spreadsheet) return;

    hasAppliedInitialLayoutRef.current = true;

    setTimeout(() => {
      applyAllLayouts(spreadsheet);
      resetViewportToTop();
    }, 0);
  }, [applyAllLayouts, getSpreadsheet, resetViewportToTop]);

  useEffect(() => {
    const onUnload = () => saveToLocal();
    window.addEventListener("beforeunload", onUnload);
    return () => {
      window.removeEventListener("beforeunload", onUnload);
      if (autoSaveTimerRef.current !== null) {
        window.clearTimeout(autoSaveTimerRef.current);
      }
    };
  }, [saveToLocal]);

  return {
    spreadsheetRef,
    fileInputRef,
    status,
    fileName,
    showNameDialog,
    tempFileName,
    setShowNameDialog,
    setTempFileName,
    handleSaveClick,
    handleLoadFromFile,
    downloadAsExcel,
    handleFileSelected,
    handleConfirmSave,
    handleDialogKeyDown,
    scheduleAutoSave,
    handleCreated,
    handleDataBound,
  };
}
