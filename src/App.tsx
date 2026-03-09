import { useCallback, useEffect, useRef, useState } from "react";
import "./App.css";
import {
  SpreadsheetComponent,
  SheetsDirective,
  SheetDirective,
  RangesDirective,
  RangeDirective,
  ColumnsDirective,
  ColumnDirective,
} from "@syncfusion/ej2-react-spreadsheet";
import {
  projectHeader,
  ecmRows,
  hvacRows,
  chillerRows,
  freezerRows,
  ecm1InputVariables,
  ecm1EquipmentRowsOrdered,
  ecm1EnergyCostSavings,
  ecm1TotalSavings,
  ecm1BinDataAHU1,
  ecm1BinDataAHU2,
  ecm1BinDataAHU3,
  ecm1BinDataAHU4,
  ecm1BinDataAHU5,
  ecm1BinDataAHU6,
  ecm1BinDataAHU7,
  ECM1_DETAIL_SHEET_NAME,
} from "./datasource";
import {
  ecm2InputVariables,
  ecm2Equations,
  ecm2EquipmentSummary,
  ecm2EnergySavings,
  ecm2TotalGasSavings,
  ecm2RtuLoadProfile,
  ecm2BinDataAHU1,
  ecm2BinDataAHU2,
  ecm2BinDataAHU3,
  ecm2BinDataAHU4,
  ecm2BinDataAHU5,
  ecm2BinDataAHU6,
  ecm2BinDataAHU7,
  ecm2BinDataRTU1,
} from "./ecm2Datasource";
import {
  saveWorkbookToLocalStorage,
  loadWorkbookFromLocalStorage,
  hasSavedWorkbook,
} from "./storage";
import { exportSpreadsheetToXlsx } from "./excelExport";
import { parseExcelFile, buildWorkbookJson } from "./excelImport";
import { applyParsedStyles } from "./styleApplicator";
import {
  setSourceWorkbook,
  cloneSourceWorkbook,
  clearSourceWorkbook,
} from "./workbookStore";
import type { ParsedWorkbook } from "./excelImport";

const ACCEPTED_EXTENSIONS = ".xlsx,.xls,.csv";
const PROJECT_INPUT_SHEET_NAME = "Project Input";
const PROJECT_INPUT_SHEET_ALIASES = [
  PROJECT_INPUT_SHEET_NAME,
  "Project Sheet",
  "Input Sheet",
];
const PROJECT_INPUT_TEMPLATE_TITLE = "PROJECT INPUT SHEET: HBS SOLUTION";
const ECM1_DETAIL_SHEET_ALIASES = [ECM1_DETAIL_SHEET_NAME];
const ECM2_DETAIL_SHEET_NAME = "ECM2";
const PUBLIC_TEMPLATE_WORKBOOK =
  "8300 Meadowbrook Ln_WGCPPS1550333380_Final ECM Calculations.xlsx";

const normalizeSheetName = (name: string | undefined) =>
  String(name ?? "")
    .trim()
    .toLowerCase();

const insertGapColumn = (data: any[], index: number) => {
  return data.map((row) => {
    const entries = Object.entries(row);
    const newEntries = [
      ...entries.slice(0, index),
      [" ", null],
      ...entries.slice(index),
    ];
    return Object.fromEntries(newEntries);
  });
};

function App() {
  const spreadsheetRef = useRef<SpreadsheetComponent | null>(null);
  const fileInputRef = useRef<HTMLInputElement | null>(null);
  const autoSaveTimerRef = useRef<number | null>(null);
  const isLoadingRef = useRef(false);
  const parsedWorkbookRef = useRef<ParsedWorkbook | null>(null);
  const applyLayoutOnDataBoundRef = useRef(true);
  const hasAppliedInitialLayoutRef = useRef(false);
  const isImportingFileRef = useRef(false);

  const [status, setStatus] = useState("");
  const [fileName, setFileName] = useState("spreadsheet");
  const [showNameDialog, setShowNameDialog] = useState(false);
  const [tempFileName, setTempFileName] = useState("spreadsheet");

  const ecm1SectionRows = {
    inputVariables: 2,
    energySavings: 7,
    totals: 7,
    equipmentSummary: 19,
  };

  const ecm1EnergyCostSavingsWithGap = insertGapColumn(
    ecm1EnergyCostSavings,
    2,
  );

  const ecm1EquipmentEndRow =
    ecm1SectionRows.equipmentSummary + ecm1EquipmentRowsOrdered.length;
  const ecm1AhuSectionFirstStartRow = ecm1EquipmentEndRow + 3;
  const ecm1AhuSectionGapRows = 1;
  const ecm1AhuSections = [
    { title: "AHU #1", data: insertGapColumn(ecm1BinDataAHU1, 7) },
    { title: "AHU #2", data: insertGapColumn(ecm1BinDataAHU2, 7) },
    { title: "AHU #3", data: insertGapColumn(ecm1BinDataAHU3, 7) },
    { title: "AHU #4", data: insertGapColumn(ecm1BinDataAHU4, 7) },
    { title: "AHU #5", data: insertGapColumn(ecm1BinDataAHU5, 7) },
    { title: "AHU #6", data: insertGapColumn(ecm1BinDataAHU6, 7) },
    { title: "AHU #7", data: insertGapColumn(ecm1BinDataAHU7, 7) },
  ].map((section, index, list) => {
    const previousRows = list
      .slice(0, index)
      .reduce((sum, prev) => sum + (prev.data.length + 1), 0);

    return {
      ...section,
      startRow: ecm1AhuSectionFirstStartRow + previousRows,
    };
  });

  const ecm2SectionRows = {
    inputVariables: 4,
    equations: 4,
    energySavings: 9,
    totalGasSavings: 9,
    equipmentSummary: 21,
    rtuLoadProfile: 31,
  };

  const ecm2AhuSectionFirstStartRow = 43;
  const ecm2AhuSectionGapRows = 1;
  const ecm2AhuSections = [
    { title: "AHU #1", data: ecm2BinDataAHU1 },
    { title: "AHU #2", data: ecm2BinDataAHU2 },
    { title: "AHU #3", data: ecm2BinDataAHU3 },
    { title: "AHU #4", data: ecm2BinDataAHU4 },
    { title: "AHU #5", data: ecm2BinDataAHU5 },
    { title: "AHU #6", data: ecm2BinDataAHU6 },
    { title: "AHU #7", data: ecm2BinDataAHU7 },
    { title: "RTU # 1", data: ecm2BinDataRTU1 },
  ].map((section, index, list) => {
    const previousRows = list
      .slice(0, index)
      .reduce((sum, prev) => sum + (prev.data.length + 1), 0);
    return { ...section, startRow: ecm2AhuSectionFirstStartRow + previousRows };
  });

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
            setStatus(
              ` Saved "${saveName}" at ${new Date().toLocaleTimeString()}`,
            );
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

  const restoreFromLocalStorage = useCallback(() => {
    const spreadsheet = getSpreadsheet();
    if (!spreadsheet) return;

    const result = loadWorkbookFromLocalStorage();
    if (!result) {
      setStatus("ℹ No saved workbook found in local storage.");
      return;
    }

    const { jsonData, meta } = result;
    isLoadingRef.current = true;
    applyLayoutOnDataBoundRef.current = false;
    hasAppliedInitialLayoutRef.current = true;

    try {
      spreadsheet.openFromJson({ file: jsonData as object });
      setFileName(meta.name);
      setStatus(
        ` Restored "${meta.name}" (saved ${new Date(meta.timestamp).toLocaleString()})`,
      );
      setTimeout(() => resetViewportToTop(), 0);
    } catch (error) {
      console.error("openFromJson error:", error);
      setStatus(" Unable to restore saved workbook.");
    } finally {
      setTimeout(() => {
        isLoadingRef.current = false;
      }, 500);
    }
  }, [getSpreadsheet, resetViewportToTop]);

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
        setStatus(` Invalid file type "${ext}". Use .xlsx, .xls, or .csv.`);
        return;
      }

      isLoadingRef.current = true;
      setStatus(`Loading "${file.name}"...`);

      try {
        const parsed = await parseExcelFile(file);
        const spreadsheet = getSpreadsheet();
        if (!spreadsheet) return;

        parsedWorkbookRef.current = parsed;
        setSourceWorkbook(parsed.workbook);
        applyLayoutOnDataBoundRef.current = false;
        hasAppliedInitialLayoutRef.current = true;
        isImportingFileRef.current = true;
        applyLayoutOnDataBoundRef.current = true;
        hasAppliedInitialLayoutRef.current = false;

        const workbookJson = buildWorkbookJson(parsed);
        spreadsheet.openFromJson({ file: workbookJson });

        setTimeout(() => {
          try {
            for (let i = 0; i < parsed.sheets.length; i++) {
              applyParsedStyles(spreadsheet, parsed, i);
            }
            if (parsed.sheets.length > 0) {
              spreadsheet.goTo(`${parsed.sheets[0].name}!A1`);
            }
            resetViewportToTop();
          } catch (styleError) {
            console.warn("Style application warning:", styleError);
          }
        }, 300);

        const nameWithoutExt = file.name.replace(/\.[^/.]+$/, "");
        setFileName(nameWithoutExt);
        setStatus(`Loaded "${file.name}" successfully.`);
      } catch (error) {
        console.error("File load error:", error);
        setStatus(
          `Failed to load "${file.name}". ${error instanceof Error ? error.message : ""}`,
        );
        clearSourceWorkbook();
        parsedWorkbookRef.current = null;
      } finally {
        setTimeout(() => {
          isLoadingRef.current = false;
        }, 800);
      }
    },
    [getSpreadsheet, resetViewportToTop],
  );

  const loadPublicTemplateWorkbook = useCallback(async () => {
    const spreadsheet = getSpreadsheet();
    if (!spreadsheet) return false;

    isLoadingRef.current = true;
    setStatus("Loading default workbook template...");

    try {
      const encodedPath = `/${encodeURIComponent(PUBLIC_TEMPLATE_WORKBOOK)}`;
      const response = await fetch(encodedPath);
      if (!response.ok) {
        throw new Error(`Template fetch failed (${response.status})`);
      }

      const buffer = await response.arrayBuffer();
      const file = new File([buffer], PUBLIC_TEMPLATE_WORKBOOK, {
        type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
      });

      const parsed = await parseExcelFile(file);
      parsedWorkbookRef.current = parsed;
      setSourceWorkbook(parsed.workbook);
      applyLayoutOnDataBoundRef.current = false;
      hasAppliedInitialLayoutRef.current = true;
      isImportingFileRef.current = true;
      applyLayoutOnDataBoundRef.current = true;
      hasAppliedInitialLayoutRef.current = false;

      const workbookJson = buildWorkbookJson(parsed);
      spreadsheet.openFromJson({ file: workbookJson });

      setTimeout(() => {
        try {
          for (let i = 0; i < parsed.sheets.length; i++) {
            applyParsedStyles(spreadsheet, parsed, i);
          }
          if (parsed.sheets.length > 0) {
            spreadsheet.goTo(`${parsed.sheets[0].name}!A1`);
          }
          resetViewportToTop();
        } catch (styleError) {
          console.warn("Style application warning:", styleError);
        }
      }, 300);

      const nameWithoutExt = PUBLIC_TEMPLATE_WORKBOOK.replace(/\.[^/.]+$/, "");
      setFileName(nameWithoutExt);
      setStatus(`Loaded template "${PUBLIC_TEMPLATE_WORKBOOK}" from public.`);
      return true;
    } catch (error) {
      console.error("Default template load error:", error);
      setStatus(
        `Using default sheet layout. Template load failed. ${error instanceof Error ? error.message : ""}`,
      );
      clearSourceWorkbook();
      parsedWorkbookRef.current = null;
      applyLayoutOnDataBoundRef.current = true;
      hasAppliedInitialLayoutRef.current = false;
      return false;
    } finally {
      setTimeout(() => {
        isLoadingRef.current = false;
      }, 800);
    }
  }, [getSpreadsheet, resetViewportToTop]);

  const downloadAsExcel = useCallback(() => {
    const spreadsheet = getSpreadsheet();
    if (!spreadsheet) return;

    setStatus(`Preparing "${fileName}.xlsx"...`);

    spreadsheet
      .saveAsJson()
      .then(async (response) => {
        try {
          const source = (await cloneSourceWorkbook()) ?? undefined;
          await exportSpreadsheetToXlsx(response, fileName, source);
          setStatus(`Downloaded "${fileName}.xlsx"`);
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

  const handleSaveClick = () => {
    setTempFileName(fileName);
    setShowNameDialog(true);
  };

  const handleConfirmSave = () => {
    const name = tempFileName.trim() || "spreadsheet";
    setFileName(name);
    setShowNameDialog(false);
    setTimeout(() => saveToLocal(name), 50);
  };

  const handleDialogKeyDown = (e: React.KeyboardEvent<HTMLInputElement>) => {
    if (e.key === "Enter") handleConfirmSave();
    if (e.key === "Escape") setShowNameDialog(false);
  };

  const scheduleAutoSave = useCallback(() => {
    if (isLoadingRef.current) return;
    if (autoSaveTimerRef.current !== null) {
      window.clearTimeout(autoSaveTimerRef.current);
    }
    autoSaveTimerRef.current = window.setTimeout(() => {
      saveToLocal();
    }, 3000);
  }, [saveToLocal]);

  const applyImageLikeLayout = useCallback(() => {
    const spreadsheet = getSpreadsheet();
    if (!spreadsheet) return;

    const targetSheetIndex = (spreadsheet.sheets ?? []).findIndex((sheet) =>
      PROJECT_INPUT_SHEET_ALIASES.some(
        (alias) =>
          normalizeSheetName(sheet?.name) === normalizeSheetName(alias),
      ),
    );

    if (targetSheetIndex < 0) {
      setStatus(" Project input sheet not found.");
      return;
    }

    const targetSheet = spreadsheet.sheets?.[targetSheetIndex] as
      | {
          rows?: Array<{
            cells?: Array<{ value?: unknown }>;
          }>;
        }
      | undefined;
    const a1Value = String(
      targetSheet?.rows?.[0]?.cells?.[0]?.value ?? "",
    ).trim();
    if (
      a1Value &&
      normalizeSheetName(a1Value) !==
        normalizeSheetName(PROJECT_INPUT_TEMPLATE_TITLE)
    ) {
      setStatus(
        " Skipped Project Input template layout to preserve existing data.",
      );
      return;
    }

    const targetSheetName =
      spreadsheet.sheets?.[targetSheetIndex]?.name ?? PROJECT_INPUT_SHEET_NAME;
    const addr = (range: string) => `${targetSheetName}!${range}`;

    const safeMerge = (range: string) => {
      try {
        spreadsheet.merge(addr(range));
      } catch {}
    };

    const isImport = isImportingFileRef.current;
    const header = projectHeader[0];

    spreadsheet.updateCell(
      { value: "PROJECT INPUT SHEET: HBS SOLUTION" },
      addr("A1"),
    );
    if (!isImport) spreadsheet.updateCell({ value: header.Date }, addr("Q1"));

    spreadsheet.updateCell({ value: "Project Utility" }, addr("A2"));
    if (!isImport)
      spreadsheet.updateCell({ value: header.ProjectUtility }, addr("E2"));
    spreadsheet.updateCell({ value: "Project Name" }, addr("I2"));
    if (!isImport)
      spreadsheet.updateCell({ value: header.ProjectName }, addr("L2"));
    spreadsheet.updateCell({ value: "Project Type" }, addr("O2"));
    if (!isImport)
      spreadsheet.updateCell({ value: header.ProjectType }, addr("Q2"));
    if (!isImport)
      spreadsheet.updateCell({ value: header.ProgramType }, addr("R2"));

    spreadsheet.updateCell({ value: "Square Footage" }, addr("A3"));
    if (!isImport)
      spreadsheet.updateCell(
        { value: String(header.SquareFootage) },
        addr("E3"),
      );
    spreadsheet.updateCell({ value: "(S.F.) Annual Therm usage" }, addr("G3"));
    if (!isImport)
      spreadsheet.updateCell(
        { value: String(header.AnnualThermUsage) },
        addr("L3"),
      );
    spreadsheet.updateCell({ value: "Therm" }, addr("N3"));

    spreadsheet.updateCell({ value: "No" }, addr("B5"));
    spreadsheet.updateCell({ value: "Energy Conservation Method" }, addr("C5"));
    spreadsheet.updateCell({ value: "Estimated Cost ($)" }, addr("D5"));
    spreadsheet.updateCell(
      { value: "Natural Gas Savings (Therm/yr)" },
      addr("E5"),
    );
    spreadsheet.updateCell(
      { value: "Natural Gas Energy Cost Savings ($/yr)" },
      addr("F5"),
    );
    spreadsheet.updateCell({ value: "Simple Payback (years gas)" }, addr("G5"));

    spreadsheet.updateCell({ value: "Estimated Cost ($)" }, addr("E9"));
    spreadsheet.updateCell(
      { value: "Natural Gas Savings (Therm/yr)" },
      addr("F9"),
    );
    spreadsheet.updateCell(
      { value: "Electrical Energy Cost Savings ($/yr)" },
      addr("G9"),
    );
    spreadsheet.updateCell({ value: "Simple Payback (years gas)" }, addr("H9"));

    spreadsheet.updateCell({ value: "Total Project Summary" }, addr("C10"));
    spreadsheet.updateCell({ formula: "=SUM(D6:D7)" }, addr("E10"));
    spreadsheet.updateCell({ formula: "=SUM(E6:E7)" }, addr("F10"));
    spreadsheet.updateCell({ formula: "=SUM(F6:F7)" }, addr("G10"));
    spreadsheet.updateCell({ formula: "=E10/G10" }, addr("H10"));

    spreadsheet.updateCell(
      { value: "Total Project Summary (After Rebate)" },
      addr("C11"),
    );
    spreadsheet.updateCell({ formula: "=E10-K15" }, addr("E11"));
    spreadsheet.updateCell({ formula: "=F10" }, addr("F11"));
    spreadsheet.updateCell({ formula: "=G10" }, addr("G11"));
    spreadsheet.updateCell({ formula: "=E11/G11" }, addr("H11"));

    spreadsheet.cellFormat({ fontWeight: "bold" }, addr("C10:C11"));

    spreadsheet.updateCell({ value: "Baseline (Therm)" }, addr("C14"));
    spreadsheet.updateCell({ value: "Adjust (Therm)" }, addr("D14"));
    spreadsheet.updateCell({ value: "% Decrease" }, addr("E14"));
    spreadsheet.updateCell({ value: "Therm/SF" }, addr("F14"));
    spreadsheet.updateCell(
      { value: "WG Incentive amount @ $3.7/Therm" },
      addr("K14"),
    );

    spreadsheet.updateCell({ value: "Cooling Tonnage" }, addr("C18"));
    spreadsheet.updateCell({ value: "EER (Existing)" }, addr("D18"));
    spreadsheet.updateCell({ value: "EER (Proposed)" }, addr("E18"));
    spreadsheet.updateCell({ value: "Cooling Eff (kW/Ton)" }, addr("F18"));
    spreadsheet.updateCell({ value: "Heat Output (MBH)" }, addr("G18"));
    spreadsheet.updateCell({ value: "Heating Efficiency" }, addr("H18"));
    spreadsheet.updateCell({ value: "Supply Fan HP" }, addr("I18"));
    spreadsheet.updateCell({ value: "Supply Fan Unit" }, addr("J18"));
    spreadsheet.updateCell(
      { value: "Supply Fan Motor Load Factor" },
      addr("K18"),
    );
    spreadsheet.updateCell(
      { value: "Supply Fan Motor Efficiency" },
      addr("L18"),
    );
    spreadsheet.updateCell({ value: "Supply Fan Max CFM" }, addr("M18"));
    spreadsheet.updateCell({ value: "Supply Fan Min CFM" }, addr("N18"));
    spreadsheet.updateCell({ value: "Exhaust fans HP" }, addr("O18"));
    spreadsheet.updateCell({ value: "Motor Efficiency" }, addr("P18"));
    spreadsheet.updateCell({ value: "Number of Exhaust fans" }, addr("Q18"));
    spreadsheet.updateCell({ value: "Electric Heating" }, addr("R18"));
    spreadsheet.updateCell({ value: "Equipment Name" }, addr("B18"));

    spreadsheet.updateCell({ value: "Cooling Tonnage" }, addr("C29"));
    spreadsheet.updateCell({ value: "EER (Existing)" }, addr("D29"));
    spreadsheet.updateCell({ value: "EER (Proposed)" }, addr("E29"));
    spreadsheet.updateCell({ value: "Cooling Eff (kW/Ton)" }, addr("F29"));
    spreadsheet.updateCell({ value: "Equipment Name" }, addr("B29"));

    spreadsheet.updateCell({ value: "MBH" }, addr("C34"));
    spreadsheet.updateCell({ value: "EER (Proposed)" }, addr("D34"));
    spreadsheet.updateCell({ value: "Cooling Eff (kW/Ton)" }, addr("E34"));
    spreadsheet.updateCell({ value: "Fan Motor HP" }, addr("F34"));
    spreadsheet.updateCell({ value: "Fan Motor kW" }, addr("G34"));
    spreadsheet.updateCell({ value: "Equipment Name" }, addr("B34"));

    spreadsheet.updateCell({ value: "HVAC" }, addr("A18"));
    spreadsheet.updateCell({ value: "CHILLER" }, addr("A29"));
    spreadsheet.updateCell({ value: "FREEZER" }, addr("A34"));

    safeMerge("A1:P1");
    safeMerge("Q1:R1");
    safeMerge("A2:D2");
    safeMerge("E2:H2");
    safeMerge("I2:K2");
    safeMerge("L2:N2");
    safeMerge("O2:P2");
    safeMerge("A3:D3");
    safeMerge("E3:F3");
    safeMerge("G3:K3");
    safeMerge("L3:M3");
    safeMerge("K14:N14");
    safeMerge("K15:N15");
    safeMerge("A18:A26");
    safeMerge("A29:A31");
    safeMerge("A34:A36");

    spreadsheet.cellFormat(
      { fontWeight: "bold", textAlign: "center", verticalAlign: "middle" },
      addr("A1:R3"),
    );
    spreadsheet.cellFormat(
      { backgroundColor: "#b9c8dc", fontWeight: "bold", textAlign: "center" },
      addr("A1:R3"),
    );
    spreadsheet.cellFormat(
      { backgroundColor: "#d6e1ef", fontWeight: "bold", textAlign: "center" },
      addr("B5:G5,E9:H9,C14:F14,K14:N14,B18:R18,B29:F29,B34:G34"),
    );
    spreadsheet.cellFormat(
      {
        backgroundColor: "#d6e1ef",
        fontWeight: "bold",
        textAlign: "center",
        verticalAlign: "middle",
      },
      addr("A18:A36"),
    );

    spreadsheet.cellFormat(
      { textAlign: "center", verticalAlign: "middle" },
      addr("A1:R36"),
    );

    spreadsheet.numberFormat("$#,##0.00", addr("D6:D7,E10:E11"));
    spreadsheet.numberFormat("#,##0.00", addr("E6:E7,F10:F11"));
    spreadsheet.numberFormat("$#,##0.00", addr("F6:F7,G10:G11"));
    spreadsheet.numberFormat("0.00", addr("G6:G7,H10:H11"));
    spreadsheet.numberFormat("#,##0", addr("L3:L3,M19:N31"));
    spreadsheet.numberFormat("#,##0", addr("C15:D15"));
    spreadsheet.numberFormat("0.00%", addr("E15"));
    spreadsheet.numberFormat("0.00", addr("F15"));
    spreadsheet.numberFormat("$#,##0.00", addr("K15"));

    spreadsheet.updateCell({ value: "1" }, addr("B6"));
    spreadsheet.updateCell({ value: "2" }, addr("B7"));

    // Add formulas for Project Input sheet calculations
    spreadsheet.updateCell({ formula: '=IF(F6>0,D6/F6,"")' }, addr("G6"));
    spreadsheet.updateCell({ formula: "=IF(F7<>0,D7/F7,0)" }, addr("G7"));

    spreadsheet.updateCell(
      { formula: '=IF(C15>0,(C15-D15)/C15,"")' },
      addr("E15"),
    );
    spreadsheet.updateCell({ formula: "=IF(E3<>0,C15/E3,0)" }, addr("F15"));
    spreadsheet.updateCell({ formula: "=MIN((E10/2),(F10*3.7))" }, addr("K15"));

    // Link Baseline to Annual Therm Usage
    spreadsheet.updateCell({ formula: "=L3" }, addr("C15"));
    // Calculate Adjust based on savings
    spreadsheet.updateCell({ formula: "=C15-SUM(E6:E7)" }, addr("D15"));

    // Link ECM-1 Savings
    const ecm1EnergyTotalsRow =
      ecm1SectionRows.energySavings + ecm1EnergyCostSavings.length;
    spreadsheet.updateCell(
      { formula: `='${ECM1_DETAIL_SHEET_NAME}'!J8` },
      addr("E6"),
    );
    spreadsheet.updateCell(
      { formula: `='${ECM1_DETAIL_SHEET_NAME}'!D${ecm1EnergyTotalsRow}` },
      addr("F6"),
    );

    // Link ECM-2 Savings
    spreadsheet.updateCell(
      { formula: `='${ECM2_DETAIL_SHEET_NAME}'!J10` },
      addr("E7"),
    );
    spreadsheet.updateCell(
      { formula: `='${ECM2_DETAIL_SHEET_NAME}'!J11` },
      addr("F7"),
    );

    // HVAC Formulas (Cooling Eff = 12 / EER)
    for (let r = 19; r <= 26; r++) {
      spreadsheet.updateCell(
        { formula: `=IF(E${r}<>0,12/E${r},0)` },
        addr(`F${r}`),
      );
    }
    // Chiller Formulas (Cooling Eff = 12 / EER)
    for (let r = 30; r <= 31; r++) {
      spreadsheet.updateCell(
        { formula: `=IF(E${r}<>0,12/E${r},0)` },
        addr(`F${r}`),
      );
    }
    // Freezer Formulas (Cooling Eff = 12 / EER, Fan kW = HP * 0.746)
    for (let r = 35; r <= 36; r++) {
      spreadsheet.updateCell(
        { formula: `=IF(D${r}<>0,12/D${r},0)` },
        addr(`E${r}`),
      );
      spreadsheet.updateCell({ formula: `=F${r}*0.746` }, addr(`G${r}`));
    }

    const tableRanges = [
      "A1:R3",
      "B5:G7",
      "E9:H11",
      "C14:F15",
      "K14:N15",
      "A18:R26",
      "A29:F31",
      "A34:G36",
    ];

    // Apply borders to all tables
    tableRanges.forEach((range) => {
      spreadsheet.setBorder(
        { border: "1px solid #000000" },
        addr(range),
        "Inner",
      );
      spreadsheet.setBorder(
        { border: "2px solid #000000" },
        addr(range),
        "Outer",
      );
    });
  }, [getSpreadsheet]);

  const applyEcmDetailLayout = useCallback(
    (sheetName: string) => {
      const spreadsheet = getSpreadsheet();
      if (!spreadsheet) return;

      const candidateSheetNames =
        normalizeSheetName(sheetName) ===
        normalizeSheetName(ECM1_DETAIL_SHEET_NAME)
          ? ECM1_DETAIL_SHEET_ALIASES
          : [sheetName];

      const targetSheetIndex = (spreadsheet.sheets ?? []).findIndex((sheet) =>
        candidateSheetNames.some(
          (candidate) =>
            normalizeSheetName(sheet?.name) === normalizeSheetName(candidate),
        ),
      );

      if (targetSheetIndex < 0) {
        setStatus(` Sheet "${sheetName}" not found.`);
        return;
      }

      const targetSheetName =
        spreadsheet.sheets?.[targetSheetIndex]?.name ?? candidateSheetNames[0];
      const targetSheet = spreadsheet.sheets?.[targetSheetIndex];
      if (!targetSheet) return;

      const previousSheetIndex = spreadsheet.activeSheetIndex;
      const previousSheetName =
        spreadsheet.sheets?.[previousSheetIndex]?.name ?? targetSheetName;

      spreadsheet.activeSheetIndex = targetSheetIndex;
      spreadsheet.goTo(`${targetSheetName}!A1`);

      const addr = (range: string) => range;
      const getEndRow = (startRow: number, dataLength: number) =>
        startRow + Math.max(dataLength, 0);
      try {
        try {
          spreadsheet.merge(addr("A1:J1"));
        } catch {}

        spreadsheet.updateCell(
          {
            value: `${sheetName} Optimize HVAC Scheduling`,
            style: {
              fontWeight: "bold",
              textAlign: "center",
              verticalAlign: "middle",
            },
          },
          addr("A1"),
        );

        spreadsheet.cellFormat(
          {
            backgroundColor: "#2563eb",
            color: "#ffffff",
            fontWeight: "bold",
            textAlign: "center",
            verticalAlign: "middle",
          },
          addr("A1:J1"),
        );

        ecm1AhuSections.forEach((section) => {
          const sectionTitleRow = section.startRow - 1;
          try {
            spreadsheet.merge(addr(`A${sectionTitleRow}:J${sectionTitleRow}`));
          } catch {}
          spreadsheet.updateCell(
            { value: section.title },
            addr(`A${sectionTitleRow}`),
          );
        });

        const headerRanges = [
          addr("A2:B2"),
          addr("A7:D7"),
          addr("I7:J7"),
          addr("A19:I19"),
          ...ecm1AhuSections.map((section) =>
            addr(`A${section.startRow - 1}:J${section.startRow - 1}`),
          ),
          ...ecm1AhuSections.map((section) =>
            addr(`A${section.startRow}:J${section.startRow}`),
          ),
        ];

        headerRanges.forEach((range) => {
          spreadsheet.cellFormat(
            {
              backgroundColor: "#d6e1ef",
              fontWeight: "bold",
              textAlign: "center",
              verticalAlign: "middle",
            },
            range,
          );
        });

        spreadsheet.cellFormat(
          { textAlign: "center", verticalAlign: "middle" },
          addr("A1:J220"),
        );

        const leftAlignRanges = [
          addr(
            `A3:A${getEndRow(ecm1SectionRows.inputVariables, ecm1InputVariables.length)}`,
          ),
          addr(
            `A8:A${getEndRow(ecm1SectionRows.energySavings, ecm1EnergyCostSavings.length)}`,
          ),
          addr(
            `I8:I${getEndRow(ecm1SectionRows.totals, ecm1TotalSavings.length)}`,
          ),
          addr(
            `A20:A${getEndRow(ecm1SectionRows.equipmentSummary, ecm1EquipmentRowsOrdered.length)}`,
          ),
        ];

        leftAlignRanges.forEach((range) => {
          spreadsheet.cellFormat({ textAlign: "left" }, range);
        });

        const tableRanges = [
          addr("A1:J1"),
          addr(
            `A2:B${getEndRow(ecm1SectionRows.inputVariables, ecm1InputVariables.length)}`,
          ),
          addr(
            `A7:D${getEndRow(ecm1SectionRows.energySavings, ecm1EnergyCostSavings.length)}`,
          ),
          addr(
            `I7:J${getEndRow(ecm1SectionRows.totals, ecm1TotalSavings.length)}`,
          ),
          addr(
            `A19:I${getEndRow(ecm1SectionRows.equipmentSummary, ecm1EquipmentRowsOrdered.length)}`,
          ),
          ...ecm1AhuSections.map((section) =>
            addr(
              `A${section.startRow - 1}:J${getEndRow(section.startRow, section.data.length)}`,
            ),
          ),
        ];

        tableRanges.forEach((range) => {
          spreadsheet.cellFormat(
            { textAlign: "center", verticalAlign: "middle" },
            range,
          );
          spreadsheet.setBorder(
            { border: "1px solid #000000" },
            range,
            "Inner",
          );
          spreadsheet.setBorder(
            { border: "2px solid #000000" },
            range,
            "Outer",
          );
        });

        spreadsheet.numberFormat(
          "$#,##0.00",
          addr(
            `D8:D${getEndRow(ecm1SectionRows.energySavings, ecm1EnergyCostSavings.length)}`,
          ),
        );
        spreadsheet.numberFormat(
          "#,##0",
          addr(
            `B8:B${getEndRow(ecm1SectionRows.energySavings, ecm1EnergyCostSavings.length)}`,
          ),
        );
        spreadsheet.numberFormat("#,##0.00", addr(`J8:J8`));
        spreadsheet.numberFormat("#,##0", addr(`J9:J9`));
        spreadsheet.numberFormat("$#,##0.00", addr(`J10:J10`));
        spreadsheet.numberFormat(
          "#,##0.00",
          addr(
            `B20:I${getEndRow(ecm1SectionRows.equipmentSummary, ecm1EquipmentRowsOrdered.length)}`,
          ),
        );

        ecm1AhuSections.forEach((section) => {
          const startRow = section.startRow;
          const endRow = getEndRow(startRow, section.data.length);
          spreadsheet.numberFormat("0.0", addr(`A${startRow + 1}:A${endRow}`));
          spreadsheet.numberFormat(
            "#,##0",
            addr(`B${startRow + 1}:B${endRow}`),
          );
          spreadsheet.numberFormat(
            "#,##0.00",
            addr(`C${startRow + 1}:C${endRow}`),
          );
          spreadsheet.numberFormat(
            "0.00%",
            addr(`D${startRow + 1}:D${endRow}`),
          );
          spreadsheet.numberFormat(
            "#,##0",
            addr(`E${startRow + 1}:E${endRow}`),
          );
          spreadsheet.numberFormat("0", addr(`F${startRow + 1}:G${endRow}`));
          spreadsheet.numberFormat("0.0", addr(`I${startRow + 1}:I${endRow}`));
          spreadsheet.numberFormat(
            "#,##0.00",
            addr(`J${startRow + 1}:J${endRow}`),
          );
        });
      } finally {
        spreadsheet.activeSheetIndex = previousSheetIndex;
        spreadsheet.goTo(`${previousSheetName}!A1`);
      }
    },
    [
      ecm1AhuSections,
      ecm1EnergyCostSavings.length,
      ecm1InputVariables.length,
      ecm1EquipmentRowsOrdered.length,
      ecm1SectionRows,
      ecm1TotalSavings.length,
      getSpreadsheet,
    ],
  );

  const applyEcm2DetailLayout = useCallback(
    (sheetName: string) => {
      const spreadsheet = getSpreadsheet();
      if (!spreadsheet) return;

      const targetSheetIndex = (spreadsheet.sheets ?? []).findIndex(
        (sheet) =>
          normalizeSheetName(sheet?.name) === normalizeSheetName(sheetName),
      );

      if (targetSheetIndex < 0) return;

      const targetSheetName =
        spreadsheet.sheets?.[targetSheetIndex]?.name ?? sheetName;
      const previousSheetIndex = spreadsheet.activeSheetIndex;
      const previousSheetName =
        spreadsheet.sheets?.[previousSheetIndex]?.name ?? targetSheetName;

      spreadsheet.activeSheetIndex = targetSheetIndex;
      spreadsheet.goTo(`${targetSheetName}!A1`);

      const addr = (range: string) => range;
      const getEndRow = (startRow: number, dataLength: number) =>
        startRow + Math.max(dataLength, 0);

      try {
        try {
          spreadsheet.merge(addr("A1:J1"));
        } catch {}

        spreadsheet.updateCell(
          {
            value: "ECM-2 AHU temperature scale adjustment",
            style: {
              fontWeight: "bold",
              textAlign: "center",
              verticalAlign: "middle",
            },
          },
          addr("A1"),
        );
        spreadsheet.cellFormat(
          {
            backgroundColor: "#2563eb",
            color: "#ffffff",
            fontWeight: "bold",
            textAlign: "center",
            verticalAlign: "middle",
          },
          addr("A1:J1"),
        );

        ecm2AhuSections.forEach((section) => {
          const sectionTitleRow = section.startRow - 1;
          try {
            spreadsheet.merge(addr(`A${sectionTitleRow}:J${sectionTitleRow}`));
          } catch {}
          spreadsheet.updateCell(
            { value: section.title },
            addr(`A${sectionTitleRow}`),
          );
        });

        const headerRanges = [
          addr("A4:B4"),
          addr("E4:F4"),
          addr("A9:D9"),
          addr("I9:J9"),
          addr("A21:I21"),
          addr("A31:C31"),
          ...ecm2AhuSections.map((section) =>
            addr(`A${section.startRow - 1}:J${section.startRow - 1}`),
          ),
          ...ecm2AhuSections.map((section) =>
            addr(`A${section.startRow}:J${section.startRow}`),
          ),
        ];

        headerRanges.forEach((range) => {
          spreadsheet.cellFormat(
            {
              backgroundColor: "#d6e1ef",
              fontWeight: "bold",
              textAlign: "center",
              verticalAlign: "middle",
            },
            range,
          );
        });

        spreadsheet.cellFormat(
          { textAlign: "center", verticalAlign: "middle" },
          addr("A1:J320"),
        );

        const leftAlignRanges = [
          addr(
            `A5:A${getEndRow(ecm2SectionRows.inputVariables, ecm2InputVariables.length)}`,
          ),
          addr(
            `E5:E${getEndRow(ecm2SectionRows.equations, ecm2Equations.length)}`,
          ),
          addr(
            `A10:A${getEndRow(ecm2SectionRows.energySavings, ecm2EnergySavings.length)}`,
          ),
          addr(
            `I10:I${getEndRow(ecm2SectionRows.totalGasSavings, ecm2TotalGasSavings.length)}`,
          ),
          addr(
            `A22:A${getEndRow(ecm2SectionRows.equipmentSummary, ecm2EquipmentSummary.length)}`,
          ),
          addr(
            `A32:A${getEndRow(ecm2SectionRows.rtuLoadProfile, ecm2RtuLoadProfile.length)}`,
          ),
        ];

        leftAlignRanges.forEach((range) => {
          spreadsheet.cellFormat({ textAlign: "left" }, range);
        });

        const tableRanges = [
          addr("A1:J1"),
          addr(
            `A4:B${getEndRow(ecm2SectionRows.inputVariables, ecm2InputVariables.length)}`,
          ),
          addr(
            `E4:F${getEndRow(ecm2SectionRows.equations, ecm2Equations.length)}`,
          ),
          addr(
            `A9:D${getEndRow(ecm2SectionRows.energySavings, ecm2EnergySavings.length)}`,
          ),
          addr(
            `I9:J${getEndRow(ecm2SectionRows.totalGasSavings, ecm2TotalGasSavings.length)}`,
          ),
          addr(
            `A21:I${getEndRow(ecm2SectionRows.equipmentSummary, ecm2EquipmentSummary.length)}`,
          ),
          addr(
            `A31:C${getEndRow(ecm2SectionRows.rtuLoadProfile, ecm2RtuLoadProfile.length)}`,
          ),
          ...ecm2AhuSections.map((section) =>
            addr(
              `A${section.startRow - 1}:J${getEndRow(section.startRow, section.data.length)}`,
            ),
          ),
        ];

        tableRanges.forEach((range) => {
          spreadsheet.cellFormat(
            { textAlign: "center", verticalAlign: "middle" },
            range,
          );
          spreadsheet.setBorder(
            { border: "1px solid #000000" },
            range,
            "Inner",
          );
          spreadsheet.setBorder(
            { border: "2px solid #000000" },
            range,
            "Outer",
          );
        });

        spreadsheet.numberFormat(
          "$#,##0.00",
          addr(
            `D10:D${getEndRow(ecm2SectionRows.energySavings, ecm2EnergySavings.length)}`,
          ),
        );
        spreadsheet.numberFormat(
          "#,##0",
          addr(
            `B10:B${getEndRow(ecm2SectionRows.energySavings, ecm2EnergySavings.length)}`,
          ),
        );
        spreadsheet.numberFormat("#,##0.00", addr(`J10:J10`));
        spreadsheet.numberFormat("$#,##0.00", addr(`J11:J11`));
        spreadsheet.numberFormat(
          "#,##0.00",
          addr(
            `B22:I${getEndRow(ecm2SectionRows.equipmentSummary, ecm2EquipmentSummary.length)}`,
          ),
        );

        ecm2AhuSections.forEach((section) => {
          const startRow = section.startRow;
          const endRow = getEndRow(startRow, section.data.length);
          spreadsheet.numberFormat("0.0", addr(`A${startRow + 1}:A${endRow}`));
          spreadsheet.numberFormat(
            "#,##0",
            addr(`B${startRow + 1}:B${endRow}`),
          );
          spreadsheet.numberFormat(
            "#,##0.00",
            addr(`C${startRow + 1}:C${endRow}`),
          );
          spreadsheet.numberFormat(
            "0.00%",
            addr(`D${startRow + 1}:D${endRow}`),
          );
          spreadsheet.numberFormat(
            "#,##0",
            addr(`E${startRow + 1}:E${endRow}`),
          );
          spreadsheet.numberFormat("0", addr(`F${startRow + 1}:G${endRow}`));
          spreadsheet.numberFormat("0.0", addr(`I${startRow + 1}:I${endRow}`));
          spreadsheet.numberFormat(
            "#,##0.00",
            addr(`J${startRow + 1}:J${endRow}`),
          );
        });
      } finally {
        spreadsheet.activeSheetIndex = previousSheetIndex;
        spreadsheet.goTo(`${previousSheetName}!A1`);
      }
    },
    [ecm2AhuSections, ecm2SectionRows, getSpreadsheet],
  );

  const handleCreated = useCallback(() => {
    clearSourceWorkbook();
    parsedWorkbookRef.current = null;
    applyLayoutOnDataBoundRef.current = true;
    hasAppliedInitialLayoutRef.current = false;

    if (hasSavedWorkbook()) {
      restoreFromLocalStorage();
    } else {
      void loadPublicTemplateWorkbook().then((loaded) => {
        if (!loaded) {
          setStatus('ℹ Using default data. Click "Save to Local" to persist.');
        }
      });
    }
  }, [loadPublicTemplateWorkbook, restoreFromLocalStorage]);

  const handleDataBound = useCallback(() => {
    if (!applyLayoutOnDataBoundRef.current) return;
    if (hasAppliedInitialLayoutRef.current) return;

    hasAppliedInitialLayoutRef.current = true;

    setTimeout(() => {
      applyImageLikeLayout();
      applyEcmDetailLayout(ECM1_DETAIL_SHEET_NAME);
      applyEcm2DetailLayout(ECM2_DETAIL_SHEET_NAME);
      window.setTimeout(() => {
        applyEcmDetailLayout(ECM1_DETAIL_SHEET_NAME);
        applyEcm2DetailLayout(ECM2_DETAIL_SHEET_NAME);
      }, 150);
      resetViewportToTop();
    }, 0);
  }, [applyEcmDetailLayout, applyImageLikeLayout, resetViewportToTop]);

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

  return (
    <div style={{ height: "100%", display: "flex", flexDirection: "column" }}>
      <div className="toolbar">
        <button type="button" className="toolbar-btn" onClick={handleSaveClick}>
          Save to Local
        </button>
        <button
          type="button"
          className="toolbar-btn"
          onClick={handleLoadFromFile}
        >
          Load Excel File
        </button>
        <button
          type="button"
          className="toolbar-btn primary"
          onClick={downloadAsExcel}
        >
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
          style={{ display: "none" }}
          onChange={handleFileSelected}
        />
      </div>

      {showNameDialog && (
        <div
          className="dialog-overlay"
          onClick={() => setShowNameDialog(false)}
        >
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
                display: "flex",
                gap: 8,
                justifyContent: "flex-end",
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
          dataBound={handleDataBound}
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
            <SheetDirective name="Project Input">
              <RangesDirective>
                <RangeDirective
                  dataSource={ecmRows}
                  startCell="C6"
                  showFieldAsHeader={false}
                />
                <RangeDirective
                  dataSource={hvacRows}
                  startCell="B19"
                  showFieldAsHeader={false}
                />
                <RangeDirective
                  dataSource={chillerRows}
                  startCell="B30"
                  showFieldAsHeader={false}
                />
                <RangeDirective
                  dataSource={freezerRows}
                  startCell="B35"
                  showFieldAsHeader={false}
                />
              </RangesDirective>
              <ColumnsDirective>
                <ColumnDirective width={70} />
                <ColumnDirective width={170} />
                <ColumnDirective width={145} />
                <ColumnDirective width={150} />
                <ColumnDirective width={145} />
                <ColumnDirective width={145} />
                <ColumnDirective width={120} />
                <ColumnDirective width={120} />
                <ColumnDirective width={95} />
                <ColumnDirective width={95} />
                <ColumnDirective width={140} />
                <ColumnDirective width={135} />
                <ColumnDirective width={115} />
                <ColumnDirective width={115} />
                <ColumnDirective width={105} />
                <ColumnDirective width={110} />
                <ColumnDirective width={100} />
                <ColumnDirective width={130} />
              </ColumnsDirective>
            </SheetDirective>
            <SheetDirective name={ECM1_DETAIL_SHEET_NAME}>
              <RangesDirective>
                <RangeDirective
                  dataSource={ecm1InputVariables}
                  startCell={`A${ecm1SectionRows.inputVariables}`}
                  showFieldAsHeader={true}
                />
                <RangeDirective
                  dataSource={ecm1EnergyCostSavingsWithGap}
                  startCell={`A${ecm1SectionRows.energySavings}`}
                  showFieldAsHeader={true}
                />
                <RangeDirective
                  dataSource={ecm1TotalSavings}
                  startCell={`I${ecm1SectionRows.totals}`}
                  showFieldAsHeader={true}
                />
                <RangeDirective
                  dataSource={ecm1EquipmentRowsOrdered}
                  startCell={`A${ecm1SectionRows.equipmentSummary}`}
                  showFieldAsHeader={true}
                />
                {ecm1AhuSections.map((section) => (
                  <RangeDirective
                    key={section.title}
                    dataSource={section.data}
                    startCell={`A${section.startRow}`}
                    showFieldAsHeader={true}
                  />
                ))}
              </RangesDirective>
              <ColumnsDirective>
                <ColumnDirective width={130} />
                <ColumnDirective width={240} />
                <ColumnDirective width={150} />
                <ColumnDirective width={170} />
                <ColumnDirective width={160} />
                <ColumnDirective width={150} />
                <ColumnDirective width={150} />
                <ColumnDirective width={130} />
                <ColumnDirective width={120} />
                <ColumnDirective width={120} />
              </ColumnsDirective>
            </SheetDirective>
            <SheetDirective name={ECM2_DETAIL_SHEET_NAME}>
              <RangesDirective>
                <RangeDirective
                  dataSource={ecm2InputVariables}
                  startCell={`A${ecm2SectionRows.inputVariables}`}
                  showFieldAsHeader={true}
                />
                <RangeDirective
                  dataSource={ecm2Equations}
                  startCell={`E${ecm2SectionRows.equations}`}
                  showFieldAsHeader={true}
                />
                <RangeDirective
                  dataSource={ecm2EnergySavings}
                  startCell={`A${ecm2SectionRows.energySavings}`}
                  showFieldAsHeader={true}
                />
                <RangeDirective
                  dataSource={ecm2TotalGasSavings}
                  startCell={`I${ecm2SectionRows.totalGasSavings}`}
                  showFieldAsHeader={true}
                />
                <RangeDirective
                  dataSource={ecm2EquipmentSummary}
                  startCell={`A${ecm2SectionRows.equipmentSummary}`}
                  showFieldAsHeader={true}
                />
                <RangeDirective
                  dataSource={ecm2RtuLoadProfile}
                  startCell={`A${ecm2SectionRows.rtuLoadProfile}`}
                  showFieldAsHeader={true}
                />
                {ecm2AhuSections.map((section) => (
                  <RangeDirective
                    key={`ecm2-${section.title}`}
                    dataSource={section.data}
                    startCell={`A${section.startRow}`}
                    showFieldAsHeader={true}
                  />
                ))}
              </RangesDirective>
              <ColumnsDirective>
                <ColumnDirective width={130} />
                <ColumnDirective width={240} />
                <ColumnDirective width={150} />
                <ColumnDirective width={170} />
                <ColumnDirective width={160} />
                <ColumnDirective width={150} />
                <ColumnDirective width={150} />
                <ColumnDirective width={130} />
                <ColumnDirective width={120} />
                <ColumnDirective width={120} />
              </ColumnsDirective>
            </SheetDirective>
          </SheetsDirective>
        </SpreadsheetComponent>
      </div>
    </div>
  );
}

export default App;
