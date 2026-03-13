import type { SpreadsheetComponent } from "@syncfusion/ej2-react-spreadsheet";
import {
  ECM1_DETAIL_SHEET_NAME,
  ecm1EnergyCostSavings,
  ecm1Equations,
  ecm1EquipmentRowsOrdered,
  ecm1InputVariables,
  ecm1TotalSavings,
  projectHeader,
} from "../../datasource";
import {
  ecm2EnergySavings,
  ecm2EquipmentSummary,
  ecm2Equations,
  ecm2InputVariables,
  ecm2RtuLoadProfile,
  ecm2TotalGasSavings,
} from "../../ecm2Datasource";
import {
  ECM1_DETAIL_SHEET_ALIASES,
  ECM2_DETAIL_SHEET_NAME,
  PROJECT_INPUT_SHEET_ALIASES,
  PROJECT_INPUT_SHEET_NAME,
  PROJECT_INPUT_TEMPLATE_TITLE,
  ecm1AhuSections,
  ecm1SectionRows,
  ecm2AhuSections,
  ecm2SectionRows,
  normalizeSheetName,
} from "./sheetConfig";

interface LayoutContext {
  spreadsheet: SpreadsheetComponent;
  setStatus: (message: string) => void;
}

const getEndRow = (startRow: number, dataLength: number) =>
  startRow + Math.max(dataLength, 0);

const safeMerge = (spreadsheet: SpreadsheetComponent, range: string) => {
  try {
    spreadsheet.merge(range);
  } catch {
  }
};

const safeUnmerge = (spreadsheet: SpreadsheetComponent, range: string) => {
  try {
    spreadsheet.unMerge(range);
  } catch {
  }
};

export function applyProjectInputLayout({
  spreadsheet,
  setStatus,
}: LayoutContext): void {
  const sheets = spreadsheet.sheets ?? [];
  const normalizedPrimaryProjectInputSheet = normalizeSheetName(
    PROJECT_INPUT_SHEET_NAME,
  );

  const exactProjectInputSheetIndex = sheets.findIndex(
    (sheet) => normalizeSheetName(sheet?.name) === normalizedPrimaryProjectInputSheet,
  );

  const targetSheetIndex =
    exactProjectInputSheetIndex >= 0
      ? exactProjectInputSheetIndex
      : sheets.findIndex((sheet) =>
          PROJECT_INPUT_SHEET_ALIASES.some(
            (alias) => normalizeSheetName(sheet?.name) === normalizeSheetName(alias),
          ),
        );

  if (targetSheetIndex < 0) {
    setStatus(" Project input sheet not found.");
    return;
  }

  const targetSheetName =
    spreadsheet.sheets?.[targetSheetIndex]?.name ?? PROJECT_INPUT_SHEET_NAME;
  const addr = (range: string) => `${targetSheetName}!${range}`;
  const header = projectHeader[0];

  // Reset any pre-existing top-area merges so row 1-3 values can be placed per cell.
  safeUnmerge(spreadsheet, addr("A1:X3"));
  safeUnmerge(spreadsheet, addr("A1:T1"));
  safeUnmerge(spreadsheet, addr("V1:X1"));
  safeUnmerge(spreadsheet, addr("A2:X2"));
  safeUnmerge(spreadsheet, addr("A3:X3"));

  
  spreadsheet.updateCell({ value: PROJECT_INPUT_TEMPLATE_TITLE }, addr("A1:T1"));
  spreadsheet.updateCell({ value: "" }, addr("U1"));
  spreadsheet.updateCell({ value: header.Date }, addr("V1:X1"));

  spreadsheet.updateCell({ value: "Project Utility" }, addr("A2"));
  spreadsheet.updateCell({ value: header.ProjectUtility }, addr("E2"));
  spreadsheet.updateCell({ value: "Project Name" }, addr("I2"));
  spreadsheet.updateCell({ value: header.ProjectName }, addr("L2"));
  spreadsheet.updateCell({ value: "Project Type" }, addr("O2"));
  spreadsheet.updateCell({ value: header.ProjectType }, addr("Q2"));
  spreadsheet.updateCell({ value: header.ProgramType }, addr("S2"));

  spreadsheet.updateCell({ value: "Square Footage" }, addr("A3"));
  spreadsheet.updateCell({ value: String(header.SquareFootage) }, addr("E3"));
  spreadsheet.updateCell({ value: "(S.F.) Annual Therm usage" }, addr("G3"));
  spreadsheet.updateCell({ value: String(header.AnnualThermUsage) }, addr("L3"));
  spreadsheet.updateCell({ value: "Therm" }, addr("N3"));

  spreadsheet.updateCell({ value: "No" }, addr("B5"));
  spreadsheet.updateCell({ value: "Energy Conservation Method" }, addr("C5"));
  spreadsheet.updateCell({ value: "Estimated Cost ($)" }, addr("D5"));
  spreadsheet.updateCell({ value: "Natural Gas Savings (Therm/yr)" }, addr("E5"));
  spreadsheet.updateCell(
    { value: "Natural Gas Energy Cost Savings ($/yr)" },
    addr("F5"),
  );
  spreadsheet.updateCell({ value: "Simple Payback (years gas)" }, addr("G5"));

  spreadsheet.updateCell({ value: "Estimated Cost ($)" }, addr("E9"));
  spreadsheet.updateCell({ value: "Natural Gas Savings (Therm/yr)" }, addr("F9"));
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
    { value: "Total Project Summary (After Rebate) " },
    addr("C11"),
  );
  spreadsheet.updateCell({ formula: "=E10-K15 " }, addr("E11"));
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
  spreadsheet.updateCell({ value: "Supply Fan Motor Load Factor" }, addr("K18"));
  spreadsheet.updateCell({ value: "Supply Fan Motor Efficiency" }, addr("L18"));
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

  safeMerge(spreadsheet, addr("A1:T1"));
  safeMerge(spreadsheet, addr("V1:X1"));
  safeMerge(spreadsheet, addr("A2:D2"));
  safeMerge(spreadsheet, addr("E2:H2"));
  safeMerge(spreadsheet, addr("I2:K2"));
  safeMerge(spreadsheet, addr("L2:N2"));
  safeMerge(spreadsheet, addr("O2:P2"));
  safeMerge(spreadsheet, addr("Q2:R2"));
  safeMerge(spreadsheet, addr("S2:X2"));
  safeMerge(spreadsheet, addr("A3:D3"));
  safeMerge(spreadsheet, addr("E3:F3"));
  safeMerge(spreadsheet, addr("G3:K3"));
  safeMerge(spreadsheet, addr("L3:M3"));
  safeMerge(spreadsheet, addr("K14:N14"));
  safeMerge(spreadsheet, addr("K15:N15"));
  safeMerge(spreadsheet, addr("A18:A26"));
  safeMerge(spreadsheet, addr("A29:A31"));
  safeMerge(spreadsheet, addr("A34:A36"));

  spreadsheet.cellFormat(
    { textAlign: "center", verticalAlign: "middle", fontWeight: "bold" },
    addr("A1:T1"),
  );

  spreadsheet.cellFormat(
    { fontWeight: "bold", textAlign: "center", verticalAlign: "middle" },
    addr("A1:X3"),
  );
  spreadsheet.cellFormat(
    { backgroundColor: "#b9c8dc", fontWeight: "bold", textAlign: "center" },
    addr("A1:X3"),
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
    addr("A1:X36"),
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

  spreadsheet.updateCell({ formula: '=IF(F6>0,D6/F6,"")' }, addr("G6"));
  spreadsheet.updateCell({ formula: "=IF(F7<>0,D7/F7,0)" }, addr("G7"));

  spreadsheet.updateCell({ formula: '=IF(C15>0,(C15-D15)/C15,"")' }, addr("E15"));
  spreadsheet.updateCell({ formula: "=IF(E3<>0,C15/E3,0)" }, addr("F15"));
  spreadsheet.updateCell({ formula: `=MIN((E10/2),(F10*3.7))` }, addr("K15"));

  spreadsheet.updateCell({ formula: "=L3" }, addr("C15"));
  spreadsheet.updateCell({ formula: "=C15-SUM(E6:E7)" }, addr("D15"));

  const ecm1EnergyTotalsRow =
    ecm1SectionRows.energySavings + ecm1EnergyCostSavings.length;
  spreadsheet.updateCell({ formula: `='${ECM1_DETAIL_SHEET_NAME}'!J8` }, addr("E6"));
  spreadsheet.updateCell(
    { formula: `='${ECM1_DETAIL_SHEET_NAME}'!D${ecm1EnergyTotalsRow}` },
    addr("F6"),
  );

  spreadsheet.updateCell({ formula: `='${ECM2_DETAIL_SHEET_NAME}'!J10` }, addr("E7"));
  spreadsheet.updateCell({ formula: `='${ECM2_DETAIL_SHEET_NAME}'!J11` }, addr("F7"));

  for (let r = 19; r <= 26; r++) {
    spreadsheet.updateCell({ formula: `=IF(E${r}<>0,12/E${r},0)` }, addr(`F${r}`));
  }
  for (let r = 30; r <= 31; r++) {
    spreadsheet.updateCell({ formula: `=IF(E${r}<>0,12/E${r},0)` }, addr(`F${r}`));
  }
  for (let r = 35; r <= 36; r++) {
    spreadsheet.updateCell({ formula: `=IF(D${r}<>0,12/D${r},0)` }, addr(`E${r}`));
    spreadsheet.updateCell({ formula: `=F${r}*0.746` }, addr(`G${r}`));
  }

  const tableRanges = [
    "A1:X3",
    "B5:G7",
    "E9:H11",
    "C14:F15",
    "K14:N15",
    "A18:R26",
    "A29:F31",
    "A34:G36",
  ];

  tableRanges.forEach((range) => {
    spreadsheet.setBorder({ border: "1px solid #000000" }, addr(range), "Inner");
    spreadsheet.setBorder({ border: "2px solid #000000" }, addr(range), "Outer");
  });
}

export function applyEcm1DetailLayout({
  spreadsheet,
  setStatus,
}: LayoutContext): void {
  const targetSheetIndex = (spreadsheet.sheets ?? []).findIndex((sheet) =>
    ECM1_DETAIL_SHEET_ALIASES.some(
      (candidate) =>
        normalizeSheetName(sheet?.name) === normalizeSheetName(candidate),
    ),
  );

  if (targetSheetIndex < 0) {
    setStatus(` Sheet "${ECM1_DETAIL_SHEET_NAME}" not found.`);
    return;
  }

  const targetSheetName =
    spreadsheet.sheets?.[targetSheetIndex]?.name ?? ECM1_DETAIL_SHEET_NAME;
  const targetSheet = spreadsheet.sheets?.[targetSheetIndex];
  if (!targetSheet) return;

  const previousSheetIndex = spreadsheet.activeSheetIndex;
  const previousSheetName =
    spreadsheet.sheets?.[previousSheetIndex]?.name ?? targetSheetName;

  spreadsheet.activeSheetIndex = targetSheetIndex;
  spreadsheet.goTo(`${targetSheetName}!A1`);

  const addr = (range: string) => range;
  const {
    inputVariables: inputVarsRow,
    energySavings: energySavingsRow,
    totals: totalsRow,
    equipmentSummary: equipmentSummaryRow,
  } = ecm1SectionRows;
  const colLetter = (colIndex: number) => {
    let n = colIndex;
    let s = "";
    while (n > 0) {
      const rem = (n - 1) % 26;
      s = String.fromCharCode(65 + rem) + s;
      n = Math.floor((n - 1) / 26);
    }
    return s;
  };

  try {
    safeMerge(spreadsheet, addr("A2:D2"));
    spreadsheet.updateCell({ value: "ECM-" }, addr("A2"));
    spreadsheet.updateCell({ formula: "='Input Sheet -W''Gas'!B6" }, addr("E2"));
    for (let col = 6; col <= 21; col++) {
      spreadsheet.updateCell(
        { formula: "='Input Sheet -W''Gas'!C6" },
        addr(`${colLetter(col)}2`),
      );
    }

    safeMerge(spreadsheet, addr(`A${inputVarsRow}:E${inputVarsRow}`));
    safeMerge(spreadsheet, addr(`G${inputVarsRow}:M${inputVarsRow}`));
    spreadsheet.updateCell({ value: "Input Variables" }, addr(`A${inputVarsRow}`));
    spreadsheet.updateCell({ value: "Equations & Variables" }, addr(`G${inputVarsRow}`));

    const inputDataStartRow = inputVarsRow + 1;
    safeMerge(spreadsheet, addr(`A${inputDataStartRow}:C${inputDataStartRow}`));
    safeMerge(spreadsheet, addr(`D${inputDataStartRow}:E${inputDataStartRow}`));
    safeMerge(spreadsheet, addr(`A${inputDataStartRow + 1}:C${inputDataStartRow + 1}`));
    safeMerge(spreadsheet, addr(`D${inputDataStartRow + 1}:E${inputDataStartRow + 1}`));
    safeMerge(spreadsheet, addr(`G${inputDataStartRow}:H${inputDataStartRow}`));
    safeMerge(spreadsheet, addr(`I${inputDataStartRow}:M${inputDataStartRow}`));
    safeMerge(spreadsheet, addr(`G${inputDataStartRow + 1}:H${inputDataStartRow + 1}`));
    safeMerge(spreadsheet, addr(`I${inputDataStartRow + 1}:M${inputDataStartRow + 1}`));

    spreadsheet.updateCell(
      { value: ecm1InputVariables[0]["Input Variables"] },
      addr(`A${inputDataStartRow}`),
    );
    spreadsheet.updateCell({ value: String(ecm1InputVariables[0].Value) }, addr(`D${inputDataStartRow}`));
    spreadsheet.updateCell(
      { value: ecm1InputVariables[1]["Input Variables"] },
      addr(`A${inputDataStartRow + 1}`),
    );
    spreadsheet.updateCell({ value: String(ecm1InputVariables[1].Value) }, addr(`D${inputDataStartRow + 1}`));

    spreadsheet.updateCell(
      { value: ecm1Equations[0]["Equations & Variables"] },
      addr(`G${inputDataStartRow}`),
    );
    spreadsheet.updateCell({ value: ecm1Equations[0].Equation }, addr(`I${inputDataStartRow}`));
    spreadsheet.updateCell(
      { value: ecm1Equations[1]["Equations & Variables"] },
      addr(`G${inputDataStartRow + 1}`),
    );
    spreadsheet.updateCell({ value: ecm1Equations[1].Equation }, addr(`I${inputDataStartRow + 1}`));

    ecm1AhuSections.forEach((section) => {
      const sectionTitleRow = section.startRow - 1;
      safeMerge(spreadsheet, addr(`A${sectionTitleRow}:J${sectionTitleRow}`));
      spreadsheet.updateCell({ value: section.title }, addr(`A${sectionTitleRow}`));
    });
    
    const headerRanges = [
      addr(`A${inputVarsRow}:E${inputVarsRow}`),
      addr(`G${inputVarsRow}:M${inputVarsRow}`),
      addr(`A${energySavingsRow}:D${energySavingsRow}`),
      addr(`I${totalsRow}:J${totalsRow}`),
      addr(`A${equipmentSummaryRow}:I${equipmentSummaryRow}`),
      ...ecm1AhuSections.map((section) =>
        addr(`A${section.startRow - 1}:J${section.startRow - 1}`),
      ),
      ...ecm1AhuSections.map((section) => addr(`A${section.startRow}:J${section.startRow}`)),
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
      addr("A1:U220"),
    );

    spreadsheet.cellFormat(
      { textAlign: "right", fontWeight: "bold" },
      addr("A2:D2"),
    );

    const inputEndRow = inputVarsRow + ecm1InputVariables.length;
    spreadsheet.setBorder({ border: "1px solid #000000" }, addr(`A${inputVarsRow}:E${inputEndRow}`), "Inner");
    spreadsheet.setBorder({ border: "2px solid #000000" }, addr(`A${inputVarsRow}:E${inputEndRow}`), "Outer");
    spreadsheet.setBorder({ border: "1px solid #000000" }, addr(`G${inputVarsRow}:M${inputEndRow}`), "Inner");
    spreadsheet.setBorder({ border: "2px solid #000000" }, addr(`G${inputVarsRow}:M${inputEndRow}`), "Outer");

    const leftAlignRanges = [
      addr(`A${inputVarsRow + 1}:A${inputEndRow}`),
      addr(`G${inputVarsRow + 1}:G${inputEndRow}`),
      addr(`A${energySavingsRow + 1}:A${getEndRow(energySavingsRow, ecm1EnergyCostSavings.length)}`),
      addr(`I${totalsRow + 1}:I${getEndRow(totalsRow, ecm1TotalSavings.length)}`),
      addr(
        `A${equipmentSummaryRow + 1}:A${getEndRow(equipmentSummaryRow, ecm1EquipmentRowsOrdered.length)}`
      ),
    ];

    leftAlignRanges.forEach((range) => {
      spreadsheet.cellFormat({ textAlign: "left" }, range);
    });

    const tableRanges = [
      addr(`A${energySavingsRow}:D${getEndRow(energySavingsRow, ecm1EnergyCostSavings.length)}`),
      addr(`I${totalsRow}:J${getEndRow(totalsRow, ecm1TotalSavings.length)}`),
      addr(`A${equipmentSummaryRow}:I${getEndRow(equipmentSummaryRow, ecm1EquipmentRowsOrdered.length)}`),
      ...ecm1AhuSections.map((section) =>
        addr(`A${section.startRow - 1}:J${getEndRow(section.startRow, section.data.length)}`),
      ),
    ];

    tableRanges.forEach((range) => {
      spreadsheet.cellFormat({ textAlign: "center", verticalAlign: "middle" }, range);
      spreadsheet.setBorder({ border: "1px solid #000000" }, range, "Inner");
      spreadsheet.setBorder({ border: "2px solid #000000" }, range, "Outer");
    });

    spreadsheet.numberFormat(
      "$#,##0.00",
      addr(`D${energySavingsRow + 1}:D${getEndRow(energySavingsRow, ecm1EnergyCostSavings.length)}`),
    );
    spreadsheet.numberFormat(
      "#,##0",
      addr(`B${energySavingsRow + 1}:B${getEndRow(energySavingsRow, ecm1EnergyCostSavings.length)}`),
    );
    spreadsheet.numberFormat("#,##0.00", addr(`J${totalsRow + 1}:J${totalsRow + 1}`));
    spreadsheet.numberFormat("$#,##0.00", addr(`J${totalsRow + 2}:J${totalsRow + 2}`));
    spreadsheet.numberFormat(
      "#,##0.00",
      addr(`B${equipmentSummaryRow + 1}:I${getEndRow(equipmentSummaryRow, ecm1EquipmentRowsOrdered.length)}`),
    );

    ecm1AhuSections.forEach((section) => {
      const startRow = section.startRow;
      const endRow = getEndRow(startRow, section.data.length);
      spreadsheet.numberFormat("0.0", addr(`A${startRow + 1}:A${endRow}`));
      spreadsheet.numberFormat("#,##0", addr(`B${startRow + 1}:B${endRow}`));
      spreadsheet.numberFormat("#,##0.00", addr(`C${startRow + 1}:C${endRow}`));
      spreadsheet.numberFormat("0.00%", addr(`D${startRow + 1}:D${endRow}`));
      spreadsheet.numberFormat("#,##0", addr(`E${startRow + 1}:E${endRow}`));
      spreadsheet.numberFormat("0", addr(`F${startRow + 1}:G${endRow}`));
      spreadsheet.numberFormat("0.0", addr(`I${startRow + 1}:I${endRow}`));
      spreadsheet.numberFormat("#,##0.00", addr(`J${startRow + 1}:J${endRow}`));
    });
  } finally {
    spreadsheet.activeSheetIndex = previousSheetIndex;
    spreadsheet.goTo(`${previousSheetName}!A1`);
  }
}

export function applyEcm2DetailLayout({
  spreadsheet,
}: Omit<LayoutContext, "setStatus">): void {
  const targetSheetIndex = (spreadsheet.sheets ?? []).findIndex(
    (sheet) => normalizeSheetName(sheet?.name) === normalizeSheetName(ECM2_DETAIL_SHEET_NAME),
  );

  if (targetSheetIndex < 0) return;

  const targetSheetName =
    spreadsheet.sheets?.[targetSheetIndex]?.name ?? ECM2_DETAIL_SHEET_NAME;
  const previousSheetIndex = spreadsheet.activeSheetIndex;
  const previousSheetName =
    spreadsheet.sheets?.[previousSheetIndex]?.name ?? targetSheetName;

  spreadsheet.activeSheetIndex = targetSheetIndex;
  spreadsheet.goTo(`${targetSheetName}!A1`);

  const addr = (range: string) => range;
  const {
    inputVariables: inputVarsRow,
    equations: equationsRow,
    energySavings: energySavingsRow,
    totalGasSavings: totalGasSavingsRow,
    equipmentSummary: equipmentSummaryRow,
    rtuLoadProfile: rtuLoadProfileRow,
  } = ecm2SectionRows;

  try {
    safeMerge(spreadsheet, addr("A1:J1"));

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
      safeMerge(spreadsheet, addr(`A${sectionTitleRow}:J${sectionTitleRow}`));
      spreadsheet.updateCell({ value: section.title }, addr(`A${sectionTitleRow}`));
    });

    const headerRanges = [
      addr(`A${inputVarsRow}:B${inputVarsRow}`),
      addr(`E${equationsRow}:F${equationsRow}`),
      addr(`A${energySavingsRow}:D${energySavingsRow}`),
      addr(`I${totalGasSavingsRow}:J${totalGasSavingsRow}`),
      addr(`A${equipmentSummaryRow}:I${equipmentSummaryRow}`),
      addr(`A${rtuLoadProfileRow}:C${rtuLoadProfileRow}`),
      ...ecm2AhuSections.map((section) =>
        addr(`A${section.startRow - 1}:J${section.startRow - 1}`),
      ),
      ...ecm2AhuSections.map((section) => addr(`A${section.startRow}:J${section.startRow}`)),
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
      addr(`A${inputVarsRow + 1}:A${getEndRow(inputVarsRow, ecm2InputVariables.length)}`),
      addr(`E${equationsRow + 1}:E${getEndRow(equationsRow, ecm2Equations.length)}`),
      addr(`A${energySavingsRow + 1}:A${getEndRow(energySavingsRow, ecm2EnergySavings.length)}`),
      addr(`I${totalGasSavingsRow + 1}:I${getEndRow(totalGasSavingsRow, ecm2TotalGasSavings.length)}`),
      addr(`A${equipmentSummaryRow + 1}:A${getEndRow(equipmentSummaryRow, ecm2EquipmentSummary.length)}`),
      addr(`A${rtuLoadProfileRow + 1}:A${getEndRow(rtuLoadProfileRow, ecm2RtuLoadProfile.length)}`),
    ];

    leftAlignRanges.forEach((range) => {
      spreadsheet.cellFormat({ textAlign: "left" }, range);
    });

    const tableRanges = [
      addr("A1:J1"),
      addr(`A${inputVarsRow}:B${getEndRow(inputVarsRow, ecm2InputVariables.length)}`),
      addr(`E${equationsRow}:F${getEndRow(equationsRow, ecm2Equations.length)}`),
      addr(`A${energySavingsRow}:D${getEndRow(energySavingsRow, ecm2EnergySavings.length)}`),
      addr(`I${totalGasSavingsRow}:J${getEndRow(totalGasSavingsRow, ecm2TotalGasSavings.length)}`),
      addr(`A${equipmentSummaryRow}:I${getEndRow(equipmentSummaryRow, ecm2EquipmentSummary.length)}`),
      addr(`A${rtuLoadProfileRow}:C${getEndRow(rtuLoadProfileRow, ecm2RtuLoadProfile.length)}`),
      ...ecm2AhuSections.map((section) =>
        addr(`A${section.startRow - 1}:J${getEndRow(section.startRow, section.data.length)}`),
      ),
    ];

    tableRanges.forEach((range) => {
      spreadsheet.cellFormat({ textAlign: "center", verticalAlign: "middle" }, range);
      spreadsheet.setBorder({ border: "1px solid #000000" }, range, "Inner");
      spreadsheet.setBorder({ border: "2px solid #000000" }, range, "Outer");
    });

    spreadsheet.numberFormat(
      "$#,##0.00",
      addr(`D${energySavingsRow + 1}:D${getEndRow(energySavingsRow, ecm2EnergySavings.length)}`),
    );
    spreadsheet.numberFormat(
      "#,##0",
      addr(`B${energySavingsRow + 1}:B${getEndRow(energySavingsRow, ecm2EnergySavings.length)}`),
    );
    spreadsheet.numberFormat("#,##0.00", addr(`J${totalGasSavingsRow + 1}:J${totalGasSavingsRow + 1}`));
    spreadsheet.numberFormat("$#,##0.00", addr(`J${totalGasSavingsRow + 2}:J${totalGasSavingsRow + 2}`));
    spreadsheet.numberFormat(
      "#,##0.00",
      addr(`B${equipmentSummaryRow + 1}:I${getEndRow(equipmentSummaryRow, ecm2EquipmentSummary.length)}`),
    );

    ecm2AhuSections.forEach((section) => {
      const startRow = section.startRow;
      const endRow = getEndRow(startRow, section.data.length);
      spreadsheet.numberFormat("0.0", addr(`A${startRow + 1}:A${endRow}`));
      spreadsheet.numberFormat("#,##0", addr(`B${startRow + 1}:B${endRow}`));
      spreadsheet.numberFormat("#,##0.00", addr(`C${startRow + 1}:C${endRow}`));
      spreadsheet.numberFormat("0.00%", addr(`D${startRow + 1}:D${endRow}`));
      spreadsheet.numberFormat("#,##0", addr(`E${startRow + 1}:E${endRow}`));
      spreadsheet.numberFormat("0", addr(`F${startRow + 1}:G${endRow}`));
      spreadsheet.numberFormat("0.0", addr(`I${startRow + 1}:I${endRow}`));
      spreadsheet.numberFormat("#,##0.00", addr(`J${startRow + 1}:J${endRow}`));
    });
  } finally {
    spreadsheet.activeSheetIndex = previousSheetIndex;
    spreadsheet.goTo(`${previousSheetName}!A1`);
  }
}
