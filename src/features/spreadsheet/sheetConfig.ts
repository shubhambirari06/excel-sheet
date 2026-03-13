import {
  ECM1_DETAIL_SHEET_NAME,
  ecm1BinDataAHU1,
  ecm1BinDataAHU2,
  ecm1BinDataAHU3,
  ecm1BinDataAHU4,
  ecm1BinDataAHU5,
  ecm1BinDataAHU6,
  ecm1EnergyCostSavings,
  ecm1EquipmentRowsOrdered,
  ecm1InputVariables,
  ecm1TotalSavings,
} from "../../datasource";
import {
  ecm2BinDataAHU1,
  ecm2BinDataAHU2,
  ecm2BinDataAHU3,
  ecm2BinDataAHU4,
  ecm2BinDataAHU5,
  ecm2BinDataAHU6,
  ecm2BinDataAHU7,
  ecm2BinDataRTU1,
  ecm2EnergySavings,
  ecm2EquipmentSummary,
  ecm2Equations,
  ecm2InputVariables,
  ecm2RtuLoadProfile,
  ecm2TotalGasSavings,
} from "../../ecm2Datasource";

export const ACCEPTED_EXTENSIONS = ".xlsx,.xls,.csv";
export const PROJECT_INPUT_SHEET_NAME = "Input Sheet -W'Gas";
export const PROJECT_INPUT_SHEET_ALIASES = [
  PROJECT_INPUT_SHEET_NAME,
  "Project Input",
  "Project Sheet",
  "Input Sheet",
];
export const PROJECT_INPUT_TEMPLATE_TITLE = "PROJECT INPUT SHEET: HBS SOLUTION";
export const ECM1_DETAIL_SHEET_ALIASES = [ECM1_DETAIL_SHEET_NAME];
export const ECM2_DETAIL_SHEET_NAME = "ECM2";

const calculateLayouts = () => {
  // ECM1
  let ecm1Row = 4;
  const ecm1 = {
    inputVariables: ecm1Row,
    equations: ecm1Row,
    equipmentSummary: 0,
    energySavings: 0,
    totals: 0,
  };
  ecm1Row += Math.max(ecm1InputVariables.length) + 1; // Section ends
  ecm1Row += 1; // Gap
  ecm1.equipmentSummary = ecm1Row;
  ecm1Row += ecm1EquipmentRowsOrdered.length + 1;
  ecm1Row += 3; // Gap
  ecm1.energySavings = ecm1Row;
  ecm1Row += ecm1EnergyCostSavings.length + 1;
  ecm1Row += 12; // Gap
  ecm1.totals = ecm1Row;
  ecm1Row += ecm1TotalSavings.length + 1;
  ecm1Row += 42; // Gap
  const ecm1AhuStart = ecm1Row;

  // ECM2
  let ecm2Row = 4;
  const ecm2 = {
    inputVariables: ecm2Row,
    equations: ecm2Row,
    energySavings: 0,
    totalGasSavings: 0,
    equipmentSummary: 0,
    rtuLoadProfile: 0,
  };
  ecm2Row += Math.max(ecm2InputVariables.length, ecm2Equations.length) + 1;
  ecm2Row += 2; // Gap
  ecm2.energySavings = ecm2Row;
  ecm2.totalGasSavings = ecm2Row;
  ecm2Row += Math.max(ecm2EnergySavings.length, ecm2TotalGasSavings.length) + 1;
  ecm2Row += 6; // Gap
  ecm2.equipmentSummary = ecm2Row;
  ecm2Row += ecm2EquipmentSummary.length + 1;
  ecm2Row += 6; // Gap
  ecm2.rtuLoadProfile = ecm2Row;
  ecm2Row += ecm2RtuLoadProfile.length + 1;
  ecm2Row += 8; // Gap
  const ecm2AhuStart = ecm2Row;

  return { ecm1, ecm1AhuStart, ecm2, ecm2AhuStart };
};
const { ecm1, ecm1AhuStart, ecm2, ecm2AhuStart } = calculateLayouts();
export const ecm1SectionRows = ecm1;
export const ecm2SectionRows = ecm2;

export interface SpreadsheetSection {
  title: string;
  data: Array<Record<string, unknown>>;
  startRow: number;
}

export const normalizeSheetName = (name: string | undefined) =>
  String(name ?? "")
    .trim()
    .toLowerCase();

export const insertGapColumn = (
  data: Array<Record<string, unknown>>,
  index: number,
): Array<Record<string, unknown>> => {
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

export const ecm1EnergyCostSavingsWithGap = insertGapColumn(
  ecm1EnergyCostSavings,
  2,
);

export const ecm1AhuSections: SpreadsheetSection[] = [
  { title: "AHU #1", data: insertGapColumn(ecm1BinDataAHU1, 7) },
  { title: "AHU #2", data: insertGapColumn(ecm1BinDataAHU2, 7) },
  { title: "AHU #3", data: insertGapColumn(ecm1BinDataAHU3, 7) },
  { title: "AHU #4", data: insertGapColumn(ecm1BinDataAHU4, 7) },
  { title: "AHU #5", data: insertGapColumn(ecm1BinDataAHU5, 7) },
  { title: "AHU #6", data: insertGapColumn(ecm1BinDataAHU6, 7) },
].map((section, index, list) => {
  const previousRows = list
    .slice(0, index)
    .reduce((sum, prev) => sum + (prev.data.length + 1), 0);

  return {
    ...section,
    startRow: ecm1AhuStart + previousRows,
  };
});

export const ecm2AhuSections: SpreadsheetSection[] = [
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
  return { ...section, startRow: ecm2AhuStart + previousRows };
});
