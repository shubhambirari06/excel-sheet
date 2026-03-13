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
  ecmRows,
  hvacRows,
  chillerRows,
  freezerRows,
  ecm1EquipmentRowsOrdered,
  ecm1TotalSavings,
  ECM1_DETAIL_SHEET_NAME,
} from "../../datasource";
import {
  ecm2InputVariables,
  ecm2Equations,
  ecm2EquipmentSummary,
  ecm2EnergySavings,
  ecm2TotalGasSavings,
  ecm2RtuLoadProfile,
} from "../../ecm2Datasource";
import {
  ECM2_DETAIL_SHEET_NAME,
  PROJECT_INPUT_SHEET_NAME,
  ecm1AhuSections,
  ecm1EnergyCostSavingsWithGap,
  ecm1SectionRows,
  ecm2AhuSections,
  ecm2SectionRows,
} from "./sheetConfig";

interface SpreadsheetShellProps {
  spreadsheetRef: React.RefObject<SpreadsheetComponent | null>;
  onCreated: () => void;
  onDataBound: () => void;
  onScheduleAutoSave: () => void;
}

export function SpreadsheetShell({
  spreadsheetRef,
  onCreated,
  onDataBound,
  onScheduleAutoSave,
}: SpreadsheetShellProps) {
  return (
    <div style={{ flex: 1, minHeight: 0 }}>
      <SpreadsheetComponent
        ref={spreadsheetRef}
        created={onCreated}
        dataBound={onDataBound}
        cellSave={onScheduleAutoSave}
        actionComplete={onScheduleAutoSave}
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
          <SheetDirective name={PROJECT_INPUT_SHEET_NAME}>
            <RangesDirective>
              <RangeDirective dataSource={ecmRows} startCell="C6" showFieldAsHeader={false} />
              <RangeDirective dataSource={hvacRows} startCell="B19" showFieldAsHeader={false} />
              <RangeDirective dataSource={chillerRows} startCell="B30" showFieldAsHeader={false} />
              <RangeDirective dataSource={freezerRows} startCell="B35" showFieldAsHeader={false} />
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
  );
}
