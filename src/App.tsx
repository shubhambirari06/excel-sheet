import "./App.css";
import { SpreadsheetToolbar } from "./components/SpreadsheetToolbar";
import { SaveWorkbookDialog } from "./components/SaveWorkbookDialog";
import { SpreadsheetShell } from "./features/spreadsheet/SpreadsheetShell";
import { ACCEPTED_EXTENSIONS } from "./features/spreadsheet/sheetConfig";
import { useSpreadsheetController } from "./hooks/useSpreadsheetController";

function App() {
  const {
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
  } = useSpreadsheetController();

  return (
    <div style={{ height: "100%", display: "flex", flexDirection: "column" }}>
      <SpreadsheetToolbar
        fileName={fileName}
        status={status}
        onSaveToLocal={handleSaveClick}
        onLoadExcelFile={handleLoadFromFile}
        onDownloadXlsx={downloadAsExcel}
        fileInputRef={fileInputRef}
        acceptedExtensions={ACCEPTED_EXTENSIONS}
        onFileSelected={handleFileSelected}
      />

      {showNameDialog && (
        <SaveWorkbookDialog
          value={tempFileName}
          onChange={setTempFileName}
          onClose={() => setShowNameDialog(false)}
          onConfirm={handleConfirmSave}
          onKeyDown={handleDialogKeyDown}
        />
      )}

      <SpreadsheetShell
        spreadsheetRef={spreadsheetRef}
        onCreated={handleCreated}
        onDataBound={handleDataBound}
        onScheduleAutoSave={scheduleAutoSave}
      />
    </div>
  );
}

export default App;
