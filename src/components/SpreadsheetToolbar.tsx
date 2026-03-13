interface SpreadsheetToolbarProps {
  fileName: string;
  status: string;
  onSaveToLocal: () => void;
  onLoadExcelFile: () => void;
  onDownloadXlsx: () => void;
  fileInputRef: React.RefObject<HTMLInputElement | null>;
  acceptedExtensions: string;
  onFileSelected: (event: React.ChangeEvent<HTMLInputElement>) => void;
}

export function SpreadsheetToolbar({
  fileName,
  status,
  onSaveToLocal,
  onLoadExcelFile,
  onDownloadXlsx,
  fileInputRef,
  acceptedExtensions,
  onFileSelected,
}: SpreadsheetToolbarProps) {
  return (
    <div className="toolbar">
      <button type="button" className="toolbar-btn" onClick={onSaveToLocal}>
        Save to Local
      </button>
      <button type="button" className="toolbar-btn" onClick={onLoadExcelFile}>
        Load Excel File
      </button>
      <button type="button" className="toolbar-btn primary" onClick={onDownloadXlsx}>
        Download .xlsx
      </button>

      <span className="toolbar-filename">
        File: <strong>{fileName}</strong>
      </span>
      {status && <span className="toolbar-status">{status}</span>}

      <input
        ref={fileInputRef}
        type="file"
        accept={acceptedExtensions}
        style={{ display: "none" }}
        onChange={onFileSelected}
      />
    </div>
  );
}
