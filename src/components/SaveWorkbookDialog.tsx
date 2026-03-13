interface SaveWorkbookDialogProps {
  value: string;
  onChange: (value: string) => void;
  onClose: () => void;
  onConfirm: () => void;
  onKeyDown: (event: React.KeyboardEvent<HTMLInputElement>) => void;
}

export function SaveWorkbookDialog({
  value,
  onChange,
  onClose,
  onConfirm,
  onKeyDown,
}: SaveWorkbookDialogProps) {
  return (
    <div className="dialog-overlay" onClick={onClose}>
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
          value={value}
          onChange={(e) => onChange(e.target.value)}
          onKeyDown={onKeyDown}
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
          <button type="button" className="dialog-btn cancel" onClick={onClose}>
            Cancel
          </button>
          <button type="button" className="dialog-btn save" onClick={onConfirm}>
            Save
          </button>
        </div>
      </div>
    </div>
  );
}
