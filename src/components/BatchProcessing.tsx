import { useState } from "react";
import { toast } from "sonner";
import { CheckboxOption } from "./LayoutSection";
import { Files, FolderOpen, X } from "lucide-react";
import { useProgress } from "./ProgressContext";

const BatchProcessing = () => {
  const [enabled, setEnabled] = useState(false);
  const { state, cancel } = useProgress();

  const handleSelectFiles = () => {
    const input = document.createElement("input");
    input.type = "file";
    input.accept = ".docx,.doc";
    input.multiple = true;
    input.onchange = (e) => {
      const files = (e.target as HTMLInputElement).files;
      if (files && files.length > 0) {
        toast.success(`Selected ${files.length} file(s)`);
      }
    };
    input.click();
  };

  const handleSelectFolder = () => {
    const input = document.createElement("input");
    input.type = "file";
    input.setAttribute("webkitdirectory", "");
    input.onchange = (e) => {
      const files = (e.target as HTMLInputElement).files;
      if (files && files.length > 0) {
        toast.success(`Selected folder with ${files.length} file(s)`);
      }
    };
    input.click();
  };

  return (
    <div className="border-t border-border" style={{ background: "hsl(200 12% 14%)" }}>
      {/* Progress bar */}
      {state.running && (
        <div className="px-3 py-2 border-b border-border">
          <div className="flex items-center justify-between gap-2 mb-1">
            <span className="text-[10px] font-medium text-muted-foreground truncate">{state.label}</span>
            <button
              onClick={cancel}
              className="shrink-0 p-0.5 rounded hover:bg-destructive/20 text-muted-foreground hover:text-destructive transition-colors"
              title="Cancel"
            >
              <X className="w-3 h-3" />
            </button>
          </div>
          <div className="w-full h-1.5 rounded-full bg-secondary overflow-hidden">
            <div
              className="h-full rounded-full bg-primary transition-all duration-300 ease-out"
              style={{ width: `${state.progress}%` }}
            />
          </div>
          <span className="text-[9px] text-muted-foreground mt-0.5 block">{Math.round(state.progress)}%</span>
        </div>
      )}

      {/* Batch controls */}
      <div className="px-3 py-2">
        <CheckboxOption label="Enable Batch Mode" checked={enabled} onChange={setEnabled} />
        {enabled && (
          <div className="flex gap-2 mt-2">
            <button
              onClick={handleSelectFiles}
              className="flex-1 flex items-center justify-center gap-1.5 px-2.5 py-1.5 text-xs font-semibold rounded-md bg-secondary border border-input text-secondary-foreground hover:brightness-110 transition-all"
            >
              <Files className="w-3.5 h-3.5" />
              Select Files
            </button>
            <button
              onClick={handleSelectFolder}
              className="flex-1 flex items-center justify-center gap-1.5 px-2.5 py-1.5 text-xs font-semibold rounded-md bg-secondary border border-input text-secondary-foreground hover:brightness-110 transition-all"
            >
              <FolderOpen className="w-3.5 h-3.5" />
              Select Folder
            </button>
          </div>
        )}
      </div>
    </div>
  );
};

export default BatchProcessing;
