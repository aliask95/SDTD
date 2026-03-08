import { useState } from "react";
import { toast } from "sonner";
import { CheckboxOption } from "./LayoutSection";
import { Files, FolderOpen } from "lucide-react";

const BatchProcessing = () => {
  const [enabled, setEnabled] = useState(false);

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
    <div className="px-3 py-2 border-t border-border" style={{ background: "hsl(200 12% 14%)" }}>
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
  );
};

export default BatchProcessing;
