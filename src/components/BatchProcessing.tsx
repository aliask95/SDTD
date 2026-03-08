import { useState } from "react";
import { toast } from "sonner";
import { CheckboxOption } from "./LayoutSection";

const BatchProcessing = () => {
  const [enabled, setEnabled] = useState(false);

  return (
    <div className="px-3 py-2.5 border-t border-border" style={{ background: "hsl(var(--section-bg))" }}>
      <CheckboxOption label="Enable Batch Mode" checked={enabled} onChange={setEnabled} />
      {enabled && (
        <div className="flex gap-2 mt-2">
          <button
            onClick={() => toast.info("Select Files: Office.js file picker required.")}
            className="flex-1 px-2.5 py-1.5 text-xs font-semibold rounded border border-input bg-secondary text-secondary-foreground hover:opacity-80 transition-opacity"
          >
            Select Files
          </button>
          <button
            onClick={() => toast.info("Select Folder: Office.js file picker required.")}
            className="flex-1 px-2.5 py-1.5 text-xs font-semibold rounded border border-input bg-secondary text-secondary-foreground hover:opacity-80 transition-opacity"
          >
            Select Folder
          </button>
        </div>
      )}
    </div>
  );
};

export default BatchProcessing;
