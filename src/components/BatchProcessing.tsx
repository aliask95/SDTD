import { useState } from "react";
import { toast } from "sonner";
import { CheckboxOption } from "./LayoutSection";

const BatchProcessing = () => {
  const [enabled, setEnabled] = useState(false);

  return (
    <div className="px-3 py-2.5" style={{ borderTop: "1px solid hsl(200 10% 22%)", background: "hsl(200 12% 14%)" }}>
      <CheckboxOption label="Enable Batch Mode" checked={enabled} onChange={setEnabled} />
      {enabled && (
        <div className="flex gap-2 mt-2">
          <button
            onClick={() => toast.info("Select Files: Office.js file picker required.")}
            className="flex-1 px-2.5 py-1.5 text-xs font-semibold rounded-md transition-all hover:brightness-110"
            style={{ background: "hsl(200 12% 22%)", border: "1px solid hsl(200 10% 28%)", color: "hsl(200 20% 75%)" }}
          >
            Select Files
          </button>
          <button
            onClick={() => toast.info("Select Folder: Office.js file picker required.")}
            className="flex-1 px-2.5 py-1.5 text-xs font-semibold rounded-md transition-all hover:brightness-110"
            style={{ background: "hsl(200 12% 22%)", border: "1px solid hsl(200 10% 28%)", color: "hsl(200 20% 75%)" }}
          >
            Select Folder
          </button>
        </div>
      )}
    </div>
  );
};

export default BatchProcessing;
