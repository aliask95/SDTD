import { useState } from "react";
import { toast } from "sonner";
import { CheckboxOption } from "./LayoutSection";

const FormattingSection = () => {
  const [mepsBookmanWTS, setMepsBookmanWTS] = useState(true);
  const [mepsBookmanUniversal, setMepsBookmanUniversal] = useState(true);

  const handleRunFormatHelper = () => {
    toast.info("Format Helper: Office.js integration required. Will detect Bold/Italic/Bold+Italic and add review comments.");
  };

  const handleRemoveMeps = () => {
    toast.info("Remove MEPS Markup: Office.js integration required. Will remove text with selected MEPS fonts.");
  };

  return (
    <div className="space-y-3">
      <div className="tool-card">
        <div className="tool-title">Format Helper</div>
        <p className="text-xs text-muted-foreground leading-relaxed">
          Detects formatted text (Italic, Bold, Bold+Italic) in the Source document and adds comments in the Target asking if formatting is correct.
        </p>
        <p className="text-xs text-muted-foreground mt-1">
          Comment author: <span className="font-semibold text-foreground">Format Assist</span>
        </p>
        <button onClick={handleRunFormatHelper} className="w-full mt-2 px-3 py-1.5 text-xs font-semibold rounded bg-primary text-primary-foreground hover:opacity-90 transition-opacity">
          Run Format Helper
        </button>
      </div>

      {/* Remove MEPS Markup */}
      <div className="tool-card">
        <div className="tool-title">Remove the MEPS Up</div>
        <p className="text-xs text-muted-foreground leading-relaxed">
          Removes all text with the selected MEPS fonts from the Target document.
        </p>
        <div className="space-y-1.5 mt-2">
          <CheckboxOption label="MEPS Bookman WTS" checked={mepsBookmanWTS} onChange={setMepsBookmanWTS} />
          <CheckboxOption label="MEPS Bookman Universal" checked={mepsBookmanUniversal} onChange={setMepsBookmanUniversal} />
        </div>
        <button onClick={handleRemoveMeps} className="w-full mt-2 px-3 py-1.5 text-xs font-semibold rounded bg-primary text-primary-foreground hover:opacity-90 transition-opacity">
           Remove the MEPS Up
        </button>
      </div>
    </div>
  );
};

export default FormattingSection;
