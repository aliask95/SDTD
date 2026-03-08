import { useState } from "react";
import { CheckboxOption } from "./LayoutSection";
import { ScanEye, Eraser } from "lucide-react";
import { useSimulatedRun } from "../hooks/useSimulatedRun";

const FormattingSection = () => {
  const [mepsBookmanWTS, setMepsBookmanWTS] = useState(true);
  const [mepsBookmanUniversal, setMepsBookmanUniversal] = useState(true);
  const run = useSimulatedRun();

  return (
    <div className="space-y-3">
      <div className="tool-card">
        <div className="tool-title flex items-center gap-1.5">
          <ScanEye className="w-4 h-4" />
          Format Helper
        </div>
        <p className="text-xs text-muted-foreground leading-relaxed">
          Detects formatted text (Italic, Bold, Bold+Italic) in the Source document and adds comments in the Target asking if formatting is correct.
        </p>
        <p className="text-xs text-muted-foreground mt-1">
          Comment author: <span className="font-semibold text-foreground">Format Assist</span>
        </p>
        <button onClick={() => run("Format Helper")} className="w-full mt-2 px-3 py-1.5 text-xs font-semibold rounded bg-primary text-primary-foreground hover:opacity-90 transition-opacity">
          Run Format Helper
        </button>
      </div>

      <div className="tool-card">
        <div className="tool-title flex items-center gap-1.5">
          <Eraser className="w-4 h-4" />
          Clean the MEPS Up
        </div>
        <p className="text-xs text-muted-foreground leading-relaxed">
          Removes all text with the selected MEPS fonts from the Target document.
        </p>
        <div className="space-y-1.5 mt-2">
          <CheckboxOption label="MEPS Bookman WTS" checked={mepsBookmanWTS} onChange={setMepsBookmanWTS} />
          <CheckboxOption label="MEPS Bookman Universal" checked={mepsBookmanUniversal} onChange={setMepsBookmanUniversal} />
        </div>
        <button onClick={() => run("Clean the MEPS Up")} className="w-full mt-2 px-3 py-1.5 text-xs font-semibold rounded bg-primary text-primary-foreground hover:opacity-90 transition-opacity">
           Clean the MEPS Up
        </button>
      </div>
    </div>
  );
};

export default FormattingSection;
