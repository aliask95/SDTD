import { useState } from "react";
import { CheckboxOption } from "./LayoutSection";
import { ScanEye, Eraser } from "lucide-react";
import { useSimulatedRun } from "../hooks/useSimulatedRun";
import InfoTooltip from "./InfoTooltip";

const FormattingSection = () => {
  const [mepsBookmanWTS, setMepsBookmanWTS] = useState(true);
  const [mepsBookmanUniversal, setMepsBookmanUniversal] = useState(true);
  const run = useSimulatedRun();

  return (
    <div className="space-y-3">
      <div className="tool-card">
        <div className="tool-title flex items-center gap-1.5">
          <ScanEye className="w-[18px] h-[18px]" />
          <span>Format Helper</span>
          <InfoTooltip text="Detects formatted text (Italic, Bold, Bold+Italic) in the Source and adds review comments in the Target. Comment author: Format Assist." />
        </div>
        <button onClick={() => run("Format Helper")} className="w-full mt-2 px-3 py-1.5 text-xs font-semibold rounded bg-primary text-primary-foreground hover:opacity-90 transition-opacity">
          Run Format Helper
        </button>
      </div>

      <div className="tool-card">
        <div className="tool-title flex items-center gap-1.5">
          <Eraser className="w-[18px] h-[18px]" />
          <span>Clean the MEPS Up</span>
          <InfoTooltip text="Removes all text with the selected MEPS fonts from the Target document." />
        </div>
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
