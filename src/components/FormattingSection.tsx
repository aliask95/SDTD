import { toast } from "sonner";

const FormattingSection = () => {
  const handleRunFormatHelper = () => {
    toast.info("Format Helper: Office.js integration required. Will detect Bold/Italic/Bold+Italic and add review comments.");
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
    </div>
  );
};

export default FormattingSection;
