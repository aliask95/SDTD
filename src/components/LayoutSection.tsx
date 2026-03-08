import { useState } from "react";
import { toast } from "sonner";

const LayoutSection = () => {
  const [pageSize, setPageSize] = useState(true);
  const [header, setHeader] = useState(true);
  const [footer, setFooter] = useState(true);
  const [copyHeaderFooterContent, setCopyHeaderFooterContent] = useState(false);


  const handleApplyLayout = () => {
    toast.info("Apply Layout: Office.js integration required. This will copy page layout from Source to Target.");
  };

  const handleCopyStyles = () => {
    toast.info("Copy Styles: Office.js integration required. This will copy styles from Source to Target.");
  };

  return (
    <div className="space-y-3">
      {/* Copy Page Layout */}
      <div className="tool-card">
        <div className="tool-title">Copy Page Layout</div>
        <div className="space-y-1.5">
          <CheckboxOption label="Page Size" checked={pageSize} onChange={setPageSize} />
          <CheckboxOption label="Header" checked={header} onChange={setHeader} />
          <CheckboxOption label="Footer" checked={footer} onChange={setFooter} />
          <CheckboxOption label="Copy Header/Footer Content" checked={copyHeaderFooterContent} onChange={setCopyHeaderFooterContent} />
        </div>
        <button onClick={handleApplyLayout} className="w-full mt-2 px-3 py-1.5 text-xs font-semibold rounded bg-primary text-primary-foreground hover:opacity-90 transition-opacity">
          Apply Layout
        </button>
      </div>

      {/* Copy Styles */}
      <div className="tool-card">
        <div className="tool-title">Copy Styles</div>
        <p className="text-xs text-muted-foreground leading-relaxed">
          Copy paragraph styles from Source to Target document.
        </p>
        <button onClick={handleCopyStyles} className="w-full mt-2 px-3 py-1.5 text-xs font-semibold rounded bg-primary text-primary-foreground hover:opacity-90 transition-opacity">
          Copy Styles
        </button>
      </div>
    </div>
  );
};

export const CheckboxOption = ({ label, checked, onChange }: { label: string; checked: boolean; onChange: (v: boolean) => void }) => (
  <label className="flex items-center gap-2 text-xs cursor-pointer">
    <input
      type="checkbox"
      checked={checked}
      onChange={(e) => onChange(e.target.checked)}
      className="rounded border-input accent-primary w-3.5 h-3.5"
    />
    <span className="text-foreground">{label}</span>
  </label>
);

export default LayoutSection;
