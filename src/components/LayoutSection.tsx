import { useState } from "react";
import { Copy, PaintBucket } from "lucide-react";
import { useSimulatedRun } from "../hooks/useSimulatedRun";
import InfoTooltip from "./InfoTooltip";

const LayoutSection = () => {
  const [pageSize, setPageSize] = useState(true);
  const [header, setHeader] = useState(true);
  const [footer, setFooter] = useState(true);
  const [copyHeaderFooterContent, setCopyHeaderFooterContent] = useState(false);
  const [deleteEmptyParas, setDeleteEmptyParas] = useState(false);
  const run = useSimulatedRun();

  return (
    <div className="space-y-3">
      <div className="tool-card">
        <div className="tool-title flex items-center gap-1.5">
          <Copy className="w-[18px] h-[18px]" />
          <span>Copy Page Layout</span>
          <InfoTooltip text="Copies page size, headers, footers, and their content from the Source document to the Target document." />
        </div>
        <div className="space-y-1.5">
          <CheckboxOption label="Page Size" checked={pageSize} onChange={setPageSize} />
          <CheckboxOption label="Header" checked={header} onChange={setHeader} />
          <CheckboxOption label="Footer" checked={footer} onChange={setFooter} />
          <CheckboxOption label="Copy Header/Footer Content" checked={copyHeaderFooterContent} onChange={setCopyHeaderFooterContent} />
          {copyHeaderFooterContent && (
            <div className="ml-5">
              <CheckboxOption label="Delete empty paragraphs in header/footer" checked={deleteEmptyParas} onChange={setDeleteEmptyParas} />
            </div>
          )}
        </div>
        <button onClick={() => run("Apply Layout")} className="w-full mt-2 px-3 py-1.5 text-xs font-semibold rounded bg-primary text-primary-foreground hover:opacity-90 transition-opacity">
          Apply Layout
        </button>
      </div>

      <div className="tool-card">
        <div className="tool-title flex items-center gap-1.5">
          <PaintBucket className="w-[18px] h-[18px]" />
          <span>Copy Styles</span>
          <InfoTooltip text="Copy paragraph styles from Source to Target document." />
        </div>
        <button onClick={() => run("Copy Styles")} className="w-full mt-2 px-3 py-1.5 text-xs font-semibold rounded bg-primary text-primary-foreground hover:opacity-90 transition-opacity">
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
