import { useState } from "react";
import { toast } from "sonner";
import { Copy, PaintBucket } from "lucide-react";

const LayoutSection = () => {
  const [pageSize, setPageSize] = useState(true);
  const [header, setHeader] = useState(true);
  const [footer, setFooter] = useState(true);
  const [copyHeaderFooterContent, setCopyHeaderFooterContent] = useState(false);
  const [deleteEmptyParas, setDeleteEmptyParas] = useState(false);

  const handleApplyLayout = () => {
    toast.info("Apply Layout: Office.js integration required. This will copy page layout from Source to Target.");
  };

  const handleCopyStyles = () => {
    toast.info("Copy Styles: Office.js integration required. This will copy styles from Source to Target.");
  };

  const handleDeleteEmptyParagraphs = () => {
    toast.info("Delete Empty Paragraphs: Office.js integration required. Will remove empty paragraphs in header and footer.");
  };

  return (
    <div className="space-y-3">
      <div className="tool-card">
        <div className="tool-title flex items-center gap-1.5">
          <Copy className="w-4 h-4" />
          Copy Page Layout
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
        <button onClick={handleApplyLayout} className="w-full mt-2 px-3 py-1.5 text-xs font-semibold rounded bg-primary text-primary-foreground hover:opacity-90 transition-opacity">
          Apply Layout
        </button>
      </div>

      <div className="tool-card">
        <div className="tool-title flex items-center gap-1.5">
          <PaintBucket className="w-4 h-4" />
          Copy Styles
        </div>
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
