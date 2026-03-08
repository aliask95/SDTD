import { useState } from "react";
import CollapsibleSection from "../components/CollapsibleSection";
import LayoutSection from "../components/LayoutSection";
import FormattingSection from "../components/FormattingSection";
import ADPSection from "../components/ADPSection";
import BatchProcessing from "../components/BatchProcessing";
import { toast } from "sonner";
import sdtdLogo from "@/assets/sdtd-logo.png";

/**
 * Design A: Clean & Bright
 * - White logo container with rounded shape on teal header
 * - Lighter overall feel, subtle shadows
 */

const MOCK_DOCS = ["Document1.docx", "Report_EN.docx", "Translation_RU.docx"];

const DesignA = () => {
  const [sourceDoc, setSourceDoc] = useState("");
  const [targetDoc, setTargetDoc] = useState("");

  const handleCheckParagraphCount = () => {
    toast.info("Check Paragraph Count: Office.js integration required.");
  };

  return (
    <div className="min-h-screen flex flex-col items-center justify-center p-4" style={{ background: "linear-gradient(135deg, hsl(190 15% 95%), hsl(190 20% 90%))" }}>
      <div className="mb-3 px-3 py-1 rounded-full text-xs font-bold tracking-widest uppercase" style={{ background: "hsl(190 79% 39% / 0.1)", color: "hsl(190 79% 32%)" }}>
        Design A — Clean & Bright
      </div>
      <div className="w-[360px] max-h-[80vh] overflow-y-auto rounded-xl shadow-xl border border-border bg-background flex flex-col">
        {/* Header */}
        <div className="flex items-center gap-3 px-4 py-3 rounded-t-xl" style={{ background: "hsl(190 79% 39%)" }}>
          <div className="w-9 h-9 rounded-lg bg-white/20 backdrop-blur-sm flex items-center justify-center p-1">
            <img src={sdtdLogo} alt="SDTD Logo" className="w-7 h-7 rounded" />
          </div>
          <div className="flex flex-col">
            <span className="text-sm font-bold tracking-wide leading-tight text-white">SDTD</span>
            <span className="text-[10px] text-white/70 leading-tight">Source → Target Document</span>
          </div>
        </div>

        {/* Document Selectors */}
        <div className="px-3 py-2.5 space-y-2 border-b border-border">
          <DocSelect label="Source Document" value={sourceDoc} onChange={setSourceDoc} docs={MOCK_DOCS} />
          <DocSelect label="Target Document" value={targetDoc} onChange={setTargetDoc} docs={MOCK_DOCS} />
          <button
            onClick={handleCheckParagraphCount}
            className="w-full px-3 py-1.5 text-xs font-semibold rounded-md transition-all hover:shadow-md"
            style={{ background: "hsl(190 79% 39%)", color: "white" }}
          >
            Check Paragraph Count
          </button>
        </div>

        {/* Sections */}
        <div className="flex-1 overflow-y-auto">
          <CollapsibleSection title="1. Layout" defaultOpen>
            <LayoutSection />
          </CollapsibleSection>
          <CollapsibleSection title="2. Formatting">
            <FormattingSection />
          </CollapsibleSection>
          <CollapsibleSection title="3. ADP (Advanced Document Processing)">
            <ADPSection />
          </CollapsibleSection>
        </div>

        <BatchProcessing />
      </div>
    </div>
  );
};

export default DesignA;

const DocSelect = ({ label, value, onChange, docs }: { label: string; value: string; onChange: (v: string) => void; docs: string[] }) => (
  <div>
    <label className="text-xs font-semibold text-muted-foreground">{label}</label>
    <select value={value} onChange={(e) => onChange(e.target.value)} className="w-full mt-0.5 px-2 py-1.5 text-xs rounded border border-input bg-background text-foreground">
      <option value="">— Select —</option>
      {docs.map((d) => <option key={d} value={d}>{d}</option>)}
    </select>
  </div>
);
