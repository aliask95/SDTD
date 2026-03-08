import { useState } from "react";
import CollapsibleSection from "../components/CollapsibleSection";
import LayoutSection from "../components/LayoutSection";
import FormattingSection from "../components/FormattingSection";
import ADPSection from "../components/ADPSection";
import BatchProcessing from "../components/BatchProcessing";
import { toast } from "sonner";
import sdtdLogo from "@/assets/sdtd-logo.png";

/**
 * Design B: Bold & Dark
 * - Dark header with large logo, sharp contrast
 * - Teal accents on dark surface
 */

const MOCK_DOCS = ["Document1.docx", "Report_EN.docx", "Translation_RU.docx"];

const DesignB = () => {
  const [sourceDoc, setSourceDoc] = useState("");
  const [targetDoc, setTargetDoc] = useState("");

  const handleCheckParagraphCount = () => {
    toast.info("Check Paragraph Count: Office.js integration required.");
  };

  return (
    <div className="min-h-screen flex flex-col items-center justify-center p-4" style={{ background: "hsl(200 15% 12%)" }}>
      <div className="mb-3 px-3 py-1 rounded-full text-xs font-bold tracking-widest uppercase" style={{ background: "hsl(190 79% 39% / 0.2)", color: "hsl(190 79% 60%)" }}>
        Design B — Bold & Dark
      </div>
      <div className="w-[360px] max-h-[80vh] overflow-y-auto rounded-xl shadow-2xl flex flex-col" style={{ background: "hsl(200 12% 16%)", border: "1px solid hsl(200 10% 22%)" }}>
        {/* Header - tall with centered logo */}
        <div className="flex items-center gap-3 px-4 py-4 rounded-t-xl" style={{ background: "linear-gradient(135deg, hsl(190 79% 35%), hsl(190 79% 28%))", borderBottom: "2px solid hsl(190 79% 45%)" }}>
          <img src={sdtdLogo} alt="SDTD Logo" className="w-10 h-10 rounded-lg shadow-lg" style={{ border: "2px solid rgba(255,255,255,0.25)" }} />
          <div className="flex flex-col">
            <span className="text-base font-bold tracking-wider leading-tight text-white">SDTD</span>
            <span className="text-[10px] text-white/60 leading-tight font-medium">Source → Target Document</span>
          </div>
        </div>

        {/* Document Selectors */}
        <div className="px-3 py-3 space-y-2" style={{ borderBottom: "1px solid hsl(200 10% 22%)" }}>
          <DocSelect label="Source Document" value={sourceDoc} onChange={setSourceDoc} docs={MOCK_DOCS} />
          <DocSelect label="Target Document" value={targetDoc} onChange={setTargetDoc} docs={MOCK_DOCS} />
          <button
            onClick={handleCheckParagraphCount}
            className="w-full px-3 py-2 text-xs font-bold rounded-md transition-all uppercase tracking-wider hover:brightness-110"
            style={{ background: "hsl(190 79% 39%)", color: "white", boxShadow: "0 2px 8px hsl(190 79% 39% / 0.3)" }}
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

export default DesignB;

const DocSelect = ({ label, value, onChange, docs }: { label: string; value: string; onChange: (v: string) => void; docs: string[] }) => (
  <div>
    <label className="text-[10px] font-bold uppercase tracking-wider" style={{ color: "hsl(190 60% 60%)" }}>{label}</label>
    <select value={value} onChange={(e) => onChange(e.target.value)} className="w-full mt-0.5 px-2 py-1.5 text-xs rounded" style={{ background: "hsl(200 12% 20%)", border: "1px solid hsl(200 10% 28%)", color: "hsl(200 20% 80%)" }}>
      <option value="">— Select —</option>
      {docs.map((d) => <option key={d} value={d}>{d}</option>)}
    </select>
  </div>
);
