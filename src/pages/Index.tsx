import { useState } from "react";
import CollapsibleSection from "../components/CollapsibleSection";
import LayoutSection from "../components/LayoutSection";
import FormattingSection from "../components/FormattingSection";
import ADPSection from "../components/ADPSection";
import BatchProcessing from "../components/BatchProcessing";
import { toast } from "sonner";
import sdtdLogo from "@/assets/sdtd-logo.png";

const MOCK_DOCS = ["Document1.docx", "Report_EN.docx", "Translation_RU.docx"];

const Index = () => {
  const [sourceDoc, setSourceDoc] = useState("");
  const [targetDoc, setTargetDoc] = useState("");
  const [openSection, setOpenSection] = useState<string | null>("1");

  const handleCheckParagraphCount = () => {
    toast.info("Check Paragraph Count: Office.js integration required.");
  };

  const handleToggle = (id: string) => {
    setOpenSection((prev) => (prev === id ? null : id));
  };

  return (
    <div className="min-h-screen flex items-center justify-center p-4" style={{ background: "hsl(200 15% 12%)" }}>
      <div className="w-[360px] max-h-[90vh] overflow-y-auto rounded-xl shadow-2xl flex flex-col" style={{ background: "hsl(200 12% 16%)", border: "1px solid hsl(200 10% 22%)" }}>
        {/* Header — seamless logo integration */}
        <div className="flex items-center gap-3 px-4 py-3 rounded-t-xl" style={{ background: "linear-gradient(135deg, hsl(190 79% 35%), hsl(190 79% 28%))" }}>
          <img src={sdtdLogo} alt="SDTD Logo" className="w-10 h-10 rounded-lg" style={{ boxShadow: "0 2px 10px rgba(0,0,0,0.3)" }} />
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

        {/* Sections — accordion style */}
        <div className="flex-1 overflow-y-auto">
          <CollapsibleSection title="1. Layout" isOpen={openSection === "1"} onToggle={() => handleToggle("1")}>
            <LayoutSection />
          </CollapsibleSection>
          <CollapsibleSection title="2. Formatting" isOpen={openSection === "2"} onToggle={() => handleToggle("2")}>
            <FormattingSection />
          </CollapsibleSection>
          <CollapsibleSection title="3. ADP (Advanced Document Processing)" isOpen={openSection === "3"} onToggle={() => handleToggle("3")}>
            <ADPSection />
          </CollapsibleSection>
        </div>

        <BatchProcessing />
      </div>
    </div>
  );
};

export default Index;

const DocSelect = ({ label, value, onChange, docs }: { label: string; value: string; onChange: (v: string) => void; docs: string[] }) => (
  <div>
    <label className="text-[10px] font-bold uppercase tracking-wider" style={{ color: "hsl(190 60% 60%)" }}>{label}</label>
    <select value={value} onChange={(e) => onChange(e.target.value)} className="w-full mt-0.5 px-2 py-1.5 text-xs rounded" style={{ background: "hsl(200 12% 20%)", border: "1px solid hsl(200 10% 28%)", color: "hsl(200 20% 80%)" }}>
      <option value="">— Select —</option>
      {docs.map((d) => <option key={d} value={d}>{d}</option>)}
    </select>
  </div>
);
