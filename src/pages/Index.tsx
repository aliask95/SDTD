import { useState } from "react";
import CollapsibleSection from "../components/CollapsibleSection";
import LayoutSection from "../components/LayoutSection";
import FormattingSection from "../components/FormattingSection";
import ADPSection from "../components/ADPSection";
import BatchProcessing from "../components/BatchProcessing";
import { FileText } from "lucide-react";

const MOCK_DOCS = ["Document1.docx", "Report_EN.docx", "Translation_RU.docx"];

const Index = () => {
  const [sourceDoc, setSourceDoc] = useState("");
  const [targetDoc, setTargetDoc] = useState("");

  return (
    <div className="min-h-screen flex items-center justify-center bg-muted/50 p-4">
      <div className="w-[360px] max-h-[90vh] overflow-y-auto rounded-lg shadow-lg border border-border bg-background flex flex-col">
        {/* Header */}
        <div className="flex items-center gap-2 px-3 py-2.5 border-b border-border bg-primary text-primary-foreground">
          <FileText className="w-4 h-4" />
          <span className="text-sm font-bold tracking-wide">SDTD</span>
          <span className="text-xs opacity-80 ml-1">Source → Target</span>
        </div>

        {/* Document Selectors */}
        <div className="px-3 py-2.5 space-y-2 border-b border-border">
          <DocSelect label="Source Document" value={sourceDoc} onChange={setSourceDoc} docs={MOCK_DOCS} />
          <DocSelect label="Target Document" value={targetDoc} onChange={setTargetDoc} docs={MOCK_DOCS} />
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

        {/* Batch Processing */}
        <BatchProcessing />
      </div>
    </div>
  );
};

const DocSelect = ({ label, value, onChange, docs }: { label: string; value: string; onChange: (v: string) => void; docs: string[] }) => (
  <div>
    <label className="text-xs font-semibold text-muted-foreground">{label}</label>
    <select
      value={value}
      onChange={(e) => onChange(e.target.value)}
      className="w-full mt-0.5 px-2 py-1.5 text-xs rounded border border-input bg-background text-foreground"
    >
      <option value="">— Select —</option>
      {docs.map((d) => (
        <option key={d} value={d}>{d}</option>
      ))}
    </select>
  </div>
);

export default Index;
