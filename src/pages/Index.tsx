import { useState, useCallback, DragEvent } from "react";
import CollapsibleSection from "../components/CollapsibleSection";
import LayoutSection from "../components/LayoutSection";
import FormattingSection from "../components/FormattingSection";
import ADPSection from "../components/ADPSection";
import BatchProcessing from "../components/BatchProcessing";
import { toast } from "sonner";
import { FolderOpen } from "lucide-react";
import sdtdLogo from "@/assets/sdtd-logo.png";

const MOCK_DOCS = ["Document1.docx", "Report_EN.docx", "Translation_RU.docx"];

const Index = () => {
  const [sourceDoc, setSourceDoc] = useState("");
  const [targetDoc, setTargetDoc] = useState("");
  const [openSection, setOpenSection] = useState<string | null>("1");
  const [sourceDrag, setSourceDrag] = useState(false);
  const [targetDrag, setTargetDrag] = useState(false);

  const handleCheckParagraphCount = () => {
    toast.info("Check Paragraph Count: Office.js integration required.");
  };

  const handleToggle = (id: string) => {
    setOpenSection((prev) => (prev === id ? null : id));
  };

  const handleFilePick = (setter: (v: string) => void) => {
    const input = document.createElement("input");
    input.type = "file";
    input.accept = ".docx,.doc";
    input.onchange = (e) => {
      const file = (e.target as HTMLInputElement).files?.[0];
      if (file) {
        setter(file.name);
        toast.success(`Selected: ${file.name}`);
      }
    };
    input.click();
  };

  const handleDrop = useCallback(
    (setter: (v: string) => void, setDrag: (v: boolean) => void) =>
      (e: DragEvent) => {
        e.preventDefault();
        setDrag(false);
        const file = e.dataTransfer.files[0];
        if (file) {
          if (!file.name.match(/\.docx?$/i)) {
            toast.error("Please drop a .doc or .docx file.");
            return;
          }
          setter(file.name);
          toast.success(`Dropped: ${file.name}`);
        }
      },
    []
  );

  const handleDragOver = (setDrag: (v: boolean) => void) => (e: DragEvent) => {
    e.preventDefault();
    setDrag(true);
  };

  const handleDragLeave = (setDrag: (v: boolean) => void) => () => {
    setDrag(false);
  };

  return (
    <div className="min-h-screen flex items-center justify-center p-2 sm:p-4 bg-background">
      <div className="w-full max-w-[400px] max-h-[95vh] overflow-y-auto rounded-xl shadow-2xl flex flex-col border border-border bg-background">
        {/* Header */}
        <div className="flex items-center gap-3 px-4 py-3 rounded-t-xl" style={{ background: "linear-gradient(135deg, hsl(190 79% 35%), hsl(190 79% 28%))" }}>
          <img src={sdtdLogo} alt="SDTD Logo" className="w-9 h-9 rounded-lg shadow-md" />
          <div className="flex flex-col">
            <span className="text-sm font-bold tracking-wider leading-tight text-primary-foreground">SDTD</span>
            <span className="text-[10px] text-primary-foreground/60 leading-tight font-medium">Source → Target Document</span>
          </div>
        </div>

        {/* Document Selectors */}
        <div className="px-3 py-3 space-y-2 border-b border-border">
          <DocSelect
            label="Source Document"
            value={sourceDoc}
            onChange={setSourceDoc}
            docs={MOCK_DOCS}
            onBrowse={() => handleFilePick(setSourceDoc)}
            isDragging={sourceDrag}
            onDrop={handleDrop(setSourceDoc, setSourceDrag)}
            onDragOver={handleDragOver(setSourceDrag)}
            onDragLeave={handleDragLeave(setSourceDrag)}
          />
          <DocSelect
            label="Target Document"
            value={targetDoc}
            onChange={setTargetDoc}
            docs={MOCK_DOCS}
            onBrowse={() => handleFilePick(setTargetDoc)}
            isDragging={targetDrag}
            onDrop={handleDrop(setTargetDoc, setTargetDrag)}
            onDragOver={handleDragOver(setTargetDrag)}
            onDragLeave={handleDragLeave(setTargetDrag)}
          />
          <button
            onClick={handleCheckParagraphCount}
            className="w-full px-3 py-2 text-xs font-bold rounded-md transition-all uppercase tracking-wider hover:brightness-110 bg-primary text-primary-foreground"
            style={{ boxShadow: "0 2px 8px hsl(190 79% 39% / 0.3)" }}
          >
            Check Paragraph Count
          </button>
        </div>

        {/* Sections */}
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

interface DocSelectProps {
  label: string;
  value: string;
  onChange: (v: string) => void;
  docs: string[];
  onBrowse: () => void;
  isDragging: boolean;
  onDrop: (e: DragEvent<HTMLDivElement>) => void;
  onDragOver: (e: DragEvent<HTMLDivElement>) => void;
  onDragLeave: () => void;
}

const DocSelect = ({ label, value, onChange, docs, onBrowse, isDragging, onDrop, onDragOver, onDragLeave }: DocSelectProps) => (
  <div
    className={`rounded-md p-2 transition-all ${isDragging ? "ring-2 ring-primary bg-accent" : ""}`}
    onDrop={onDrop as any}
    onDragOver={onDragOver as any}
    onDragLeave={onDragLeave}
  >
    <label className="text-[10px] font-bold uppercase tracking-wider text-accent-foreground">{label}</label>
    <div className="flex items-center gap-1.5 mt-0.5">
      <select
        value={value}
        onChange={(e) => onChange(e.target.value)}
        className="flex-1 min-w-0 px-2 py-1.5 text-xs rounded bg-secondary border border-input text-secondary-foreground"
      >
        <option value="">— Select or drop file —</option>
        {docs.map((d) => (
          <option key={d} value={d}>{d}</option>
        ))}
        {value && !docs.includes(value) && <option value={value}>{value}</option>}
      </select>
      <button
        onClick={onBrowse}
        className="shrink-0 p-1.5 rounded-md bg-secondary border border-input text-secondary-foreground hover:bg-muted transition-colors"
        title="Browse for file"
      >
        <FolderOpen className="w-3.5 h-3.5" />
      </button>
    </div>
    {isDragging && (
      <p className="text-[10px] text-accent-foreground mt-1 text-center animate-pulse">Drop .docx file here</p>
    )}
  </div>
);
