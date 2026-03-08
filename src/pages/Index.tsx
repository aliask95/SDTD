import { useState, useCallback, DragEvent } from "react";
import { ProgressProvider } from "../components/ProgressContext";
import LayoutSection from "../components/LayoutSection";
import FormattingSection from "../components/FormattingSection";
import ADPSection from "../components/ADPSection";
import BatchProcessing from "../components/BatchProcessing";
import { toast } from "sonner";
import { FolderOpen, LayoutTemplate, Type, Cpu } from "lucide-react";
import sdtdLogo from "@/assets/sdtd-logo.png";

const MOCK_DOCS = ["Document1.docx", "Report_EN.docx", "Translation_RU.docx"];

const TABS = [
  { id: "layout", label: "Layout", icon: LayoutTemplate },
  { id: "format", label: "Format", icon: Type },
  { id: "adp", label: "ADP", icon: Cpu },
] as const;

type TabId = (typeof TABS)[number]["id"];

const Index = () => {
  const [sourceDoc, setSourceDoc] = useState("");
  const [targetDoc, setTargetDoc] = useState("");
  const [activeTab, setActiveTab] = useState<TabId>("layout");
  const [sourceDrag, setSourceDrag] = useState(false);
  const [targetDrag, setTargetDrag] = useState(false);

  const handleCheckParagraphCount = () => {
    toast.info("Check Paragraph Count: Office.js integration required.");
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
    <ProgressProvider>
    <div className="min-h-screen flex items-center justify-center p-2 sm:p-4 bg-background">
      <div className="w-full max-w-[400px] max-h-[95vh] overflow-y-auto rounded-xl shadow-2xl flex flex-col border border-border bg-background">
        {/* Header */}
        <div className="flex items-center gap-3 px-4 py-2.5 rounded-t-xl" style={{ background: "linear-gradient(135deg, hsl(190 79% 35%), hsl(190 79% 28%))" }}>
          <img src={sdtdLogo} alt="SDTD Logo" className="w-8 h-8 rounded-lg shadow-md" />
          <div className="flex flex-col">
            <span className="text-sm font-bold tracking-wider leading-tight text-primary-foreground">SDTD</span>
            <span className="text-[10px] text-primary-foreground/60 leading-tight font-medium">Source → Target Document</span>
          </div>
        </div>

        {/* Document Selectors — compact */}
        <div className="px-3 py-2 space-y-1.5 border-b border-border">
          <DocSelect
            label="Source"
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
            label="Target"
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
            className="w-full px-3 py-1.5 text-[11px] font-bold rounded-md transition-all uppercase tracking-wider hover:brightness-110 bg-primary text-primary-foreground"
            style={{ boxShadow: "0 2px 8px hsl(190 79% 39% / 0.3)" }}
          >
            Check Paragraph Count
          </button>
        </div>

        {/* Tab Bar */}
        <div className="flex border-b border-border bg-card">
          {TABS.map(({ id, label, icon: Icon }) => (
            <button
              key={id}
              onClick={() => setActiveTab(id)}
              className={`flex-1 flex items-center justify-center gap-1.5 px-2 py-2.5 text-[11px] font-semibold transition-all relative ${
                activeTab === id
                  ? "text-accent-foreground"
                  : "text-muted-foreground hover:text-foreground"
              }`}
            >
              <Icon className="w-4 h-4" />
              <span>{label}</span>
              {activeTab === id && (
                <span className="absolute bottom-0 left-2 right-2 h-[2px] rounded-full bg-primary" />
              )}
            </button>
          ))}
        </div>

        {/* Tab Content */}
        <div className="flex-1 overflow-y-auto px-3 py-2.5" style={{ background: "hsl(200 12% 14%)" }}>
          {activeTab === "layout" && <LayoutSection />}
          {activeTab === "format" && <FormattingSection />}
          {activeTab === "adp" && <ADPSection />}
        </div>

        <BatchProcessing />
      </div>
    </div>
    </ProgressProvider>
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
    className={`rounded p-1.5 transition-all ${isDragging ? "ring-2 ring-primary bg-accent" : ""}`}
    onDrop={onDrop as any}
    onDragOver={onDragOver as any}
    onDragLeave={onDragLeave}
  >
    <div className="flex items-center gap-1.5">
      <label className="text-[10px] font-bold uppercase tracking-wider text-accent-foreground w-12 shrink-0">{label}</label>
      <select
        value={value}
        onChange={(e) => onChange(e.target.value)}
        className="flex-1 min-w-0 px-2 py-1 text-xs rounded bg-secondary border border-input text-secondary-foreground"
      >
        <option value="">— Select or drop —</option>
        {docs.map((d) => (
          <option key={d} value={d}>{d}</option>
        ))}
        {value && !docs.includes(value) && <option value={value}>{value}</option>}
      </select>
      <button
        onClick={onBrowse}
        className="shrink-0 p-1 rounded bg-secondary border border-input text-secondary-foreground hover:bg-muted transition-colors"
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
