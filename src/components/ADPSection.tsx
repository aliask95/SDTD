import { useState } from "react";
import { toast } from "sonner";
import { CheckboxOption } from "./LayoutSection";

const ADPSection = () => {
  const [findId, setFindId] = useState("=E");
  const [replaceWith, setReplaceWith] = useState("=BL");
  const [excludeIt, setExcludeIt] = useState(false);
  const [excelFile, setExcelFile] = useState<string | null>(
    localStorage.getItem("sdtd_excel_file")
  );

  // Text Cleanup
  const [replacePar, setReplacePar] = useState(true);
  const [replaceSpaces, setReplaceSpaces] = useState(true);

  // Metadata
  const [cleanFirst, setCleanFirst] = useState(true);
  const [sentenceCase, setSentenceCase] = useState(true);
  const [setTitle, setSetTitle] = useState(true);

  const stub = (msg: string) => () => toast.info(msg + ": Office.js integration required.");

  const handleSelectExcel = () => {
    const name = "dictionary.xlsx";
    setExcelFile(name);
    localStorage.setItem("sdtd_excel_file", name);
    toast.success("Excel file selected (simulated): " + name);
  };

  return (
    <div className="space-y-3">
      {/* Transfer DP Comments */}
      <div className="tool-card">
        <div className="tool-title">Transfer DP Comments</div>
        <p className="text-xs text-muted-foreground leading-relaxed">
          Copies comments from author "Digital Publishing" in Source to corresponding paragraphs in Target.
        </p>
        <div className="grid grid-cols-2 gap-2 mt-2">
          <div>
            <label className="text-xs text-muted-foreground">Find identifier</label>
            <input value={findId} onChange={(e) => setFindId(e.target.value)} className="w-full mt-0.5 px-2 py-1 text-xs rounded border border-input bg-background text-foreground" />
          </div>
          <div>
            <label className="text-xs text-muted-foreground">Replace with</label>
            <input value={replaceWith} onChange={(e) => setReplaceWith(e.target.value)} className="w-full mt-0.5 px-2 py-1 text-xs rounded border border-input bg-background text-foreground" />
          </div>
        </div>
        <div className="mt-2">
          <CheckboxOption label='Exclude comments containing "it"' checked={excludeIt} onChange={setExcludeIt} />
        </div>
        <button onClick={stub("Transfer Comments")} className="w-full mt-2 px-3 py-1.5 text-xs font-semibold rounded bg-primary text-primary-foreground hover:opacity-90 transition-opacity">
          Transfer Comments
        </button>
      </div>

      {/* Detect Missing Sections */}
      <div className="tool-card">
        <div className="tool-title">Detect Missing Sections</div>
        <p className="text-xs text-muted-foreground leading-relaxed">
          Compares headings between Source and Target. Reports headings present in Source but missing from Target.
        </p>
        <button onClick={stub("Detect Missing Sections")} className="w-full mt-2 px-3 py-1.5 text-xs font-semibold rounded bg-primary text-primary-foreground hover:opacity-90 transition-opacity">
          Run Detection
        </button>
      </div>

      {/* Replace Key Phrases */}
      <div className="tool-card">
        <div className="tool-title">Replace Key Phrases</div>
        <p className="text-xs text-muted-foreground leading-relaxed">
          Uses an Excel dictionary (Column A → Column B) to replace phrases in the Target document.
        </p>
        <div className="flex items-center gap-2 mt-2">
          <button onClick={handleSelectExcel} className="px-2.5 py-1 text-xs rounded border border-input bg-secondary text-secondary-foreground hover:opacity-80 transition-opacity">
            Select Excel File
          </button>
          <span className="text-xs text-muted-foreground truncate">
            {excelFile || "No file selected"}
          </span>
        </div>
        <button onClick={stub("Replace Key Phrases")} className="w-full mt-2 px-3 py-1.5 text-xs font-semibold rounded bg-primary text-primary-foreground hover:opacity-90 transition-opacity">
          Run Replacement
        </button>
      </div>

      {/* Text Cleanup */}
      <div className="tool-card">
        <div className="tool-title">Text Cleanup</div>
        <div className="space-y-1.5">
          <CheckboxOption label='Replace "par." and "pars." with "абз."' checked={replacePar} onChange={setReplacePar} />
          <CheckboxOption label="Replace spaces before numbers with en spaces" checked={replaceSpaces} onChange={setReplaceSpaces} />
        </div>
        <button onClick={stub("Text Cleanup")} className="w-full mt-2 px-3 py-1.5 text-xs font-semibold rounded bg-primary text-primary-foreground hover:opacity-90 transition-opacity">
          Run Cleanup
        </button>
      </div>

      {/* Document Metadata */}
      <div className="tool-card">
        <div className="tool-title">Document Metadata</div>
        <div className="space-y-1.5">
          <CheckboxOption label="Clean first paragraph until first Cyrillic character" checked={cleanFirst} onChange={setCleanFirst} />
          <CheckboxOption label="Convert first paragraph to sentence case" checked={sentenceCase} onChange={setSentenceCase} />
          <CheckboxOption label="Set as document Title property" checked={setTitle} onChange={setSetTitle} />
        </div>
        <button onClick={stub("Apply Metadata")} className="w-full mt-2 px-3 py-1.5 text-xs font-semibold rounded bg-primary text-primary-foreground hover:opacity-90 transition-opacity">
          Apply Metadata
        </button>
      </div>
    </div>
  );
};

export default ADPSection;
