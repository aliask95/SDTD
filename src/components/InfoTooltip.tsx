import { useState, useRef, useEffect } from "react";
import { Info } from "lucide-react";

const InfoTooltip = ({ text }: { text: string }) => {
  const [open, setOpen] = useState(false);
  const ref = useRef<HTMLDivElement>(null);

  useEffect(() => {
    if (!open) return;
    const handler = (e: MouseEvent) => {
      if (ref.current && !ref.current.contains(e.target as Node)) setOpen(false);
    };
    document.addEventListener("mousedown", handler);
    return () => document.removeEventListener("mousedown", handler);
  }, [open]);

  return (
    <div className="relative inline-flex" ref={ref}>
      <button
        onClick={() => setOpen(!open)}
        className="text-muted-foreground hover:text-accent-foreground transition-colors"
        title="Info"
      >
        <Info className="w-3.5 h-3.5" />
      </button>
      {open && (
        <div className="absolute left-0 top-full mt-1 z-50 w-56 p-2 text-[11px] leading-relaxed text-foreground rounded-md border border-border bg-card shadow-lg">
          {text}
        </div>
      )}
    </div>
  );
};

export default InfoTooltip;
