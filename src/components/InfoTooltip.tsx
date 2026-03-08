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
        <div className="fixed inset-x-0 mx-3 z-50 p-2.5 text-[11px] leading-relaxed normal-case tracking-normal font-normal text-foreground rounded-md border border-border bg-card shadow-lg" style={{ top: ref.current?.getBoundingClientRect().bottom ?? 0 + 4 }}>
          {text}
        </div>
      )}
    </div>
  );
};

export default InfoTooltip;
