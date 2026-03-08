import { useState, useRef, useEffect } from "react";
import { Info } from "lucide-react";
import { createPortal } from "react-dom";

const InfoTooltip = ({ text }: { text: string }) => {
  const [open, setOpen] = useState(false);
  const btnRef = useRef<HTMLButtonElement>(null);
  const tooltipRef = useRef<HTMLDivElement>(null);
  const [pos, setPos] = useState({ top: 0, left: 0, width: 0 });

  useEffect(() => {
    if (!open || !btnRef.current) return;

    // Find the closest scrollable panel (the add-in container)
    const container = btnRef.current.closest(".max-w-\\[400px\\]") as HTMLElement | null;
    if (!container) return;

    const containerRect = container.getBoundingClientRect();
    const btnRect = btnRef.current.getBoundingClientRect();

    setPos({
      top: btnRect.bottom + 4,
      left: containerRect.left + 12,
      width: containerRect.width - 24,
    });

    const handler = (e: MouseEvent) => {
      if (
        btnRef.current?.contains(e.target as Node) ||
        tooltipRef.current?.contains(e.target as Node)
      )
        return;
      setOpen(false);
    };
    document.addEventListener("mousedown", handler);
    return () => document.removeEventListener("mousedown", handler);
  }, [open]);

  return (
    <>
      <button
        ref={btnRef}
        onClick={() => setOpen(!open)}
        className="text-muted-foreground hover:text-accent-foreground transition-colors"
        title="Info"
      >
        <Info className="w-3.5 h-3.5" />
      </button>
      {open &&
        createPortal(
          <div
            ref={tooltipRef}
            className="fixed z-[9999] p-2.5 text-[11px] leading-relaxed normal-case tracking-normal font-normal text-foreground rounded-md border border-border bg-card shadow-lg"
            style={{ top: pos.top, left: pos.left, width: pos.width }}
          >
            {text}
          </div>,
          document.body
        )}
    </>
  );
};

export default InfoTooltip;
