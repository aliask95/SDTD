import { ReactNode } from "react";
import { ChevronDown, ChevronRight } from "lucide-react";

interface CollapsibleSectionProps {
  title: string;
  /** Controlled mode */
  isOpen?: boolean;
  onToggle?: () => void;
  /** Uncontrolled fallback */
  defaultOpen?: boolean;
  children: ReactNode;
}

const CollapsibleSection = ({ title, isOpen, onToggle, defaultOpen = false, children }: CollapsibleSectionProps) => {
  // If controlled props are provided, use them; otherwise fall back to uncontrolled
  const open = isOpen !== undefined ? isOpen : defaultOpen;
  const handleClick = onToggle ?? (() => {});

  return (
    <div>
      <div
        className="flex items-center justify-between cursor-pointer select-none px-3 py-2.5 text-sm font-semibold transition-colors"
        style={{
          background: open ? "hsl(190 79% 39% / 0.12)" : "hsl(200 12% 18%)",
          borderBottom: "1px solid hsl(200 10% 22%)",
          color: open ? "hsl(190 60% 70%)" : "hsl(200 20% 70%)",
        }}
        onClick={handleClick}
      >
        <span>{title}</span>
        {open ? <ChevronDown className="w-4 h-4" /> : <ChevronRight className="w-4 h-4" />}
      </div>
      {open && (
        <div
          className="px-3 py-2.5 space-y-3"
          style={{
            background: "hsl(200 12% 14%)",
            borderBottom: "1px solid hsl(200 10% 22%)",
          }}
        >
          {children}
        </div>
      )}
    </div>
  );
};

export default CollapsibleSection;
