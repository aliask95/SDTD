import { useState, ReactNode } from "react";
import { ChevronDown, ChevronRight } from "lucide-react";

interface CollapsibleSectionProps {
  title: string;
  defaultOpen?: boolean;
  children: ReactNode;
}

const CollapsibleSection = ({ title, defaultOpen = false, children }: CollapsibleSectionProps) => {
  const [open, setOpen] = useState(defaultOpen);

  return (
    <div>
      <div className="section-header" onClick={() => setOpen(!open)}>
        <span>{title}</span>
        {open ? <ChevronDown className="w-4 h-4" /> : <ChevronRight className="w-4 h-4" />}
      </div>
      {open && <div className="section-body">{children}</div>}
    </div>
  );
};

export default CollapsibleSection;
