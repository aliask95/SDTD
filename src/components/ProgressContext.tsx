import { createContext, useContext, useState, useCallback, useRef, ReactNode } from "react";

interface ProgressState {
  running: boolean;
  label: string;
  progress: number; // 0-100
}

interface ProgressContextType {
  state: ProgressState;
  start: (label: string) => void;
  update: (progress: number) => void;
  cancel: () => void;
  isCancelled: () => boolean;
}

const ProgressContext = createContext<ProgressContextType | null>(null);

export const useProgress = () => {
  const ctx = useContext(ProgressContext);
  if (!ctx) throw new Error("useProgress must be inside ProgressProvider");
  return ctx;
};

export const ProgressProvider = ({ children }: { children: ReactNode }) => {
  const [state, setState] = useState<ProgressState>({ running: false, label: "", progress: 0 });
  const cancelledRef = useRef(false);

  const start = useCallback((label: string) => {
    cancelledRef.current = false;
    setState({ running: true, label, progress: 0 });
  }, []);

  const update = useCallback((progress: number) => {
    setState((s) => ({ ...s, progress: Math.min(100, progress) }));
  }, []);

  const cancel = useCallback(() => {
    cancelledRef.current = true;
    setState({ running: false, label: "", progress: 0 });
  }, []);

  const isCancelled = useCallback(() => cancelledRef.current, []);

  return (
    <ProgressContext.Provider value={{ state, start, update, cancel, isCancelled }}>
      {children}
    </ProgressContext.Provider>
  );
};
