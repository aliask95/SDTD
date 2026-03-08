import { useCallback } from "react";
import { useProgress } from "../components/ProgressContext";
import { toast } from "sonner";

/** Simulates a running operation with progress. Replace with real Office.js calls later. */
export const useSimulatedRun = () => {
  const { start, update, cancel, isCancelled, state } = useProgress();

  const run = useCallback(
    async (label: string) => {
      if (state.running) {
        toast.warning("Another operation is already running.");
        return;
      }
      start(label);
      for (let i = 0; i <= 100; i += 5) {
        if (isCancelled()) {
          toast.info(`${label} cancelled.`);
          return;
        }
        update(i);
        await new Promise((r) => setTimeout(r, 120));
      }
      update(100);
      // auto-clear after a short delay
      setTimeout(() => cancel(), 600);
      toast.success(`${label} completed.`);
    },
    [start, update, cancel, isCancelled, state.running]
  );

  return run;
};
