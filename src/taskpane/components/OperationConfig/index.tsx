import { Operation } from "../OperationPicker";
import { SelectionInfo } from "../../../excel/dataHandler";
import { RankConfig } from "./RankConfig";
import { ScreenConfig } from "./ScreenConfig";
import { DedupeConfig } from "./DedupeConfig";
import { AgentConfig } from "./AgentConfig";
import { MergeConfig } from "./MergeConfig";

interface OperationConfigProps {
  operation: Operation;
  selection: SelectionInfo;
  apiKey: string;
  onRunning: () => void;
  onComplete: (success: boolean, message: string, sessionUrl?: string) => void;
}

export function OperationConfig({
  operation,
  selection,
  apiKey,
  onRunning,
  onComplete,
}: OperationConfigProps) {
  const props = { selection, apiKey, onRunning, onComplete };

  switch (operation) {
    case "rank":
      return <RankConfig {...props} />;
    case "screen":
      return <ScreenConfig {...props} />;
    case "dedupe":
      return <DedupeConfig {...props} />;
    case "agent":
      return <AgentConfig {...props} />;
    case "merge":
      return <MergeConfig {...props} />;
    default:
      return null;
  }
}

export type { OperationConfigProps };
