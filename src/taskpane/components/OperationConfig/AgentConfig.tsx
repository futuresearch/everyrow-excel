import { useState } from "react";
import {
  makeStyles,
  tokens,
  Textarea,
  Button,
  Label,
} from "@fluentui/react-components";
import { SelectionInfo } from "../../../excel/dataHandler";
import { runAgentOperation } from "../../../api/operations";

const useStyles = makeStyles({
  container: {
    display: "flex",
    flexDirection: "column",
    gap: tokens.spacingVerticalM,
  },
  field: {
    display: "flex",
    flexDirection: "column",
    gap: tokens.spacingVerticalXS,
  },
});

interface AgentConfigProps {
  selection: SelectionInfo;
  apiKey: string;
  onRunning: () => void;
  onComplete: (success: boolean, message: string, sessionUrl?: string) => void;
}

export function AgentConfig({
  selection,
  apiKey,
  onRunning,
  onComplete,
}: AgentConfigProps) {
  const styles = useStyles();
  const [task, setTask] = useState("");

  const handleRun = async () => {
    if (!task.trim()) {
      onComplete(false, "Please describe what to research");
      return;
    }

    onRunning();

    try {
      const result = await runAgentOperation({
        apiKey,
        selection,
        task: task.trim(),
      });
      onComplete(
        true,
        `Created sheet with ${result.rowCount} researched rows`,
        result.sessionUrl
      );
    } catch (error) {
      onComplete(
        false,
        error instanceof Error ? error.message : "An error occurred"
      );
    }
  };

  return (
    <div className={styles.container}>
      <div className={styles.field}>
        <Label htmlFor="task">Research Task</Label>
        <Textarea
          id="task"
          placeholder="e.g., Find LinkedIn page, headquarters location, and founding year"
          value={task}
          onChange={(_, data) => setTask(data.value)}
          rows={3}
        />
      </div>

      <Button appearance="primary" onClick={handleRun}>
        Run Agent
      </Button>
    </div>
  );
}
