import { useState } from "react";
import {
  makeStyles,
  tokens,
  Textarea,
  Button,
  Label,
} from "@fluentui/react-components";
import { SelectionInfo } from "../../../excel/dataHandler";
import { runScreenOperation } from "../../../api/operations";

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

interface ScreenConfigProps {
  selection: SelectionInfo;
  apiKey: string;
  onRunning: () => void;
  onComplete: (success: boolean, message: string, sessionUrl?: string) => void;
}

export function ScreenConfig({
  selection,
  apiKey,
  onRunning,
  onComplete,
}: ScreenConfigProps) {
  const styles = useStyles();
  const [task, setTask] = useState("");

  const handleRun = async () => {
    if (!task.trim()) {
      onComplete(false, "Please enter screening criteria");
      return;
    }

    onRunning();

    try {
      const result = await runScreenOperation({
        apiKey,
        selection,
        task: task.trim(),
      });
      onComplete(true, `Created sheet with ${result.rowCount} matching rows`, result.sessionUrl);
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
        <Label htmlFor="task">Screening Criteria</Label>
        <Textarea
          id="task"
          placeholder="e.g., B2B SaaS companies with more than 50 employees"
          value={task}
          onChange={(_, data) => setTask(data.value)}
          rows={3}
        />
      </div>

      <Button appearance="primary" onClick={handleRun}>
        Run Screen
      </Button>
    </div>
  );
}
