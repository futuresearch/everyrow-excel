import { useState } from "react";
import {
  makeStyles,
  tokens,
  Textarea,
  Input,
  Button,
  Checkbox,
  Label,
} from "@fluentui/react-components";
import { SelectionInfo } from "../../../excel/dataHandler";
import { runRankOperation } from "../../../api/operations";

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
  row: {
    display: "flex",
    alignItems: "center",
    gap: tokens.spacingHorizontalM,
  },
});

interface RankConfigProps {
  selection: SelectionInfo;
  apiKey: string;
  onRunning: () => void;
  onComplete: (success: boolean, message: string, sessionUrl?: string) => void;
}

export function RankConfig({
  selection,
  apiKey,
  onRunning,
  onComplete,
}: RankConfigProps) {
  const styles = useStyles();
  const [task, setTask] = useState("");
  const [fieldName, setFieldName] = useState("score");
  const [ascending, setAscending] = useState(false);

  const handleRun = async () => {
    if (!task.trim()) {
      onComplete(false, "Please enter a task description");
      return;
    }

    onRunning();

    try {
      const result = await runRankOperation({
        apiKey,
        selection,
        task: task.trim(),
        fieldName,
        ascending,
      });
      onComplete(true, `Created sheet with ${result.rowCount} rows`, result.sessionUrl);
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
        <Label htmlFor="task">Task Description</Label>
        <Textarea
          id="task"
          placeholder="e.g., Rank companies by growth potential and market opportunity"
          value={task}
          onChange={(_, data) => setTask(data.value)}
          rows={3}
        />
      </div>

      <div className={styles.row}>
        <div className={styles.field}>
          <Label htmlFor="fieldName">Score Field Name</Label>
          <Input
            id="fieldName"
            value={fieldName}
            onChange={(_, data) => setFieldName(data.value)}
            placeholder="score"
          />
        </div>

        <Checkbox
          checked={ascending}
          onChange={(_, data) => setAscending(data.checked === true)}
          label="Ascending order"
        />
      </div>

      <Button appearance="primary" onClick={handleRun}>
        Run Rank
      </Button>
    </div>
  );
}
