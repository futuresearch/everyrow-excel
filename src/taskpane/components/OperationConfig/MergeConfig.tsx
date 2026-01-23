import { useState } from "react";
import {
  makeStyles,
  tokens,
  Textarea,
  Input,
  Button,
  Label,
  Body1,
  Caption1,
} from "@fluentui/react-components";
import { SelectionInfo } from "../../../excel/dataHandler";
import { runMergeOperation } from "../../../api/operations";

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
  infoBox: {
    backgroundColor: tokens.colorNeutralBackground3,
    padding: tokens.spacingHorizontalS,
    borderRadius: tokens.borderRadiusSmall,
  },
});

interface MergeConfigProps {
  selection: SelectionInfo;
  apiKey: string;
  onRunning: () => void;
  onComplete: (success: boolean, message: string, sessionUrl?: string) => void;
}

export function MergeConfig({
  selection,
  apiKey,
  onRunning,
  onComplete,
}: MergeConfigProps) {
  const styles = useStyles();
  const [task, setTask] = useState("");
  const [table2Range, setTable2Range] = useState("");
  const [mergeOnLeft, setMergeOnLeft] = useState("");
  const [mergeOnRight, setMergeOnRight] = useState("");

  const handleRun = async () => {
    if (!task.trim()) {
      onComplete(false, "Please describe how to match rows");
      return;
    }
    if (!table2Range.trim()) {
      onComplete(false, "Please specify the second table range");
      return;
    }

    onRunning();

    try {
      const result = await runMergeOperation({
        apiKey,
        selection,
        task: task.trim(),
        table2Range: table2Range.trim(),
        mergeOnLeft: mergeOnLeft.trim() || undefined,
        mergeOnRight: mergeOnRight.trim() || undefined,
      });
      onComplete(
        true,
        `Created sheet with ${result.rowCount} merged rows`,
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
      <div className={styles.infoBox}>
        <Body1>Table 1: {selection.range}</Body1>
        <Caption1>({selection.rowCount} rows)</Caption1>
      </div>

      <div className={styles.field}>
        <Label htmlFor="table2">Table 2 Range</Label>
        <Input
          id="table2"
          placeholder="e.g., Sheet2!A1:D100"
          value={table2Range}
          onChange={(_, data) => setTable2Range(data.value)}
        />
      </div>

      <div className={styles.field}>
        <Label htmlFor="task">Merge Criteria</Label>
        <Textarea
          id="task"
          placeholder="e.g., Match each software product to its parent company"
          value={task}
          onChange={(_, data) => setTask(data.value)}
          rows={2}
        />
      </div>

      <div className={styles.field}>
        <Label htmlFor="mergeLeft">Match Column (Table 1) - Optional</Label>
        <Input
          id="mergeLeft"
          placeholder="e.g., company_name"
          value={mergeOnLeft}
          onChange={(_, data) => setMergeOnLeft(data.value)}
        />
      </div>

      <div className={styles.field}>
        <Label htmlFor="mergeRight">Match Column (Table 2) - Optional</Label>
        <Input
          id="mergeRight"
          placeholder="e.g., supplier_name"
          value={mergeOnRight}
          onChange={(_, data) => setMergeOnRight(data.value)}
        />
      </div>

      <Button appearance="primary" onClick={handleRun}>
        Run Merge
      </Button>
    </div>
  );
}
