import { useState, useEffect } from "react";
import {
  makeStyles,
  tokens,
  Textarea,
  Input,
  Button,
  Checkbox,
  Label,
} from "@fluentui/react-components";
import { SheetInfo, getSheetInfo } from "../../../excel/dataHandler";
import { runRankOperation } from "../../../api/operations";
import { SheetSelector } from "../SheetSelector";

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
  sheets: SheetInfo[];
  currentSheet: string;
  apiKey: string;
  onRefreshSheets: () => void;
  onRunning: () => void;
  onComplete: (success: boolean, message: string, sessionUrl?: string) => void;
}

export function RankConfig({
  sheets,
  currentSheet,
  apiKey,
  onRefreshSheets,
  onRunning,
  onComplete,
}: RankConfigProps) {
  const styles = useStyles();
  const [selectedSheet, setSelectedSheet] = useState(currentSheet);
  const [rowCount, setRowCount] = useState<number | undefined>(undefined);
  const [task, setTask] = useState("");
  const [fieldName, setFieldName] = useState("score");
  const [ascending, setAscending] = useState(false);

  useEffect(() => {
    setSelectedSheet(currentSheet);
  }, [currentSheet]);

  useEffect(() => {
    const loadSheetInfo = async () => {
      try {
        const info = await getSheetInfo(selectedSheet);
        setRowCount(info.rowCount);
      } catch {
        setRowCount(undefined);
      }
    };
    loadSheetInfo();
  }, [selectedSheet]);

  const handleRun = async () => {
    if (!task.trim()) {
      onComplete(false, "Please enter a task description");
      return;
    }

    onRunning();

    try {
      const result = await runRankOperation({
        apiKey,
        sheetName: selectedSheet,
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
      <SheetSelector
        label="Input Sheet"
        sheets={sheets}
        selectedSheet={selectedSheet}
        onSheetChange={setSelectedSheet}
        onRefresh={onRefreshSheets}
        rowCount={rowCount}
      />

      <div className={styles.field}>
        <Label htmlFor="task">Task Description</Label>
        <Textarea
          id="task"
          placeholder="e.g., Rank companies by growth potential and market opportunity"
          value={task}
          onChange={(_, data) => setTask(data.value)}
          rows={5}
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
