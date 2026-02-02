import { useState, useEffect } from "react";
import {
  makeStyles,
  tokens,
  Textarea,
  Button,
  Label,
} from "@fluentui/react-components";
import { SheetInfo, getSheetInfo } from "../../../excel/dataHandler";
import { runScreenOperation } from "../../../api/operations";
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
});

interface ScreenConfigProps {
  sheets: SheetInfo[];
  currentSheet: string;
  apiKey: string;
  onRefreshSheets: () => void;
  onRunning: () => void;
  onComplete: (success: boolean, message: string, sessionUrl?: string) => void;
}

export function ScreenConfig({
  sheets,
  currentSheet,
  apiKey,
  onRefreshSheets,
  onRunning,
  onComplete,
}: ScreenConfigProps) {
  const styles = useStyles();
  const [selectedSheet, setSelectedSheet] = useState(currentSheet);
  const [rowCount, setRowCount] = useState<number | undefined>(undefined);
  const [task, setTask] = useState("");

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
      onComplete(false, "Please enter screening criteria");
      return;
    }

    onRunning();

    try {
      const result = await runScreenOperation({
        apiKey,
        sheetName: selectedSheet,
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
      <SheetSelector
        label="Input Sheet"
        sheets={sheets}
        selectedSheet={selectedSheet}
        onSheetChange={setSelectedSheet}
        onRefresh={onRefreshSheets}
        rowCount={rowCount}
      />

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
