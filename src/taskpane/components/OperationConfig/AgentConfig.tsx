import { useState, useEffect } from "react";
import {
  makeStyles,
  tokens,
  Textarea,
  Button,
  Label,
} from "@fluentui/react-components";
import { SheetInfo, getSheetInfo } from "../../../excel/dataHandler";
import { runAgentOperation } from "../../../api/operations";
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

interface AgentConfigProps {
  sheets: SheetInfo[];
  currentSheet: string;
  apiKey: string;
  onRefreshSheets: () => void;
  onRunning: () => void;
  onComplete: (success: boolean, message: string, sessionUrl?: string) => void;
}

export function AgentConfig({
  sheets,
  currentSheet,
  apiKey,
  onRefreshSheets,
  onRunning,
  onComplete,
}: AgentConfigProps) {
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
      onComplete(false, "Please describe what to research");
      return;
    }

    onRunning();

    try {
      const result = await runAgentOperation({
        apiKey,
        sheetName: selectedSheet,
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
      <SheetSelector
        label="Input Sheet"
        sheets={sheets}
        selectedSheet={selectedSheet}
        onSheetChange={setSelectedSheet}
        onRefresh={onRefreshSheets}
        rowCount={rowCount}
      />

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
