import { useState, useEffect } from "react";
import {
  makeStyles,
  tokens,
  Textarea,
  Input,
  Button,
  Label,
} from "@fluentui/react-components";
import { SheetInfo, getSheetInfo } from "../../../excel/dataHandler";
import { runMergeOperation } from "../../../api/operations";
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
  sheetRow: {
    display: "flex",
    flexWrap: "wrap",
    gap: tokens.spacingHorizontalM,
    "& > *": {
      flex: "1 1 120px",
      minWidth: "120px",
    },
  },
});

interface MergeConfigProps {
  sheets: SheetInfo[];
  currentSheet: string;
  apiKey: string;
  onRefreshSheets: () => void;
  onRunning: () => void;
  onComplete: (success: boolean, message: string, sessionUrl?: string) => void;
}

export function MergeConfig({
  sheets,
  currentSheet,
  apiKey,
  onRefreshSheets,
  onRunning,
  onComplete,
}: MergeConfigProps) {
  const styles = useStyles();
  const [leftSheet, setLeftSheet] = useState(currentSheet);
  const [rightSheet, setRightSheet] = useState(
    sheets.length > 1 && sheets[1].name !== currentSheet
      ? sheets[1].name
      : sheets.length > 0
      ? sheets[0].name
      : ""
  );
  const [leftRowCount, setLeftRowCount] = useState<number | undefined>(undefined);
  const [rightRowCount, setRightRowCount] = useState<number | undefined>(undefined);
  const [task, setTask] = useState("");
  const [mergeOnLeft, setMergeOnLeft] = useState("");
  const [mergeOnRight, setMergeOnRight] = useState("");

  useEffect(() => {
    setLeftSheet(currentSheet);
  }, [currentSheet]);

  useEffect(() => {
    const loadLeftInfo = async () => {
      try {
        const info = await getSheetInfo(leftSheet);
        setLeftRowCount(info.rowCount);
      } catch {
        setLeftRowCount(undefined);
      }
    };
    loadLeftInfo();
  }, [leftSheet]);

  useEffect(() => {
    const loadRightInfo = async () => {
      if (!rightSheet) {
        setRightRowCount(undefined);
        return;
      }
      try {
        const info = await getSheetInfo(rightSheet);
        setRightRowCount(info.rowCount);
      } catch {
        setRightRowCount(undefined);
      }
    };
    loadRightInfo();
  }, [rightSheet]);

  const handleRun = async () => {
    if (!task.trim()) {
      onComplete(false, "Please describe how to match rows");
      return;
    }
    if (!leftSheet) {
      onComplete(false, "Please select the left table sheet");
      return;
    }
    if (!rightSheet) {
      onComplete(false, "Please select the right table sheet");
      return;
    }

    onRunning();

    try {
      const result = await runMergeOperation({
        apiKey,
        leftSheetName: leftSheet,
        rightSheetName: rightSheet,
        task: task.trim(),
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
      <div className={styles.sheetRow}>
        <SheetSelector
          label="Left Table"
          sheets={sheets}
          selectedSheet={leftSheet}
          onSheetChange={setLeftSheet}
          onRefresh={onRefreshSheets}
          rowCount={leftRowCount}
        />
        <SheetSelector
          label="Right Table"
          sheets={sheets}
          selectedSheet={rightSheet}
          onSheetChange={setRightSheet}
          rowCount={rightRowCount}
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
        <Label htmlFor="mergeLeft">Match Column (Left Table) - Optional</Label>
        <Input
          id="mergeLeft"
          placeholder="e.g., company_name"
          value={mergeOnLeft}
          onChange={(_, data) => setMergeOnLeft(data.value)}
        />
      </div>

      <div className={styles.field}>
        <Label htmlFor="mergeRight">Match Column (Right Table) - Optional</Label>
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
