import { useState, useEffect } from "react";
import {
  makeStyles,
  tokens,
  Textarea,
  Button,
  Label,
} from "@fluentui/react-components";
import { SheetInfo, getSheetInfo } from "../../../excel/dataHandler";
import { runDedupeOperation } from "../../../api/operations";
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

interface DedupeConfigProps {
  sheets: SheetInfo[];
  currentSheet: string;
  apiKey: string;
  onRefreshSheets: () => void;
  onRunning: () => void;
  onComplete: (success: boolean, message: string, sessionUrl?: string) => void;
}

export function DedupeConfig({
  sheets,
  currentSheet,
  apiKey,
  onRefreshSheets,
  onRunning,
  onComplete,
}: DedupeConfigProps) {
  const styles = useStyles();
  const [selectedSheet, setSelectedSheet] = useState(currentSheet);
  const [rowCount, setRowCount] = useState<number | undefined>(undefined);
  const [equivalenceRelation, setEquivalenceRelation] = useState("");

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
    if (!equivalenceRelation.trim()) {
      onComplete(false, "Please describe what makes rows equivalent");
      return;
    }

    onRunning();

    try {
      const result = await runDedupeOperation({
        apiKey,
        sheetName: selectedSheet,
        equivalenceRelation: equivalenceRelation.trim(),
      });
      onComplete(
        true,
        `Created sheet with ${result.rowCount} unique rows`,
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
        <Label htmlFor="equivalence">Equivalence Relation</Label>
        <Textarea
          id="equivalence"
          placeholder="e.g., Same company, possibly with different name variations or abbreviations"
          value={equivalenceRelation}
          onChange={(_, data) => setEquivalenceRelation(data.value)}
          rows={3}
        />
      </div>

      <Button appearance="primary" onClick={handleRun}>
        Run Dedupe
      </Button>
    </div>
  );
}
