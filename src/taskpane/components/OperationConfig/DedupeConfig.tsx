import { useState } from "react";
import {
  makeStyles,
  tokens,
  Textarea,
  Button,
  Label,
} from "@fluentui/react-components";
import { SelectionInfo } from "../../../excel/dataHandler";
import { runDedupeOperation } from "../../../api/operations";

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
  selection: SelectionInfo;
  apiKey: string;
  onRunning: () => void;
  onComplete: (success: boolean, message: string, sessionUrl?: string) => void;
}

export function DedupeConfig({
  selection,
  apiKey,
  onRunning,
  onComplete,
}: DedupeConfigProps) {
  const styles = useStyles();
  const [equivalenceRelation, setEquivalenceRelation] = useState("");

  const handleRun = async () => {
    if (!equivalenceRelation.trim()) {
      onComplete(false, "Please describe what makes rows equivalent");
      return;
    }

    onRunning();

    try {
      const result = await runDedupeOperation({
        apiKey,
        selection,
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
