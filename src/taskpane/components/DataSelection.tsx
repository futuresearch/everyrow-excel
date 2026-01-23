import {
  makeStyles,
  tokens,
  Body1,
  Caption1,
  Badge,
} from "@fluentui/react-components";
import { SelectionInfo } from "../../excel/dataHandler";

const useStyles = makeStyles({
  container: {
    display: "flex",
    flexDirection: "column",
    gap: tokens.spacingVerticalXS,
  },
  row: {
    display: "flex",
    alignItems: "center",
    gap: tokens.spacingHorizontalS,
  },
  columns: {
    display: "flex",
    flexWrap: "wrap",
    gap: tokens.spacingHorizontalXS,
    marginTop: tokens.spacingVerticalXS,
  },
  noSelection: {
    color: tokens.colorNeutralForeground3,
    fontStyle: "italic",
  },
});

interface DataSelectionProps {
  selection: SelectionInfo | null;
}

export function DataSelection({ selection }: DataSelectionProps) {
  const styles = useStyles();

  if (!selection || selection.rowCount === 0) {
    return (
      <div className={styles.container}>
        <Body1 className={styles.noSelection}>
          Select a range with headers and at least one data row
        </Body1>
      </div>
    );
  }

  return (
    <div className={styles.container}>
      <div className={styles.row}>
        <Body1>
          <strong>{selection.rowCount}</strong> rows selected
        </Body1>
        <Caption1>({selection.range})</Caption1>
      </div>

      <Caption1>Columns:</Caption1>
      <div className={styles.columns}>
        {selection.headers.map((header, idx) => (
          <Badge key={idx} appearance="outline" size="small">
            {header}
          </Badge>
        ))}
      </div>
    </div>
  );
}
