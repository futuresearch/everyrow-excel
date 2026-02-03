import {
  makeStyles,
  tokens,
  Dropdown,
  Option,
  Label,
  Button,
  Caption1,
} from "@fluentui/react-components";
import { SheetInfo } from "../../excel/dataHandler";

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
  dropdown: {
    minWidth: "80px",
    flex: 1,
  },
  refreshButton: {
    minWidth: "auto",
    padding: tokens.spacingHorizontalXS,
  },
  rowInfo: {
    color: tokens.colorNeutralForeground3,
  },
});

interface SheetSelectorProps {
  label: string;
  sheets: SheetInfo[];
  selectedSheet: string;
  onSheetChange: (sheetName: string) => void;
  onRefresh?: () => void;
  rowCount?: number;
}

export function SheetSelector({
  label,
  sheets,
  selectedSheet,
  onSheetChange,
  onRefresh,
  rowCount,
}: SheetSelectorProps) {
  const styles = useStyles();

  return (
    <div className={styles.container}>
      <Label>{label}</Label>
      <div className={styles.row}>
        <Dropdown
          className={styles.dropdown}
          value={selectedSheet}
          selectedOptions={[selectedSheet]}
          onOptionSelect={(_, data) => {
            if (data.optionValue) {
              onSheetChange(data.optionValue);
            }
          }}
        >
          {sheets.map((sheet) => (
            <Option key={sheet.name} value={sheet.name}>
              {sheet.name}
            </Option>
          ))}
        </Dropdown>
        {onRefresh && (
          <Button
            className={styles.refreshButton}
            appearance="subtle"
            onClick={onRefresh}
            title="Refresh sheet list"
          >
            ↻
          </Button>
        )}
      </div>
      {rowCount !== undefined && rowCount > 0 && (
        <Caption1 className={styles.rowInfo}>{rowCount} rows</Caption1>
      )}
    </div>
  );
}
