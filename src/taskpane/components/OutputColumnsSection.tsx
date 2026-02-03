import { useState, useCallback } from "react";
import {
  makeStyles,
  tokens,
  Input,
  Dropdown,
  Option,
  Button,
  Text,
} from "@fluentui/react-components";
import { Add16Regular, Dismiss16Regular } from "@fluentui/react-icons";

const useStyles = makeStyles({
  container: {
    marginBottom: tokens.spacingVerticalM,
  },
  header: {
    display: "flex",
    alignItems: "center",
    gap: tokens.spacingHorizontalXS,
    cursor: "pointer",
    padding: `${tokens.spacingVerticalS} 0`,
    color: tokens.colorBrandForeground1,
    fontSize: tokens.fontSizeBase300,
  },
  headerIcon: {
    fontSize: tokens.fontSizeBase200,
    width: "12px",
  },
  hint: {
    color: tokens.colorNeutralForeground3,
    fontSize: tokens.fontSizeBase200,
    marginLeft: tokens.spacingHorizontalXS,
  },
  content: {
    backgroundColor: tokens.colorNeutralBackground3,
    borderRadius: tokens.borderRadiusMedium,
    padding: tokens.spacingHorizontalM,
    marginTop: tokens.spacingVerticalXS,
  },
  columnRow: {
    display: "flex",
    gap: tokens.spacingHorizontalS,
    marginBottom: tokens.spacingVerticalS,
    alignItems: "center",
  },
  nameInput: {
    flex: 2,
    minWidth: 0,
  },
  typeDropdown: {
    flex: 1,
    minWidth: "80px",
  },
  removeButton: {
    minWidth: "auto",
    padding: tokens.spacingHorizontalXS,
  },
  addButton: {
    width: "100%",
    marginTop: tokens.spacingVerticalXS,
  },
});

interface OutputColumn {
  id: string;
  name: string;
  type: "string" | "number" | "boolean";
}

// API expects flat schema: { fieldName: { type: "str" | "float" | "bool" } }
export interface OutputSchema {
  [key: string]: { type: "str" | "float" | "bool"; description?: string };
}

interface OutputColumnsSectionProps {
  onChange: (schema: OutputSchema | null) => void;
}

let columnIdCounter = 0;

export function OutputColumnsSection({ onChange }: OutputColumnsSectionProps) {
  const styles = useStyles();
  const [isExpanded, setIsExpanded] = useState(false);
  const [columns, setColumns] = useState<OutputColumn[]>([]);

  const updateSchema = useCallback((cols: OutputColumn[]) => {
    const validColumns = cols.filter((c) => c.name.trim());
    if (validColumns.length === 0) {
      onChange(null);
      return;
    }

    // Convert UI types to API types (use float for numbers)
    const typeMap: Record<string, "str" | "float" | "bool"> = {
      string: "str",
      number: "float",
      boolean: "bool",
    };

    const schema: OutputSchema = {};
    for (const col of validColumns) {
      schema[col.name.trim()] = { type: typeMap[col.type] };
    }

    onChange(schema);
  }, [onChange]);

  const handleToggle = () => {
    const newExpanded = !isExpanded;
    setIsExpanded(newExpanded);

    // Add first column when expanding if empty
    if (newExpanded && columns.length === 0) {
      const newColumn: OutputColumn = {
        id: `col-${++columnIdCounter}`,
        name: "",
        type: "string",
      };
      setColumns([newColumn]);
    }
  };

  const addColumn = () => {
    const newColumn: OutputColumn = {
      id: `col-${++columnIdCounter}`,
      name: "",
      type: "string",
    };
    const newColumns = [...columns, newColumn];
    setColumns(newColumns);
  };

  const removeColumn = (id: string) => {
    const newColumns = columns.filter((c) => c.id !== id);
    setColumns(newColumns);
    updateSchema(newColumns);
  };

  const updateColumn = (id: string, field: "name" | "type", value: string) => {
    const newColumns = columns.map((c) =>
      c.id === id
        ? { ...c, [field]: field === "type" ? value as OutputColumn["type"] : value }
        : c
    );
    setColumns(newColumns);
    updateSchema(newColumns);
  };

  return (
    <div className={styles.container}>
      <div className={styles.header} onClick={handleToggle}>
        <span className={styles.headerIcon}>{isExpanded ? "▼" : "▶"}</span>
        <Text>Define output columns</Text>
        <span className={styles.hint}>(optional)</span>
      </div>

      {isExpanded && (
        <div className={styles.content}>
          {columns.map((column) => (
            <div key={column.id} className={styles.columnRow}>
              <Input
                className={styles.nameInput}
                placeholder="Column name"
                value={column.name}
                onChange={(_, data) => updateColumn(column.id, "name", data.value)}
              />
              <Dropdown
                className={styles.typeDropdown}
                value={column.type === "string" ? "Text" : column.type === "number" ? "Number" : "Yes/No"}
                selectedOptions={[column.type]}
                onOptionSelect={(_, data) => {
                  if (data.optionValue) {
                    updateColumn(column.id, "type", data.optionValue);
                  }
                }}
              >
                <Option value="string">Text</Option>
                <Option value="number">Number</Option>
                <Option value="boolean">Yes/No</Option>
              </Dropdown>
              <Button
                className={styles.removeButton}
                appearance="subtle"
                icon={<Dismiss16Regular />}
                onClick={() => removeColumn(column.id)}
              />
            </div>
          ))}

          <Button
            className={styles.addButton}
            appearance="subtle"
            icon={<Add16Regular />}
            onClick={addColumn}
          >
            Add column
          </Button>
        </div>
      )}
    </div>
  );
}
