import {
  makeStyles,
  tokens,
  Card,
  CardHeader,
  Text,
} from "@fluentui/react-components";

const useStyles = makeStyles({
  container: {
    display: "flex",
    flexDirection: "column",
    gap: tokens.spacingVerticalS,
  },
  card: {
    cursor: "pointer",
    "&:hover": {
      backgroundColor: tokens.colorNeutralBackground1Hover,
    },
  },
  cardSelected: {
    cursor: "pointer",
    backgroundColor: tokens.colorBrandBackground2,
    border: `1px solid ${tokens.colorBrandStroke1}`,
  },
  description: {
    color: tokens.colorNeutralForeground3,
    fontSize: tokens.fontSizeBase200,
  },
});

export type Operation = "rank" | "screen" | "dedupe" | "agent" | "merge";

interface OperationDef {
  id: Operation;
  name: string;
  description: string;
}

const operations: OperationDef[] = [
  {
    id: "rank",
    name: "Rank",
    description: "Score and sort rows by qualitative criteria",
  },
  {
    id: "screen",
    name: "Screen",
    description: "Filter rows that match specific conditions",
  },
  {
    id: "dedupe",
    name: "Dedupe",
    description: "Remove semantic duplicates",
  },
  {
    id: "merge",
    name: "Merge",
    description: "Join tables using AI-powered matching",
  },
  {
    id: "agent",
    name: "Agent",
    description: "Run AI web research on each row",
  },
];

interface OperationPickerProps {
  selected: Operation | null;
  onSelect: (op: Operation) => void;
}

export function OperationPicker({ selected, onSelect }: OperationPickerProps) {
  const styles = useStyles();

  return (
    <div className={styles.container}>
      {operations.map((op) => (
        <Card
          key={op.id}
          className={selected === op.id ? styles.cardSelected : styles.card}
          onClick={() => onSelect(op.id)}
          size="small"
        >
          <CardHeader
            header={<Text weight="semibold">{op.name}</Text>}
            description={<span className={styles.description}>{op.description}</span>}
          />
        </Card>
      ))}
    </div>
  );
}
