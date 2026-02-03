import {
  makeStyles,
  tokens,
  Dialog,
  DialogSurface,
  DialogBody,
  DialogTitle,
  DialogContent,
  Button,
} from "@fluentui/react-components";
import { Dismiss24Regular } from "@fluentui/react-icons";
import { Operation } from "./OperationPicker";

const useStyles = makeStyles({
  surface: {
    maxWidth: "400px",
  },
  header: {
    display: "flex",
    justifyContent: "space-between",
    alignItems: "center",
  },
  closeButton: {
    minWidth: "auto",
  },
  content: {
    display: "flex",
    flexDirection: "column",
    gap: tokens.spacingVerticalM,
  },
  section: {
    marginBottom: tokens.spacingVerticalS,
  },
  example: {
    backgroundColor: tokens.colorNeutralBackground3,
    padding: tokens.spacingHorizontalS,
    borderRadius: tokens.borderRadiusSmall,
    fontStyle: "italic",
    fontSize: tokens.fontSizeBase200,
    marginBottom: tokens.spacingVerticalXS,
  },
  hint: {
    fontSize: tokens.fontSizeBase200,
    color: tokens.colorNeutralForeground3,
    marginTop: tokens.spacingVerticalM,
    paddingTop: tokens.spacingVerticalS,
    borderTop: `1px solid ${tokens.colorNeutralStroke1}`,
  },
});

interface HelpModalProps {
  open: boolean;
  onClose: () => void;
  operation: Operation | null;
}

export function HelpModal({ open, onClose, operation }: HelpModalProps) {
  const styles = useStyles();

  const getTitle = () => {
    if (!operation) return "How to use EveryRow";
    const titles: Record<Operation, string> = {
      rank: "How to use Rank",
      screen: "How to use Screen",
      dedupe: "How to use Dedupe",
      agent: "How to use Agent",
      merge: "How to use Merge",
    };
    return titles[operation];
  };

  const renderContent = () => {
    if (!operation) {
      return (
        <div className={styles.content}>
          <div className={styles.section}>
            <p><strong>Rank</strong> — Score and sort rows based on criteria you describe. Great for prioritizing leads, ranking candidates, or sorting by relevance.</p>
          </div>
          <div className={styles.section}>
            <p><strong>Screen</strong> — Filter rows that match specific conditions. Use natural language to describe which rows to keep.</p>
          </div>
          <div className={styles.section}>
            <p><strong>Dedupe</strong> — Find and remove duplicate rows, even when names are spelled differently or data varies slightly.</p>
          </div>
          <div className={styles.section}>
            <p><strong>Merge</strong> — Combine two tables by matching rows intelligently, even without exact key matches.</p>
          </div>
          <div className={styles.section}>
            <p><strong>Agent</strong> — Run AI-powered web research on each row. Find LinkedIn profiles, company info, or any data available online.</p>
          </div>
        </div>
      );
    }

    switch (operation) {
      case "rank":
        return (
          <div className={styles.content}>
            <p>Rank assigns a score to each row based on criteria you describe, then sorts the results.</p>
            <p><strong>Example prompts:</strong></p>
            <div className={styles.example}>"Rank companies by growth potential based on their description and funding"</div>
            <div className={styles.example}>"Score candidates by relevance to a senior engineering role"</div>
            <p><strong>Options:</strong></p>
            <p>• <em>Score column name</em> — The name of the new column containing scores</p>
            <p>• <em>Sort ascending</em> — Check to put lowest scores first</p>
          </div>
        );

      case "screen":
        return (
          <div className={styles.content}>
            <p>Screen filters your data, keeping only rows that match your criteria. Each row is evaluated and marked as passing or not.</p>
            <p><strong>Example prompts:</strong></p>
            <div className={styles.example}>"Companies with more than 100 employees in the technology sector"</div>
            <div className={styles.example}>"Candidates with at least 5 years of Python experience"</div>
            <p>Results include a <em>passes</em> column (true/false) and a <em>reason</em> explaining the decision.</p>
          </div>
        );

      case "dedupe":
        return (
          <div className={styles.content}>
            <p>Dedupe finds and groups duplicate rows, even when data isn't exactly the same. It uses AI to understand when two rows represent the same entity.</p>
            <p><strong>Example prompts:</strong></p>
            <div className={styles.example}>"Same company, possibly with different name spellings or abbreviations"</div>
            <div className={styles.example}>"Same person based on name and email, ignoring case differences"</div>
            <p>Results include <em>selected</em> (the canonical row), <em>equivalence_class_id</em>, and <em>equivalence_class_name</em> columns.</p>
          </div>
        );

      case "agent":
        return (
          <div className={styles.content}>
            <p>Agent performs AI-powered web research for each row. It can browse the web, find information, and add new columns with the results.</p>
            <p><strong>Example prompts:</strong></p>
            <div className={styles.example}>"Find this company's LinkedIn page, headquarters location, and founding year"</div>
            <div className={styles.example}>"Look up the CEO's name and their Twitter handle"</div>
            <p><strong>Define output columns (optional):</strong></p>
            <p>Click to specify exactly which columns to return. For each column, provide a name and type (Text, Number, or Yes/No). If left empty, the agent will determine the output structure automatically.</p>
          </div>
        );

      case "merge":
        return (
          <div className={styles.content}>
            <p>Merge combines two tables by intelligently matching rows, even without exact key matches. It's like a fuzzy VLOOKUP.</p>
            <p><strong>Example prompts:</strong></p>
            <div className={styles.example}>"Match companies from both tables by name, even if spellings differ"</div>
            <div className={styles.example}>"Join on company name, accounting for 'Inc', 'LLC' variations"</div>
            <p><strong>Options:</strong></p>
            <p>• <em>Match column (optional)</em> — Specify which columns to match on, or leave blank for auto-detection</p>
          </div>
        );
    }
  };

  return (
    <Dialog open={open} onOpenChange={(_, data) => !data.open && onClose()}>
      <DialogSurface className={styles.surface}>
        <DialogBody>
          <DialogTitle
            action={
              <Button
                appearance="subtle"
                className={styles.closeButton}
                icon={<Dismiss24Regular />}
                onClick={onClose}
              />
            }
          >
            {getTitle()}
          </DialogTitle>
          <DialogContent>
            {renderContent()}
            <p className={styles.hint}>
              Select a sheet, describe your task, and click Run. Results appear in a new sheet tab.
            </p>
          </DialogContent>
        </DialogBody>
      </DialogSurface>
    </Dialog>
  );
}
