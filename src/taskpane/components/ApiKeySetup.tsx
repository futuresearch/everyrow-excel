import { useState } from "react";
import {
  makeStyles,
  tokens,
  Body1,
  Input,
  Button,
  Link,
  MessageBar,
  MessageBarBody,
} from "@fluentui/react-components";
import { saveApiKey, validateApiKeyFormat } from "../../config/settings";

const useStyles = makeStyles({
  container: {
    display: "flex",
    flexDirection: "column",
    gap: tokens.spacingVerticalM,
    padding: tokens.spacingHorizontalM,
  },
  inputGroup: {
    display: "flex",
    flexDirection: "column",
    gap: tokens.spacingVerticalS,
  },
});

interface ApiKeySetupProps {
  onSaved: (key: string) => void;
}

export function ApiKeySetup({ onSaved }: ApiKeySetupProps) {
  const styles = useStyles();
  const [apiKey, setApiKey] = useState("");
  const [error, setError] = useState<string | null>(null);
  const [saving, setSaving] = useState(false);

  const handleSave = () => {
    setError(null);

    if (!apiKey.trim()) {
      setError("Please enter an API key");
      return;
    }

    if (!validateApiKeyFormat(apiKey.trim())) {
      setError('API key should start with "sk-cho-"');
      return;
    }

    setSaving(true);
    try {
      saveApiKey(apiKey.trim());
      onSaved(apiKey.trim());
    } catch (err) {
      setError("Failed to save API key");
      setSaving(false);
    }
  };

  return (
    <div className={styles.container}>
      <Body1>
        Enter your EveryRow API key to get started. Get your key at{" "}
        <Link href="https://everyrow.io/api-key" target="_blank">
          everyrow.io/api-key
        </Link>
      </Body1>

      {error && (
        <MessageBar intent="error">
          <MessageBarBody>{error}</MessageBarBody>
        </MessageBar>
      )}

      <div className={styles.inputGroup}>
        <Input
          placeholder="sk-cho-..."
          value={apiKey}
          onChange={(_, data) => setApiKey(data.value)}
          onKeyDown={(e) => e.key === "Enter" && handleSave()}
          type="password"
        />
        <Button appearance="primary" onClick={handleSave} disabled={saving}>
          {saving ? "Saving..." : "Save API Key"}
        </Button>
      </div>
    </div>
  );
}
