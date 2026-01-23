import { useState, useEffect, useCallback } from "react";
import {
  makeStyles,
  tokens,
  Title2,
  Body1,
  Spinner,
  MessageBar,
  MessageBarBody,
  MessageBarTitle,
  Link,
} from "@fluentui/react-components";
import { ApiKeySetup } from "./components/ApiKeySetup";
import { DataSelection } from "./components/DataSelection";
import { OperationPicker, Operation } from "./components/OperationPicker";
import { OperationConfig } from "./components/OperationConfig";
import { getApiKey } from "../config/settings";
import { getSelectionInfo, SelectionInfo } from "../excel/dataHandler";

const useStyles = makeStyles({
  container: {
    display: "flex",
    flexDirection: "column",
    height: "100%",
    padding: tokens.spacingHorizontalM,
  },
  header: {
    marginBottom: tokens.spacingVerticalM,
    paddingBottom: tokens.spacingVerticalS,
    borderBottom: `1px solid ${tokens.colorNeutralStroke1}`,
  },
  title: {
    color: tokens.colorBrandForeground1,
  },
  content: {
    flex: 1,
    display: "flex",
    flexDirection: "column",
    gap: tokens.spacingVerticalM,
    overflowY: "auto",
  },
  step: {
    backgroundColor: tokens.colorNeutralBackground1,
    borderRadius: tokens.borderRadiusMedium,
    padding: tokens.spacingHorizontalM,
    boxShadow: tokens.shadow4,
  },
  stepTitle: {
    marginBottom: tokens.spacingVerticalS,
    fontWeight: tokens.fontWeightSemibold,
  },
  loadingContainer: {
    display: "flex",
    flexDirection: "column",
    alignItems: "center",
    justifyContent: "center",
    height: "100%",
    gap: tokens.spacingVerticalM,
  },
  footer: {
    marginTop: tokens.spacingVerticalM,
    paddingTop: tokens.spacingVerticalS,
    borderTop: `1px solid ${tokens.colorNeutralStroke1}`,
    fontSize: tokens.fontSizeBase200,
    color: tokens.colorNeutralForeground3,
  },
});

type AppState = "loading" | "setup" | "ready" | "running" | "success" | "error";

export default function App() {
  const styles = useStyles();
  const [appState, setAppState] = useState<AppState>("loading");
  const [apiKey, setApiKey] = useState<string | null>(null);
  const [selection, setSelection] = useState<SelectionInfo | null>(null);
  const [selectedOperation, setSelectedOperation] = useState<Operation | null>(
    null
  );
  const [statusMessage, setStatusMessage] = useState<string>("");
  const [sessionUrl, setSessionUrl] = useState<string | null>(null);
  const [errorMessage, setErrorMessage] = useState<string>("");

  // Check for API key on mount
  useEffect(() => {
    const key = getApiKey();
    if (key) {
      setApiKey(key);
      setAppState("ready");
    } else {
      setAppState("setup");
    }
  }, []);

  // Poll for selection changes
  useEffect(() => {
    if (appState !== "ready") return;

    const updateSelection = async () => {
      try {
        const info = await getSelectionInfo();
        setSelection(info);
      } catch (error) {
        console.error("Error getting selection:", error);
      }
    };

    updateSelection();
    const interval = setInterval(updateSelection, 1500);
    return () => clearInterval(interval);
  }, [appState]);

  const handleApiKeySaved = useCallback((key: string) => {
    setApiKey(key);
    setAppState("ready");
  }, []);

  const handleOperationComplete = useCallback(
    (success: boolean, message: string, url?: string) => {
      if (success) {
        setAppState("success");
        setStatusMessage(message);
        setSessionUrl(url || null);
      } else {
        setAppState("error");
        setErrorMessage(message);
      }
    },
    []
  );

  const handleReset = useCallback(() => {
    setAppState("ready");
    setSelectedOperation(null);
    setStatusMessage("");
    setErrorMessage("");
    setSessionUrl(null);
  }, []);

  if (appState === "loading") {
    return (
      <div className={styles.loadingContainer}>
        <Spinner size="large" />
        <Body1>Loading...</Body1>
      </div>
    );
  }

  if (appState === "setup" || !apiKey) {
    return (
      <div className={styles.container}>
        <div className={styles.header}>
          <Title2 className={styles.title}>EveryRow</Title2>
        </div>
        <ApiKeySetup onSaved={handleApiKeySaved} />
      </div>
    );
  }

  return (
    <div className={styles.container}>
      <div className={styles.header}>
        <Title2 className={styles.title}>EveryRow</Title2>
      </div>

      <div className={styles.content}>
        {appState === "success" && (
          <MessageBar intent="success">
            <MessageBarBody>
              <MessageBarTitle>Success</MessageBarTitle>
              {statusMessage}
              {sessionUrl && (
                <>
                  {" "}
                  <Link href={sessionUrl} target="_blank">
                    View session
                  </Link>
                </>
              )}
            </MessageBarBody>
          </MessageBar>
        )}

        {appState === "error" && (
          <MessageBar intent="error">
            <MessageBarBody>
              <MessageBarTitle>Error</MessageBarTitle>
              {errorMessage}
            </MessageBarBody>
          </MessageBar>
        )}

        {(appState === "success" || appState === "error") && (
          <Link onClick={handleReset}>Run another operation</Link>
        )}

        {appState === "running" && (
          <div className={styles.loadingContainer}>
            <Spinner size="large" />
            <Body1>Processing... This may take several minutes.</Body1>
          </div>
        )}

        {appState === "ready" && (
          <>
            <div className={styles.step}>
              <div className={styles.stepTitle}>1. Select Data</div>
              <DataSelection selection={selection} />
            </div>

            <div className={styles.step}>
              <div className={styles.stepTitle}>2. Choose Operation</div>
              <OperationPicker
                selected={selectedOperation}
                onSelect={setSelectedOperation}
              />
            </div>

            {selectedOperation && selection && selection.rowCount > 0 && (
              <div className={styles.step}>
                <div className={styles.stepTitle}>3. Configure & Run</div>
                <OperationConfig
                  operation={selectedOperation}
                  selection={selection}
                  apiKey={apiKey}
                  onRunning={() => setAppState("running")}
                  onComplete={handleOperationComplete}
                />
              </div>
            )}
          </>
        )}
      </div>

      <div className={styles.footer}>
        API Key: {apiKey.substring(0, 10)}...{" "}
        <Link onClick={() => setAppState("setup")}>Change</Link>
      </div>
    </div>
  );
}
