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
import { OperationPicker, Operation } from "./components/OperationPicker";
import { OperationConfig } from "./components/OperationConfig";
import { getApiKey } from "../config/settings";
import { getAvailableSheets, SheetInfo } from "../excel/dataHandler";

const useStyles = makeStyles({
  container: {
    display: "flex",
    flexDirection: "column",
    height: "100%",
    minHeight: 0,
    padding: tokens.spacingHorizontalM,
  },
  header: {
    flexShrink: 0,
    marginBottom: tokens.spacingVerticalM,
    paddingBottom: tokens.spacingVerticalS,
    borderBottom: `1px solid ${tokens.colorNeutralStroke1}`,
  },
  title: {
    color: tokens.colorBrandForeground1,
  },
  content: {
    flex: 1,
    minHeight: 0,
    display: "flex",
    flexDirection: "column",
    gap: tokens.spacingVerticalM,
    overflowY: "auto",
    paddingBottom: tokens.spacingVerticalM,
  },
  step: {
    flexShrink: 0,
    backgroundColor: tokens.colorNeutralBackground1,
    borderRadius: tokens.borderRadiusMedium,
    padding: tokens.spacingHorizontalM,
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
    flexShrink: 0,
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
  const [sheets, setSheets] = useState<SheetInfo[]>([]);
  const [currentSheet, setCurrentSheet] = useState<string>("");
  const [selectedOperation, setSelectedOperation] = useState<Operation | null>(
    null
  );
  const [statusMessage, setStatusMessage] = useState<string>("");
  const [sessionUrl, setSessionUrl] = useState<string | null>(null);
  const [errorMessage, setErrorMessage] = useState<string>("");

  // Load available sheets
  const loadSheets = useCallback(async () => {
    try {
      const result = await getAvailableSheets();
      setSheets(result.sheets);
      setCurrentSheet(result.currentSheet);
    } catch (error) {
      console.error("Error getting sheets:", error);
    }
  }, []);

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

  // Load sheets when ready
  useEffect(() => {
    if (appState !== "ready") return;
    loadSheets();
  }, [appState, loadSheets]);

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
              <div className={styles.stepTitle}>1. Choose Operation</div>
              <OperationPicker
                selected={selectedOperation}
                onSelect={setSelectedOperation}
              />
            </div>

            {selectedOperation && sheets.length > 0 && (
              <div className={styles.step}>
                <div className={styles.stepTitle}>2. Configure & Run</div>
                <OperationConfig
                  operation={selectedOperation}
                  sheets={sheets}
                  currentSheet={currentSheet}
                  apiKey={apiKey}
                  onRefreshSheets={loadSheets}
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
