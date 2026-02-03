import { useState, useEffect, useCallback } from "react";
import {
  makeStyles,
  tokens,
  Body1,
  Spinner,
  MessageBar,
  MessageBarBody,
  MessageBarTitle,
  Link,
  Button,
} from "@fluentui/react-components";
import { Question24Regular } from "@fluentui/react-icons";
import { ApiKeySetup } from "./components/ApiKeySetup";
import { OperationPicker, Operation } from "./components/OperationPicker";
import { OperationConfig } from "./components/OperationConfig";
import { HelpModal } from "./components/HelpModal";
import { getApiKey } from "../config/settings";
import { getAvailableSheets, SheetInfo } from "../excel/dataHandler";

// EveryRow logo as base64
const LOGO_BASE64 = "data:image/png;base64,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";

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
    display: "flex",
    alignItems: "center",
    gap: tokens.spacingHorizontalM,
    marginBottom: tokens.spacingVerticalM,
    paddingBottom: tokens.spacingVerticalS,
    borderBottom: `1px solid ${tokens.colorNeutralStroke1}`,
  },
  logo: {
    height: "32px",
  },
  helpButton: {
    minWidth: "auto",
    padding: tokens.spacingHorizontalXS,
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
  const [helpModalOpen, setHelpModalOpen] = useState(false);

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
          <img src={LOGO_BASE64} alt="EveryRow" className={styles.logo} />
          <Button
            className={styles.helpButton}
            appearance="subtle"
            icon={<Question24Regular />}
            onClick={() => setHelpModalOpen(true)}
          />
        </div>
        <ApiKeySetup onSaved={handleApiKeySaved} />
        <HelpModal
          open={helpModalOpen}
          onClose={() => setHelpModalOpen(false)}
          operation={null}
        />
      </div>
    );
  }

  return (
    <div className={styles.container}>
      <div className={styles.header}>
        <img src={LOGO_BASE64} alt="EveryRow" className={styles.logo} />
        <Button
          className={styles.helpButton}
          appearance="subtle"
          icon={<Question24Regular />}
          onClick={() => setHelpModalOpen(true)}
        />
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

      <HelpModal
        open={helpModalOpen}
        onClose={() => setHelpModalOpen(false)}
        operation={selectedOperation}
      />
    </div>
  );
}
