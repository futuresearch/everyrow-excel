// Use proxy in development to avoid CORS issues
const API_BASE_URL = import.meta.env.DEV
  ? "/api"
  : "https://engine.futuresearch.ai";

export interface ApiResponse<T> {
  data?: T;
  error?: string;
}

export interface CreateSessionResponse {
  session_id: string;
}

export interface SubmitTaskResponse {
  task_id: string;
}

export interface TaskStatusResponse {
  status: "pending" | "running" | "completed" | "failed";
  artifact_id?: string;
  error?: string;
}

export interface Artifact {
  id: string;
  type: string;
  data?: unknown[];
  artifacts?: Artifact[];
}

async function makeRequest<T>(
  endpoint: string,
  apiKey: string,
  options: RequestInit = {}
): Promise<T> {
  let response: Response;
  try {
    response = await fetch(`${API_BASE_URL}${endpoint}`, {
      ...options,
      headers: {
        Authorization: `Bearer ${apiKey}`,
        "Content-Type": "application/json",
        ...options.headers,
      },
    });
  } catch (error) {
    // Network error - fetch failed
    const message = error instanceof Error ? error.message : "Unknown error";
    throw new Error(`Network error: ${message}. Check your internet connection.`);
  }

  if (!response.ok) {
    if (response.status === 401) {
      throw new Error("Invalid API key");
    }
    if (response.status === 403) {
      throw new Error("Access denied");
    }
    const errorText = await response.text();
    throw new Error(`API error (${response.status}): ${errorText}`);
  }

  return response.json();
}

export async function createSession(
  apiKey: string,
  name: string
): Promise<CreateSessionResponse> {
  return makeRequest<CreateSessionResponse>("/sessions/create", apiKey, {
    method: "POST",
    body: JSON.stringify({ name }),
  });
}

export async function submitTask(
  apiKey: string,
  sessionId: string,
  payload: unknown
): Promise<SubmitTaskResponse> {
  return makeRequest<SubmitTaskResponse>("/tasks", apiKey, {
    method: "POST",
    body: JSON.stringify({
      payload,
      session_id: sessionId,
    }),
  });
}

export async function getTaskStatus(
  apiKey: string,
  taskId: string
): Promise<TaskStatusResponse> {
  return makeRequest<TaskStatusResponse>(`/tasks/${taskId}/status`, apiKey);
}

export async function getArtifacts(
  apiKey: string,
  artifactIds: string[]
): Promise<Artifact[]> {
  const params = artifactIds.map((id) => `artifact_ids=${id}`).join("&");
  return makeRequest<Artifact[]>(`/artifacts?${params}`, apiKey);
}

/**
 * Poll for task completion with exponential backoff
 */
export async function pollTaskCompletion(
  apiKey: string,
  taskId: string,
  maxWaitMs: number = 300000 // 5 minutes
): Promise<TaskStatusResponse> {
  const startTime = Date.now();
  let interval = 2000; // Start with 2 seconds
  const maxInterval = 10000; // Max 10 seconds
  const backoffMultiplier = 1.5;

  while (Date.now() - startTime < maxWaitMs) {
    const status = await getTaskStatus(apiKey, taskId);

    if (status.status === "completed") {
      return status;
    }

    if (status.status === "failed") {
      throw new Error(status.error || "Task failed");
    }

    // Wait before next poll
    await new Promise((resolve) => setTimeout(resolve, interval));

    // Increase interval with exponential backoff
    interval = Math.min(interval * backoffMultiplier, maxInterval);
  }

  throw new Error("Task timed out. It may still complete - check the session.");
}

export function getSessionUrl(sessionId: string): string {
  return `https://everyrow.io/sessions/${sessionId}`;
}
