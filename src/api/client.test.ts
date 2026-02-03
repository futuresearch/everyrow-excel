import { describe, it, expect, vi, beforeEach } from "vitest";

// Mock fetch
const mockFetch = vi.fn();
globalThis.fetch = mockFetch;

// Import after mocking
import {
  createSession,
  submitTask,
  getTaskStatus,
  getArtifacts,
  getSessionUrl,
} from "./client";

describe("API Client", () => {
  beforeEach(() => {
    mockFetch.mockReset();
  });

  describe("createSession", () => {
    it("creates a session successfully", async () => {
      mockFetch.mockResolvedValueOnce({
        ok: true,
        json: async () => ({ session_id: "test-session-123" }),
      });

      const result = await createSession("sk-cho-test", "Test Session");

      expect(result.session_id).toBe("test-session-123");
      expect(mockFetch).toHaveBeenCalledWith(
        expect.stringContaining("/sessions/create"),
        expect.objectContaining({
          method: "POST",
          headers: expect.objectContaining({
            Authorization: "Bearer sk-cho-test",
            "Content-Type": "application/json",
          }),
        })
      );
    });

    it("throws on 401 unauthorized", async () => {
      mockFetch.mockResolvedValueOnce({
        ok: false,
        status: 401,
      });

      await expect(createSession("invalid-key", "Test")).rejects.toThrow(
        "Invalid API key"
      );
    });

    it("throws on network error", async () => {
      mockFetch.mockRejectedValueOnce(new Error("Network failure"));

      await expect(createSession("sk-cho-test", "Test")).rejects.toThrow(
        "Network error"
      );
    });
  });

  describe("submitTask", () => {
    it("submits a task successfully", async () => {
      mockFetch.mockResolvedValueOnce({
        ok: true,
        json: async () => ({ task_id: "task-456" }),
      });

      const result = await submitTask("sk-cho-test", "session-123", {
        task_type: "deep_rank",
        query: { task: "rank by score" },
      });

      expect(result.task_id).toBe("task-456");
    });
  });

  describe("getTaskStatus", () => {
    it("returns task status", async () => {
      mockFetch.mockResolvedValueOnce({
        ok: true,
        json: async () => ({
          status: "completed",
          artifact_id: "artifact-789",
        }),
      });

      const result = await getTaskStatus("sk-cho-test", "task-456");

      expect(result.status).toBe("completed");
      expect(result.artifact_id).toBe("artifact-789");
    });
  });

  describe("getArtifacts", () => {
    it("fetches artifacts by IDs", async () => {
      mockFetch.mockResolvedValueOnce({
        ok: true,
        json: async () => [
          { id: "artifact-1", type: "group", data: [] },
        ],
      });

      const result = await getArtifacts("sk-cho-test", ["artifact-1"]);

      expect(result).toHaveLength(1);
      expect(result[0].id).toBe("artifact-1");
    });
  });

  describe("getSessionUrl", () => {
    it("returns correct session URL", () => {
      const url = getSessionUrl("session-123");
      expect(url).toBe("https://everyrow.io/sessions/session-123");
    });
  });
});
