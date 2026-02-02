import {
  createSession,
  submitTask,
  pollTaskCompletion,
  getArtifacts,
  getSessionUrl,
} from "./client";
import { sheetToRecords, writeResultsToSheet, Record } from "../excel/dataHandler";
import { TaskPayload, OperationResult } from "./types";

interface BaseOperationParams {
  apiKey: string;
  sheetName: string;
}

/**
 * Create an input artifact from records
 */
async function createInputArtifact(
  apiKey: string,
  sessionId: string,
  records: Record[]
): Promise<string> {
  const payload: TaskPayload = {
    task_type: "create_group",
    processing_mode: "map",
    query: {},
    input_artifacts: [],
  };

  // Create group task expects records in a special format
  const response = await submitTask(apiKey, sessionId, {
    ...payload,
    input_data: records,
  });

  const status = await pollTaskCompletion(apiKey, response.task_id);
  if (!status.artifact_id) {
    throw new Error("Failed to create input artifact");
  }

  return status.artifact_id;
}

/**
 * Extract result records from artifacts
 */
async function extractResults(
  apiKey: string,
  artifactId: string
): Promise<Record[]> {
  const artifacts = await getArtifacts(apiKey, [artifactId]);
  if (artifacts.length === 0) {
    throw new Error("No result artifacts found");
  }

  const artifact = artifacts[0];

  // Handle group artifacts
  if (artifact.type === "group" && artifact.artifacts) {
    const records: Record[] = [];
    for (const child of artifact.artifacts) {
      if (child.data && Array.isArray(child.data)) {
        records.push(...(child.data as Record[]));
      }
    }
    return records;
  }

  // Handle standalone artifacts
  if (artifact.data && Array.isArray(artifact.data)) {
    return artifact.data as Record[];
  }

  throw new Error("Unexpected artifact format");
}

// ============ RANK ============

interface RankParams extends BaseOperationParams {
  task: string;
  fieldName: string;
  ascending: boolean;
}

export async function runRankOperation(params: RankParams): Promise<OperationResult> {
  const { apiKey, sheetName, task, fieldName, ascending } = params;

  // Create session
  const session = await createSession(apiKey, `Excel Rank: ${task.slice(0, 50)}`);
  const sessionUrl = getSessionUrl(session.session_id);

  // Get sheet data
  const records = await sheetToRecords(sheetName);
  if (records.length === 0) {
    throw new Error(`No data found in sheet "${sheetName}"`);
  }

  // Create input artifact
  const inputArtifactId = await createInputArtifact(apiKey, session.session_id, records);

  // Submit rank task
  const payload: TaskPayload = {
    task_type: "deep_rank",
    processing_mode: "map",
    query: {
      task,
      response_schema: {
        [fieldName]: {
          type: "int",
          description: `Score for ${task}`,
        },
      },
      field_to_sort_by: fieldName,
      ascending_order: ascending,
    },
    input_artifacts: [inputArtifactId],
    context_artifacts: [],
    join_with_input: true,
  };

  const taskResponse = await submitTask(apiKey, session.session_id, payload);
  const status = await pollTaskCompletion(apiKey, taskResponse.task_id);

  if (!status.artifact_id) {
    throw new Error("No result artifact returned");
  }

  // Get results and write to sheet
  const results = await extractResults(apiKey, status.artifact_id);
  const resultSheetName = await writeResultsToSheet(results, "Ranked");

  return {
    rowCount: results.length,
    sessionUrl,
    sheetName: resultSheetName,
  };
}

// ============ SCREEN ============

interface ScreenParams extends BaseOperationParams {
  task: string;
}

export async function runScreenOperation(params: ScreenParams): Promise<OperationResult> {
  const { apiKey, sheetName, task } = params;

  const session = await createSession(apiKey, `Excel Screen: ${task.slice(0, 50)}`);
  const sessionUrl = getSessionUrl(session.session_id);

  const records = await sheetToRecords(sheetName);
  if (records.length === 0) {
    throw new Error(`No data found in sheet "${sheetName}"`);
  }

  const inputArtifactId = await createInputArtifact(apiKey, session.session_id, records);

  const payload: TaskPayload = {
    task_type: "deep_screen",
    processing_mode: "map",
    query: {
      task,
      response_schema: {
        passes: {
          type: "bool",
          description: "True if meets the screening criteria",
        },
        reason: {
          type: "str",
          description: "Brief explanation of why this row passes or fails",
        },
      },
    },
    input_artifacts: [inputArtifactId],
    context_artifacts: [],
    join_with_input: true,
  };

  const taskResponse = await submitTask(apiKey, session.session_id, payload);
  const status = await pollTaskCompletion(apiKey, taskResponse.task_id);

  if (!status.artifact_id) {
    throw new Error("No result artifact returned");
  }

  const results = await extractResults(apiKey, status.artifact_id);
  // Filter to only passing rows
  const passingResults = results.filter((r) => r.passes === true);
  const resultSheetName = await writeResultsToSheet(
    passingResults.length > 0 ? passingResults : results,
    "Screened"
  );

  return {
    rowCount: passingResults.length || results.length,
    sessionUrl,
    sheetName: resultSheetName,
  };
}

// ============ DEDUPE ============

interface DedupeParams extends BaseOperationParams {
  equivalenceRelation: string;
}

export async function runDedupeOperation(params: DedupeParams): Promise<OperationResult> {
  const { apiKey, sheetName, equivalenceRelation } = params;

  const session = await createSession(
    apiKey,
    `Excel Dedupe: ${equivalenceRelation.slice(0, 50)}`
  );
  const sessionUrl = getSessionUrl(session.session_id);

  const records = await sheetToRecords(sheetName);
  if (records.length === 0) {
    throw new Error(`No data found in sheet "${sheetName}"`);
  }

  const inputArtifactId = await createInputArtifact(apiKey, session.session_id, records);

  const payload: TaskPayload = {
    task_type: "dedupe",
    processing_mode: "map",
    query: {
      equivalence_relation: equivalenceRelation,
    },
    input_artifacts: [inputArtifactId],
    context_artifacts: [],
    join_with_input: true,
  };

  const taskResponse = await submitTask(apiKey, session.session_id, payload);
  const status = await pollTaskCompletion(apiKey, taskResponse.task_id);

  if (!status.artifact_id) {
    throw new Error("No result artifact returned");
  }

  const results = await extractResults(apiKey, status.artifact_id);
  // Filter to selected/canonical records
  const uniqueResults = results.filter((r) => r.selected === true);
  const resultSheetName = await writeResultsToSheet(
    uniqueResults.length > 0 ? uniqueResults : results,
    "Deduped"
  );

  return {
    rowCount: uniqueResults.length || results.length,
    sessionUrl,
    sheetName: resultSheetName,
  };
}

// ============ AGENT ============

interface AgentParams extends BaseOperationParams {
  task: string;
}

export async function runAgentOperation(params: AgentParams): Promise<OperationResult> {
  const { apiKey, sheetName, task } = params;

  const session = await createSession(apiKey, `Excel Agent: ${task.slice(0, 50)}`);
  const sessionUrl = getSessionUrl(session.session_id);

  const records = await sheetToRecords(sheetName);
  if (records.length === 0) {
    throw new Error(`No data found in sheet "${sheetName}"`);
  }

  const inputArtifactId = await createInputArtifact(apiKey, session.session_id, records);

  const payload: TaskPayload = {
    task_type: "agent",
    processing_mode: "map",
    query: {
      task,
      effort_level: "low",
    },
    input_artifacts: [inputArtifactId],
    context_artifacts: [],
    join_with_input: true,
  };

  const taskResponse = await submitTask(apiKey, session.session_id, payload);
  const status = await pollTaskCompletion(apiKey, taskResponse.task_id);

  if (!status.artifact_id) {
    throw new Error("No result artifact returned");
  }

  const results = await extractResults(apiKey, status.artifact_id);
  const resultSheetName = await writeResultsToSheet(results, "Researched");

  return {
    rowCount: results.length,
    sessionUrl,
    sheetName: resultSheetName,
  };
}

// ============ MERGE ============

interface MergeParams {
  apiKey: string;
  leftSheetName: string;
  rightSheetName: string;
  task: string;
  mergeOnLeft?: string;
  mergeOnRight?: string;
}

export async function runMergeOperation(params: MergeParams): Promise<OperationResult> {
  const { apiKey, leftSheetName, rightSheetName, task, mergeOnLeft, mergeOnRight } = params;

  const session = await createSession(apiKey, `Excel Merge: ${task.slice(0, 50)}`);
  const sessionUrl = getSessionUrl(session.session_id);

  // Get both tables
  const leftRecords = await sheetToRecords(leftSheetName);
  if (leftRecords.length === 0) {
    throw new Error(`No data found in left table sheet "${leftSheetName}"`);
  }

  const rightRecords = await sheetToRecords(rightSheetName);
  if (rightRecords.length === 0) {
    throw new Error(`No data found in right table sheet "${rightSheetName}"`);
  }

  // Create input artifacts for both tables
  const leftArtifactId = await createInputArtifact(apiKey, session.session_id, leftRecords);
  const rightArtifactId = await createInputArtifact(apiKey, session.session_id, rightRecords);

  const query: Record = { task };
  if (mergeOnLeft) query.merge_on_left = mergeOnLeft;
  if (mergeOnRight) query.merge_on_right = mergeOnRight;

  const payload: TaskPayload = {
    task_type: "deep_merge",
    processing_mode: "map",
    query,
    input_artifacts: [leftArtifactId, rightArtifactId],
    context_artifacts: [],
    join_with_input: true,
  };

  const taskResponse = await submitTask(apiKey, session.session_id, payload);
  const status = await pollTaskCompletion(apiKey, taskResponse.task_id);

  if (!status.artifact_id) {
    throw new Error("No result artifact returned");
  }

  const results = await extractResults(apiKey, status.artifact_id);
  const resultSheetName = await writeResultsToSheet(results, "Merged");

  return {
    rowCount: results.length,
    sessionUrl,
    sheetName: resultSheetName,
  };
}
