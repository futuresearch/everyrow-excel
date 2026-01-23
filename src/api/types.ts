export interface TaskPayload {
  task_type: string;
  processing_mode: string;
  query: unknown;
  input_artifacts: string[];
  context_artifacts?: string[];
  join_with_input?: boolean;
}

export interface RankQuery {
  task: string;
  response_schema: {
    [field: string]: {
      type: string;
      description: string;
    };
  };
  field_to_sort_by: string;
  ascending_order: boolean;
}

export interface ScreenQuery {
  task: string;
  response_schema: {
    passes: {
      type: string;
      description: string;
    };
    reason?: {
      type: string;
      description: string;
    };
  };
}

export interface DedupeQuery {
  equivalence_relation: string;
}

export interface AgentQuery {
  task: string;
  effort_level?: "low" | "medium" | "high";
}

export interface MergeQuery {
  task: string;
  merge_on_left?: string;
  merge_on_right?: string;
}

export interface OperationResult {
  rowCount: number;
  sessionUrl: string;
  sheetName: string;
}
