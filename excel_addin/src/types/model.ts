/**
 * Type definitions matching the architecture specification
 * Source: DEPLOYMENT.md
 */

/**
 * Tracked Range
 * Represents a named cell range that is monitored
 */
export interface TrackedRange {
  name: string;
  range: string;
}

/**
 * Workbook Model
 * Complete model metadata stored in dbo.workbook_model
 */
export interface WorkbookModel {
  model_name: string;
  tracked_ranges: TrackedRange[];
  model_id: string;
  version: number;
}

/**
 * Workbook Trace
 * Individual trace log entry stored in dbo.workbook_trace
 */
export interface WorkbookTrace {
  model_id: string;
  timestamp: string;
  tracked_range_name: string;
  username: string;
  value: any;
}

/**
 * Request: PUT /wb/upsert-model
 */
export interface UpsertModelRequest {
  model_name: string;
  tracked_ranges: TrackedRange[];
  model_id?: string;  // Optional: omit for new model
  version?: number;   // Optional: omit for new model
}

/**
 * Response: PUT /wb/upsert-model
 * Response: GET /wb/load-model
 */
export interface UpsertModelResponse {
  model_name: string;
  tracked_ranges: TrackedRange[];
  model_id: string;
  version: number;
}

/**
 * Request: GET /wb/load-model
 */
export interface LoadModelRequest {
  model_id: string;
}

/**
 * Request: POST /wb/create-model-trace
 */
export interface CreateTraceRequest {
  model_id: string;
  timestamp: string;
  tracked_range_name: string;
  username: string;
  value: any;
}

/**
 * Request: POST /wb/create-model-trace (Batch version)
 */
export interface CreateTraceBatchRequest {
  model_id: string;
  timestamp: string;
  changes: Array<{
    tracked_range_name: string;
    value: any;
  }>;
  username: string;
}

/**
 * Response: POST /wb/create-model-trace
 */
export interface CreateTraceResponse {
  success: boolean;
}
