-- Database initialization script for Excel Model Tracker

-- Create workbook_model table
CREATE TABLE IF NOT EXISTS workbook_model (
    model_id VARCHAR(255) PRIMARY KEY,
    model_name VARCHAR(500) NOT NULL,
    tracked_ranges JSONB NOT NULL,
    version INTEGER NOT NULL DEFAULT 1,
    created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
    updated_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
);

-- Create index on model_name for faster lookups
CREATE INDEX IF NOT EXISTS idx_workbook_model_name ON workbook_model(model_name);

-- Create workbook_trace table
CREATE TABLE IF NOT EXISTS workbook_trace (
    trace_id SERIAL PRIMARY KEY,
    model_id VARCHAR(255) NOT NULL,
    timestamp TIMESTAMP NOT NULL,
    tracked_range_name VARCHAR(500) NOT NULL,
    username VARCHAR(255) NOT NULL,
    value JSONB,
    created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
    FOREIGN KEY (model_id) REFERENCES workbook_model(model_id) ON DELETE CASCADE
);

-- Create indexes for better query performance
CREATE INDEX IF NOT EXISTS idx_workbook_trace_model_id ON workbook_trace(model_id);
CREATE INDEX IF NOT EXISTS idx_workbook_trace_timestamp ON workbook_trace(timestamp DESC);
CREATE INDEX IF NOT EXISTS idx_workbook_trace_range_name ON workbook_trace(tracked_range_name);

-- Insert sample data for testing (optional)
-- INSERT INTO workbook_model (model_id, model_name, tracked_ranges, version)
-- VALUES (
--     'sample-model-123',
--     'Sample Financial Model',
--     '[{"name": "Inputs", "range": "Sheet1!A1:B10"}, {"name": "Outputs", "range": "Sheet1!D1:E10"}]'::jsonb,
--     1
-- );

COMMENT ON TABLE workbook_model IS 'Stores Excel workbook model metadata including tracked ranges';
COMMENT ON TABLE workbook_trace IS 'Stores change history for tracked cell ranges in models';
