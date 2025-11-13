-- Excel Add-In Model Tracking Database
-- PostgreSQL Schema

-- Drop tables if they exist (for clean restarts)
DROP TABLE IF EXISTS workbook_trace CASCADE;
DROP TABLE IF EXISTS workbook_model CASCADE;

-- Table: workbook_model
-- Stores metadata about registered Excel workbook models
CREATE TABLE workbook_model (
    model_id VARCHAR(255) PRIMARY KEY,
    model_name VARCHAR(500) NOT NULL,
    version INTEGER NOT NULL DEFAULT 1,
    tracked_ranges JSONB NOT NULL DEFAULT '[]'::jsonb,
    created_at TIMESTAMP NOT NULL DEFAULT CURRENT_TIMESTAMP,
    updated_at TIMESTAMP NOT NULL DEFAULT CURRENT_TIMESTAMP
);

-- Table: workbook_trace
-- Stores change history for tracked cell ranges
CREATE TABLE workbook_trace (
    trace_id BIGSERIAL PRIMARY KEY,
    model_id VARCHAR(255) NOT NULL REFERENCES workbook_model(model_id) ON DELETE CASCADE,
    timestamp TIMESTAMP NOT NULL,
    tracked_range_name VARCHAR(255) NOT NULL,
    username VARCHAR(500) NOT NULL,
    value TEXT,
    created_at TIMESTAMP NOT NULL DEFAULT CURRENT_TIMESTAMP
);

-- Indexes for performance
CREATE INDEX idx_workbook_model_name ON workbook_model(model_name);
CREATE INDEX idx_workbook_model_version ON workbook_model(version);

CREATE INDEX idx_workbook_trace_model_id ON workbook_trace(model_id);
CREATE INDEX idx_workbook_trace_timestamp ON workbook_trace(timestamp);
CREATE INDEX idx_workbook_trace_range_name ON workbook_trace(tracked_range_name);
CREATE INDEX idx_workbook_trace_username ON workbook_trace(username);

-- Function to automatically update updated_at timestamp
CREATE OR REPLACE FUNCTION update_updated_at_column()
RETURNS TRIGGER AS $$
BEGIN
    NEW.updated_at = CURRENT_TIMESTAMP;
    RETURN NEW;
END;
$$ LANGUAGE plpgsql;

-- Trigger to auto-update updated_at on workbook_model
CREATE TRIGGER tr_workbook_model_updated_at
    BEFORE UPDATE ON workbook_model
    FOR EACH ROW
    EXECUTE FUNCTION update_updated_at_column();

-- Success message
DO $$
BEGIN
    RAISE NOTICE 'Excel Add-In database schema initialized successfully!';
END $$;
