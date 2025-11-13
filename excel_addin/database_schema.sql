-- Database Schema for Excel Add-In Architecture
-- Source: DEPLOYMENT.md Section 8

-- =====================================================
-- Table: dbo.workbook_model
-- Stores model metadata and tracked range definitions
-- =====================================================
CREATE TABLE dbo.workbook_model (
    -- Primary key
    model_id VARCHAR(255) PRIMARY KEY,

    -- Model metadata
    model_name VARCHAR(500) NOT NULL,
    version INT NOT NULL DEFAULT 1,

    -- Tracked ranges (stored as JSON array)
    -- Format: [{"name": "Revenue", "range": "A1:A10"}, ...]
    tracked_ranges NVARCHAR(MAX) NOT NULL,

    -- Audit fields
    created_at DATETIME2 DEFAULT GETUTCDATE(),
    updated_at DATETIME2 DEFAULT GETUTCDATE(),

    -- Indexes
    INDEX IX_workbook_model_name (model_name),
    INDEX IX_workbook_model_version (version)
);

-- =====================================================
-- Table: dbo.workbook_trace
-- Stores all tracked range change events
-- =====================================================
CREATE TABLE dbo.workbook_trace (
    -- Primary key
    trace_id BIGINT IDENTITY(1,1) PRIMARY KEY,

    -- Foreign key to model
    model_id VARCHAR(255) NOT NULL,

    -- Trace data
    timestamp DATETIME2 NOT NULL,
    tracked_range_name VARCHAR(255) NOT NULL,
    username VARCHAR(500) NOT NULL,
    value NVARCHAR(MAX),  -- Can store any value (string, number, JSON, etc.)

    -- Foreign key constraint
    CONSTRAINT FK_workbook_trace_model
        FOREIGN KEY (model_id)
        REFERENCES dbo.workbook_model(model_id)
        ON DELETE CASCADE,

    -- Indexes for common queries
    INDEX IX_workbook_trace_model_id (model_id),
    INDEX IX_workbook_trace_timestamp (timestamp DESC),
    INDEX IX_workbook_trace_range_name (tracked_range_name),
    INDEX IX_workbook_trace_username (username),
    INDEX IX_workbook_trace_composite (model_id, timestamp DESC)
);

-- =====================================================
-- Trigger: Auto-update updated_at on model changes
-- =====================================================
CREATE TRIGGER TR_workbook_model_updated_at
ON dbo.workbook_model
AFTER UPDATE
AS
BEGIN
    SET NOCOUNT ON;

    UPDATE dbo.workbook_model
    SET updated_at = GETUTCDATE()
    FROM dbo.workbook_model m
    INNER JOIN inserted i ON m.model_id = i.model_id;
END;
GO

-- =====================================================
-- Sample Queries
-- =====================================================

-- Get model with all tracked ranges
-- SELECT
--     model_id,
--     model_name,
--     version,
--     JSON_QUERY(tracked_ranges) as tracked_ranges,
--     created_at,
--     updated_at
-- FROM dbo.workbook_model
-- WHERE model_id = 'excel_abc123_xyz';

-- Get recent traces for a model
-- SELECT
--     trace_id,
--     model_id,
--     timestamp,
--     tracked_range_name,
--     username,
--     value
-- FROM dbo.workbook_trace
-- WHERE model_id = 'excel_abc123_xyz'
-- ORDER BY timestamp DESC;

-- Get all changes to a specific tracked range
-- SELECT
--     t.timestamp,
--     t.tracked_range_name,
--     t.username,
--     t.value,
--     m.model_name
-- FROM dbo.workbook_trace t
-- INNER JOIN dbo.workbook_model m ON t.model_id = m.model_id
-- WHERE t.tracked_range_name = 'Revenue'
-- ORDER BY t.timestamp DESC;

-- Get audit trail: all versions of a model
-- SELECT
--     model_id,
--     model_name,
--     version,
--     updated_at
-- FROM dbo.workbook_model
-- WHERE model_name LIKE '%Revenue%'
-- ORDER BY version DESC;
