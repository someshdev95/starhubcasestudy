
-- Create a table to flatten service and order data for analysis
-- This table includes new signups, transfers, churn events, and current status

CREATE TABLE flatten_service AS 

-- First case: Capture the first active instance of each service
SELECT 
    MIN(a.SNAPSHOT_DATE) AS DateKey,         -- Earliest snapshot date
    a.SERVICE_ID AS ServiceId,               -- Unique service ID
    a.SERVICE_NAME AS ServiceName,           -- Name of the service
    a.CUSTOMER_ID AS CustomerId,             -- Customer associated with the service
    'Yes' AS CheckNewSignup,                 -- Mark as a new signup
    CASE
        WHEN o.ORDER_TYPE_L2 = 'transfer' THEN 'Yes' -- Mark as a transfer
        WHEN o.ORDER_TYPE_L2 = 'new' THEN 'No'       -- Not a transfer
        ELSE NULL                                    -- Handle gaps in data
    END AS CheckTransfer,
    'No' AS CheckChurn,                    -- No churn for the first active instance
    'Active' AS CurrentStatus              -- Mark the service as active
FROM 
    active_final a
LEFT JOIN 
    order_final o 
ON 
    a.SERVICE_ID = o.SERVICE_ID             -- Join on Service ID
GROUP BY 
    a.SERVICE_ID,                          -- Group by unique service attributes
    a.SERVICE_NAME, 
    a.CUSTOMER_ID, 
    o.ORDER_TYPE_L2

UNION 

-- Second case: Capture subsequent active snapshots for services
SELECT 
    t.DateKey,                             -- Snapshot date
    t.ServiceId,                           -- Unique service ID
    t.ServiceName,                         -- Name of the service
    t.CustomerId,                          -- Associated customer
    t.CheckNewSignup,                      -- Not a new signup
    t.CheckTransfer,                       -- Transfer check is NULL
    t.CheckChurn,                          -- Churn status remains 'No'
    t.CurrentStatus                        -- Status remains 'Active'
FROM (
    SELECT 
        a.SNAPSHOT_DATE AS DateKey,        -- Snapshot date
        a.SERVICE_ID AS ServiceId,         -- Unique service ID
        a.SERVICE_NAME AS ServiceName,     -- Name of the service
        a.CUSTOMER_ID AS CustomerId,       -- Associated customer
        'No' AS CheckNewSignup,            -- Not a new signup
        NULL AS CheckTransfer,             -- Transfer check is NULL
        'No' AS CheckChurn,                -- Churn status remains 'No'
        'Active' AS CurrentStatus,         -- Status remains active
        ROW_NUMBER() OVER (
            PARTITION BY a.SERVICE_ID 
            ORDER BY a.SNAPSHOT_DATE
        ) AS checker                       -- Assign row numbers per service
    FROM 
        active_final a
) t
WHERE 
    t.checker != 1                         -- Exclude the first instance

UNION 

-- Third case: Capture churn events from orders
SELECT 
    o.REPORT_DATE AS DateKey,              -- Order report date
    o.SERVICE_ID AS ServiceId,             -- Unique service ID
    a.SERVICE_NAME AS ServiceName,         -- Name of the service
    a.CUSTOMER_ID AS CustomerId,           -- Associated customer
    'No' AS CheckNewSignup,                -- Not a new signup
    NULL AS CheckTransfer,                 -- Transfer check is NULL
    'Yes' AS CheckChurn,                   -- Mark as a churn event
    'Inactive' AS CurrentStatus            -- Mark the service as inactive
FROM 
    order_final o 
LEFT JOIN (
    SELECT DISTINCT 
        SERVICE_ID, 
        CUSTOMER_ID, 
        SERVICE_NAME
    FROM 
        active_final                      -- Extract distinct service details
) a 
ON 
    o.SERVICE_ID = a.SERVICE_ID;           -- Join on Service ID
