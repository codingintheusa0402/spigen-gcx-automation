-- Caspi registered query for sync.py notifications (param: created_date, 'YYYY-MM-DD').
-- Zendesk tickets created on that date, all statuses — no channel / agent-reply filter.
-- Register via Caspi data_api (action=register, params=["created_date"]), then
-- `sync.py setup --intake-query-id pq_…` (same API key as query.sql).
SELECT "status" status, COUNT(*) n
FROM S3.ZENDESK.TICKETS
WHERE TO_CHAR("created_at",'YYYY-MM-DD') = ?
GROUP BY 1
ORDER BY 1
