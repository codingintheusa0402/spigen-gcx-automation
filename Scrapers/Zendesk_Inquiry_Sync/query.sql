-- Caspi registered query for sync.py (param: min_ticket_id).
-- Register once per user via Caspi data_api (action=register, params=["min_ticket_id"]):
-- execution always runs under the registrant's own Caspi permissions, so every teammate
-- registers their own copy + issues their own API key (never share keys).
-- Output columns map 1:1 onto '26년 전체문의' A:AD (see build_row in sync.py).
WITH opt AS (
  SELECT f."id" fid, o.value:value::string v, o.value:name::string nm
  FROM S3.ZENDESK.TICKET_FIELDS f, LATERAL FLATTEN(input => f."payload":custom_field_options) o
), t AS (
  SELECT "id","status","created_at","updated_at","custom_fields" FROM S3.ZENDESK.TICKETS
  WHERE "id" >= ? AND "status" IN ('solved','closed')
), cf AS (
  SELECT t."id" tid, c.value:field_id::number fid, c.value:value::string v
  FROM t, LATERAL FLATTEN(input => t."custom_fields") c WHERE c.value:value IS NOT NULL
), d AS (
  SELECT cf.tid, cf.fid, COALESCE(opt.nm, cf.v) val FROM cf LEFT JOIN opt ON opt.fid = cf.fid AND opt.v = cf.v
), p AS (
  SELECT tid,
   MAX(IFF(fid=4513936822297,val,NULL)) country, MAX(IFF(fid=5495572594201,val,NULL)) brand,
   MAX(IFF(fid=900006613446,val,NULL)) category, MAX(IFF(fid=360019639831,val,NULL)) channel,
   MAX(IFF(fid=360022182831,val,NULL)) defect1, MAX(IFF(fid=5274834603289,val,NULL)) defect2,
   MAX(IFF(fid=360022185671,val,NULL)) device, MAX(IFF(fid=360022185891,val,NULL)) product,
   MAX(IFF(fid=360022192791,val,NULL)) pacc_cat, MAX(IFF(fid=360022301351,val,NULL)) caseology,
   MAX(IFF(fid=360019586172,val,NULL)) purchase, MAX(IFF(fid=900007557523,val,NULL)) esc,
   MAX(IFF(fid=360021363112,val,NULL)) t2, MAX(IFF(fid=26936618247577,val,NULL)) photo,
   MAX(IFF(fid=5109079191833,val,NULL)) final_resp, MAX(IFF(fid=360021934132,val,NULL)) order_id,
   MAX(IFF(fid=360021934312,val,NULL)) asin, MAX(IFF(fid=900008676703,val,NULL)) sku,
   MAX(IFF(fid=17592900949401,val,NULL)) t1t2, MAX(IFF(fid=17600391188249,val,NULL)) t2t3,
   MAX(IFF(fid=21714421937305,val,NULL)) orders_total, MAX(IFF(fid=21745453864345,val,NULL)) refunds_total,
   MAX(IFF(fid=21745465897369,val,NULL)) refunds_spigen
  FROM d GROUP BY tid
), m AS (
  SELECT "ticket_id" mtid, "replies" replies, "agent_wait_time_in_minutes_calendar" agent_wait, "on_hold_time_in_minutes_calendar" on_hold
  FROM S3.ZENDESK.TICKET_METRIC_SETS QUALIFY ROW_NUMBER() OVER (PARTITION BY "ticket_id" ORDER BY "updated_at" DESC) = 1
)
SELECT t."id" ticket_id, t."status" status, TO_CHAR(t."updated_at",'YYYY-MM-DD') updated, TO_CHAR(t."created_at",'YYYY-MM-DD') created,
  p.country, p.brand, p.category, p.channel, p.defect1, p.defect2, p.device, p.product, p.pacc_cat, p.caseology,
  p.purchase, p.esc, p.t2, p.photo, p.final_resp, p.order_id, p.asin, p.sku, p.t1t2, p.t2t3,
  p.orders_total, p.refunds_total, p.refunds_spigen, m.replies, m.agent_wait, m.on_hold
FROM t LEFT JOIN p ON p.tid = t."id" LEFT JOIN m ON m.mtid = t."id"
ORDER BY t."id"
