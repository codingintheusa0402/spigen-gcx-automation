"""Step 9a — print the Caspi SQL for Amazon EU units (2026, non-cancelled, dedup by order+sku, amzn.gr resale excluded)
for every SKU of the summary-deck cases. Run it with mcp__claude_ai_CaspiLM__run_query (limit 1000), then
python3 09b_save_eu_sales.py to pull the result out of this session's transcript into eu_sales.json.
EU = DE·FR·IT·ES·NL·SE·BE·IE·PL + UK (user: 'Amazon EU only, no JP/US/IN'; UK counted as Pan-EU — confirm if unsure).
Usage: python3 09a_eu_sales_sql.py [--include-sp]"""
import sys, json, glob, datetime, common  # noqa
inc = '--include-sp' in sys.argv
skus = set()
for f in sorted(glob.glob('case_[0-9][0-9].json')):
    d = json.load(open(f))
    if not inc and '(SP)' in d['product']['product_name'] + d['issue_title']: pass
    skus |= {x['sku'] for x in d['product']['skus']}
ch = ",".join(f"'{c}'" for c in common.EU_CHANNELS)
print(f"""WITH o AS (
  SELECT DISTINCT "amazon-order-id" AS oid, "sku" AS sku, "sales-channel" AS ch, TRY_TO_NUMBER("quantity") AS q
  FROM S3.AMAZON_SELLER.FLAT_FILE_ALL_ORDERS_DATA_BY_ORDER_DATE_GENERAL
  WHERE "purchase-date" >= '{datetime.date.today().year}-01-01' AND "order-status" <> 'Cancelled'
    AND "sales-channel" IN ({ch}) AND "sku" NOT ILIKE 'amzn.gr.%'
)
SELECT REGEXP_SUBSTR(sku,'[A-Z]{{3}}[0-9]{{5}}') AS base_sku, SUM(q) AS units, SUM(IFF(ch='Amazon.co.uk', q, 0)) AS uk_units
FROM o WHERE REGEXP_SUBSTR(sku,'[A-Z]{{3}}[0-9]{{5}}') IN ({",".join(f"'{s}'" for s in sorted(skus))})
GROUP BY 1 ORDER BY 1""")
