// /************************************************************
//  *  MCFGen — ASIN-safe version (COL_EMAIL renamed)
//  ************************************************************/

// // ===== Output columns (W ~ AC) =====
// const OUT_RUN      = 23;
// const OUT_PAYLOAD  = 24;
// const OUT_DETAIL   = 25;
// const OUT_TRACKING = 26;
// const OUT_TIME     = 27;
// const OUT_ERROR    = 28;
// const OUT_TOTAL    = 29;

// // ===== Input columns =====
// const COL_COUNTRY = 2;   // B
// const COL_NAME    = 6;   // F
// const COL_POST    = 7;   // G
// const COL_ADDR1   = 8;   // H
// const COL_ADDR2   = 9;   // I
// const COL_PHONE   = 10;  // J
// const COL_STATE   = 11;  // K

// // 🔥 FIXED: renamed to avoid collision with TamperMonkey.gs
// const MCF_COL_EMAIL = 12;  // L

// const COL_ASIN    = 13;    // M
// const COL_SKU     = 14;    // N


// /************************************************************
//  * Marketplace resolver (EU + JP)
//  ************************************************************/
// function resolveMarketplace(country) {
//   const c = String(country || "").trim().toUpperCase();

//   if (c === "JP" || c === "JAPAN")
//     return { region: "FE", marketplaceId: "A1VC38T7YXB528" };

//   const EU_MAP = {
//     UK: "A1F83G8C2ARO7P",
//     GB: "A1F83G8C2ARO7P",
//     DE: "A1PA6795UKMFR9",
//     FR: "A13V1IB3VIYZZH",
//     IT: "APJ6JRA9NG5V4",
//     ES: "A1RKKUPIHCS9HS",
//     NL: "A1805IZSGTT6HS",
//     SE: "A2NODRKZP88ZB9",
//     PL: "A1C3SOZRARQ6R3",
//     BE: "AMEN7PMS3EDWL",
//     TR: "A33AVAJ2PDY3EV"
//   };

//   return { region: "EU", marketplaceId: EU_MAP[c] || EU_MAP["DE"] };
// }


// /************************************************************
//  * MAIN RUN
//  ************************************************************/
// function processMCFRow(sh, row) {
//   clearOutputs(sh, row);

//   try {
//     const input = readInput(sh, row);
//     const market = resolveMarketplace(input.country);

//     sh.getRange(row, OUT_PAYLOAD).setValue(JSON.stringify(input, null, 2));

//     const stock = getStockForAsin(input.asin, market);

//     sh.getRange(row, OUT_DETAIL).setValue(stock.detail);
//     sh.getRange(row, OUT_TOTAL).setValue(stock.total);
//     sh.getRange(row, OUT_TIME).setValue(timestamp());

//   } catch (err) {
//     sh.getRange(row, OUT_ERROR).setValue("Processing Error: " + err);
//   }
// }


// /************************************************************
//  * STOCK ONLY button
//  ************************************************************/
// function runStockCheckOnly(sh, row) {
//   try {
//     const input = readInput(sh, row);
//     const market = resolveMarketplace(input.country);
//     const stock = getStockForAsin(input.asin, market);

//     sh.getRange(row, OUT_DETAIL).setValue(stock.detail);
//     sh.getRange(row, OUT_TOTAL).setValue(stock.total);
//     sh.getRange(row, OUT_TIME).setValue(timestamp());

//   } catch (err) {
//     sh.getRange(row, OUT_ERROR).setValue("Stock Error: " + err);
//   }
// }


// /************************************************************
//  * Read input row
//  ************************************************************/
// function readInput(sh, row) {
//   return {
//     country:  sh.getRange(row, COL_COUNTRY).getValue(),
//     asin:     sh.getRange(row, COL_ASIN).getValue(),
//     name:     sh.getRange(row, COL_NAME).getValue(),
//     email:    sh.getRange(row, MCF_COL_EMAIL).getValue(), // <-- FIXED
//     address1: sh.getRange(row, COL_ADDR1).getValue(),
//     address2: sh.getRange(row, COL_ADDR2).getValue(),
//     postcode: sh.getRange(row, COL_POST).getValue(),
//     phone:    sh.getRange(row, COL_PHONE).getValue(),
//     state:    sh.getRange(row, COL_STATE).getValue()
//   };
// }


// /************************************************************
//  * Clear outputs
//  ************************************************************/
// function clearOutputs(sh, row) {
//   sh.getRange(row, OUT_PAYLOAD, 1, 7).clearContent();
// }


// /************************************************************
//  * Timestamp
//  ************************************************************/
// function timestamp() {
//   return Utilities.formatDate(new Date(), "Asia/Seoul", "yyyy-MM-dd HH:mm:ss");
// }


// /************************************************************
//  * Fetch inventory
//  ************************************************************/
// function fetchInventory(marketplaceId, region) {
//   const query =
//     "granularityType=Marketplace" +
//     "&granularityId=" + marketplaceId +
//     "&marketplaceIds=" + marketplaceId +
//     "&details=true";

//   const path = "/fba/inventory/v1/summaries?" + query;

//   const res = spapiFetchWithRetry("GET", path, { endpoint: region });
//   return res?.payload?.inventorySummaries || [];
// }


// /************************************************************
//  * Load prefixes
//  ************************************************************/
// function loadPrefixes(sh) {
//   const last = sh.getLastRow();
//   if (last < 2) return [];

//   return [
//     ...new Set(
//       sh.getRange(2, COL_SKU, last - 1, 1)
//         .getValues()
//         .flat()
//         .map(v => String(v || "").toUpperCase().replace(/[^A-Z0-9]/g, ""))
//         .filter(v => v)
//     )
//   ];
// }


// /************************************************************
//  * Prefix match
//  ************************************************************/
// function matchesPrefix(sku, prefixes) {
//   sku = String(sku || "").toUpperCase();
//   return prefixes.some(p => sku.startsWith(p));
// }


// /************************************************************
//  * ASIN → SKU → stock (ASIN filter + prefix filter)
//  ************************************************************/
// function getStockForAsin(asin, market) {
//   asin = String(asin || "").toUpperCase();

//   const sh = SpreadsheetApp.getActiveSheet();
//   const prefixes = loadPrefixes(sh);

//   if (!prefixes.length)
//     throw new Error("Column N has no SKU prefixes.");

//   const inventory = fetchInventory(market.marketplaceId, market.region);

//   if (!inventory.length)
//     throw new Error("Inventory API returned no items.");

//   let total = 0;
//   let lines = [];

//   inventory.forEach(item => {
//     if (String(item.asin || "").toUpperCase() !== asin) return;

//     const sku = item.sellerSku;
//     if (!sku) return;

//     if (!matchesPrefix(sku, prefixes)) return;

//     const qty =
//       item?.inventoryDetails?.fulfillableQuantity ??
//       item?.inventoryDetails?.available?.quantity ??
//       item?.totalQuantity ??
//       0;

//     total += qty;
//     lines.push(`${sku}: ${qty}`);
//   });

//   if (!lines.length)
//     throw new Error("No matching SKUs found in Amazon inventory for this ASIN.");

//   return {
//     total,
//     detail: lines.join("\n")
//   };
// }
