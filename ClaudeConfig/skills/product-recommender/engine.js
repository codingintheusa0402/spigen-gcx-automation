// Runs the Chat app's own matcher (GAS_Operations/TicketReporterCard/ProductSearch.gs) on a product index,
// so /products-similar and this skill always agree. stdin: {index, query}  argv[2]: path to ProductSearch.gs
// stdout: {base, label, cats, sameKind, same: [sku…]} or {error}
const fs = require('fs');
const api = new Function(fs.readFileSync(process.argv[2], 'utf8') + ';return {similarProducts_};')();
const inp = JSON.parse(fs.readFileSync(0, 'utf8'));
const r = api.similarProducts_(inp.index, inp.query, () => 0);
process.stdout.write(JSON.stringify(r.error ? { error: r.error } : {
  base: r.base && r.base.sku, label: r.label, cats: r.cats.map(c => c.label), sameKind: r.sameKind, same: r.same.map(p => p.sku)
}));
