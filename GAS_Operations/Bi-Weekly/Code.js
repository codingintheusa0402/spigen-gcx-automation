// TEMP debug — read-only structure dump for the Claim/Review card template
// (slides 9-41 of this deck) and for the two source decks. Delete once the
// slide-maker feature is built and verified.
function _describeTextRange(tr) {
  const txt = tr.asString();
  if (!txt) return '';
  const lineStarts = [0];
  for (let i = 0; i < txt.length; i++) if (txt[i] === '\n') lineStarts.push(i + 1);
  const perLine = lineStarts
    .filter(function(idx) { return idx < txt.length; })
    .map(function(idx) {
      const parts = [];
      let st;
      try { st = tr.getRange(idx, idx + 1).getTextStyle(); } catch (e) { return 'styleErr'; }
      try { parts.push(st.getFontFamily() + '/' + st.getFontSize()); } catch (e) {}
      try { if (st.isBold()) parts.push('bold'); } catch (e) {}
      try {
        const fg = st.getForegroundColor();
        if (fg) {
          try { parts.push(fg.asRgbColor().asHexString()); }
          catch (e2) { try { parts.push('theme:' + fg.asThemeColor().getThemeColorType()); } catch (e3) {} }
        }
      } catch (e) {}
      try {
        const link = st.getLink();
        if (link) parts.push('LINK=' + (link.getUrl() || link.getSlideId() || '?'));
      } catch (e) {}
      try {
        const pAlign = tr.getRange(idx, idx + 1).getParagraphStyle().getParagraphAlignment();
        parts.push('align=' + pAlign);
      } catch (e) {}
      return parts.join(' ');
    });
  return 'text="' + txt.replace(/\n/g, '\\n') + '" [' + perLine.join('] [') + ']';
}

function _describeEl(el, depth, lines) {
  const type = el.getPageElementType();
  const l = Math.round(el.getLeft()), t = Math.round(el.getTop());
  const w = Math.round(el.getWidth()), h = Math.round(el.getHeight());
  let line = '  '.repeat(depth) + String(type) + ' L' + l + ' T' + t + ' W' + w + ' H' + h;
  if (el.getTitle && el.getTitle()) line += ' title="' + el.getTitle() + '"';

  if (type === SlidesApp.PageElementType.GROUP) {
    lines.push(line);
    el.asGroup().getChildren().forEach(function(child) { _describeEl(child, depth + 1, lines); });
    return;
  }

  if (type === SlidesApp.PageElementType.SHAPE) {
    const shape = el.asShape();
    line += ' shape=' + shape.getShapeType();
    try {
      const sf = shape.getFill() && shape.getFill().getSolidFill();
      line += ' fill=' + (sf ? sf.getColor().asRgbColor().asHexString() : 'none');
    } catch (e) {}
    try {
      const border = shape.getBorder();
      if (border && !border.isTransparent()) {
        const bsf = border.getLineFill() && border.getLineFill().getSolidFill();
        line += ' border=' + (bsf ? bsf.getColor().asRgbColor().asHexString() : '?');
      }
    } catch (e) {}
    try {
      const link = shape.getLink();
      if (link) line += ' LINK=' + (link.getUrl() || link.getSlideId() || '?');
    } catch (e) {}
    line += ' ' + _describeTextRange(shape.getText());
  } else if (type === SlidesApp.PageElementType.IMAGE) {
    const img = el.asImage();
    try {
      const link = img.getLink();
      if (link) line += ' LINK=' + (link.getUrl() || link.getSlideId() || '?');
    } catch (e) {}
  } else if (type === SlidesApp.PageElementType.LINE) {
    // decorative divider lines — position only, already captured above
  } else if (type === SlidesApp.PageElementType.TABLE) {
    const table = el.asTable();
    line += ' rows=' + table.getNumRows() + ' cols=' + table.getNumColumns();
    lines.push(line);
    for (let r = 0; r < table.getNumRows(); r++) {
      for (let c = 0; c < table.getNumColumns(); c++) {
        try {
          const cellText = table.getCell(r, c).getText().asString().replace(/\n/g, '\\n');
          lines.push('  '.repeat(depth + 1) + 'cell[' + r + '][' + c + ']="' + cellText + '"');
        } catch (e) {
          lines.push('  '.repeat(depth + 1) + 'cell[' + r + '][' + c + ']=(merged, not head)');
        }
      }
    }
    return;
  }
  lines.push(line);
}

// Dumps into `lines` (does not log directly — caller writes lines to a Doc,
// since the editor's log panel is too small/awkward to scroll through for
// a dump this size).
function _dumpSlideInto(slide, slideIndex, lines) {
  lines.push('=== SLIDE index ' + slideIndex + ' id=' + slide.getObjectId() +
    ' pageElementCount=' + slide.getPageElements().length + ' ===');
  slide.getPageElements().forEach(function(el) { _describeEl(el, 0, lines); });
  lines.push('=== LAYOUT (' + slide.getLayout().getLayoutName() + ') ===');
  slide.getLayout().getPageElements().forEach(function(el) { _describeEl(el, 0, lines); });
}

// Writes `lines` to a fresh Google Doc and logs just its URL (short enough
// to read from the small execution-log panel), instead of logging every
// line directly.
function _dumpLinesToDoc(lines, title) {
  const doc = DocumentApp.create(title + ' ' + new Date().toISOString());
  const body = doc.getBody();
  body.setText(lines.join('\n'));
  doc.saveAndClose();
  Logger.log('DOC: ' + doc.getUrl());
  return doc.getUrl();
}

function debugDumpTemplate() {
  const pres = SlidesApp.openById('1qHQoYAOvmI-X1rQrRlbzxWtkFQtyNH1lqjpB2szG9vc');
  const lines = [];
  lines.push('TOTAL SLIDES: ' + pres.getSlides().length);
  lines.push('PAGE: ' + pres.getPageWidth() + ' x ' + pres.getPageHeight());
  // Sample across the 9-41 range (0-indexed 8-40) to catch different card
  // types (claim vs review, Zendesk vs Amazon, single vs multi-photo, etc.)
  [8, 9, 12, 16, 20, 24, 28, 32, 36, 40].forEach(function(idx) {
    if (idx < pres.getSlides().length) _dumpSlideInto(pres.getSlides()[idx], idx, lines);
  });
  _dumpLinesToDoc(lines, 'TemplateDump');
}

// Fast catalog pass: for every slide 9-41 (0-idx 8-40), log just the title
// text + any hyperlink host, to classify family section + claim/review/link
// type without a full element dump.
function debugSurveyTemplateFamilies() {
  const pres = SlidesApp.openById('1qHQoYAOvmI-X1rQrRlbzxWtkFQtyNH1lqjpB2szG9vc');
  const slides = pres.getSlides();
  const lines = [];
  for (let i = 8; i <= 40 && i < slides.length; i++) {
    const slide = slides[i];
    let titleText = '';
    let linkInfo = '';
    let headingText = '';
    slide.getPageElements().forEach(function(el) {
      const type = el.getPageElementType();
      const l = Math.round(el.getLeft()), t = Math.round(el.getTop());
      if (type === SlidesApp.PageElementType.SHAPE) {
        const shape = el.asShape();
        const txt = shape.getText().asString();
        if (l < 100 && t < 30 && txt) titleText = txt.replace(/\n/g, '');
        if (l > 400 && l < 600 && t > 60 && t < 90 && txt) headingText = txt.replace(/\n/g, '');
      } else if (type === SlidesApp.PageElementType.IMAGE) {
        try {
          const link = el.asImage().getLink();
          if (link && link.getUrl() && t < 90 && l > 300) {
            const url = link.getUrl();
            let host = 'other';
            if (url.indexOf('zendesk') !== -1) host = 'ZENDESK';
            else if (url.indexOf('docs.google.com/presentation') !== -1) host = 'SLIDES-LINK';
            else if (url.indexOf('amazon') !== -1) host = 'AMAZON';
            else if (url.indexOf('pstatic') !== -1 || url.indexOf('naver') !== -1) host = 'NAVER/PSTATIC';
            linkInfo = host + '(' + url.substring(0, 70) + ')';
          }
        } catch (e) {}
      }
    });
    lines.push('slide ' + (i + 1) + ' (idx ' + i + '): title="' + titleText + '" heading="' + headingText + '" link=' + linkInfo);
  }
  _dumpLinesToDoc(lines, 'TemplateFamilySurvey');
}

// Full raw dump of the first N real (non-blank) slides of a source deck,
// so we can see the actual table/shape structure (not the "템플릿 예시" row).
function debugDumpSourceDeck(deckId, label, startIdx, count) {
  const pres = SlidesApp.openById(deckId);
  const slides = pres.getSlides();
  const lines = [];
  lines.push(label + ' TOTAL SLIDES: ' + slides.length);
  for (let i = startIdx; i < Math.min(startIdx + count, slides.length); i++) {
    _dumpSlideInto(slides[i], i, lines);
  }
  return _dumpLinesToDoc(lines, label + 'SourceDump');
}

function debugDumpBothSourceDecks() {
  debugDumpSourceDeck('1VC5WAoiufinAPz9bPn1OrBnAef9JkDZEZxlAGF6DDho', 'GlxZ8', 2, 5);
  debugDumpSourceDeck('1JJKzzBnm9no89mocr6Xqzwqgz8YWoiU5S44Em7gJYSc', 'Pixel11', 2, 5);
}

// TEMP — trial run of the slide maker from the editor (2 cards max, one
// per deck), so the result can be checked before a full run. Delete after.
function debugTrialClaimSlides() {
  const pres = SlidesApp.openById('1qHQoYAOvmI-X1rQrRlbzxWtkFQtyNH1lqjpB2szG9vc');
  const opts = { sources: ['glxZ8', 'pixel11'], startDate: '2026-08-28', endDate: '2026-09-10', maxCount: 1 };
  Logger.log('PREVIEW: ' + JSON.stringify(_previewClaimSlides(pres, opts).sources.map(function(s) {
    return s.label + ': ' + s.newCount + ' new / ' + s.rows.length + ' in range, templates=' + s.existingCards;
  })));
  const res = _generateClaimSlides(pres, opts);
  Logger.log('RESULT: ' + JSON.stringify(res));
}

// TEMP: logs parsed header + media details of the first in-range row per source.
function debugInspectTrialRows() {
  ['glxZ8', 'pixel11'].forEach(function(key) {
    const rows = _listSourceRows(key, '2026-08-28', '2026-09-10');
    if (!rows.length) { Logger.log(key + ': no rows'); return; }
    const r = rows[0];
    const media = r.media.map(function(m) {
      const info = { kind: m.kind, w: Math.round(m.el.getWidth()), h: Math.round(m.el.getHeight()), top: Math.round(m.el.getTop()) };
      try { info.contentUrl = m.kind === 'image' ? (m.el.getContentUrl() || '').slice(0, 60) : (m.el.getThumbnailUrl() || '').slice(0, 60); } catch (e) { info.urlErr = e.message; }
      if (m.kind === 'image') { try { info.blobBytes = m.el.getBlob().getBytes().length; } catch (e) { info.blobErr = e.message; } }
      return info;
    });
    Logger.log(key + ': ' + JSON.stringify({ date: r.date, sku: r.sku, model: r.model, device: r.device, type: r.type, link: r.link, media: media }));
  });
}

// TEMP: removes the cards inserted by debugTrialClaimSlides (matched by button
// URL against the first in-range row of each source). Never touches idx <= 8.
function debugDeleteTrialSlides() {
  const pres = SlidesApp.openById('1qHQoYAOvmI-X1rQrRlbzxWtkFQtyNH1lqjpB2szG9vc');
  const removed = [];
  Object.keys(CLAIM_SLIDE_SOURCES).forEach(function(key) {
    const src = CLAIM_SLIDE_SOURCES[key];
    const rows = _listSourceRows(key, '2026-08-28', '2026-09-10');
    if (!rows.length) return;
    const urls = [rows[0].link, rows[0].sourceSlideUrl].filter(Boolean);
    _findFamilyCards(pres, src.familyMatch).forEach(function(c) {
      if (c.index > 8 && c.url && urls.indexOf(c.url) !== -1) { removed.push(key + '@' + c.index); c.slide.remove(); }
    });
  });
  pres.saveAndClose();
  Logger.log('REMOVED: ' + JSON.stringify(removed));
}

// Counts how many content slides in a source deck have 작성 날짜 within
// [startDateStr, endDateStr] (inclusive, 'YYYY-MM-DD'), read-only.
function debugCountInRange(deckId, label, startDateStr, endDateStr) {
  const pres = SlidesApp.openById(deckId);
  const slides = pres.getSlides();
  const start = new Date(startDateStr + 'T00:00:00');
  const end = new Date(endDateStr + 'T23:59:59');
  let matched = 0, total = 0, noDate = 0;
  const matchedDates = [];
  slides.forEach(function(slide, i) {
    const rows = _readSourceTable(slide);
    if (!rows || rows.length < 2) return; // no table = title/template slide
    total++;
    const dateStr = rows[1][0]; // row1, col0 = 작성 날짜
    const d = new Date(dateStr + 'T00:00:00');
    if (isNaN(d.getTime())) { noDate++; return; }
    if (d >= start && d <= end) { matched++; matchedDates.push('idx' + i + ':' + dateStr); }
  });
  Logger.log(label + ': total content slides=' + total + ' noDate=' + noDate +
    ' MATCHED[' + startDateStr + '..' + endDateStr + ']=' + matched);
  Logger.log(label + ' matched: ' + JSON.stringify(matchedDates));
}

function debugCountBothInRange() {
  debugCountInRange('1VC5WAoiufinAPz9bPn1OrBnAef9JkDZEZxlAGF6DDho', 'GlxZ8', '2026-08-28', '2026-09-10');
  debugCountInRange('1JJKzzBnm9no89mocr6Xqzwqgz8YWoiU5S44Em7gJYSc', 'Pixel11', '2026-08-28', '2026-09-10');
}

// ─── CLAIM / REVIEW SLIDE MAKER (sidebar) ─────────────────────────────────────
// Pulls dated rows out of the per-product "클레임 및 배드리뷰 고객사진 모음" decks
// and appends one Claims/Reviews card per row into this deck, in the exact
// visual style of the existing cards (slides 9-34).
//
// Source deck layout (one row per slide):
//   header shape  "■ SKU: {sku} | {model} | {device}"
//   TABLE 3x6     row0 = 작성 날짜/국가/인입사유/클레임/배드리뷰/Zendesk·Review link
//                 row1 = values, row2 = "클레임 상세내용" + detail text (merged)
//   IMAGE/VIDEO   customer photos below the table
//
// Strategy: duplicate the LAST existing card of the same family + type
// (claim = Zendesk-linked, review = Amazon/Naver-linked) so every icon,
// badge, colour and font is inherited, then swap only the data fields, the
// photo, and the "바로가기" button link. ASIN is looked up by SKU from the
// matching Amazon 1-3점 sheet (BAD_REVIEW_SOURCES), with the family sheet's
// "DE" tab as fallback; the 아마존 리뷰 평점 / 갯수 card comes from that DE tab
// (Score / Global Ratings per SKU) and stays blank when it has no score.
// Re-runs are idempotent: rows already present in the deck are skipped.

const CLAIM_SLIDE_SOURCES = {
  glxZ8:   { id: '1VC5WAoiufinAPz9bPn1OrBnAef9JkDZEZxlAGF6DDho', label: 'Galaxy Z8 Series', familyMatch: 'Galaxy Z8', badReviewKey: 'glxZ8' },
  pixel11: { id: '1JJKzzBnm9no89mocr6Xqzwqgz8YWoiU5S44Em7gJYSc', label: 'Pixel 11 Series',  familyMatch: 'Pixel 11', badReviewKey: 'pixel11' }
};

// Card field positions (pt) measured off the live template. Matching is by
// position because these cards carry real values, not {{placeholders}}.
const CLAIM_CARD_FIELDS = {
  title:   { l: [60, 110],  t: [5, 35] },
  product: { l: [80, 100],  t: [50, 75] },
  date:    { l: [380, 405], t: [55, 100] },
  heading: { l: [485, 510], t: [65, 90] },
  content: { l: [485, 510], t: [95, 120] },
  country: { l: [485, 510], t: [160, 190] },
  reason:  { l: [485, 510], t: [205, 235] },
  asin:    { l: [485, 510], t: [250, 280] },
  sku:     { l: [485, 510], t: [295, 325] },
  rating:  { l: [485, 510], t: [340, 370] }
};

function showSlideMakerSidebar() {
  const html = HtmlService.createHtmlOutputFromFile('SlideMaker')
    .setTitle('Claim / Review Slide Maker');
  SlidesApp.getUi().showSidebar(html);
}

// Finds the (first) TABLE on a slide and returns its cells as trimmed
// strings; merged non-head cells come back as ''.
function _readSourceTable(slide) {
  let table = null;
  slide.getPageElements().forEach(function(el) {
    if (!table && el.getPageElementType() === SlidesApp.PageElementType.TABLE) table = el.asTable();
  });
  if (!table) return null;
  const rows = [];
  for (let r = 0; r < table.getNumRows(); r++) {
    const row = [];
    for (let c = 0; c < table.getNumColumns(); c++) {
      try { row.push(table.getCell(r, c).getText().asString().trim()); }
      catch (e) { row.push(''); }
    }
    rows.push(row);
  }
  return rows;
}

// "■ SKU: ACS11561  |  Air Skin Aramid MagFit  |  Glx. Z Fold 8"
function _parseSourceHeader(text) {
  // Source headers separate fields with ASCII '|', fullwidth/box-drawing bars,
  // or — most often — the Hangul letter ㅣ (U+3163) typed as a bar.
  const parts = String(text || '').replace(/^[\s■]+/, '').split(/[|｜│┃ㅣ]/).map(function(s) { return s.trim(); });
  return {
    sku:    (parts[0] || '').replace(/^SKU\s*:\s*/i, '').trim(),
    model:  parts[1] || '',
    device: parts[2] || ''
  };
}

function _normalizeDate(s) {
  const m = String(s || '').match(/(\d{4})[.\-\/]\s*(\d{1,2})[.\-\/]\s*(\d{1,2})/);
  if (!m) return null;
  const mm = ('0' + m[2]).slice(-2), dd = ('0' + m[3]).slice(-2);
  return { iso: m[1] + '-' + mm + '-' + dd, dotted: m[1] + '.' + mm + '.' + dd };
}

function _isChecked(glyph) {
  return /[☑✓✔■◼]/.test(String(glyph || ''));
}

// Extracts one row object from a source content slide, or null if the
// slide has no data table (title / template-example slides).
function _readSourceSlide(slide, sourceKey, deckId, index) {
  const rows = _readSourceTable(slide);
  if (!rows || rows.length < 2 || !rows[1][0]) return null;
  const date = _normalizeDate(rows[1][0]);
  if (!date) return null;

  let header = { sku: '', model: '', device: '' };
  let link = rows[1][5] || '';
  const media = [];
  slide.getPageElements().forEach(function(el) {
    const type = el.getPageElementType();
    if (type === SlidesApp.PageElementType.SHAPE) {
      const txt = el.asShape().getText().asString();
      if (/SKU\s*:/i.test(txt) && el.getTop() < 40) header = _parseSourceHeader(txt);
    } else if (type === SlidesApp.PageElementType.IMAGE && el.getTop() > 100) {
      media.push({ kind: 'image', el: el.asImage() });
    } else if (type === SlidesApp.PageElementType.VIDEO) {
      media.push({ kind: 'video', el: el.asVideo() });
    }
  });
  // Link cell may hold a hyperlink whose display text isn't the URL.
  if (!/^https?:\/\//i.test(link)) {
    try {
      let table = null;
      slide.getPageElements().forEach(function(el) {
        if (!table && el.getPageElementType() === SlidesApp.PageElementType.TABLE) table = el.asTable();
      });
      const runs = table.getCell(1, 5).getText().getRuns();
      for (let i = 0; i < runs.length; i++) {
        const l = runs[i].getTextStyle().getLink();
        if (l && l.getUrl()) { link = l.getUrl(); break; }
      }
    } catch (e) {}
  }
  const isReview = link ? !/zendesk/i.test(link) : (_isChecked(rows[1][4]) && !_isChecked(rows[1][3]));

  return {
    sourceKey: sourceKey,
    sourceIndex: index,
    sourceSlideUrl: 'https://docs.google.com/presentation/d/' + deckId + '/edit#slide=id.' + slide.getObjectId(),
    date: date,
    country: rows[1][1] || '',
    reason: rows[1][2] || '',
    detail: rows.length > 2 ? (rows[2][1] || '') : '',
    link: link,
    type: isReview ? 'review' : 'claim',
    sku: header.sku, model: header.model, device: header.device,
    media: media
  };
}

function _listSourceRows(sourceKey, startIso, endIso) {
  const src = CLAIM_SLIDE_SOURCES[sourceKey];
  if (!src) throw new Error('Unknown source: ' + sourceKey);
  const start = new Date(startIso + 'T00:00:00'), end = new Date(endIso + 'T23:59:59');
  const slides = SlidesApp.openById(src.id).getSlides();
  const rows = [];
  slides.forEach(function(slide, i) {
    const row = _readSourceSlide(slide, sourceKey, src.id, i);
    if (!row) return;
    const d = new Date(row.date.iso + 'T00:00:00');
    if (d >= start && d <= end) rows.push(row);
  });
  return rows;
}

function _shapeInBox(el, box) {
  const l = el.getLeft(), t = el.getTop();
  return l >= box.l[0] && l <= box.l[1] && t >= box.t[0] && t <= box.t[1];
}

// Maps a card's page elements to field roles by position.
function _getCardShapes(slide) {
  const out = { photos: [], button: null };
  slide.getPageElements().forEach(function(el) {
    const type = el.getPageElementType();
    if (type === SlidesApp.PageElementType.SHAPE) {
      const shape = el.asShape();
      Object.keys(CLAIM_CARD_FIELDS).forEach(function(role) {
        if (!out[role] && _shapeInBox(el, CLAIM_CARD_FIELDS[role])) out[role] = shape;
      });
    } else if (type === SlidesApp.PageElementType.IMAGE) {
      const img = el.asImage();
      let hasLink = false;
      try { hasLink = !!(img.getLink() && img.getLink().getUrl()); } catch (e) {}
      if (hasLink && el.getTop() < 100 && el.getLeft() > 300) out.button = img;
      else if (_isPhotoSlot(el)) out.photos.push(img); // may carry a link (video thumbnails)
    }
  });
  out.photos.sort(function(a, b) { return a.getLeft() - b.getLeft(); }); // left → right
  return out;
}

// The card background panel (~600x330 image at L90/T60) and a few small
// decorative icons are IMAGEs too, so a photo slot is recognised by sitting
// inside the photo area at a plausible photo size.
const CLAIM_PHOTO_AREA = { l: [95, 480], t: [110, 360], minSide: 60, maxSide: 330 };
function _isPhotoSlot(el) {
  const w = el.getWidth(), h = el.getHeight();
  return _shapeInBox(el, CLAIM_PHOTO_AREA) &&
    w >= CLAIM_PHOTO_AREA.minSide && w <= CLAIM_PHOTO_AREA.maxSide &&
    h >= CLAIM_PHOTO_AREA.minSide && h <= CLAIM_PHOTO_AREA.maxSide;
}

function _cardText(shape) {
  return shape ? shape.getText().asString().replace(/\n+$/, '').trim() : '';
}

// Existing cards of one family: [{ index, slide, type, key }], in deck order.
function _findFamilyCards(presentation, familyMatch) {
  const cards = [];
  presentation.getSlides().forEach(function(slide, i) {
    const s = _getCardShapes(slide);
    if (!s.title || !s.heading) return;
    if (_cardText(s.title).indexOf(familyMatch) === -1) return;
    const heading = _cardText(s.heading);
    if (heading.indexOf('클레임') === -1 && heading.indexOf('리뷰') === -1) return;
    let url = '';
    try { url = s.button ? (s.button.getLink().getUrl() || '') : ''; } catch (e) {}
    cards.push({
      index: i, slide: slide, title: _cardText(s.title),
      type: heading.indexOf('리뷰') !== -1 ? 'review' : 'claim',
      url: url,
      key: [_cardText(s.date), _cardText(s.sku), _cardText(s.country), _cardText(s.reason)].join('|')
    });
  });
  return cards;
}

function _rowKey(row) {
  return [row.date.dotted, row.sku, row.country, row.reason].join('|');
}

// SKU -> ASIN from the matching Amazon 1-3점 sheet, cached 30 min.
// SKU -> { asin, score, count } from the family sheet's "DE" tab (one row per
// ASIN with Amazon Score / Global Ratings); feeds the ASIN fallback and the
// 아마존 리뷰 평점 / 갯수 card. Cached 30 min.
function _deInfoMapForSource(badReviewKey) {
  const cache = CacheService.getScriptCache();
  const cacheKey = 'deInfo_' + badReviewKey;
  const cached = cache.get(cacheKey);
  if (cached) return JSON.parse(cached);
  const src = BAD_REVIEW_SOURCES[badReviewKey];
  const map = {};
  try {
    const sheet = SpreadsheetApp.openById(src.id).getSheetByName('DE');
    const values = sheet.getRange(1, 1, sheet.getLastRow(), Math.min(sheet.getLastColumn(), 12)).getDisplayValues();
    let hi = -1;
    for (let i = 0; i < Math.min(values.length, 6); i++) if (values[i].indexOf('SKU') !== -1 && values[i].indexOf('ASIN') !== -1) { hi = i; break; }
    if (hi === -1) throw new Error('header row with SKU/ASIN not found');
    const h = values[hi];
    const ix = { sku: h.indexOf('SKU'), asin: h.indexOf('ASIN'), score: h.indexOf('Score'), count: h.indexOf('Global Ratings') };
    for (let r = hi + 1; r < values.length; r++) {
      const row = values[r];
      const sku = String(row[ix.sku] || '').trim();
      if (!sku || map[sku]) continue;
      map[sku] = { asin: String(row[ix.asin] || '').trim(), score: String(ix.score >= 0 ? row[ix.score] || '' : '').trim(),
                   count: String(ix.count >= 0 ? row[ix.count] || '' : '').trim() };
    }
  } catch (e) {
    Logger.log('DE map error (' + badReviewKey + '): ' + e.message);
  }
  try { cache.put(cacheKey, JSON.stringify(map), 1800); } catch (e) {}
  return map;
}

// "4.5점  Global Ratings: 1,562" in the deck's own format; '' when the DE
// sheet has no score for the SKU.
function _ratingText(info) {
  if (!info || !info.score) return '';
  const n = parseFloat(info.score);
  const score = isNaN(n) ? info.score : n.toFixed(1);
  const c = parseFloat(String(info.count).replace(/,/g, ''));
  const count = isNaN(c) ? info.count : Math.round(c).toString().replace(/\B(?=(\d{3})+(?!\d))/g, ',');
  return count ? score + '점  Global Ratings: ' + count : score + '점';
}

// ── SIREN badge ─────────────────────────────────────────────────────────
// Cards whose SKU + 인입사유 match a registered row (SIREN 등록 = O) of the
// 26년 SIREN sheet get a red "SIREN 등록됨" chip linked to that row's SIREN
// review deck (the C-column title, looked up in Drive by name).
const SIREN_SHEET = { id: '15Jh6ZFDBIbpv4OANVtD3g4wFBJxoof9SHWDUEU3GiXI', name: '26년 SIREN', gid: 1840076165, headerRow: 18 };
const SIREN_BADGE = { l: 391, t: 63.8, w: 81, h: 14, fill: '#4A1426', text: '#FF5252', label: 'SIREN 등록됨' };

function _sirenEntries() {
  const cache = CacheService.getScriptCache();
  const cached = cache.get('sirenEntries');
  if (cached) return JSON.parse(cached);
  const out = [];
  try {
    const sheet = SpreadsheetApp.openById(SIREN_SHEET.id).getSheetByName(SIREN_SHEET.name);
    const first = SIREN_SHEET.headerRow + 1;
    const n = sheet.getLastRow() - first + 1;
    const rows = n > 0 ? sheet.getRange(first, 1, n, 11).getDisplayValues() : [];
    rows.forEach(function(r, i) {
      const skus = String(r[4] || '').split(/[\s,\/]+/).map(function(s) { return s.trim().toUpperCase(); }).filter(Boolean);
      if (!skus.length) return;
      const title = String(r[2] || '').trim();
      out.push({ row: first + i, skus: skus, issue: String(r[5] || '').trim(), title: title,
                 registered: String(r[7] || '').trim().toUpperCase() === 'O', link: _sirenDeckUrl(title) ||
                 ('https://docs.google.com/spreadsheets/d/' + SIREN_SHEET.id + '/edit#gid=' + SIREN_SHEET.gid + '&range=C' + (first + i)) });
    });
  } catch (e) { Logger.log('SIREN sheet error: ' + e.message); }
  try { cache.put('sirenEntries', JSON.stringify(out), 1800); } catch (e) {}
  return out;
}

// The SIREN review decks are Google Slides files named exactly like column C.
function _sirenDeckUrl(title) {
  if (!title) return '';
  try {
    let it = DriveApp.getFilesByName(title);
    if (!it.hasNext()) it = DriveApp.searchFiles("title contains '" + title.slice(0, 30).replace(/'/g, "\\'") + "' and trashed = false");
    if (it.hasNext()) return it.next().getUrl();
  } catch (e) {}
  return '';
}

const _SIREN_STOP = { '이슈': 1, '불량': 1, '문제': 1, '발생': 1, '현상': 1 };
function _sirenNorm(s) { return String(s || '').toLowerCase().replace(/[\s\(\)\[\]\-_\/,.:·]+/g, '').replace(/이슈/g, ''); }
function _sirenTokens(s) {
  return String(s || '').toLowerCase().split(/[\s\(\)\[\]\-_\/,.:·]+/).filter(function(t) { return t.length >= 2 && !_SIREN_STOP[t]; });
}
// "very similar" 인입사유: equal / substring / shared keyword / bigram overlap.
function _issueMatches(a, b) {
  const na = _sirenNorm(a), nb = _sirenNorm(b);
  if (!na || !nb) return false;
  if (na === nb || (na.length >= 2 && nb.indexOf(na) !== -1) || (nb.length >= 2 && na.indexOf(nb) !== -1)) return true;
  if (_sirenTokens(a).some(function(t) { return nb.indexOf(t) !== -1; })) return true;
  if (_sirenTokens(b).some(function(t) { return na.indexOf(t) !== -1; })) return true;
  const bg = function(s) { const o = {}; for (let i = 0; i < s.length - 1; i++) o[s.substr(i, 2)] = 1; return o; };
  const A = bg(na), B = bg(nb), ka = Object.keys(A), kb = Object.keys(B);
  const both = ka.filter(function(k) { return B[k]; }).length;
  return ka.length && kb.length && (2 * both) / (ka.length + kb.length) >= 0.5;
}

function _findSiren(entries, sku, reason) {
  const key = String(sku || '').trim().toUpperCase();
  for (let i = 0; i < entries.length; i++) {
    const e = entries[i];
    if (e.skus.indexOf(key) !== -1 && _issueMatches(reason, e.issue)) return e;
  }
  return null;
}

function _slideHasSirenBadge(slide) {
  return slide.getShapes().some(function(sh) { try { return /SIREN/.test(sh.getText().asString()); } catch (e) { return false; } });
}

function _addSirenBadge(slide, url) {
  const b = SIREN_BADGE;
  const sh = slide.insertShape(SlidesApp.ShapeType.ROUND_RECTANGLE, b.l, b.t, b.w, b.h);
  sh.getFill().setSolidFill(b.fill);
  sh.getBorder().setTransparent();
  sh.setContentAlignment(SlidesApp.ContentAlignment.MIDDLE);
  const tr = sh.getText();
  tr.setText(b.label);
  tr.getTextStyle().setFontFamily('Roboto Mono').setFontSize(7.5).setBold(true).setForegroundColor(b.text);
  tr.getParagraphStyle().setParagraphAlignment(SlidesApp.ParagraphAlignment.CENTER);
  if (url) sh.setLinkUrl(url);
  return sh;
}

// Menu action: badge every existing Z8 / Pixel 11 card that matches the SIREN
// sheet and doesn't carry a badge yet. Safe to re-run.
function applySirenBadges() {
  const presentation = SlidesApp.getActivePresentation();
  const entries = _sirenEntries();
  let added = 0, skippedX = 0;
  _findFamilyCards(presentation, '').forEach(function(c) {
    if (!/Galaxy Z8|Pixel 11/.test(c.title)) return;
    const s = _getCardShapes(c.slide);
    const m = _findSiren(entries, _cardText(s.sku), _cardText(s.reason));
    if (!m) return;
    if (!m.registered) { skippedX++; return; }
    if (_slideHasSirenBadge(c.slide)) return;
    _addSirenBadge(c.slide, m.link); added++;
  });
  SlidesApp.getUi().alert('SIREN badges added: ' + added + (skippedX ? '\n(matched but SIREN 등록 = X, no badge: ' + skippedX + ')' : ''));
}

function _asinMapForSource(badReviewKey) {
  const cache = CacheService.getScriptCache();
  const cacheKey = 'asinMap_' + badReviewKey;
  const cached = cache.get(cacheKey);
  if (cached) return JSON.parse(cached);
  const src = BAD_REVIEW_SOURCES[badReviewKey];
  const map = {};
  try {
    const sheet = SpreadsheetApp.openById(src.id).getSheetByName(src.sheetName);
    const rowCount = Math.max(sheet.getLastRow() - 1, 0);
    const skuCol = getColumnIndexByHeader(sheet, 'SKU');
    const asinCol = getColumnIndexByHeader(sheet, 'ASIN');
    const skus = rowCount ? sheet.getRange(2, skuCol, rowCount, 1).getValues().flat() : [];
    const asins = rowCount ? sheet.getRange(2, asinCol, rowCount, 1).getValues().flat() : [];
    for (let i = 0; i < skus.length; i++) {
      const k = String(skus[i]).trim(), a = String(asins[i]).trim();
      if (k && a && !map[k]) map[k] = a;
    }
  } catch (e) {
    Logger.log('ASIN map error (' + badReviewKey + '): ' + e.message);
  }
  try { cache.put(cacheKey, JSON.stringify(map), 1800); } catch (e) {}
  return map;
}

// Replaces a shape's text while keeping its formatting: swapping the old
// value in place preserves the run's font/colour; setText is only the
// fallback for an empty shape.
function _setCardText(shape, newText, autofit) {
  if (!shape) return;
  const tr = shape.getText();
  const old = _cardText(shape);
  if (old) tr.replaceAllText(old, newText); else tr.setText(newText);
  if (autofit) { try { shape.getAutofit().setAutofitType(SlidesApp.AutofitType.TEXT_AUTOFIT); } catch (e) {} }
}

function _fillCard(slide, row, asin, presId, ratingText) {
  const s = _getCardShapes(slide);
  _setCardText(s.product, [row.model, row.device].filter(Boolean).join(' ㅣ'), true); // template style: "Air Skin MagFit ㅣGlx. Z Fold 8"
  _setCardText(s.date, row.date.dotted);
  _setCardText(s.content, row.detail, true);
  _setCardText(s.country, row.country);
  _setCardText(s.reason, row.reason, true);
  _setCardText(s.asin, asin || '');
  _setCardText(s.sku, row.sku);
  _setCardText(s.rating, ratingText || '');
  if (s.button) s.button.setLinkUrl(row.link || row.sourceSlideUrl);
  // Autofit set via the API isn't computed until someone edits the box in
  // the UI, so shrink long values ourselves (inner widths ≈ box − padding).
  _shrinkToFit(s.product, 240, 18, 8, 1.4);   // one line
  _shrinkToFit(s.content, 170, 60, 6, 1.5);   // room above the 국가 box
  _shrinkToFit(s.reason, 160, 30, 8, 1.4);
  _shrinkToFit(s.sku, 160, 30, 8, 1.4);

  // Photos: fill template slots left→right from the source media, drop any
  // template photos left over so a stale photo never survives.
  const notes = [];
  let slots = s.photos;
  if (row.media.length === 1 && slots.length > 1) {
    // Single photo: keep the bigger slot, centred across the photo area.
    const areaL = Math.min.apply(null, slots.map(function(p) { return p.getLeft(); }));
    const areaR = Math.max.apply(null, slots.map(function(p) { return p.getLeft() + p.getWidth(); }));
    const keep = slots.reduce(function(a, b) { return b.getWidth() * b.getHeight() > a.getWidth() * a.getHeight() ? b : a; });
    slots.forEach(function(p) { if (p !== keep) p.remove(); });
    keep.setLeft(areaL + (areaR - areaL - keep.getWidth()) / 2);
    slots = [keep];
  }
  slots.forEach(function(img, i) {
    const m = row.media[i];
    if (!m) { img.remove(); return; }
    try { _replaceSlotImage(img, m, presId); }
    catch (e) {
      // Never leave the template's own photo on a new card.
      try { img.remove(); } catch (e2) {}
      notes.push('photo ' + (i + 1) + ' (' + m.kind + '): ' + e.message + ' — slot left empty');
    }
  });
  if (row.media.length > slots.length) notes.push((row.media.length - slots.length) + ' extra source photo(s) not placed');
  return notes;
}

// Swaps the slot's picture for the source media, cropping to fill the slot
// (replace(url, true)) so the card keeps its layout. Falls back to a plain
// blob replace (fit-inside) when the URL route is unavailable.
function _replaceSlotImage(img, m, presId) {
  let url = '', blob = null;
  if (m.kind === 'video') {
    // Videos can't be embedded from Apps Script (insertVideo is YouTube-only),
    // so the slot shows the video's thumbnail, linked to the video.
    try { url = m.el.getThumbnailUrl() || ''; } catch (e) {}
    if (!url && m.el.getSource() === SlidesApp.VideoSourceType.DRIVE) {
      const t = _driveThumbnail(m.el.getVideoId());
      url = t.url; blob = t.blob;
    }
    if (!url && !blob) throw new Error('video thumbnail unavailable');
  } else {
    try { url = m.el.getContentUrl() || m.el.getSourceUrl() || ''; } catch (e) {}
  }
  _swapPicture(img, url, blob, m.kind === 'image' ? m.el : null, presId);
  if (m.kind === 'video') { try { img.setLinkUrl(m.el.getUrl()); } catch (e) {} }
}

// Puts a new picture into an existing image element, keeping the element's
// exact box. Slides API replaceImage CENTER_CROP is the only route that
// scales-and-crops without moving the box (Image.replace(url, true) resizes
// it) — but the API can't see elements SlidesApp hasn't flushed yet, so URL
// swaps are queued and applied after saveAndClose() (see _flushCropQueue).
// Blob-only sources are placed immediately, fitted inside the box.
const _CROP_QUEUE = [];
function _swapPicture(img, url, blob, srcImage, presId) {
  if (url) { _CROP_QUEUE.push({ id: img.getObjectId(), url: url, blob: blob }); return; }
  if (blob) { img.replace(blob); return; }
  if (srcImage) { img.replace(srcImage.getBlob()); return; }
  throw new Error('no picture source');
}

function _flushCropQueue(presId, notes) {
  const items = _CROP_QUEUE.splice(0);
  if (!items.length) return;
  let pres = null;
  const fallback = function(it) {
    try {
      pres = pres || SlidesApp.openById(presId);
      const el = pres.getPageElementById(it.id);
      if (!el) return;
      if (it.blob) el.asImage().replace(it.blob); else el.asImage().replace(it.url);
    } catch (e) { notes.push('photo could not be placed (' + e.message + ')'); }
  };
  const req = function(it) { return { replaceImage: { imageObjectId: it.id, url: it.url, imageReplaceMethod: 'CENTER_CROP' } }; };
  for (let i = 0; i < items.length; i += 40) {
    const chunk = items.slice(i, i + 40);
    try { Slides.Presentations.batchUpdate({ requests: chunk.map(req) }, presId); }
    catch (e) {
      // The batch is atomic; redo this chunk one by one, fitting inside on failure.
      chunk.forEach(function(it) {
        try { Slides.Presentations.batchUpdate({ requests: [req(it)] }, presId); }
        catch (e2) { fallback(it); }
      });
    }
  }
  if (pres) pres.saveAndClose();
}

// Drive's own preview frame for a video file: a short-lived public link
// (usable by replaceImage) plus the bytes as a fallback. DriveApp's
// getThumbnail() returns null for most videos, hence the REST call.
function _driveThumbnail(fileId) {
  const out = { url: '', blob: null };
  try {
    const auth = { headers: { Authorization: 'Bearer ' + ScriptApp.getOAuthToken() }, muteHttpExceptions: true };
    const meta = UrlFetchApp.fetch('https://www.googleapis.com/drive/v3/files/' + encodeURIComponent(fileId) +
      '?fields=thumbnailLink&supportsAllDrives=true', auth);
    if (meta.getResponseCode() !== 200) return out;
    const link = JSON.parse(meta.getContentText()).thumbnailLink;
    if (!link) return out;
    out.url = link.replace(/=s\d+(-[a-z]+)?$/, '=s1600');
    const res = UrlFetchApp.fetch(out.url, auth);
    if (res.getResponseCode() === 200) out.blob = res.getBlob();
  } catch (e) {}
  return out;
}

// Reduces the font size of a single-style text box until an estimated
// line-wrap fits in maxW × maxH points (CJK glyphs ≈ 1em, others ≈ 0.55em).
function _shrinkToFit(shape, maxW, maxH, minSize, lineK) {
  if (!shape) return;
  lineK = lineK || 1.5;
  const tr = shape.getText();
  const text = tr.asString().replace(/\n+$/, '');
  if (!text) return;
  let size = null;
  try { size = tr.getTextStyle().getFontSize(); } catch (e) {}
  if (!size) { try { size = tr.getRange(0, 1).getTextStyle().getFontSize(); } catch (e) {} }
  if (!size) return;
  const units = Array.from(text).reduce(function(sum, ch) {
    return sum + (/[ᄀ-ᇿ㄰-㆏가-힯　-ヿ一-鿿＀-￯]/.test(ch) ? 1 : 0.55);
  }, 0);
  for (let s = size; s >= minSize; s -= 0.5) {
    const lines = text.split('\n').reduce(function(n, para) {
      const u = Array.from(para).reduce(function(sum, ch) { return sum + (/[ᄀ-ᇿ㄰-㆏가-힯　-ヿ一-鿿＀-￯]/.test(ch) ? 1 : 0.55); }, 0);
      return n + Math.max(1, Math.ceil(u / Math.max(1, maxW / s)));
    }, 0);
    if (lines * s * lineK <= maxH) { if (s !== size) tr.getTextStyle().setFontSize(s); return; }
  }
  tr.getTextStyle().setFontSize(minSize);
}

function _indexOfSlide(presentation, objectId) {
  const slides = presentation.getSlides();
  for (let i = 0; i < slides.length; i++) if (slides[i].getObjectId() === objectId) return i;
  return -1;
}

// Moves `slide` to sit immediately after `afterSlide`, verifying the result
// (Slide.move index semantics are easy to get off by one).
function _placeAfter(presentation, slide, afterSlide) {
  for (let attempt = 0; attempt < 3; attempt++) {
    const want = _indexOfSlide(presentation, afterSlide.getObjectId()) + 1;
    const have = _indexOfSlide(presentation, slide.getObjectId());
    if (have === want) return;
    slide.move(have < want ? want + 1 : want);
  }
}

// opts = { sources: ['glxZ8','pixel11'], startDate: 'YYYY-MM-DD', endDate: 'YYYY-MM-DD' }
function previewClaimSlides(opts) {
  const presentation = SlidesApp.getActivePresentation();
  return _previewClaimSlides(presentation, opts);
}

function _previewClaimSlides(presentation, opts) {
  const result = { sources: [] };
  (opts.sources || []).forEach(function(key) {
    const src = CLAIM_SLIDE_SOURCES[key];
    const existing = _findFamilyCards(presentation, src.familyMatch);
    const seenUrl = {}, seenKey = {};
    existing.forEach(function(c) { if (c.url) seenUrl[c.url] = true; seenKey[c.key] = true; });
    const rows = _listSourceRows(key, opts.startDate, opts.endDate);
    const items = rows.map(function(r) {
      const dup = (r.link && seenUrl[r.link]) || seenKey[_rowKey(r)];
      return {
        date: r.date.dotted, sku: r.sku, model: r.model, device: r.device,
        country: r.country, reason: r.reason, type: r.type,
        photos: r.media.length, hasLink: !!r.link, alreadyInDeck: !!dup
      };
    });
    result.sources.push({
      key: key, label: src.label,
      existingCards: existing.length,
      templateFound: existing.length > 0,
      rows: items,
      newCount: items.filter(function(i) { return !i.alreadyInDeck; }).length
    });
  });
  return result;
}

// Same opts as preview, plus optional maxCount (for a small trial run).
function generateClaimSlides(opts) {
  const presentation = SlidesApp.getActivePresentation();
  return _generateClaimSlides(presentation, opts);
}

function _generateClaimSlides(presentation, opts) {
  const startedAt = Date.now();
  const TIME_BUDGET_MS = 300 * 1000; // stay under the 6-min execution cap
  const summary = { inserted: 0, skippedExisting: 0, remaining: 0, notes: [], perSource: [] };
  const allCards = _findFamilyCards(presentation, ''); // every family, for cross-family fallback

  (opts.sources || []).forEach(function(key) {
    const src = CLAIM_SLIDE_SOURCES[key];
    let budgetLeft = Number(opts.maxCount) > 0 ? Number(opts.maxCount) : Infinity; // per source deck
    const cards = allCards.filter(function(c) { return c.title.indexOf(src.familyMatch) !== -1; });
    if (!cards.length) { summary.notes.push(src.label + ': no existing card found to use as template — skipped'); return; }

    const seenUrl = {}, seenKey = {};
    cards.forEach(function(c) { if (c.url) seenUrl[c.url] = true; seenKey[c.key] = true; });
    const templateByType = {};
    cards.forEach(function(c) { templateByType[c.type] = c; }); // last of each type wins
    let afterSlide = cards[cards.length - 1].slide;

    const asinMap = _asinMapForSource(src.badReviewKey);
    const deMap = _deInfoMapForSource(src.badReviewKey);
    const sirenEntries = _sirenEntries();
    const rows = _listSourceRows(key, opts.startDate, opts.endDate);
    let inserted = 0, skipped = 0, remaining = 0;

    rows.forEach(function(row) {
      if ((row.link && seenUrl[row.link]) || seenKey[_rowKey(row)]) { skipped++; return; }
      if (budgetLeft <= 0 || Date.now() - startedAt > TIME_BUDGET_MS) { remaining++; return; }

      // Template preference: same family + same type; else same type from
      // another family (background/badge/pill match, only the title differs);
      // else the other type within the family (heading text patched).
      let template = templateByType[row.type], crossFamily = false;
      if (!template) {
        const alt = allCards.filter(function(c) { return c.type === row.type; });
        if (alt.length) { template = alt[alt.length - 1]; crossFamily = true; }
        else template = templateByType[row.type === 'claim' ? 'review' : 'claim'];
      }
      const dup = template.slide.duplicate();
      _placeAfter(presentation, dup, afterSlide);
      if (crossFamily) {
        _setCardText(_getCardShapes(dup).title, cards[0].title);
      } else if (template.type !== row.type) {
        const s = _getCardShapes(dup);
        _setCardText(s.heading, row.type === 'review' ? '리뷰 내용' : '클레임 내용');
        summary.notes.push(src.label + ' ' + row.date.dotted + ' ' + row.sku + ': no ' + row.type + ' card in deck, used a ' + template.type + ' card');
      }
      const de = deMap[row.sku];
      const asin = asinMap[row.sku] || (de && de.asin) || '';
      const rating = _ratingText(de);
      const notes = _fillCard(dup, row, asin, presentation.getId(), rating);
      notes.forEach(function(n) { summary.notes.push(src.label + ' ' + row.date.dotted + ' ' + row.sku + ': ' + n); });
      if (!asin) summary.notes.push(src.label + ' ' + row.date.dotted + ' ' + row.sku + ': ASIN not found (1-3점 / DE sheets), left blank');
      if (!rating) summary.notes.push(src.label + ' ' + row.date.dotted + ' ' + row.sku + ': no Score in DE sheet, 평점/갯수 left blank');
      const siren = _findSiren(sirenEntries, row.sku, row.reason);
      if (siren && siren.registered) {
        if (!_slideHasSirenBadge(dup)) _addSirenBadge(dup, siren.link);
        summary.notes.push(src.label + ' ' + row.date.dotted + ' ' + row.sku + ': SIREN 등록됨 badge (sheet row ' + siren.row + ')');
      } else if (siren) {
        summary.notes.push(src.label + ' ' + row.date.dotted + ' ' + row.sku + ': in SIREN sheet (row ' + siren.row + ') but SIREN 등록 = X — no badge');
      }

      afterSlide = dup;
      seenKey[_rowKey(row)] = true;
      if (row.link) seenUrl[row.link] = true;
      inserted++; budgetLeft--;
    });

    summary.perSource.push({ label: src.label, inserted: inserted, skippedExisting: skipped, remaining: remaining });
    summary.inserted += inserted; summary.skippedExisting += skipped; summary.remaining += remaining;
  });

  presentation.saveAndClose();
  _flushCropQueue(presentation.getId(), summary.notes);
  return summary;
}

// ─── SELF-CONTAINED CHART CARD BUILDER ────────────────────────────────────────
// Draws a full TOP3 card grid (2 rows x 3 cards, dark-navy arc-chart cards on a
// light background) from scratch on a brand-new slide — no pre-existing
// placeholder shapes required. Geometry/colors/fonts below were measured
// directly off the live template (slide 4 of the real deck) so the output
// matches it exactly: card size, arc chart colors/shape, count/title/legend
// fonts and colors, rank-number styling.
const CHART_CARD_GEOM = {
  cardW: 160, cardH: 123,
  count:  { l: 10,  t: 30, w: 135, h: 29, fontSize: 13.5, color: '#FFFFFF', align: 'CENTER', valign: 'MIDDLE' },
  title:  { l: 10,  t: 47, w: 141, h: 31, fontSize: 8.5,  color: '#7680A2', align: 'CENTER', valign: 'MIDDLE' },
  legend: { l: 23,  t: 59, w: 90,  h: 62, fontSize: 8.5,  color: '#FFFFFF', align: 'START',  valign: 'TOP' },
  value:  { l: 118, t: 59, w: 40,  h: 62, fontSize: 8.5,  color: '#FFFFFF', align: 'END',    valign: 'TOP' },
  rank:   { l: -22, t: 98, w: 26,  h: 33, fontSize: 11,   color: '#7278B2', align: 'START',  valign: 'MIDDLE' }
};
const CHART_GRID_COL_LEFTS = [142, 326, 509];
const CHART_GRID_HEADER_LEFT = 139;
const CHART_GRID_BG_COLOR = '#F1F1F3';

// Estimates a font size (never above `baseFontSize`, never below `minFontSize`)
// small enough that every line in `lines` fits on one row within `boxWidthPt`,
// so long product/reason names shrink instead of getting cut with "…".
// Korean/wide characters are weighted ~1.7x a Latin character since they're
// roughly full-width at any given font size.
function _fitFontSizeForLines(lines, boxWidthPt, baseFontSize, minFontSize) {
  let maxWeighted = 0;
  lines.forEach(function(line) {
    let weighted = 0;
    for (let i = 0; i < line.length; i++) {
      weighted += /[ㄱ-힝]/.test(line[i]) ? 1.7 : 1;
    }
    if (weighted > maxWeighted) maxWeighted = weighted;
  });
  if (maxWeighted === 0) return baseFontSize;
  const avgCharWidthPerPt = 0.52; // approx Arial width-per-point per weighted char unit
  const fitSize = boxWidthPt / (avgCharWidthPerPt * maxWeighted);
  return Math.max(minFontSize, Math.min(baseFontSize, fitSize));
}

// Draws one card (arc chart image + count/title/legend/legend-value/rank text)
// with its top-left corner at (left, top). `legendFontSize` is computed once
// for the whole grid (see insertGeneratedChartGrid) so every card's legend
// text renders at the same size instead of each shrinking independently.
function insertGeneratedChartCard(slide, left, top, chartData, chartBlobTitle, cardTitle, countText, legendLabelText, legendValueText, rank, legendFontSize) {
  const g = CHART_CARD_GEOM;
  const blob = buildDefectModelChartBlob(chartData, chartBlobTitle);
  const image = slide.insertImage(blob, left, top, g.cardW, g.cardH);
  image.setTitle(chartBlobTitle);

  function addText(text, spec, fontSizeOverride, autofit) {
    const box = slide.insertTextBox(text, left + spec.l, top + spec.t, spec.w, spec.h);
    const tr = box.getText();
    tr.getTextStyle().setFontFamily('Arial').setFontSize(fontSizeOverride || spec.fontSize).setForegroundColor(spec.color);
    tr.getParagraphStyle().setParagraphAlignment(SlidesApp.ParagraphAlignment[spec.align]);
    box.setContentAlignment(SlidesApp.ContentAlignment[spec.valign]);
    // Shrinks the font (never truncates) so the full product/reason name
    // always shows, even when longer than the box at the base font size.
    if (autofit) {
      try { box.getAutofit().setAutofitType(SlidesApp.AutofitType.TEXT_AUTOFIT); } catch (e) {}
    }
    return box;
  }

  addText(countText, g.count);
  addText(cardTitle, g.title, null, true);
  addText(legendLabelText, g.legend, legendFontSize);
  addText(legendValueText, g.value, legendFontSize);
  addText(String(rank), g.rank);
}

// Draws one full row: a "{boldPrefix} 클레임" header followed by up to 3 cards.
function insertGeneratedChartRow(slide, top, headerBoldText, items, blobTitlePrefix, legendFontSize) {
  const boldBox = slide.insertTextBox(headerBoldText, CHART_GRID_HEADER_LEFT, top, 106, 26);
  boldBox.getText().getTextStyle().setFontFamily('Arial').setFontSize(11.5).setBold(true).setForegroundColor('#121735');

  const lightBox = slide.insertTextBox('클레임', CHART_GRID_HEADER_LEFT + 108, top + 1, 47, 27);
  lightBox.getText().getTextStyle().setFontFamily('Arial').setFontSize(10.5).setForegroundColor('#7278B2');

  const cardsTop = top + 31;
  items.slice(0, 3).forEach(function(item, i) {
    insertGeneratedChartCard(
      slide, CHART_GRID_COL_LEFTS[i], cardsTop,
      item.chartData, blobTitlePrefix + '_' + (i + 1),
      item.title, item.count, item.legendLabel, item.legendValue, i + 1, legendFontSize
    );
  });
}

// Appends a new blank slide with the full 모델별 TOP3 / 인입사유별 TOP3 grid,
// built from the already-computed topProducts / topReasons arrays.
function insertGeneratedChartGrid(presentation, topProducts, topReasons, blobTitlePrefix) {
  const slide = presentation.appendSlide(SlidesApp.PredefinedLayout.BLANK);
  try { slide.getBackground().setSolidFill(CHART_GRID_BG_COLOR); } catch (e) {}

  const productItems = topProducts.map(function(p) {
    return {
      chartData: p,
      title: p.productName,
      count: p.total.toLocaleString() + '건',
      legendLabel: buildLegendText(p),
      legendValue: buildLegendValues(p)
    };
  });
  const reasonItems = topReasons.map(function(r) {
    return {
      chartData: { reasons: r.models, other: r.other },
      title: r.reasonName,
      count: r.total.toLocaleString() + '건',
      legendLabel: buildModelLegendText(r),
      legendValue: buildModelLegendValues(r)
    };
  });

  // One font size for every legend box across all 6 cards — computed from the
  // single longest line anywhere in the grid, so nothing truncates and no
  // card's text is a different size from another's.
  const allLegendLines = productItems.concat(reasonItems)
    .reduce(function(lines, item) { return lines.concat(item.legendLabel.split('\n')); }, []);
  const legendFontSize = _fitFontSizeForLines(allLegendLines, CHART_CARD_GEOM.legend.w - 4, CHART_CARD_GEOM.legend.fontSize, 5.5);

  insertGeneratedChartRow(slide, 51, '모델별 TOP3', productItems, blobTitlePrefix + '_Defect_Model', legendFontSize);
  insertGeneratedChartRow(slide, 225, '인입사유별 TOP3', reasonItems, blobTitlePrefix + '_Model_Defect', legendFontSize);

  return slide;
}

function onOpen() {
  SlidesApp.getUi()
    .createMenu('Slide Updater')
    .addItem('Update Slide Text', 'updateSlideTextBoxes')
    .addItem('Custom Chart Maker...', 'showChartMakerSidebar')
    .addItem('Claim / Review Slide Maker...', 'showSlideMakerSidebar')
    .addItem('Apply SIREN badges to existing cards', 'applySirenBadges')
    .addItem('Create 260618 Report', 'createReport260618')
    .addItem('Get YoY Stats (260618)', 'getYoYStats')
    .addToUi();
}

// ─── YoY STATS ────────────────────────────────────────────────────────────────
// Counts all rows (no category filter) up to the date cutoffs,
// computes YoY%, and lists top-10 1차 Defect Reason or Inquiries for 2026.
function getYoYStats() {
  var ZD26_ID   = '1sjcCj_P4DRD8rywkmYJhbsrzwFfgiJQuF9nIKwCiKlc';
  var ZD25_ID   = '1t5CJLsVw1hAVPspt_gBVCiM6vWSrEfvqQ8aV7xvoGD8';
  var CUT26     = new Date('2026-06-18T23:59:59');
  var CUT25     = new Date('2025-06-18T23:59:59');

  // ── 2026 sheet ──────────────────────────────────────────────────────────────
  var sheet26   = SpreadsheetApp.openById(ZD26_ID).getSheetByName('26년 전체문의');
  if (!sheet26) { Logger.log('26년 전체문의 not found'); return; }

  var lastRow26 = sheet26.getLastRow();
  var numRows26 = Math.max(lastRow26 - 1, 0);
  var headers26 = sheet26.getRange(1, 1, 1, sheet26.getLastColumn()).getDisplayValues()[0];

  function colIdx26(name) {
    var i = headers26.indexOf(name);
    if (i === -1) throw new Error('Col not found: ' + name);
    return i + 1;
  }

  var dateCol26   = colIdx26('Ticket created - Date');
  var reasonCol26 = colIdx26('1차 Defect Reason or Inquiries');

  var dates26   = sheet26.getRange(2, dateCol26,   numRows26, 1).getDisplayValues().flat();
  var reasons26 = sheet26.getRange(2, reasonCol26, numRows26, 1).getDisplayValues().flat();

  var count26 = 0;
  var reasonMap = {};
  for (var i = 0; i < numRows26; i++) {
    var d = new Date(dates26[i]);
    if (isNaN(d.getTime()) || d > CUT26) continue;
    count26++;
    var r = (reasons26[i] || '').trim();
    if (r) reasonMap[r] = (reasonMap[r] || 0) + 1;
  }

  // Top-10 reasons
  var top10 = Object.entries(reasonMap)
    .sort(function(a, b) { return b[1] - a[1]; })
    .slice(0, 10);

  // ── 2025 sheet ──────────────────────────────────────────────────────────────
  var sheet25   = SpreadsheetApp.openById(ZD25_ID).getSheetByName('25년 전체문의');
  if (!sheet25) { Logger.log('25년 전체문의 not found'); return; }

  var lastRow25 = sheet25.getLastRow();
  var numRows25 = Math.max(lastRow25 - 1, 0);
  var headers25 = sheet25.getRange(1, 1, 1, sheet25.getLastColumn()).getDisplayValues()[0];

  function colIdx25(name) {
    var i = headers25.indexOf(name);
    if (i === -1) throw new Error('Col not found in 25시트: ' + name);
    return i + 1;
  }

  var dateCol25 = colIdx25('Ticket created - Date');
  var dates25   = sheet25.getRange(2, dateCol25, numRows25, 1).getDisplayValues().flat();

  var count25 = 0;
  for (var j = 0; j < numRows25; j++) {
    var d2 = new Date(dates25[j]);
    if (isNaN(d2.getTime()) || d2 > CUT25) continue;
    count25++;
  }

  // ── YoY % ───────────────────────────────────────────────────────────────────
  var yoy = count25 > 0
    ? (((count26 - count25) / count25) * 100).toFixed(2) + '%'
    : 'N/A (no 2025 data)';

  // ── Output ──────────────────────────────────────────────────────────────────
  var lines = [
    '=== GCX Bi-Weekly 260618 Stats ===',
    '',
    '2026 누적 클레임 (~6.18): ' + count26.toLocaleString() + '건',
    '2025 누적 클레임 (~6.18): ' + count25.toLocaleString() + '건',
    'YoY: ' + (count26 - count25 >= 0 ? '+' : '') + (count26 - count25).toLocaleString() + '건  (' + (count26 - count25 >= 0 ? '+' : '') + yoy + ')',
    '',
    '─── TOP 10  1차 인입사유 (2026, ~6.18) ───'
  ];
  top10.forEach(function(entry, idx) {
    lines.push((idx + 1) + '. ' + entry[0] + '  →  ' + entry[1] + '건');
  });

  var msg = lines.join('\n');
  Logger.log(msg);
  SpreadsheetApp.getUi
    ? null
    : Browser.msgBox(msg);
  SlidesApp.getUi().alert(msg);
}

// ─── CREATE 260618 ────────────────────────────────────────────────────────────

function createReport260618() {
  const SOURCE_ID  = '1Dp5A7RU7uPFGWZK-a7S7z6tIS69nz2XmU-YhHB9GPxk'; // 260605
  const ZD_ID      = '1sjcCj_P4DRD8rywkmYJhbsrzwFfgiJQuF9nIKwCiKlc';  // Zendesk claims
  const GLX26_ID   = '1fpv9TEDPGR8D6QRRc0ll-WzF7sOkfxe9UNBCmdBSE9g';  // Glx26 Amazon
  const CUT_DATE   = '2026.06.18';
  const CUT_LABEL  = '~6.18';

  // 1. Copy presentation (preserves all formatting, fonts, colors, images)
  const newFile = DriveApp.getFileById(SOURCE_ID).makeCopy('260618 GCX Bi-weekly Report');
  const pres    = SlidesApp.openById(newFile.getId());

  // 2. Static text replacements across all slides
  pres.replaceAllText('2026.06.05',       '2026.06.18');
  pres.replaceAllText('~6.4)',            '~6.18)');          // inside parentheses like (~6.4)
  pres.replaceAllText('\\~6.4',           '\\~6.18');         // markdown-escaped variant
  pres.replaceAllText('26.2.27~26.6.11', '26.2.27~26.6.11 (종료)');
  pres.replaceAllText('26.3.9~26.6.9',   '26.3.9~26.6.9 (종료)');

  // 3. Compute overview stats from Zendesk sheet
  try {
    var stats = computeOverviewStats(ZD_ID);
    pres.replaceAllText('12,022',         stats.total);
    pres.replaceAllText('-4.62%(YoY)',    stats.yoy);
    pres.replaceAllText('황변 1,520',     '황변 ' + stats.yellowing);
    pres.replaceAllText('충전불량 1,071', '충전불량 ' + stats.charging);
    pres.replaceAllText('파손 466',       '파손 ' + stats.breakage);
  } catch (e) {
    Logger.log('Overview stats error: ' + e.message);
  }

  // 4. Update claim/review example slides with new June 5-18 data
  try {
    updateExampleSlides(pres, GLX26_ID, ZD_ID);
  } catch (e) {
    Logger.log('Example slides error: ' + e.message);
  }

  pres.saveAndClose();
  Logger.log('260618 created: ' + newFile.getUrl());
  SpreadsheetApp.getUi
    ? null
    : SlidesApp.getUi().alert('Done!\n' + newFile.getUrl());
}

// ─── OVERVIEW STATS ───────────────────────────────────────────────────────────

function computeOverviewStats(zdId) {
  const sheet    = SpreadsheetApp.openById(zdId).getSheetByName('26년 전체문의');
  const lastRow  = sheet.getLastRow();
  const rowCount = Math.max(lastRow - 1, 0);

  const catCol    = getColumnIndexByHeader(sheet, 'Category');
  const reasonCol = getColumnIndexByHeader(sheet, '인입사유');
  const dateCol   = getColumnIndexByHeader(sheet, 'Ticket created - Date');

  const cats    = sheet.getRange(2, catCol,    rowCount, 1).getDisplayValues().flat();
  const reasons = sheet.getRange(2, reasonCol, rowCount, 1).getDisplayValues().flat();
  const dates   = sheet.getRange(2, dateCol,   rowCount, 1).getDisplayValues().flat();

  // Cut-off: June 18, 2026
  const cutOff = new Date('2026-06-18T23:59:59');

  var total = 0, yellowing = 0, charging = 0, breakage = 0;
  var total2025 = 0; // same period YoY: Jan 1 – Jun 18

  for (var i = 0; i < rowCount; i++) {
    var d = new Date(dates[i]);
    if (isNaN(d.getTime()) || d > cutOff) continue;
    if (cats[i].trim() !== '4. Product Issue') continue;

    total++;
    var r = reasons[i];
    if (r.indexOf('황변') !== -1)    yellowing++;
    if (r.indexOf('충전불량') !== -1) charging++;
    if (r.indexOf('파손') !== -1)    breakage++;
  }

  // YoY: count same sheet rows from 2025 Jan-Jun period
  // The YoY sheet (1t5CJLsVw1hAVPspt_gBVCiM6vWSrEfvqQ8aV7xvoGD8) has 2025 data
  try {
    var yoySheet = SpreadsheetApp.openById('1t5CJLsVw1hAVPspt_gBVCiM6vWSrEfvqQ8aV7xvoGD8')
                    .getSheets()[0];
    var yoyRows  = Math.max(yoySheet.getLastRow() - 1, 0);
    var yCatCol  = getColumnIndexByHeader(yoySheet, 'Category');
    var yDateCol = getColumnIndexByHeader(yoySheet, 'Ticket created - Date');
    var yCats    = yoySheet.getRange(2, yCatCol,  yoyRows, 1).getDisplayValues().flat();
    var yDates   = yoySheet.getRange(2, yDateCol, yoyRows, 1).getDisplayValues().flat();
    var cutOff2025 = new Date('2025-06-18T23:59:59');
    for (var j = 0; j < yoyRows; j++) {
      var yd = new Date(yDates[j]);
      if (isNaN(yd.getTime()) || yd > cutOff2025) continue;
      if (yCats[j].trim() === '4. Product Issue') total2025++;
    }
  } catch (e) { total2025 = 0; }

  var yoySuffix = '';
  if (total2025 > 0) {
    var pct = ((total - total2025) / total2025 * 100).toFixed(2);
    yoySuffix = (pct >= 0 ? '+' : '') + pct + '%(YoY)';
  } else {
    yoySuffix = '-4.62%(YoY)'; // fallback to 260605 value if no 2025 data
  }

  return {
    total:     total.toLocaleString(),
    yoy:       yoySuffix,
    yellowing: yellowing.toLocaleString(),
    charging:  charging.toLocaleString(),
    breakage:  breakage.toLocaleString()
  };
}

// ─── EXAMPLE SLIDES UPDATE ────────────────────────────────────────────────────

// Structure of 260605 example slides (0-indexed):
//   slides 5-10:  Galaxy S26 Case Claims (클레임 내용)
//   slides 11-12: Galaxy S26 Case Reviews (리뷰 내용, Case)
//   slide  13:    Galaxy S26 SP Reviews   (리뷰 내용, SP)
//   slides 14-15: Screen Protector Claims (Pixel 10a, 클레임 내용)
//   slides 17-19: SIREN entries           (리뷰 내용, SIREN)

function updateExampleSlides(pres, glx26Id, zdId) {
  var glxSheet = SpreadsheetApp.openById(glx26Id).getSheetByName('1-3점');
  if (!glxSheet) return;

  var newCaseExamples = getRecentGlx26Examples(glxSheet, '2026-06-05', '2026-06-18', 'case', 8);
  var newSpExamples   = getRecentGlx26Examples(glxSheet, '2026-06-05', '2026-06-18', 'sp',   3);

  var slides = pres.getSlides();

  // Old case claim data from 260605 (used as find-anchors for replacement)
  var oldCaseClaims = [
    { asin:'B0FVB24LQ4', sku:'ACS11060', date:'2026.06.01', text:'케이스 장착 시 측면 부분이 잘 맞지 않음',                                         product:'Galaxy S26용 Ultra Hybrid',          rating:'4.4점  Global Ratings: 7,965',    country:'FR', defect:'형합'        },
    { asin:'B0GDHBDG47', sku:'ACS11394', date:'2026.05.30', text:'케이스의 코팅이 벗겨짐 .',                                                       product:'Galaxy S26 Ultra용 Thin Fit Magfit', rating:'4.1점  Global Ratings: 1,177',    country:'IN', defect:'코팅벗겨짐'  },
    { asin:'B0FVBHT64R', sku:'ACS11033', date:'2026.05.26', text:'케이스 측면 볼륨 버튼 주변을 약 1.5cm 정도를 커버하지 못함.',                      product:'Galaxy S26 Ultra용 Tough Armor Magfit',rating:'4.6점  Global Ratings: 46,564', country:'DE', defect:'형합'        },
    { asin:'B0GDHBDG47', sku:'ACS11394', date:'2026.05.25', text:'3개월밖에 안 됐는데 벌써 뒷면 코팅이 벗겨짐.',                                     product:'Galaxy S26 Ultra용 Thin Fit MagFit', rating:'4.2점  Global Ratings: 1,175',    country:'UK', defect:'코팅벗겨짐'  },
    { asin:'B0GDHBDG47', sku:'ACS11394', date:'2026.05.24', text:'케이스 외부 무광 코팅이 벗겨지고 마모되기 시작했음.',                               product:'Galaxy S26 Ultra용 Thin Fit MagFit', rating:'4.1점  Global Ratings: 1,177(IN)',country:'IN', defect:'코팅벗겨짐'  },
    { asin:'B0GDHBDG47', sku:'ACS11394', date:'2026.05.17', text:'코팅이 벗겨짐',                                                                   product:'Galaxy S26 Ultra용 Thin Fit MagFit', rating:'4.2점  Global Ratings: 1,175',    country:'UK', defect:'코팅벗겨짐'  },
  ];
  var oldCaseReviews = [
    { asin:'B0GDHBDG47', sku:'ACS11394', date:'2026.05.26', text:'사용 몇 개월 만에 후면 코팅이 벗겨짐',                                             product:'Galaxy S26 Ultra용 Thin Fit MagFit', rating:'4.1점  Global Ratings: 1,177(IN)',country:'IN', defect:'코팅 벗겨짐' },
    { asin:'B0FVBHT64R', sku:'ACS11033', date:'2026.05.19', text:'Galaxy S26 Ultra용이라고 되어 있는데, 작은 카메라 두 개의 위치를 보면 그 설명은 맞지 않는 것 같다고 함.', product:'Galaxy S26 Ultra용 Tough Armor MagFit', rating:'4.6점  Global Ratings: 46,564', country:'DE', defect:'컷아웃' },
  ];
  var oldSpReviews = [
    { asin:'B0G7STMFKY', sku:'AGL11071', date:'2026.05.19', text:'제품 내부에 습기가 차 있었으며, 제품을 떼어내는 과정에서 카메라 렌즈가 긁혔습니다.', product:'Galaxy S26 Ultra용 Glas.tR EZ Fit Optik Pro', rating:'4.1점  Global Ratings: 2,906', country:'US', defect:'기기손상' },
  ];

  // Replace case claim slides (index 5-10 in 260605 → indices still 5-10 in the copy)
  var claimExamples = newCaseExamples.filter(function(e){ return e.type === 'claim'; }).slice(0, 6);
  for (var i = 0; i < oldCaseClaims.length; i++) {
    var old = oldCaseClaims[i];
    var neo = claimExamples[i];
    if (!neo) continue;
    replaceExampleOnSlide(slides[5 + i], old, neo);
  }

  // Replace case review slides (index 11-12)
  var reviewExamples = newCaseExamples.filter(function(e){ return e.type === 'review'; }).slice(0, 2);
  for (var i = 0; i < oldCaseReviews.length; i++) {
    var old = oldCaseReviews[i];
    var neo = reviewExamples[i];
    if (!neo) continue;
    replaceExampleOnSlide(slides[11 + i], old, neo);
  }

  // Replace SP review slide (index 13)
  var spReviewExamples = newSpExamples.filter(function(e){ return e.type === 'review'; });
  if (spReviewExamples[0]) {
    replaceExampleOnSlide(slides[13], oldSpReviews[0], spReviewExamples[0]);
  }

  // SP Claims (Pixel 10a, slides 14-15): monitoring ended June 9 — keep as-is but mark date range
  // SIREN slides (17-19): keep as-is unless new SIREN entries exist
}

// Reads the Glx26 1-3점 sheet and returns up to `limit` entries
// with Update 날짜 between startDate and endDate (YYYY-MM-DD strings).
// typeFilter: 'case' | 'sp' | 'all'
function getRecentGlx26Examples(sheet, startDate, endDate, typeFilter, limit) {
  var lastRow  = sheet.getLastRow();
  var rowCount = Math.max(lastRow - 1, 0);
  if (rowCount === 0) return [];

  var headers  = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getDisplayValues()[0];
  function col(name) {
    var idx = headers.indexOf(name);
    return idx === -1 ? null : idx + 1;
  }

  var dateC    = col('Update 날짜') || col('Exported Date');
  var modelC   = col('모델명');
  var reasonC  = col('인입사유(tag)') || col('인입사유(AI)');
  var textC    = col('본문');
  var countryC = col('Country') || col('국가');
  var asinC    = col('ASIN');
  var skuC     = col('Review ID'); // closest to SKU available
  var ratingC  = col('Rating') || col('별점');

  if (!dateC || !modelC) return [];

  var colsNeeded = [dateC, modelC, reasonC, textC, countryC, asinC, skuC, ratingC].filter(Boolean);
  var maxCol = Math.max.apply(null, colsNeeded);

  var data = sheet.getRange(2, 1, rowCount, maxCol).getDisplayValues();

  var start = new Date(startDate);
  var end   = new Date(endDate + 'T23:59:59');

  var results = [];
  for (var i = data.length - 1; i >= 0 && results.length < limit; i--) {
    var row  = data[i];
    var d    = new Date(row[dateC - 1]);
    if (isNaN(d.getTime())) continue;
    if (d < start || d > end) continue;

    var model  = (row[modelC - 1]   || '').trim();
    var reason = (reasonC  ? row[reasonC  - 1] : '').trim();
    var text   = (textC    ? row[textC    - 1] : '').trim();
    var cntry  = (countryC ? row[countryC - 1] : '').trim();
    var asin   = (asinC    ? row[asinC    - 1] : '').trim();
    var sku    = (skuC     ? row[skuC     - 1] : '').trim();
    var rating = (ratingC  ? row[ratingC  - 1] : '').trim();

    if (!model || !text) continue;

    var isSp = /Glas\.tR|SP_|EZ Fit|AlignMaster|Optik Pro/i.test(model);
    var type = isSp ? 'review' : (reason.indexOf('클레임') !== -1 ? 'claim' : 'review');

    if (typeFilter === 'case' && isSp) continue;
    if (typeFilter === 'sp'   && !isSp) continue;

    // Format date as 2026.06.XX
    var mm = String(d.getMonth() + 1).padStart(2, '0');
    var dd = String(d.getDate()).padStart(2, '0');
    var dateStr = d.getFullYear() + '.' + mm + '.' + dd;

    // Build product label
    var deviceMatch = model.match(/(Galaxy S26\s*\w*)/i) || model.match(/(iPhone\s*\w*)/i) || model.match(/(Pixel\s*\w*)/i);
    var device = deviceMatch ? deviceMatch[1] : 'Galaxy S26';
    var modelShort = model.replace(device, '').trim();
    var product = device + '용 ' + modelShort;

    var ratingNum = parseFloat(rating) || 4.1;
    var ratingStr = ratingNum.toFixed(1) + '점';

    results.push({
      type:    type,
      asin:    asin || 'B0GDHBDG47',
      sku:     sku  || 'ACS11394',
      date:    dateStr,
      text:    text.substring(0, 120),
      product: product,
      rating:  ratingStr + '  Global Ratings: -',
      country: cntry || 'UN',
      defect:  reason || '코팅벗겨짐'
    });
  }

  return results;
}

// Replace all recognisable fields on an example slide using replaceAllText.
// We target specific text patterns unique to that slide to avoid cross-slide bleed.
function replaceExampleOnSlide(slide, old, neo) {
  if (!slide || !old || !neo) return;

  var pairs = [
    [old.asin,    neo.asin],
    [old.sku,     neo.sku],
    [old.date,    neo.date],
    [old.product, neo.product],
    [old.country, neo.country],
    [old.defect,  neo.defect],
    [old.text,    neo.text],
  ];

  // replaceAllText on the slide's page elements directly
  var elements = slide.getPageElements();
  pairs.forEach(function(pair) {
    var find = pair[0], replace = pair[1];
    if (!find || find === replace) return;
    elements.forEach(function(el) {
      if (el.getPageElementType() !== SlidesApp.PageElementType.SHAPE) return;
      var shape = el.asShape();
      var t = shape.getText().asString();
      if (t.indexOf(find) !== -1) {
        shape.getText().setText(t.split(find).join(replace));
      }
    });
  });

  // Rating needs special handling since old rating has variable spacing
  if (old.rating && neo.rating) {
    var rFind = old.rating.split('  Global Ratings: ')[0]; // just the "X.Xpoint" part
    elements.forEach(function(el) {
      if (el.getPageElementType() !== SlidesApp.PageElementType.SHAPE) return;
      var shape = el.asShape();
      var t = shape.getText().asString();
      if (t.indexOf(rFind) !== -1) {
        shape.getText().setText(t.split(rFind).join(neo.rating.split('  Global Ratings: ')[0]));
      }
    });
  }
}

// ─── UPDATE SLIDE TEXT (existing, unchanged) ──────────────────────────────────

function updateSlideTextBoxes() {
  const sheetId = '1sjcCj_P4DRD8rywkmYJhbsrzwFfgiJQuF9nIKwCiKlc';
  const sheetName = '26년 전체문의';

  const sheet = SpreadsheetApp.openById(sheetId).getSheetByName(sheetName);

  if (!sheet) {
    throw new Error('Sheet not found: ' + sheetName);
  }

  const lastRow = sheet.getLastRow();
  const rowCount = Math.max(lastRow - 1, 0);

  const categoryCol = getColumnIndexByHeader(sheet, 'Category');
  const reasonCol = getColumnIndexByHeader(sheet, '인입사유');

  const categoryValues = sheet.getRange(2, categoryCol, rowCount, 1).getDisplayValues().flat();
  const reasonValues = sheet.getRange(2, reasonCol, rowCount, 1).getDisplayValues().flat();

  const filteredReasons = [];

  for (let i = 0; i < rowCount; i++) {
    const category = String(categoryValues[i]).trim();
    const reason = String(reasonValues[i]).trim();

    if (category === '4. Product Issue' && reason) {
      filteredReasons.push(reason);
    }
  }

  const reasonCounts = {};

  filteredReasons.forEach(function(reason) {
    reasonCounts[reason] = (reasonCounts[reason] || 0) + 1;
  });

  const sortedReasons = Object.entries(reasonCounts).sort(function(a, b) {
    return b[1] - a[1];
  });

  const presentation = SlidesApp.getActivePresentation();
  const activeSlide = presentation.getSelection().getCurrentPage().asSlide();

  const replacements = {
    '{{TOTAL_INQUIRIES}}': rowCount.toLocaleString()
  };

  for (let i = 1; i <= 5; i++) {
    const item = sortedReasons[i - 1];

    replacements[`{{Defect_Reason_${i}}}`] = item ? item[0] : '';
    replacements[`{{Defect_Reason_${i}_Count}}`] = item ? item[1].toLocaleString() : '';
  }

  const keywordPlaceholders = extractKeywordPlaceholders(activeSlide, 'Defect_Reason_');

  keywordPlaceholders.forEach(function(placeholder) {
    const keyword = placeholder
      .replace('{{Defect_Reason_', '')
      .replace('}}', '')
      .trim();

    if (/^\d+(_Count)?$/i.test(keyword)) return;

    const count = filteredReasons.filter(function(reason) {
      return reason.indexOf(keyword) !== -1;
    }).length;

    replacements[placeholder] = count.toLocaleString();
  });

  const topProducts = buildTopProductsData(sheet, rowCount);
  for (let n = 1; n <= 3; n++) {
    const p = topProducts[n - 1];
    replacements['{{Defect_Model_Chart_Title_' + n + '}}']        = p ? p.productName : '';
    replacements['{{Defect_Model_Chart_Count_' + n + '}}']        = p ? p.total.toLocaleString() + '건' : '';
    replacements['{{Defect_Model_Chart_Legend_' + n + '}}']       = p ? buildLegendText(p) : '';
    replacements['{{Defect_Model_Chart_Legend_Value_' + n + '}}'] = p ? buildLegendValues(p) : '';
  }

  const topReasons = buildTopReasonsData(sheet, rowCount);
  for (let n = 1; n <= 3; n++) {
    const r = topReasons[n - 1];
    replacements['{{Model_Defect_Chart_Title_' + n + '}}']        = r ? r.reasonName : '';
    replacements['{{Model_Defect_Chart_Count_' + n + '}}']        = r ? r.total.toLocaleString() + '건' : '';
    replacements['{{Model_Defect_Chart_Legend_' + n + '}}']       = r ? buildModelLegendText(r) : '';
    replacements['{{Model_Defect_Chart_Legend_Value_' + n + '}}'] = r ? buildModelLegendValues(r) : '';
  }

  const topProductsGlx26 = buildTopProductsData(sheet, rowCount, 'Galaxy S26');
  for (let n = 1; n <= 3; n++) {
    const p = topProductsGlx26[n - 1];
    replacements['{{Defect_Model_Chart_Title_Glx26_' + n + '}}']        = p ? p.productName : '';
    replacements['{{Defect_Model_Chart_Count_Glx26_' + n + '}}']        = p ? p.total.toLocaleString() + '건' : '';
    replacements['{{Defect_Model_Chart_Legend_Glx26_' + n + '}}']       = p ? buildLegendText(p) : '';
    replacements['{{Defect_Model_Chart_Legend_Value_Glx26_' + n + '}}'] = p ? buildLegendValues(p) : '';
  }

  const topReasonsGlx26 = buildTopReasonsData(sheet, rowCount, 'Galaxy S26');
  for (let n = 1; n <= 3; n++) {
    const r = topReasonsGlx26[n - 1];
    replacements['{{Model_Defect_Chart_Title_Glx26_' + n + '}}']        = r ? r.reasonName : '';
    replacements['{{Model_Defect_Chart_Count_Glx26_' + n + '}}']        = r ? r.total.toLocaleString() + '건' : '';
    replacements['{{Model_Defect_Chart_Legend_Glx26_' + n + '}}']       = r ? buildModelLegendText(r) : '';
    replacements['{{Model_Defect_Chart_Legend_Value_Glx26_' + n + '}}'] = r ? buildModelLegendValues(r) : '';
  }

  let topProductsAmzGlx26 = [];
  let topReasonsAmzGlx26 = [];
  const amzSheet = SpreadsheetApp.openById('1fpv9TEDPGR8D6QRRc0ll-WzF7sOkfxe9UNBCmdBSE9g').getSheetByName('1-3점');
  if (amzSheet) {
    const amzRowCount = Math.max(amzSheet.getLastRow() - 1, 0);
    topProductsAmzGlx26 = buildAmzTopProductsData(amzSheet, amzRowCount);
    topReasonsAmzGlx26  = buildAmzTopReasonsData(amzSheet, amzRowCount);
  }
  for (let n = 1; n <= 3; n++) {
    const p = topProductsAmzGlx26[n - 1];
    replacements['{{AMZ_Defect_Model_Chart_Title_Glx26_' + n + '}}']        = p ? p.productName : '';
    replacements['{{AMZ_Defect_Model_Chart_Count_Glx26_' + n + '}}']        = p ? p.total.toLocaleString() + '건' : '';
    replacements['{{AMZ_Defect_Model_Chart_Legend_Glx26_' + n + '}}']       = p ? buildLegendText(p) : '';
    replacements['{{AMZ_Defect_Model_Chart_Legend_Value_Glx26_' + n + '}}'] = p ? buildLegendValues(p) : '';
  }
  for (let n = 1; n <= 3; n++) {
    const r = topReasonsAmzGlx26[n - 1];
    replacements['{{AMZ_Model_Defect_Chart_Title_Glx26_' + n + '}}']        = r ? r.reasonName : '';
    replacements['{{AMZ_Model_Defect_Chart_Count_Glx26_' + n + '}}']        = r ? r.total.toLocaleString() + '건' : '';
    replacements['{{AMZ_Model_Defect_Chart_Legend_Glx26_' + n + '}}']       = r ? buildModelLegendText(r) : '';
    replacements['{{AMZ_Model_Defect_Chart_Legend_Value_Glx26_' + n + '}}'] = r ? buildModelLegendValues(r) : '';
  }

  replaceTextOnSlide(activeSlide, replacements);

  updateDefectModelCharts(activeSlide, topProducts);
  updateModelDefectCharts(activeSlide, topReasons);
  updateDefectModelChartsGlx26(activeSlide, topProductsGlx26);
  updateModelDefectChartsGlx26(activeSlide, topReasonsGlx26);
  updateDefectModelChartsAmzGlx26(activeSlide, topProductsAmzGlx26);
  updateModelDefectChartsAmzGlx26(activeSlide, topReasonsAmzGlx26);

  refreshLinkedCharts(activeSlide);

  presentation.saveAndClose();
}

// ─── CUSTOM CHART MAKER (sidebar) ─────────────────────────────────────────────
// Lets the user pick Category / Device / Product Name filters from a sidebar
// and generate the same 모델별 TOP3 / 인입사유별 TOP3 donut-chart cards as
// updateSlideTextBoxes(), but scoped to whatever subset the user selects
// (e.g. Device = Galaxy Z Fold8 Ultra, Galaxy Z Fold8, Galaxy Z Flip8).
//
// Placeholder family used on the active slide:
//   {{Defect_Model_Chart_Title_<sfx>N}} / _Count_ / _Legend_ / _Legend_Value_ / {{Defect_Model_Chart_<sfx>N}}
//   {{Model_Defect_Chart_Title_<sfx>N}} / _Count_ / _Legend_ / _Legend_Value_ / {{Model_Defect_Chart_<sfx>N}}
// where <sfx> is "" (base family, e.g. the existing TOP3 cards) or a custom
// suffix like "Fold8_" typed into the sidebar (after duplicating a chart-card
// slide and renaming its placeholders to match, e.g. {{Defect_Model_Chart_Fold8_1}}).
// This reuses the exact same arc-chart rendering as the base family, so a new
// device/category/product combo never needs new chart-drawing code — only a
// duplicated slide + a suffix.

const CHART_MAKER_SHEET_ID   = '1sjcCj_P4DRD8rywkmYJhbsrzwFfgiJQuF9nIKwCiKlc';
const CHART_MAKER_SHEET_NAME = '26년 전체문의';

// "Bad Review" data sources — Amazon 1-3점 (critical review) sheets, one per
// product line. Already pre-scoped to critical reviews, so there's no
// Category column; device filtering uses 기종명 (device line, e.g.
// "Galaxy Z Fold 8" / "Pixel 11 Pro") instead of the Zendesk sheet's
// separate Category/Device columns. All share the same column schema
// (모델명/인입사유(tag)/기종명), so adding a new product line here is enough —
// no other code changes needed.
const BAD_REVIEW_SOURCES = {
  glxZ8:   { id: '19OhswglYMx_dxSFFDtWI1WYPWq2jONJn6RK84KITwy4', sheetName: '1-3점', label: 'GlxZ8' },
  pixel11: { id: '12I6z_FFmDIMHa0rLanltKKFp7kI_yREQj3adkMamPgI', sheetName: '1-3점', label: 'Pixel 11' }
};

function _distinctSorted(arr) {
  const seen = {};
  arr.forEach(function(v) {
    const t = String(v).trim();
    if (t) seen[t] = true;
  });
  return Object.keys(seen).sort();
}

function showChartMakerSidebar() {
  const html = HtmlService.createHtmlOutputFromFile('ChartMaker')
    .setTitle('Custom Chart Maker');
  SlidesApp.getUi().showSidebar(html);
}

// Returns distinct Category / Device values from the source sheet so the
// sidebar can populate its filter dropdown + checklist live.
// Cached for 30 min (script cache) — on a sheet with tens of thousands of
// rows, a cold read can take 30+ seconds, long enough for the sidebar's
// google.script.run bridge to come back with an HTTP 503. Cache keeps repeat
// sidebar opens instant; only the first open per 30-min window pays the cost.
// dataSource: 'zendesk' (default) | 'badReview:<key>' where <key> is one of
// BAD_REVIEW_SOURCES (e.g. 'badReview:glxZ8', 'badReview:pixel11').
function getSidebarFilterOptions(dataSource) {
  if (dataSource && dataSource.indexOf('badReview:') === 0) {
    return _getBadReviewFilterOptions(dataSource.slice('badReview:'.length));
  }

  const cache = CacheService.getScriptCache();
  const cacheKey = 'chartMakerFilterOptions';
  const cached = cache.get(cacheKey);
  if (cached) return JSON.parse(cached);

  const sheet = SpreadsheetApp.openById(CHART_MAKER_SHEET_ID).getSheetByName(CHART_MAKER_SHEET_NAME);
  if (!sheet) throw new Error('Sheet not found: ' + CHART_MAKER_SHEET_NAME);
  const rowCount = Math.max(sheet.getLastRow() - 1, 0);

  const catCol = getColumnIndexByHeader(sheet, 'Category');
  const devCol = getColumnIndexByHeader(sheet, 'Device');

  // getValues() (raw) instead of getDisplayValues() (formatted) — noticeably
  // faster on large ranges since it skips per-cell locale/format resolution;
  // Category/Device are plain text so the values are identical either way.
  const cats = rowCount ? sheet.getRange(2, catCol, rowCount, 1).getValues().flat() : [];
  const devs = rowCount ? sheet.getRange(2, devCol, rowCount, 1).getValues().flat() : [];

  const result = {
    categories: _distinctSorted(cats),
    devices: _distinctSorted(devs)
  };

  try {
    cache.put(cacheKey, JSON.stringify(result), 1800);
  } catch (e) {
    // Result too large for the 100KB cache limit — just skip caching.
  }

  return result;
}

// Device list (via 기종명) for a Bad Review data source. No Category concept
// there — the sheet is already pre-scoped to 1-3점 critical reviews.
function _getBadReviewFilterOptions(sourceKey) {
  const source = BAD_REVIEW_SOURCES[sourceKey];
  if (!source) throw new Error('Unknown bad-review source: ' + sourceKey);

  const cache = CacheService.getScriptCache();
  const cacheKey = 'chartMakerBadReviewFilterOptions_' + sourceKey;
  const cached = cache.get(cacheKey);
  if (cached) return JSON.parse(cached);

  const sheet = SpreadsheetApp.openById(source.id).getSheetByName(source.sheetName);
  if (!sheet) throw new Error('Sheet not found: ' + source.sheetName);
  const rowCount = Math.max(sheet.getLastRow() - 1, 0);

  const devCol = getColumnIndexByHeader(sheet, '기종명');
  const devs = rowCount ? sheet.getRange(2, devCol, rowCount, 1).getValues().flat() : [];

  const result = { categories: [], devices: _distinctSorted(devs) };

  try {
    cache.put(cacheKey, JSON.stringify(result), 1800);
  } catch (e) {}

  return result;
}

// Generic row-level filter check shared by buildTopProductsDataV2 / buildTopReasonsDataV2.
// devices / productSubstrings match if the row's value CONTAINS any of the given strings.
function _rowMatchesFilters(category, device, product, filters) {
  if (filters.category && category !== filters.category) return false;

  if (filters.devices && filters.devices.length) {
    const ok = filters.devices.some(function(d) { return device.indexOf(d) !== -1; });
    if (!ok) return false;
  }

  if (filters.productSubstrings && filters.productSubstrings.length) {
    const ok = filters.productSubstrings.some(function(p) { return product.indexOf(p) !== -1; });
    if (!ok) return false;
  }

  return true;
}

// Reads Category / Product Name / 인입사유 / Device once via getValues()
// (faster than getDisplayValues() on large ranges — these are plain text
// columns so the values are identical either way). Shared by
// buildTopProductsDataV2 / buildTopReasonsDataV2 so generateCustomCharts can
// read the sheet once instead of twice.
function _readChartMakerColumns(sheet, rowCount) {
  const categoryCol = getColumnIndexByHeader(sheet, 'Category');
  const productCol  = getColumnIndexByHeader(sheet, 'Product Name');
  const reasonCol   = getColumnIndexByHeader(sheet, '인입사유');
  const deviceCol   = getColumnIndexByHeader(sheet, 'Device');

  return {
    categories: rowCount ? sheet.getRange(2, categoryCol, rowCount, 1).getValues().flat() : [],
    products:   rowCount ? sheet.getRange(2, productCol,  rowCount, 1).getValues().flat() : [],
    reasons:    rowCount ? sheet.getRange(2, reasonCol,   rowCount, 1).getValues().flat() : [],
    devices:    rowCount ? sheet.getRange(2, deviceCol,   rowCount, 1).getValues().flat() : []
  };
}

// Same shape as _readChartMakerColumns, sourced from the Bad Review sheet
// (모델명/인입사유(tag)/기종명). `categories` is filled with blanks since this
// source has no Category column — filters.category must stay unset
// (null/'__ALL__') by the caller so _rowMatchesFilters never checks it.
function _readBadReviewColumns(sheet, rowCount) {
  const productCol = getColumnIndexByHeader(sheet, '모델명');
  const reasonCol  = getColumnIndexByHeader(sheet, '인입사유(tag)');
  const deviceCol  = getColumnIndexByHeader(sheet, '기종명');

  return {
    categories: new Array(rowCount).fill(''),
    products:   rowCount ? sheet.getRange(2, productCol, rowCount, 1).getValues().flat() : [],
    reasons:    rowCount ? sheet.getRange(2, reasonCol,  rowCount, 1).getValues().flat() : [],
    devices:    rowCount ? sheet.getRange(2, deviceCol,  rowCount, 1).getValues().flat() : []
  };
}

// Generalized version of buildTopProductsData: takes a filters object
// { category, devices: [...], productSubstrings: [...] } instead of a single
// hardcoded deviceFilter string, and an explicit topN (default 3).
// `columns`, if passed (see _readChartMakerColumns), skips re-reading the sheet.
function buildTopProductsDataV2(sheet, rowCount, filters, topN, columns) {
  topN = topN || 3;
  filters = filters || {};

  const cols = columns || _readChartMakerColumns(sheet, rowCount);
  const categories = cols.categories;
  const products   = cols.products;
  const reasons    = cols.reasons;
  const devices    = cols.devices;

  const productMap = {};
  for (let i = 0; i < rowCount; i++) {
    const category = String(categories[i]).trim();
    const product  = String(products[i]).trim();
    const reason   = String(reasons[i]).trim();
    const device   = String(devices[i]).trim();
    if (!product || !reason) continue;
    if (!_rowMatchesFilters(category, device, product, filters)) continue;

    if (!productMap[product]) productMap[product] = { total: 0, reasons: {} };
    productMap[product].total++;
    productMap[product].reasons[reason] = (productMap[product].reasons[reason] || 0) + 1;
  }

  return Object.entries(productMap)
    .sort(function(a, b) { return b[1].total - a[1].total; })
    .slice(0, topN)
    .map(function(entry) {
      const productName = entry[0];
      const total       = entry[1].total;
      const topReasons  = Object.entries(entry[1].reasons)
        .sort(function(a, b) { return b[1] - a[1]; })
        .slice(0, 3);
      const topTotal = topReasons.reduce(function(s, r) { return s + r[1]; }, 0);
      return { productName: productName, total: total, reasons: topReasons, other: total - topTotal };
    });
}

// Generalized version of buildTopReasonsData — same filters object + topN.
// `columns`, if passed (see _readChartMakerColumns), skips re-reading the sheet.
function buildTopReasonsDataV2(sheet, rowCount, filters, topN, columns) {
  topN = topN || 3;
  filters = filters || {};

  const cols = columns || _readChartMakerColumns(sheet, rowCount);
  const categories = cols.categories;
  const products   = cols.products;
  const reasons    = cols.reasons;
  const devices    = cols.devices;

  const reasonMap = {};
  for (let i = 0; i < rowCount; i++) {
    const category = String(categories[i]).trim();
    const product  = String(products[i]).trim();
    const reason   = String(reasons[i]).trim();
    const device   = String(devices[i]).trim();
    if (!product || !reason) continue;
    if (!_rowMatchesFilters(category, device, product, filters)) continue;

    if (!reasonMap[reason]) reasonMap[reason] = { total: 0, models: {} };
    reasonMap[reason].total++;
    reasonMap[reason].models[product] = (reasonMap[reason].models[product] || 0) + 1;
  }

  return Object.entries(reasonMap)
    .sort(function(a, b) { return b[1].total - a[1].total; })
    .slice(0, topN)
    .map(function(entry) {
      const reasonName = entry[0];
      const total      = entry[1].total;
      const topModels  = Object.entries(entry[1].models)
        .sort(function(a, b) { return b[1] - a[1]; })
        .slice(0, 3);
      const topTotal = topModels.reduce(function(s, m) { return s + m[1]; }, 0);
      return { reasonName: reasonName, total: total, models: topModels, other: total - topTotal };
    });
}

// Main sidebar entry point. opts = {
//   category: string | '__ALL__',
//   devices: string[],            // OR-matched substrings on Device col
//   products: string[],           // OR-matched substrings on Product Name col
//   suffix: string,               // '' = base {{Defect_Model_Chart_N}} family
//   topN: number                  // how many top-level cards to fill (1-10, default 3)
// }
// Fills whichever placeholders exist on the CURRENTLY SELECTED slide — same
// preserve-on-rerun behavior as insertChartAtPlaceholder (once a chart image
// replaces a placeholder text box, re-running skips that slot until the
// placeholder text is retyped).
function generateCustomCharts(opts) {
  opts = opts || {};

  const isBadReview = !!(opts.dataSource && opts.dataSource.indexOf('badReview:') === 0);
  const badReviewSource = isBadReview ? BAD_REVIEW_SOURCES[opts.dataSource.slice('badReview:'.length)] : null;
  if (isBadReview && !badReviewSource) throw new Error('Unknown bad-review source: ' + opts.dataSource);

  const sheet = isBadReview
    ? SpreadsheetApp.openById(badReviewSource.id).getSheetByName(badReviewSource.sheetName)
    : SpreadsheetApp.openById(CHART_MAKER_SHEET_ID).getSheetByName(CHART_MAKER_SHEET_NAME);
  if (!sheet) throw new Error('Sheet not found: ' + (isBadReview ? badReviewSource.sheetName : CHART_MAKER_SHEET_NAME));
  const rowCount = Math.max(sheet.getLastRow() - 1, 0);

  const filters = {
    // Bad Review source has no Category column — always unset there.
    category: (!isBadReview && opts.category && opts.category !== '__ALL__') ? opts.category : null,
    devices: (opts.devices || []).map(function(s) { return String(s).trim(); }).filter(Boolean),
    productSubstrings: (opts.products || []).map(function(s) { return String(s).trim(); }).filter(Boolean)
  };
  const topN = Math.max(1, Math.min(10, Number(opts.topN) || 3));
  const suffixRaw = String(opts.suffix || '').trim();
  const sfx = suffixRaw ? (suffixRaw + '_') : '';

  const presentation = SlidesApp.getActivePresentation();

  const columns = isBadReview
    ? _readBadReviewColumns(sheet, rowCount)
    : _readChartMakerColumns(sheet, rowCount);
  const topProducts = buildTopProductsDataV2(sheet, rowCount, filters, topN, columns);
  const topReasons  = buildTopReasonsDataV2(sheet, rowCount, filters, topN, columns);

  // Default path: draw a brand-new slide with the full 2-row/3-card grid,
  // matching the live template's colors/fonts/shape exactly — no manual
  // slide prep needed. Pass insertAsNewSlide:false to fall back to the
  // legacy behavior of filling {{...}} placeholders on the active slide.
  if (opts.insertAsNewSlide !== false) {
    const newSlide = insertGeneratedChartGrid(presentation, topProducts, topReasons, sfx || 'GEN');
    const slideIndex = presentation.getSlides().indexOf(newSlide) + 1;
    presentation.saveAndClose();
    return {
      productCards: topProducts.map(function(p) { return p.productName + ' (' + p.total + '건)'; }),
      reasonCards: topReasons.map(function(r) { return r.reasonName + ' (' + r.total + '건)'; }),
      slideIndex: slideIndex
    };
  }

  const slide = presentation.getSelection().getCurrentPage().asSlide();
  if (!slide) throw new Error('No active slide selected. Click a slide in the deck first.');

  const replacements = {};
  for (let n = 1; n <= topN; n++) {
    const p = topProducts[n - 1];
    replacements['{{Defect_Model_Chart_Title_' + sfx + n + '}}']        = p ? p.productName : '';
    replacements['{{Defect_Model_Chart_Count_' + sfx + n + '}}']        = p ? p.total.toLocaleString() + '건' : '';
    replacements['{{Defect_Model_Chart_Legend_' + sfx + n + '}}']       = p ? buildLegendText(p) : '';
    replacements['{{Defect_Model_Chart_Legend_Value_' + sfx + n + '}}'] = p ? buildLegendValues(p) : '';

    const r = topReasons[n - 1];
    replacements['{{Model_Defect_Chart_Title_' + sfx + n + '}}']        = r ? r.reasonName : '';
    replacements['{{Model_Defect_Chart_Count_' + sfx + n + '}}']        = r ? r.total.toLocaleString() + '건' : '';
    replacements['{{Model_Defect_Chart_Legend_' + sfx + n + '}}']       = r ? buildModelLegendText(r) : '';
    replacements['{{Model_Defect_Chart_Legend_Value_' + sfx + n + '}}'] = r ? buildModelLegendValues(r) : '';
  }

  replaceTextOnSlide(slide, replacements);

  topProducts.forEach(function(p, idx) {
    const rank = idx + 1;
    insertChartAtPlaceholder(slide, '{{Defect_Model_Chart_' + sfx + rank + '}}', p, 'AUTO_Defect_Model_Chart_' + sfx + rank);
  });
  topReasons.forEach(function(r, idx) {
    const rank = idx + 1;
    insertChartAtPlaceholder(slide, '{{Model_Defect_Chart_' + sfx + rank + '}}', { reasons: r.models, other: r.other }, 'AUTO_Model_Defect_Chart_' + sfx + rank);
  });

  refreshLinkedCharts(slide);
  presentation.saveAndClose();

  return {
    productCards: topProducts.map(function(p) { return p.productName + ' (' + p.total + '건)'; }),
    reasonCards: topReasons.map(function(r) { return r.reasonName + ' (' + r.total + '건)'; })
  };
}

// ─── SHARED HELPERS ───────────────────────────────────────────────────────────

function replaceTextOnSlide(slide, replacements) {
  _replaceInElements(slide.getPageElements(), replacements);
}

function _replaceInElements(elements, replacements) {
  elements.forEach(function(element) {
    const type = element.getPageElementType();
    if (type === SlidesApp.PageElementType.GROUP) {
      _replaceInElements(element.asGroup().getChildren(), replacements);
      return;
    }
    if (type !== SlidesApp.PageElementType.SHAPE) return;

    const shape = element.asShape();
    const textRange = shape.getText();
    let text = textRange.asString();

    let changed = false;
    Object.keys(replacements).forEach(function(key) {
      if (text.indexOf(key) !== -1) {
        text = text.split(key).join(replacements[key]);
        changed = true;
      }
    });

    if (changed) {
      textRange.setText(text);
    }
  });
}

function buildTopProductsData(sheet, rowCount, deviceFilter) {
  return buildTopProductsDataV2(sheet, rowCount, {
    category: '4. Product Issue',
    devices: deviceFilter ? [deviceFilter] : []
  }, 3);
}

function buildLegendText(item) {
  const lines = item.reasons.map(function(r) { return r[0]; });
  if (item.other > 0) lines.push('그 외');
  return lines.join('\n');
}

function buildLegendValues(item) {
  const lines = item.reasons.map(function(r) { return String(r[1]); });
  if (item.other > 0) lines.push(String(item.other));
  return lines.join('\n');
}

function buildTopReasonsData(sheet, rowCount, deviceFilter) {
  return buildTopReasonsDataV2(sheet, rowCount, {
    category: '4. Product Issue',
    devices: deviceFilter ? [deviceFilter] : []
  }, 3);
}

function buildModelLegendText(item) {
  const lines = item.models.map(function(m) { return m[0]; });
  if (item.other > 0) lines.push('그 외');
  return lines.join('\n');
}

function buildModelLegendValues(item) {
  const lines = item.models.map(function(m) { return String(m[1]); });
  if (item.other > 0) lines.push(String(item.other));
  return lines.join('\n');
}

function updateModelDefectCharts(slide, topReasons) {
  topReasons.forEach(function(r, index) {
    const rank = index + 1;
    insertChartAtPlaceholder(slide, '{{Model_Defect_Chart_' + rank + '}}', { reasons: r.models, other: r.other }, 'AUTO_Model_Defect_Chart_' + rank);
  });
}

function updateDefectModelChartsGlx26(slide, topProducts) {
  topProducts.forEach(function(p, index) {
    const rank = index + 1;
    insertChartAtPlaceholder(slide, '{{Defect_Model_Chart_Glx26_' + rank + '}}', p, 'AUTO_Defect_Model_Chart_Glx26_' + rank);
  });
}

function updateModelDefectChartsGlx26(slide, topReasons) {
  topReasons.forEach(function(r, index) {
    const rank = index + 1;
    insertChartAtPlaceholder(slide, '{{Model_Defect_Chart_Glx26_' + rank + '}}', { reasons: r.models, other: r.other }, 'AUTO_Model_Defect_Chart_Glx26_' + rank);
  });
}

function updateDefectModelChartsAmzGlx26(slide, topProducts) {
  topProducts.forEach(function(p, index) {
    const rank = index + 1;
    insertChartAtPlaceholder(slide, '{{AMZ_Defect_Model_Chart_Glx26_' + rank + '}}', p, 'AUTO_AMZ_Defect_Model_Chart_Glx26_' + rank);
  });
}

function updateModelDefectChartsAmzGlx26(slide, topReasons) {
  topReasons.forEach(function(r, index) {
    const rank = index + 1;
    insertChartAtPlaceholder(slide, '{{AMZ_Model_Defect_Chart_Glx26_' + rank + '}}', { reasons: r.models, other: r.other }, 'AUTO_AMZ_Model_Defect_Chart_Glx26_' + rank);
  });
}

function buildAmzTopProductsData(sheet, rowCount) {
  const productCol = getColumnIndexByHeader(sheet, '모델명');
  const reasonCol  = getColumnIndexByHeader(sheet, '인입사유(tag)');

  const products = sheet.getRange(2, productCol, rowCount, 1).getDisplayValues().flat();
  const reasons  = sheet.getRange(2, reasonCol,  rowCount, 1).getDisplayValues().flat();

  const productMap = {};
  for (let i = 0; i < rowCount; i++) {
    const product = String(products[i]).trim();
    const reason  = String(reasons[i]).trim();
    if (!product || !reason) continue;
    if (!productMap[product]) productMap[product] = { total: 0, reasons: {} };
    productMap[product].total++;
    productMap[product].reasons[reason] = (productMap[product].reasons[reason] || 0) + 1;
  }

  return Object.entries(productMap)
    .sort(function(a, b) { return b[1].total - a[1].total; })
    .slice(0, 3)
    .map(function(entry) {
      const productName = entry[0];
      const total       = entry[1].total;
      const topReasons  = Object.entries(entry[1].reasons)
        .sort(function(a, b) { return b[1] - a[1]; })
        .slice(0, 3);
      const topTotal = topReasons.reduce(function(s, r) { return s + r[1]; }, 0);
      return { productName: productName, total: total, reasons: topReasons, other: total - topTotal };
    });
}

function buildAmzTopReasonsData(sheet, rowCount) {
  const productCol = getColumnIndexByHeader(sheet, '모델명');
  const reasonCol  = getColumnIndexByHeader(sheet, '인입사유(tag)');

  const products = sheet.getRange(2, productCol, rowCount, 1).getDisplayValues().flat();
  const reasons  = sheet.getRange(2, reasonCol,  rowCount, 1).getDisplayValues().flat();

  const reasonMap = {};
  for (let i = 0; i < rowCount; i++) {
    const product = String(products[i]).trim();
    const reason  = String(reasons[i]).trim();
    if (!product || !reason) continue;
    if (!reasonMap[reason]) reasonMap[reason] = { total: 0, models: {} };
    reasonMap[reason].total++;
    reasonMap[reason].models[product] = (reasonMap[reason].models[product] || 0) + 1;
  }

  return Object.entries(reasonMap)
    .sort(function(a, b) { return b[1].total - a[1].total; })
    .slice(0, 3)
    .map(function(entry) {
      const reasonName = entry[0];
      const total      = entry[1].total;
      const topModels  = Object.entries(entry[1].models)
        .sort(function(a, b) { return b[1] - a[1]; })
        .slice(0, 3);
      const topTotal = topModels.reduce(function(s, m) { return s + m[1]; }, 0);
      return { reasonName: reasonName, total: total, models: topModels, other: total - topTotal };
    });
}

function updateDefectModelCharts(slide, topProducts) {
  topProducts.forEach(function(p, index) {
    const rank = index + 1;
    insertChartAtPlaceholder(slide, '{{Defect_Model_Chart_' + rank + '}}', p, 'AUTO_Defect_Model_Chart_' + rank);
  });
}

function insertChartAtPlaceholder(slide, placeholder, chartData, title) {
  const anchor = findPlaceholderShape(slide, placeholder);
  if (!anchor) return;

  slide.getPageElements().forEach(function(el) {
    if ((el.getTitle ? el.getTitle() : '') === title) el.remove();
  });

  const shape  = anchor;
  const left   = shape.getLeft();
  const top    = shape.getTop();
  const width  = shape.getWidth();
  const height = shape.getHeight();

  shape.getText().setText('');

  const blob  = buildDefectModelChartBlob(chartData, title);
  const image = slide.insertImage(blob, left, top, width, height);
  image.setTitle(title);
}

function buildDefectModelChartBlob(data, title) {
  const labels = data.reasons.map(function(r) { return r[0] || '기타'; });
  const values = data.reasons.map(function(r) { return r[1]; });
  if (data.other > 0) {
    labels.push('그 외');
    values.push(data.other);
  }

  const visibleSum = values.reduce(function(a, b) { return a + b; }, 0);
  labels.push('');
  values.push(visibleSum);

  const dt = Charts.newDataTable()
    .addColumn(Charts.ColumnType.STRING, 'Label')
    .addColumn(Charts.ColumnType.NUMBER, 'Value');
  for (let i = 0; i < labels.length; i++) {
    dt.addRow([labels[i], values[i]]);
  }

  const spacerOpt = {};
  spacerOpt[labels.length - 1] = { color: '#11162d' };

  return Charts.newPieChart()
    .setDataTable(dt.build())
    .setDimensions(440, 340)
    .setColors(['#d336f4', '#1554ff', '#19c7f3', '#8790b5'])
    .setOption('pieHole', 0.9)
    .setOption('pieStartAngle', -90)
    .setOption('slices', spacerOpt)
    .setOption('pieSliceBorderColor', '#11162d')
    .setOption('backgroundColor', '#11162d')
    .setOption('chartArea', { left: 110, top: 45, width: 220, height: 220 })
    .setOption('pieSliceText', 'none')
    .setOption('legend', { position: 'none' })
    .build()
    .getBlob()
    .setName(title + '.png');
}

function refreshLinkedCharts(slide) {
  slide.getPageElements().forEach(function(element) {
    if (element.getPageElementType() === SlidesApp.PageElementType.SHEETS_CHART) {
      element.asSheetsChart().refresh();
    }
  });
}

function findPlaceholderShape(slide, placeholder) {
  return _findShapeInElements(slide.getPageElements(), placeholder);
}

function _findShapeInElements(elements, placeholder) {
  for (let i = 0; i < elements.length; i++) {
    const element = elements[i];
    const type = element.getPageElementType();
    if (type === SlidesApp.PageElementType.GROUP) {
      const found = _findShapeInElements(element.asGroup().getChildren(), placeholder);
      if (found) return found;
    }
    if (type !== SlidesApp.PageElementType.SHAPE) continue;
    if (element.asShape().getText().asString().indexOf(placeholder) !== -1) {
      return element.asShape();
    }
  }
  return null;
}

function findPlaceholderShapes(presentation, placeholder) {
  const results = [];
  presentation.getSlides().forEach(function(slide) {
    slide.getPageElements().forEach(function(element) {
      if (element.getPageElementType() !== SlidesApp.PageElementType.SHAPE) return;
      const shape = element.asShape();
      const text = shape.getText().asString();
      if (text.indexOf(placeholder) !== -1) {
        results.push({ slide: slide, shape: shape });
      }
    });
  });
  return results;
}

function removeOldAutoCharts(presentation) {
  presentation.getSlides().forEach(function(slide) {
    slide.getPageElements().forEach(function(element) {
      const title = element.getTitle ? element.getTitle() : '';
      if (title && title.indexOf('AUTO_Defect_Model_Chart_') === 0) {
        element.remove();
      }
    });
  });
}

function getColumnIndexByHeader(sheet, headerName) {
  const headers = sheet
    .getRange(1, 1, 1, sheet.getLastColumn())
    .getDisplayValues()[0]
    .map(function(header) { return String(header).trim(); });

  const index = headers.indexOf(headerName);
  if (index === -1) throw new Error('Header not found: ' + headerName);
  return index + 1;
}

function extractKeywordPlaceholders(slide, prefix) {
  const placeholders = new Set();
  const pattern = new RegExp('{{' + prefix + '[^}]+}}', 'g');

  slide.getPageElements().forEach(function(element) {
    if (element.getPageElementType() !== SlidesApp.PageElementType.SHAPE) return;
    const shape = element.asShape();
    const text = shape.getText().asString();
    const matches = text.match(pattern);
    if (matches) {
      matches.forEach(function(match) { placeholders.add(match); });
    }
  });

  return Array.from(placeholders);
}
