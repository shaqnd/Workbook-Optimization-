const fs = require('fs');
const {
  Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell, WidthType,
  AlignmentType, BorderStyle, ShadingType, PageOrientation, PageBreak,
  Footer, PageNumber, VerticalAlign,
} = require('docx');

// ---------- page geometry (US Letter, 1" margins) ----------
const CONTENT = 9360; // 12240 - 2*1440

// ---------- palette ----------
const NAVY = '1F3864';   // table title band
const HDR = 'D9E2F3';    // column header band
const BAND = 'F2F2F2';   // alternating data row
const WHITE = 'FFFFFF';

// ---------- small builders ----------
const thin = { style: BorderStyle.SINGLE, size: 4, color: 'BFBFBF' };
const cellBorders = { top: thin, bottom: thin, left: thin, right: thin };

function runs(text, opts = {}) {
  return new TextRun({ text: String(text), font: 'Calibri', size: opts.size || 20, bold: !!opts.bold,
    italics: !!opts.italics, color: opts.color, highlight: opts.highlight });
}

function cellPara(text, opts = {}) {
  return new Paragraph({
    alignment: opts.align || AlignmentType.LEFT,
    spacing: { before: 20, after: 20, line: 240, lineRule: 'auto' },
    children: (Array.isArray(text) ? text : [text]).map((t) => runs(t, opts)),
  });
}

function cell(text, opts = {}) {
  return new TableCell({
    width: { size: opts.w, type: WidthType.DXA },
    shading: { type: ShadingType.CLEAR, fill: opts.fill || WHITE, color: 'auto' },
    borders: cellBorders,
    verticalAlign: VerticalAlign.CENTER,
    margins: { top: 40, bottom: 40, left: 100, right: 100 },
    columnSpan: opts.span,
    children: [cellPara(text, opts)],
  });
}

// title band spanning the whole table
function titleRow(text, widths) {
  return new TableRow({
    tableHeader: true,
    children: [
      new TableCell({
        width: { size: widths.reduce((a, b) => a + b, 0), type: WidthType.DXA },
        columnSpan: widths.length,
        shading: { type: ShadingType.CLEAR, fill: NAVY, color: 'auto' },
        borders: cellBorders,
        margins: { top: 60, bottom: 60, left: 100, right: 100 },
        children: [cellPara(text, { bold: true, color: 'FFFFFF', size: 22, align: AlignmentType.CENTER })],
      }),
    ],
  });
}

function headerRow(cells, widths, aligns) {
  return new TableRow({
    tableHeader: true,
    children: cells.map((t, i) => cell(t, {
      w: widths[i], fill: HDR, bold: true,
      align: (aligns && aligns[i]) || AlignmentType.LEFT,
    })),
  });
}

function dataRow(cells, widths, opts = {}) {
  return new TableRow({
    children: cells.map((t, i) => cell(t, {
      w: widths[i],
      fill: opts.fill || WHITE,
      bold: opts.bold || (opts.boldFirst && i === 0),
      align: (opts.aligns && opts.aligns[i]) || AlignmentType.LEFT,
      highlight: opts.highlight,
    })),
  });
}

function table(widths, rows) {
  return new Table({
    columnWidths: widths,
    width: { size: widths.reduce((a, b) => a + b, 0), type: WidthType.DXA },
    alignment: AlignmentType.CENTER,
    rows,
  });
}

// ---------- narrative helpers ----------
function h1(text) {
  return new Paragraph({
    spacing: { before: 240, after: 120 },
    children: [new TextRun({ text, font: 'Calibri', size: 32, bold: true })],
  });
}
function body(text) {
  return new Paragraph({
    spacing: { after: 160, line: 259, lineRule: 'auto' },
    children: [new TextRun({ text, font: 'Calibri', size: 22 })],
  });
}
function bodyRuns(children) {
  return new Paragraph({ spacing: { after: 160, line: 259, lineRule: 'auto' }, children });
}
function gap(n = 1) {
  return Array.from({ length: n }, () => new Paragraph({ children: [] }));
}
function caption(text) {
  return new Paragraph({
    alignment: AlignmentType.CENTER,
    spacing: { before: 60, after: 200 },
    children: [new TextRun({ text, font: 'Calibri', size: 18, italics: true, color: '595959' })],
  });
}
// boxed, yellow-highlighted placeholder the reviewer can search for
const dashed = { style: BorderStyle.DASHED, size: 8, color: '808080' };
function todo(text) {
  return new Table({
    columnWidths: [CONTENT],
    width: { size: CONTENT, type: WidthType.DXA },
    alignment: AlignmentType.CENTER,
    rows: [new TableRow({
      children: [new TableCell({
        width: { size: CONTENT, type: WidthType.DXA },
        borders: { top: dashed, bottom: dashed, left: dashed, right: dashed },
        margins: { top: 240, bottom: 240, left: 120, right: 120 },
        verticalAlign: VerticalAlign.CENTER,
        children: [new Paragraph({
          alignment: AlignmentType.CENTER,
          spacing: { before: 0, after: 0 },
          children: [new TextRun({ text, font: 'Calibri', size: 20, bold: true, highlight: 'yellow' })],
        })],
      })],
    })],
  });
}
function pageBreak() {
  return new Paragraph({ children: [new PageBreak()] });
}

// =====================================================================
// COVER PAGE
// =====================================================================
const right = (children, spacing) => new Paragraph({ alignment: AlignmentType.RIGHT, spacing, children });

const cover = [
  right([new TextRun({ text: 'County Board of Commissioners', font: 'Calibri', size: 52, bold: true })]),
  right([new TextRun({ text: 'Appeal Report ', font: 'Calibri', size: 40 })]),
  right([
    new TextRun({ text: 'Tax Year:', font: 'Calibri', size: 32, bold: true }),
    new TextRun({ text: ' 2025', font: 'Calibri', size: 32 }),
  ]),
  right([
    new TextRun({ text: 'Schedule Numbers:', font: 'Calibri', size: 28, bold: true }),
    new TextRun({ text: '', font: 'Calibri', size: 28, break: 1 }),
    new TextRun({ text: 'R0003928 & R0003929', font: 'Calibri', size: 24 }),
  ]),
  right([
    new TextRun({ text: 'Address:', font: 'Calibri', size: 28, bold: true }),
    new TextRun({ text: '', font: 'Calibri', size: 28, break: 1 }),
    new TextRun({ text: '[STREET ADDRESS — confirm from assessor record]', font: 'Calibri', size: 24, highlight: 'yellow' }),
    new TextRun({ text: 'Brighton, CO 80601', font: 'Calibri', size: 24, break: 1 }),
    new TextRun({ text: 'Adams County, CO', font: 'Calibri', size: 24, break: 1 }),
  ]),
  right([new TextRun({ text: 'Prepared For', font: 'Calibri', size: 28, bold: true })]),
  right([
    new TextRun({ text: 'Board of County Commissioners', font: 'Calibri', size: 28, bold: true }),
    new TextRun({ text: '4430 S Adams County Pkwy', font: 'Calibri', size: 24, break: 1 }),
    new TextRun({ text: 'Brighton, CO 80601', font: 'Calibri', size: 24, break: 1 }),
  ]),
  pageBreak(),
];

// =====================================================================
// SUBJECT PROPERTY
// =====================================================================
const subjectPage = [
  h1('Subject Property'),
  body('The subject of this report is a single-tenant industrial/storage building situated on two adjoining '
    + 'Adams County parcels within the City of Brighton, immediately adjacent to the Interstate 76 corridor. '
    + 'Schedule R0003928 carries the improvements and its supporting site; Schedule R0003929 is a contiguous '
    + 'land-only parcel that was added to this appeal. Both schedules are appraised together as a single '
    + 'economic unit.'),
  todo('[INSERT: subject aerial / plat map — Schedules R0003928 & R0003929]'),
  caption('Assessor field note: “25 – R0003928+ – MEARLEY – 160515 – BTW I-76”'),
  pageBreak(),
];

// =====================================================================
// EXECUTIVE SUMMARY
// =====================================================================
const W2 = [3560, 5800];
const execSummary = table(W2, [
  titleRow('EXECUTIVE SUMMARY', W2),
  dataRow(['Account Number', 'R0003928 & R0003929'], W2, { boldFirst: true }),
  dataRow(['Address', '[STREET ADDRESS — confirm from assessor record]'], W2, { boldFirst: true, fill: BAND }),
  dataRow(['County', 'Adams County'], W2, { boldFirst: true }),
  dataRow(['City, State', 'Brighton, Colorado'], W2, { boldFirst: true, fill: BAND }),
  dataRow(['Zoning', 'C-3 — City of Brighton'], W2, { boldFirst: true }),
  dataRow(['Effective Date', 'June 30, 2024'], W2, { boldFirst: true, fill: BAND }),
  dataRow(['Report Date', '[REPORT DATE]'], W2, { boldFirst: true }),
  dataRow(['Hearing Date', '[HEARING DATE]'], W2, { boldFirst: true, fill: BAND }),
  dataRow(['Property Type', 'Industrial — Storage / Warehouse'], W2, { boldFirst: true }),
  dataRow(['Property Size — GBA (SF)', '4,920'], W2, { boldFirst: true, fill: BAND }),
  dataRow(['Number of Buildings', '1'], W2, { boldFirst: true }),
  dataRow(['Year Built / Condition', '1976 / Average'], W2, { boldFirst: true, fill: BAND }),
  dataRow(['Site Size (SF)', '15,450 (0.351 AC — both parcels combined)'], W2, { boldFirst: true }),
  dataRow(['Property Rights Appraised', 'Fee Simple'], W2, { boldFirst: true, fill: BAND }),
  dataRow(['Assumed Occupancy', '100% — owner / single occupant'], W2, { boldFirst: true }),
]);

// ---------- Petition summary ----------
const petitionSummary = table(W2, [
  titleRow('PETITION SUMMARY', W2),
  dataRow(['Petition Type', 'Petition for Abatement or Refund of Taxes'], W2, { boldFirst: true }),
  dataRow(['Property Tax Year', '2025'], W2, { boldFirst: true, fill: BAND }),
  dataRow(['Date Completed / Date Received', 'February 26, 2026'], W2, { boldFirst: true }),
  dataRow(['Owner Name', 'Chapman (Gallaspy Chapman)'], W2, { boldFirst: true, fill: BAND }),
  dataRow(['Person Signing', 'Gallaspy C. Chapman — signing as Agent'], W2, { boldFirst: true }),
  dataRow(['Phone Number', '303-946-4743'], W2, { boldFirst: true, fill: BAND }),
  dataRow(['Protest Filed in Abatement/Refund Tax Year', 'No'], W2, { boldFirst: true }),
  dataRow(['Petitioner Estimate of Value', 'Not provided'], W2, { boldFirst: true, fill: BAND }),
  dataRow(['Stated Reason for Request',
    '“You are charging for different use of building. This is storage use only.”'], W2, { boldFirst: true }),
  dataRow(['Assessor Recommendation', 'Deny'], W2, { boldFirst: true, fill: BAND }),
]);

// ---------- Petition scenario ----------
const W3 = [4560, 2400, 2400];
const petitionScenario = table(W3, [
  titleRow('PETITION SCENARIO', W3),
  headerRow(['', 'Value', '$/SF'], W3, [AlignmentType.LEFT, AlignmentType.RIGHT, AlignmentType.RIGHT]),
  dataRow(['Assessor’s Assigned Value', '$941,552', '$191.37'], W3,
    { boldFirst: true, aligns: [AlignmentType.LEFT, AlignmentType.RIGHT, AlignmentType.RIGHT] }),
  dataRow(['Petitioner’s Requested Value', 'Not provided', '—'], W3,
    { boldFirst: true, fill: BAND, aligns: [AlignmentType.LEFT, AlignmentType.RIGHT, AlignmentType.RIGHT] }),
]);

// ---------- 2025 value allocation ----------
const W4 = [2340, 2340, 2340, 2340];
const allocation = table(W4, [
  titleRow('2025 ASSESSOR VALUE ALLOCATION', W4),
  headerRow(['Schedule', 'Land', 'Improvements', 'Total Actual Value'], W4,
    [AlignmentType.LEFT, AlignmentType.RIGHT, AlignmentType.RIGHT, AlignmentType.RIGHT]),
  dataRow(['R0003928', '$84,000', '$827,852', '$911,852'], W4,
    { boldFirst: true, aligns: [AlignmentType.LEFT, AlignmentType.RIGHT, AlignmentType.RIGHT, AlignmentType.RIGHT] }),
  dataRow(['R0003929 (land only)', '$29,700', '$0', '$29,700'], W4,
    { boldFirst: true, fill: BAND, aligns: [AlignmentType.LEFT, AlignmentType.RIGHT, AlignmentType.RIGHT, AlignmentType.RIGHT] }),
  dataRow(['Total — Both Parcels', '$113,700', '$827,852', '$941,552'], W4,
    { bold: true, fill: HDR, aligns: [AlignmentType.LEFT, AlignmentType.RIGHT, AlignmentType.RIGHT, AlignmentType.RIGHT] }),
  dataRow(['Indicated $/SF of GBA (4,920 SF)', '—', '$168.26', '$191.37'], W4,
    { boldFirst: true, aligns: [AlignmentType.LEFT, AlignmentType.RIGHT, AlignmentType.RIGHT, AlignmentType.RIGHT] }),
]);

// ---------- land analysis ----------
const W5 = [2160, 1560, 1800, 1800, 2040];
const landTable = table(W5, [
  titleRow('LAND ANALYSIS', W5),
  headerRow(['Schedule', 'Acres', 'Site (SF)', 'Rate / SF', 'Land Value'], W5,
    [AlignmentType.LEFT, AlignmentType.RIGHT, AlignmentType.RIGHT, AlignmentType.RIGHT, AlignmentType.RIGHT]),
  dataRow(['R0003928', '0.241', '10,500', '$8.00', '$84,000'], W5,
    { boldFirst: true, aligns: [AlignmentType.LEFT, AlignmentType.RIGHT, AlignmentType.RIGHT, AlignmentType.RIGHT, AlignmentType.RIGHT] }),
  dataRow(['R0003929', '0.110', '4,950', '$6.00', '$29,700'], W5,
    { boldFirst: true, fill: BAND, aligns: [AlignmentType.LEFT, AlignmentType.RIGHT, AlignmentType.RIGHT, AlignmentType.RIGHT, AlignmentType.RIGHT] }),
  dataRow(['Total', '0.351', '15,450', '—', '$113,700'], W5,
    { bold: true, fill: HDR, aligns: [AlignmentType.LEFT, AlignmentType.RIGHT, AlignmentType.RIGHT, AlignmentType.RIGHT, AlignmentType.RIGHT] }),
]);

// =====================================================================
// SALES COMPARISON APPROACH
// =====================================================================
const W6 = [3360, 2000, 2000, 2000];
const salesStats = table(W6, [
  titleRow('MARKET SALES SUMMARY — $/SF OF GBA', W6),
  headerRow(['Metric', 'Low', 'Mean / Median', 'High'], W6,
    [AlignmentType.LEFT, AlignmentType.RIGHT, AlignmentType.RIGHT, AlignmentType.RIGHT]),
  dataRow(['Sale Price per SF — Range', '$188', '—', '$263'], W6,
    { boldFirst: true, aligns: [AlignmentType.LEFT, AlignmentType.RIGHT, AlignmentType.RIGHT, AlignmentType.RIGHT] }),
  dataRow(['Mean', '—', '$232', '—'], W6,
    { boldFirst: true, fill: BAND, aligns: [AlignmentType.LEFT, AlignmentType.RIGHT, AlignmentType.RIGHT, AlignmentType.RIGHT] }),
  dataRow(['Median', '—', '$240', '—'], W6,
    { boldFirst: true, aligns: [AlignmentType.LEFT, AlignmentType.RIGHT, AlignmentType.RIGHT, AlignmentType.RIGHT] }),
  dataRow(['Concluded Unit Value', '—', '$235', '—'], W6,
    { bold: true, fill: HDR, aligns: [AlignmentType.LEFT, AlignmentType.RIGHT, AlignmentType.RIGHT, AlignmentType.RIGHT] }),
]);

const W7 = [3960, 1800, 1800, 1800];
const subjectSale = table(W7, [
  titleRow('SUBJECT SALE — TIME ADJUSTMENT TO JUNE 30, 2024', W7),
  headerRow(['', 'Rate / Month', 'Indicated Value', '$/SF'], W7,
    [AlignmentType.LEFT, AlignmentType.RIGHT, AlignmentType.RIGHT, AlignmentType.RIGHT]),
  dataRow(['Confirmed sale — 4/30/2020 (both parcels)', '—', '$685,000', '$139.23'], W7,
    { boldFirst: true, aligns: [AlignmentType.LEFT, AlignmentType.RIGHT, AlignmentType.RIGHT, AlignmentType.RIGHT] }),
  dataRow(['Adjusted 50 months — low', '0.25%', '$776,086', '$157.74'], W7,
    { boldFirst: true, fill: BAND, aligns: [AlignmentType.LEFT, AlignmentType.RIGHT, AlignmentType.RIGHT, AlignmentType.RIGHT] }),
  dataRow(['Adjusted 50 months — most probable', '0.50%', '$879,010', '$178.66'], W7,
    { boldFirst: true, aligns: [AlignmentType.LEFT, AlignmentType.RIGHT, AlignmentType.RIGHT, AlignmentType.RIGHT] }),
  dataRow(['Adjusted 50 months — high', '0.75%', '$995,276', '$202.29'], W7,
    { boldFirst: true, fill: BAND, aligns: [AlignmentType.LEFT, AlignmentType.RIGHT, AlignmentType.RIGHT, AlignmentType.RIGHT] }),
]);

const W8 = [5160, 2100, 2100];
const salesConclusion = table(W8, [
  titleRow('SALES COMPARISON APPROACH VALUE CONCLUSION', W8),
  dataRow(['Indicated Value per SF', '$235.00', ''], W8,
    { boldFirst: true, aligns: [AlignmentType.LEFT, AlignmentType.RIGHT, AlignmentType.RIGHT] }),
  dataRow(['Building Area (SF)', 'x  4,920', ''], W8,
    { boldFirst: true, fill: BAND, aligns: [AlignmentType.LEFT, AlignmentType.RIGHT, AlignmentType.RIGHT] }),
  dataRow(['Indicated Value', '$1,156,200', ''], W8,
    { boldFirst: true, aligns: [AlignmentType.LEFT, AlignmentType.RIGHT, AlignmentType.RIGHT] }),
  dataRow(['Rounded to nearest $5,000', '$1,155,000', 'Per SF   $234.76'], W8,
    { bold: true, fill: HDR, aligns: [AlignmentType.LEFT, AlignmentType.RIGHT, AlignmentType.RIGHT] }),
]);

const salesSection = [
  h1('Sales Comparison Approach '),
  body('The Sales Comparison Approach analyzes closed sales of comparable industrial and storage properties '
    + 'located within the subject’s competitive market area during the base period ending June 30, 2024. '
    + 'Because the petitioner did not submit an estimate of value or any supporting sales, the analysis below '
    + 'relies on the sales investigated by the Assessor’s Office for this schedule, together with the subject’s '
    + 'own confirmed 2020 transaction, time-adjusted to the statutory valuation date.'),
  todo('[INSERT: comparable sales grid / CoStar Sale Comps Map & List Report for the subject market area]'),
  ...gap(1),
  salesStats,
  caption('Source: Adams County Assessor sales investigation, base period ending June 30, 2024.'),
  body('The subject last transferred on April 30, 2020 for $685,000 ($139.23 per square foot of gross building '
    + 'area), a confirmed arm’s-length sale that included both Schedule R0003928 and Schedule R0003929. That '
    + 'sale predates the June 30, 2024 valuation date by 50 months and therefore requires a market-conditions '
    + 'adjustment. Three monthly appreciation rates were applied on a compounded basis to bracket the '
    + 'indication:'),
  subjectSale,
  caption('50 months from April 30, 2020 to June 30, 2024; rates applied on a compounded monthly basis.'),
  ...gap(1),
  h1('Sales Comparison Approach - Conclusion'),
  body('Based on our analysis of the Sales Comparison Approach, it is our opinion that the value of the subject '
    + 'property as of the stated retrospective valuation date is as follows:'),
  salesConclusion,
  body('The subject’s own time-adjusted sale supports a range of $776,086 to $995,276 ($157.74 to $202.29 per '
    + 'square foot), with the most probable indication of $879,010 ($178.66 per square foot) at a 0.50% monthly '
    + 'rate. The 2025 assigned value of $941,552 ($191.37 per square foot) falls within that bracket and well '
    + 'below the $235 per square foot indicated by the market sales, notwithstanding the petitioner’s assertion '
    + 'that the improvements are limited to storage use.'),
  pageBreak(),
];

// =====================================================================
// INCOME CAPITALIZATION APPROACH
// =====================================================================
const W9 = [3960, 1800, 1800, 1800];
const proForma = table(W9, [
  titleRow('PRO FORMA OPERATING STATEMENT', W9),
  headerRow(['Category', 'Projected Expenses', 'Projected Income', 'PSF'], W9,
    [AlignmentType.LEFT, AlignmentType.RIGHT, AlignmentType.RIGHT, AlignmentType.RIGHT]),
  ...[
    ['Potential Gross Rental Income (PGRI)', '', '[    ]', '[    ]'],
    ['Plus:  Other Income', '', '[    ]', '[    ]'],
    ['Plus:  Pass-Through Income', '', '[    ]', '[    ]'],
    ['Total Potential Gross Rental Income', '', '[    ]', '[    ]'],
    ['Less: Market Vacancy', '[    ]', '[    ]', '[    ]'],
    ['Effective Gross Income (EGI)', '', '[    ]', '[    ]'],
    ['Owner’s Expenses', '[    ]', '', '[    ]'],
    ['Operating Expense Recovery Loss During Vacancy', '[    ]', '', '[    ]'],
    ['Total Expenses', '[    ]', '', '[    ]'],
    ['Net Operating Income', '', '[    ]', '[    ]'],
  ].map((r, i) => dataRow(r, W9, {
    boldFirst: true,
    fill: i % 2 ? BAND : WHITE,
    aligns: [AlignmentType.LEFT, AlignmentType.RIGHT, AlignmentType.RIGHT, AlignmentType.RIGHT],
  })),
]);

const W10 = [4560, 2400, 2400];
const dcConclusion = table(W10, [
  titleRow('DIRECT CAPITALIZATION METHOD VALUE CONCLUSION', W10),
  headerRow(['', 'Value', '$/SF'], W10, [AlignmentType.LEFT, AlignmentType.RIGHT, AlignmentType.RIGHT]),
  dataRow(['Net Operating Income (NOI)', '[    ]', '[    ]'], W10,
    { boldFirst: true, aligns: [AlignmentType.LEFT, AlignmentType.RIGHT, AlignmentType.RIGHT] }),
  dataRow(['Based on Low-Range OAR', '[    ]', '[    ]'], W10,
    { boldFirst: true, fill: BAND, aligns: [AlignmentType.LEFT, AlignmentType.RIGHT, AlignmentType.RIGHT] }),
  dataRow(['Based on Most Probable OAR', '[    ]', '[    ]'], W10,
    { boldFirst: true, aligns: [AlignmentType.LEFT, AlignmentType.RIGHT, AlignmentType.RIGHT] }),
  dataRow(['Based on High-Range OAR', '[    ]', '[    ]'], W10,
    { boldFirst: true, fill: BAND, aligns: [AlignmentType.LEFT, AlignmentType.RIGHT, AlignmentType.RIGHT] }),
  dataRow(['Reconciled Value', '[    ]', '[    ]'], W10,
    { bold: true, fill: HDR, aligns: [AlignmentType.LEFT, AlignmentType.RIGHT, AlignmentType.RIGHT] }),
]);

const incomeSection = [
  h1('Income Capitalization Approach'),
  body('The Income Capitalization Approach converts anticipated market-level income into an indication of value '
    + 'for the Fee Simple estate. In this analysis a market rental rate is applied to the entire building at '
    + 'stabilized occupancy, market vacancy and collection loss are deducted, and market-level operating '
    + 'expenses are applied to arrive at net operating income, which is then capitalized at a market-derived '
    + 'overall rate.'),
  body('Our estimate of market rent, occupancy, vacancy, collection loss, and a capitalization rate for the '
    + 'subject property are based upon firsthand data collected by staff as well as third-party service '
    + 'providers such as the following: CoStar, CompStak, Marcus & Millichap, PWHC – Real Estate Investor '
    + 'Survey. Etc.'),
  bodyRuns([
    new TextRun({ text: 'Status for this appeal: ', font: 'Calibri', size: 22, bold: true }),
    new TextRun({
      text: 'the abatement petition and the Assessor’s appeal file for Schedules R0003928 and R0003929 contain '
        + 'no market rent, operating expense, or capitalization rate data for the subject. The Income '
        + 'Capitalization Approach has therefore not been developed, and the Sales Comparison Approach is the '
        + 'sole approach relied upon. The schedules below are retained so this section can be completed if '
        + 'storage/warehouse rent and expense data are added to the file prior to the hearing.',
      font: 'Calibri', size: 22,
    }),
  ]),
  proForma,
  ...gap(1),
  h1('CoStar Cap Rate support'),
  todo('[INSERT: CoStar Market Cap Rate exhibit — Denver industrial / storage submarket, 2024 Q2]'),
  caption('Source: CoStar'),
  h1('Conclusion- Income Capitalization Approach'),
  body('An opinion of market value is indicated by the Direct Capitalization Method by dividing the net '
    + 'operating income (NOI), derived earlier in this section by the appropriate capitalization rate. Our '
    + 'conclusion via the Direct Capitalization Method is as follows'),
  dcConclusion,
  pageBreak(),
];

// =====================================================================
// RECONCILIATION / FINAL VALUE
// =====================================================================
const W11 = [3600, 2880, 2880];
const finalValue = table(W11, [
  titleRow('VALUE INDICATIONS — As Is, Retrospective Market Value, Fee Simple Estate', W11),
  dataRow(['Cost Approach', 'N/A', 'Not Developed'], W11,
    { boldFirst: true, aligns: [AlignmentType.LEFT, AlignmentType.RIGHT, AlignmentType.RIGHT] }),
  dataRow(['Sales Comparison Approach', '$1,155,000', '$234.76 Per Square Foot GBA'], W11,
    { boldFirst: true, fill: BAND, aligns: [AlignmentType.LEFT, AlignmentType.RIGHT, AlignmentType.RIGHT] }),
  dataRow(['Income Capitalization Approach — Direct Capitalization', 'N/A', 'Not Developed'], W11,
    { boldFirst: true, aligns: [AlignmentType.LEFT, AlignmentType.RIGHT, AlignmentType.RIGHT] }),
  dataRow(['Approach Reliance', 'Sales Comparison Approach', ''], W11,
    { boldFirst: true, fill: BAND, aligns: [AlignmentType.LEFT, AlignmentType.RIGHT, AlignmentType.RIGHT] }),
  dataRow(['Value Conclusion', '$1,155,000', '$234.76 Per Square Foot GBA'], W11,
    { bold: true, fill: HDR, aligns: [AlignmentType.LEFT, AlignmentType.RIGHT, AlignmentType.RIGHT] }),
]);

const W12 = [2340, 2340, 2340, 2340];
const marketValue = table(W12, [
  titleRow('MARKET VALUE CONCLUSION', W12),
  headerRow(['Appraisal Premise', 'Interest Appraised', 'Date of Value', 'Value Conclusion'], W12,
    [AlignmentType.CENTER, AlignmentType.CENTER, AlignmentType.CENTER, AlignmentType.CENTER]),
  dataRow(['Retrospective Market Value', 'Fee Simple', 'June 30, 2024', '$1,155,000'], W12,
    { bold: true, aligns: [AlignmentType.CENTER, AlignmentType.CENTER, AlignmentType.CENTER, AlignmentType.CENTER] }),
]);

const W13 = [4560, 4800];
const recommendation = table(W13, [
  titleRow('ASSESSOR RECOMMENDATION', W13),
  dataRow(['2025 Assigned Actual Value (both schedules)', '$941,552  ($191.37/SF)'], W13,
    { boldFirst: true, aligns: [AlignmentType.LEFT, AlignmentType.RIGHT] }),
  dataRow(['Petitioner’s Requested Value', 'Not provided'], W13,
    { boldFirst: true, fill: BAND, aligns: [AlignmentType.LEFT, AlignmentType.RIGHT] }),
  dataRow(['Value Concluded in This Report', '$1,155,000  ($234.76/SF)'], W13,
    { boldFirst: true, aligns: [AlignmentType.LEFT, AlignmentType.RIGHT] }),
  dataRow(['Recommended Action', 'DENY — sustain the 2025 assigned value of $941,552'], W13,
    { bold: true, fill: HDR, aligns: [AlignmentType.LEFT, AlignmentType.RIGHT] }),
]);

const reconciliation = [
  h1('Reconciliation of Value'),
  body('The Sales Comparison Approach is primarily used by owner-users in making purchase decisions.  It is the '
    + 'secondary approach used by investors. The subject is an owner-occupied storage/industrial building, which '
    + 'increases the relevance of this approach. Therefore, the Sales Comparison Approach was developed and '
    + 'given primary consideration.'),
  body('The Income Capitalization Approach was considered but was not developed. The appeal file contains no '
    + 'market rental rate, operating expense, or capitalization rate data specific to the subject’s storage use, '
    + 'and the subject is not operated as an investment property. The Cost Approach was likewise considered and '
    + 'not developed; given the 1976 year of construction, the depreciation estimate required would not produce '
    + 'a reliable independent indication of value.'),
  h1('Final Value Indication'),
  finalValue,
  ...gap(1),
  marketValue,
  ...gap(1),
  h1('Recommendation to the Board'),
  body('The petitioner’s stated basis for abatement is that the improvements are used for storage only and that '
    + 'the Assessor has valued the property for a different use. No estimate of value and no market data were '
    + 'submitted in support of the petition, and no protest was filed for the 2025 tax year. The subject’s own '
    + 'confirmed April 30, 2020 sale, time-adjusted 50 months to the June 30, 2024 valuation date, brackets the '
    + 'assigned value; the Assessor’s market sales indicate $235 per square foot, or $1,156,200, which exceeds '
    + 'the assigned value of $941,552 by approximately 23%. The assigned value is therefore supported and no '
    + 'abatement is warranted.'),
  recommendation,
  pageBreak(),
  h1('Subject Photos'),
  todo('[INSERT: subject photographs — Schedules R0003928 & R0003929]'),
  ...gap(1),
];

// =====================================================================
// DOCUMENT
// =====================================================================
const doc = new Document({
  styles: {
    default: {
      document: { run: { font: 'Calibri', size: 22 } },
    },
  },
  sections: [{
    properties: {
      page: {
        size: { width: 12240, height: 15840, orientation: PageOrientation.PORTRAIT },
        margin: { top: 1440, right: 1440, bottom: 1440, left: 1440, header: 720, footer: 720 },
      },
    },
    footers: {
      default: new Footer({
        children: [new Paragraph({
          alignment: AlignmentType.RIGHT,
          children: [new TextRun({ children: [PageNumber.CURRENT, ' | Page'], font: 'Calibri', size: 20 })],
        })],
      }),
    },
    children: [
      ...cover,
      ...subjectPage,
      execSummary,
      ...gap(1),
      petitionSummary,
      ...gap(1),
      petitionScenario,
      ...gap(1),
      allocation,
      ...gap(1),
      landTable,
      pageBreak(),
      ...salesSection,
      ...incomeSection,
      ...reconciliation,
    ],
  }],
});

Packer.toBuffer(doc).then((buf) => {
  fs.writeFileSync(process.argv[2] || 'out.docx', buf);
  console.log('wrote', process.argv[2] || 'out.docx', buf.length, 'bytes');
});
