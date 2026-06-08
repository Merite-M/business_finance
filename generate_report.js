const {
  Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell,
  HeadingLevel, AlignmentType, BorderStyle, WidthType, ShadingType,
  VerticalAlign, PageNumber, Footer, Header, LevelFormat, PageBreak
} = require('docx');
const fs = require('fs');

// ── helpers ──────────────────────────────────────────────────────────────────
const BLUE  = "1F4E79";
const LBLUE = "2E75B6";
const PALE  = "D6E4F0";
const PALEG = "EAF2FB";
const HDR   = "DEEAF1";
const GRAY  = "F2F2F2";
const WHITE = "FFFFFF";
const GOLD  = "C9A84C";
const DARK  = "1A1A2E";
const GREEN = "1D6A3A";
const RED   = "C00000";

const border = { style: BorderStyle.SINGLE, size: 4, color: "AAAAAA" };
const borders = { top: border, bottom: border, left: border, right: border };
const noBorder = { style: BorderStyle.NONE, size: 0, color: "FFFFFF" };
const noBorders = { top: noBorder, bottom: noBorder, left: noBorder, right: noBorder };
const thickBottom = { top: noBorder, bottom: { style: BorderStyle.SINGLE, size: 8, color: LBLUE }, left: noBorder, right: noBorder };

function para(text, opts = {}) {
  return new Paragraph({
    children: [new TextRun({ text, font: "Calibri", size: opts.size || 22, bold: opts.bold, color: opts.color || "000000", italics: opts.italic })],
    alignment: opts.align || AlignmentType.LEFT,
    spacing: { before: opts.before || 80, after: opts.after || 80 },
    ...(opts.heading ? { heading: opts.heading } : {}),
  });
}

function heading1(text) {
  return new Paragraph({
    children: [new TextRun({ text, font: "Calibri", size: 32, bold: true, color: WHITE })],
    heading: HeadingLevel.HEADING_1,
    spacing: { before: 240, after: 160 },
    shading: { fill: BLUE, type: ShadingType.CLEAR },
    indent: { left: 180, right: 180 },
  });
}

function heading2(text) {
  return new Paragraph({
    children: [new TextRun({ text, font: "Calibri", size: 26, bold: true, color: BLUE })],
    heading: HeadingLevel.HEADING_2,
    spacing: { before: 200, after: 100 },
    border: { bottom: { style: BorderStyle.SINGLE, size: 6, color: LBLUE } },
  });
}

function heading3(text) {
  return new Paragraph({
    children: [new TextRun({ text, font: "Calibri", size: 23, bold: true, color: LBLUE })],
    spacing: { before: 160, after: 60 },
  });
}

function body(text, opts = {}) {
  return new Paragraph({
    children: [new TextRun({ text, font: "Calibri", size: 21, color: "1A1A1A", italics: opts.italic || false })],
    spacing: { before: 60, after: 60 },
    alignment: opts.center ? AlignmentType.CENTER : AlignmentType.JUSTIFIED,
  });
}

function spacer(n = 1) {
  return new Paragraph({ children: [new TextRun({ text: "" })], spacing: { before: 60 * n, after: 60 * n } });
}

function cell(text, opts = {}) {
  return new TableCell({
    borders: opts.noBorder ? noBorders : borders,
    width: { size: opts.width || 1800, type: WidthType.DXA },
    shading: opts.fill ? { fill: opts.fill, type: ShadingType.CLEAR } : undefined,
    margins: { top: 80, bottom: 80, left: 120, right: 120 },
    verticalAlign: VerticalAlign.CENTER,
    children: [new Paragraph({
      children: [new TextRun({
        text: String(text),
        font: "Calibri",
        size: opts.size || 18,
        bold: opts.bold || false,
        color: opts.color || "1A1A1A",
        italics: opts.italic || false,
      })],
      alignment: opts.align || AlignmentType.LEFT,
      spacing: { before: 0, after: 0 },
    })],
  });
}

function row(cells) { return new TableRow({ children: cells }); }

// ── IS common-size table helper ───────────────────────────────────────────────
function isRow(label, v25, p25, v24, p24, isHeader = false, isSub = false, highlight = false) {
  const fill = isHeader ? HDR : (highlight ? PALE : (isSub ? PALEG : WHITE));
  const bold = isHeader || highlight;
  const fmt = (v) => v === null ? "" : (v < 0 ? `(${Math.abs(v).toLocaleString()})` : v.toLocaleString());
  const pct = (p) => p === null || p === undefined || typeof p === "string" ? p : `${Number(p).toFixed(2)}%`;
  const labelWidth = 3200;
  const numWidth = 1600;
  return row([
    cell(label,        { width: labelWidth, bold, fill, color: isHeader ? BLUE : "1A1A1A", size: isHeader ? 19 : 18 }),
    cell(fmt(v25),     { width: numWidth, bold, fill, align: AlignmentType.RIGHT }),
    cell(pct(p25),     { width: numWidth, bold, fill, align: AlignmentType.RIGHT, color: p25 !== null && p25 < 0 ? RED : (highlight ? GREEN : "1A1A1A") }),
    cell(fmt(v24),     { width: numWidth, bold, fill, align: AlignmentType.RIGHT }),
    cell(pct(p24),     { width: numWidth, bold, fill, align: AlignmentType.RIGHT, color: p24 !== null && p24 < 0 ? RED : (highlight ? GREEN : "1A1A1A") }),
  ]);
}

function isTable(rows) {
  return new Table({
    width: { size: 9360, type: WidthType.DXA },
    columnWidths: [3200, 1600, 1600, 1600, 1360],
    rows,
  });
}

// scorecard row
function scRow(indicator, score, just, fill) {
  const starFill = score >= 4 ? GREEN : (score >= 3 ? GOLD : RED);
  return row([
    cell(indicator, { width: 2200, bold: true, fill }),
    cell("★".repeat(score) + "☆".repeat(5 - score) + ` (${score}/5)`, { width: 1800, fill, align: AlignmentType.CENTER, bold: true, color: starFill }),
    cell(just, { width: 5360, fill }),
  ]);
}

// ── DOCUMENT ─────────────────────────────────────────────────────────────────
const sections = [];

// ─ COVER PAGE ────────────────────────────────────────────────────────────────
const cover = [
  spacer(8),
  new Paragraph({
    children: [new TextRun({ text: "MTN RWANDACELL PLC", font: "Calibri", size: 52, bold: true, color: BLUE })],
    alignment: AlignmentType.CENTER,
    spacing: { before: 0, after: 200 },
  }),
  new Paragraph({
    children: [new TextRun({ text: "VERTICAL ANALYSIS REPORT", font: "Calibri", size: 44, bold: true, color: GOLD })],
    alignment: AlignmentType.CENTER,
    spacing: { before: 0, after: 120 },
  }),
  new Paragraph({
    children: [new TextRun({ text: "Financial Year Ended 31 December 2025", font: "Calibri", size: 26, bold: false, color: "555555" })],
    alignment: AlignmentType.CENTER,
    spacing: { before: 0, after: 400 },
  }),
  new Paragraph({
    children: [new TextRun({ text: "──────────────────────────────────────────────", font: "Calibri", size: 22, color: LBLUE })],
    alignment: AlignmentType.CENTER,
    spacing: { before: 0, after: 400 },
  }),
  new Paragraph({
    children: [new TextRun({ text: "Prepared by: Group BA-05", font: "Calibri", size: 26, bold: true, color: DARK })],
    alignment: AlignmentType.CENTER,
    spacing: { before: 0, after: 100 },
  }),
  new Paragraph({
    children: [new TextRun({ text: "Kepler College · Business Analytics Programme", font: "Calibri", size: 22, color: "555555" })],
    alignment: AlignmentType.CENTER,
    spacing: { before: 0, after: 100 },
  }),
  new Paragraph({
    children: [new TextRun({ text: "Business Finance · Group Assignment", font: "Calibri", size: 22, color: "555555" })],
    alignment: AlignmentType.CENTER,
    spacing: { before: 0, after: 100 },
  }),
  new Paragraph({
    children: [new TextRun({ text: "June 2026", font: "Calibri", size: 22, color: "555555" })],
    alignment: AlignmentType.CENTER,
    spacing: { before: 0, after: 0 },
  }),
];

sections.push({ properties: { page: { size: { width: 11906, height: 16838 }, margin: { top: 1440, right: 1440, bottom: 1440, left: 1440 } } }, children: cover });

// ─ MAIN REPORT ───────────────────────────────────────────────────────────────
const main = [];

// Executive Summary
main.push(heading1("EXECUTIVE SUMMARY"));
main.push(spacer());
main.push(body("This report presents a vertical (common-size) analysis of MTN Rwandacell Plc's audited consolidated financial statements for the year ended 31 December 2025, benchmarked against the restated 2024 figures. By expressing every income statement line as a percentage of revenue and every balance sheet line as a percentage of total assets, the analysis strips out the effect of scale and reveals how the company's cost structure and financial architecture have shifted over the two-year period."));
main.push(spacer());
main.push(body("The headline finding is a meaningful improvement in operating efficiency. The operating profit margin widened from 14.48% in 2024 to 18.87% in 2025, and the company swung from a net loss of 2.07% of revenue to a net profit margin of 3.63%. This turnaround was driven by strong fintech revenue growth, a reduction in interconnect and regulatory cost ratios, and tighter depreciation ratios — partially offset by a sharp rise in other operating expenses (from 11.89% to 14.64% of revenue) and a jump in credit loss expenses."));
main.push(spacer());
main.push(body("On the balance sheet, the dominant structural feature is mobile money: the combined mobile money assets and customer deposit liability each represent roughly 40% of total assets and total equity-plus-liabilities respectively, reflecting the scale of MTN's fintech operations. Equity remains thin at 8% of total assets, and the total lease liability — spanning both current and non-current — accounts for 22% of the balance sheet. Understanding these structural weights is essential for any assessment of MTN's financial risk and capital allocation choices."));

main.push(spacer(2));

// Introduction
main.push(heading1("1. INTRODUCTION"));
main.push(spacer());
main.push(body("MTN Rwandacell Plc is Rwanda's leading telecommunications and fintech company. It provides mobile voice, data, SMS, digital services, and mobile money (MoMo) through its subsidiary Mobile Money Rwanda Limited. The Group has been listed on the Rwanda Stock Exchange since May 2021, with MTN International (Mauritius) Limited holding 55% and MTN REL (Mauritius) holding 25%, leaving 20% in public float."));
main.push(spacer());
main.push(body("For FY 2025, MTN reported group revenue of Rwf 296.6 billion, representing a 14.3% increase over restated 2024 revenue of Rwf 259.6 billion. More significantly, the Group recovered from a net loss of Rwf 5.4 billion in 2024 to record net profit of Rwf 10.8 billion in 2025 — a major turnaround that this analysis contextualises through the lens of structural cost and asset ratios."));
main.push(spacer());
main.push(body("Vertical analysis is a technique that converts absolute monetary values into proportional percentages relative to a base figure — revenue for the income statement and total assets for the balance sheet. This makes it possible to compare financial structure across periods of different scale, identify dominant cost categories, and spot structural shifts that absolute numbers alone might obscure."));

main.push(spacer(2));

// Methodology
main.push(heading1("2. METHODOLOGY"));
main.push(spacer());
main.push(body("Data source: MTN Rwandacell Plc Annual Consolidated Financial Statements for the year ended 31 December 2025, audited by Ernst & Young Rwanda Limited (unqualified opinion, 11 March 2026). The 2024 comparative figures are restated to correct the treatment of non-lease components previously included in lease liabilities and right-of-use assets (Note 36 of the financial statements)."));
main.push(spacer());
main.push(body("Income statement items are expressed as a percentage of total revenue (Rwf 296.6 billion in 2025; Rwf 259.6 billion in 2024). We use revenue — not total income — as the base because it represents the core trading activity and provides a cleaner view of cost efficiency. Other income (rental income) is small and non-recurring, so treating it separately is analytically cleaner."));
main.push(spacer());
main.push(body("Balance sheet items are expressed as a percentage of total assets (Rwf 723.0 billion in 2025; Rwf 614.0 billion in 2024). Both years are presented side-by-side to allow direct comparison of structural shifts. Rounding is to two decimal places throughout."));

main.push(spacer(2));

// Section 3 – IS Analysis
main.push(heading1("3. INCOME STATEMENT — COMMON-SIZE ANALYSIS"));
main.push(spacer());
main.push(body("Table 1 below shows every income statement line expressed as a percentage of revenue for both years."));
main.push(spacer());
main.push(heading3("Table 1: Common-Size Income Statement (% of Revenue)"));
main.push(spacer());

main.push(isTable([
  isRow("Line Item", "2025 (Rwf '000)", "2025 % of Rev", "2024 (Rwf '000)", "2024 % of Rev", true),
  isRow("Revenue", 296621727, 100.00, 259564827, 100.00, false, false, true),
  isRow("Other Income", 1609039, 0.54, 2062840, 0.79),
  isRow("Total Income", 298230766, 100.54, 261627667, 100.79, false, false, true),
  isRow("Direct Network Operating Costs", -48645982, -16.40, -40905965, -15.76),
  isRow("Government & Regulatory Costs", -13977647, -4.71, -16122868, -6.21),
  isRow("Cost of Handsets & Accessories", -5804937, -1.96, -8746344, -3.37),
  isRow("Interconnect & Roaming Fees", -6642811, -2.24, -11051887, -4.26),
  isRow("Employee Benefits Expense", -26214639, -8.84, -23582210, -9.09),
  isRow("Sales, Distribution & Marketing", -43472795, -14.66, -38721931, -14.92),
  isRow("Other Operating Expenses", -43434254, -14.64, -30860386, -11.89),
  isRow("Credit Loss Expense", -3222394, -1.09, -570541, -0.22),
  isRow("Depreciation — PPE", -20560215, -6.93, -23855592, -9.19),
  isRow("Depreciation — Right-of-Use Asset", -15183388, -5.12, -16504072, -6.36),
  isRow("Amortisation of Intangibles", -15104578, -5.09, -13114484, -5.05),
  isRow("Operating Profit", 55967126, 18.87, 37591387, 14.48, false, false, true),
  isRow("Finance Income", 947334, 0.32, 1475095, 0.57),
  isRow("Finance Costs", -38614661, -13.02, -38780715, -14.94),
  isRow("Profit Before Tax", 18299799, 6.17, 285767, 0.11, false, false, true),
  isRow("Income Tax Expense", -7522444, -2.54, -5657609, -2.18),
  isRow("Net Profit / (Loss)", 10777355, 3.63, -5371842, -2.07, false, false, true),
]));

main.push(spacer(2));
main.push(heading2("3.1 Revenue and Income Structure"));
main.push(body("Revenue is the 100% base. Other income — consisting entirely of site rental receipts — contributed 0.54% of revenue in 2025 (down from 0.79%), confirming its marginal role. Total income therefore equalled 100.54% of revenue, essentially unchanged from 100.79%."));
main.push(spacer());
main.push(body("Within revenue, mobile money commissions are by far the largest driver. The segment reporting (Note 34) shows fintech revenue of Rwf 153.4 billion in 2025 — that is roughly 52% of consolidated revenue — and it grew from Rwf 119.1 billion in 2024, a 29% rise. Telecommunications revenue grew more modestly from Rwf 163.6 billion to Rwf 167.3 billion (before inter-segment eliminations). This structural shift towards fintech is the most significant business trend visible in the income statement and is the main reason costs did not rise proportionally with revenue."));

main.push(spacer());
main.push(heading2("3.2 Major Cost Drivers"));
main.push(body("Direct network operating costs rose from 15.76% to 16.40% of revenue — a modest deterioration driven by network expansion capex and power costs. However, this is still a relatively contained cost line and reflects the fact that incremental mobile data and voice traffic carries low marginal network cost once infrastructure is in place."));
main.push(spacer());
main.push(body("Sales, distribution and marketing held almost flat at 14.66% versus 14.92%, which is positive. Attracting and retaining customers at roughly the same proportional spend while growing revenue by 14% suggests improving marketing efficiency, possibly linked to MoMo customer acquisition dynamics where digital channels reduce agent commission costs per new user at the margin."));
main.push(spacer());
main.push(body("Other operating expenses is the sharpest cost deterioration: 14.64% in 2025 versus 11.89% in 2024, an increase of 2.75 percentage points. The absolute increase was Rwf 12.6 billion. Note 7 explains this: general administration expenses (including community-based health insurance levied at 3% of revenue) jumped from Rwf 10.8 billion to Rwf 20.5 billion. The 3%-of-revenue CBHI levy alone would account for roughly Rwf 8.9 billion in 2025 versus Rwf 7.8 billion in 2024 — but the scale of the jump in note 7 suggests broader cost pressures in administrative overhead as the business scaled. A fair value loss on the Ericsson financial liability (Rwf 850 million) also entered this line for the first time in 2025."));
main.push(spacer());
main.push(body("Two costs improved significantly in ratio terms. Government and regulatory costs fell from 6.21% to 4.71% — partly because they are somewhat fixed in nature and revenue grew strongly. Interconnect and roaming fees dropped from 4.26% to 2.24%, a drop of over 2 percentage points. Note 5(a) shows interconnect revenue grew from Rwf 6.3 billion to Rwf 10.5 billion — MTN's interconnect position improved, shifting it from a net payer to a stronger position, and routing optimisation likely contributed."));

main.push(spacer());
main.push(heading2("3.3 Depreciation and Amortisation"));
main.push(body("Depreciation on PPE fell from 9.19% to 6.93% of revenue — a 2.26 pp improvement. This is partly because revenue grew, but also because PPE additions in 2025 (Rwf 30.7 billion) were similar in scale to 2024 (Rwf 25.3 billion) while some older assets became fully depreciated. ROU asset depreciation fell from 6.36% to 5.12%, also benefiting from the revenue denominator effect. Amortisation of intangibles held steady near 5%, reflecting the GSM licence (10-year straight-line, Rwf 9.1 billion per year) and growing software amortisation. Combined, depreciation and amortisation represented 17.14% of revenue in 2025 — a large but slightly improving burden."));

main.push(spacer());
main.push(heading2("3.4 Profitability Margins"));
main.push(body("Operating profit margin expanded from 14.48% to 18.87% — a gain of 4.39 percentage points. This is a meaningful improvement and reflects genuine operating leverage: revenue grew faster than most cost lines. Finance costs, though still large at 13.02% of revenue (down from 14.94%), absorbed a huge portion of operating profit. The ratio improved because revenue grew while the absolute finance cost was roughly flat (Rwf 38.6 billion in both years). Pre-tax margin recovered from near-zero (0.11%) to 6.17%."));
main.push(spacer());
main.push(body("Net margin settled at 3.63% in 2025 versus negative 2.07% in 2024. This is encouraging but still modest, and it is partly supported by a deferred tax credit (a Rwf 9.7 billion deferred tax asset was recognised, mostly from tax losses carried forward). Strip out the deferred tax benefit and the underlying cash profitability is thinner. That said, the operating turnaround is real: EBITDA — operating profit before depreciation and amortisation — would be approximately Rwf 106.8 billion, or around 36% of revenue, compared to Rwf 91.1 billion and 35% in 2024."));

main.push(spacer(2));
// Section 4 – BS
main.push(heading1("4. BALANCE SHEET — COMMON-SIZE ANALYSIS"));
main.push(spacer());
main.push(body("Table 2 presents all balance sheet items as a percentage of total assets (Rwf 723.0 billion in 2025; Rwf 614.0 billion in 2024)."));
main.push(spacer());
main.push(heading3("Table 2: Common-Size Balance Sheet — Assets (% of Total Assets)"));
main.push(spacer());

// Assets table
main.push(new Table({
  width: { size: 9360, type: WidthType.DXA },
  columnWidths: [3200, 1540, 1080, 1540, 1000, 1000],
  rows: [
    row([
      cell("Asset Item", { width: 3200, bold: true, fill: HDR, color: BLUE, size: 18 }),
      cell("2025 (Rwf '000)", { width: 1540, bold: true, fill: HDR, color: BLUE, size: 18, align: AlignmentType.RIGHT }),
      cell("2025 %", { width: 1080, bold: true, fill: HDR, color: BLUE, size: 18, align: AlignmentType.CENTER }),
      cell("2024 (Rwf '000)", { width: 1540, bold: true, fill: HDR, color: BLUE, size: 18, align: AlignmentType.RIGHT }),
      cell("2024 %", { width: 1000, bold: true, fill: HDR, color: BLUE, size: 18, align: AlignmentType.CENTER }),
    ]),
    row([cell("NON-CURRENT ASSETS", { width: 3200, bold: true, fill: PALE, color: BLUE }), cell("", { width: 1540, fill: PALE }), cell("", { width: 1080, fill: PALE }), cell("", { width: 1540, fill: PALE }), cell("", { width: 1000, fill: PALE })]),
    ...[
      ["IRU Assets (NC)", "6,902,353", "0.95%", "8,295,835", "1.35%"],
      ["Intangible Assets", "69,137,902", "9.56%", "81,016,479", "13.19%"],
      ["Property, Plant & Equipment", "131,133,340", "18.14%", "120,961,726", "19.70%"],
      ["Capitalised Contract Costs (NC)", "1,109,723", "0.15%", "1,611,397", "0.26%"],
      ["Contract Assets (NC)", "1,480,184", "0.20%", "1,498,577", "0.24%"],
      ["Non-Current Prepayments", "252,798", "0.03%", "430,688", "0.07%"],
      ["Right-of-Use Assets", "114,719,190", "15.87%", "116,579,055", "18.99%"],
      ["Deferred Tax Assets", "15,936,152", "2.20%", "6,272,524", "1.02%"],
    ].map(([label, v25, p25, v24, p24]) => row([
      cell(label, { width: 3200 }),
      cell(v25, { width: 1540, align: AlignmentType.RIGHT }),
      cell(p25, { width: 1080, align: AlignmentType.CENTER }),
      cell(v24, { width: 1540, align: AlignmentType.RIGHT }),
      cell(p24, { width: 1000, align: AlignmentType.CENTER }),
    ])),
    row([
      cell("TOTAL NON-CURRENT ASSETS", { width: 3200, bold: true, fill: PALEG }),
      cell("340,671,642", { width: 1540, bold: true, fill: PALEG, align: AlignmentType.RIGHT }),
      cell("47.12%", { width: 1080, bold: true, fill: PALEG, align: AlignmentType.CENTER }),
      cell("336,666,281", { width: 1540, bold: true, fill: PALEG, align: AlignmentType.RIGHT }),
      cell("54.83%", { width: 1000, bold: true, fill: PALEG, align: AlignmentType.CENTER }),
    ]),
    row([cell("CURRENT ASSETS", { width: 3200, bold: true, fill: PALE, color: BLUE }), cell("", { width: 1540, fill: PALE }), cell("", { width: 1080, fill: PALE }), cell("", { width: 1540, fill: PALE }), cell("", { width: 1000, fill: PALE })]),
    ...[
      ["Cash & Cash Equivalents", "58,832,086", "8.14%", "14,514,695", "2.36%"],
      ["Mobile Money Deposits", "286,621,094", "39.65%", "201,517,642", "32.82%"],
      ["Collective Savings Deposits", "7,307,871", "1.01%", "6,319,489", "1.03%"],
      ["Trade & Other Receivables", "25,276,610", "3.50%", "49,729,222", "8.10%"],
      ["Inventories", "320,438", "0.04%", "2,065,585", "0.34%"],
      ["Other Current Assets", "4,933,952", "0.68%", "4,697,031", "0.77%"],
    ].map(([label, v25, p25, v24, p24]) => row([
      cell(label, { width: 3200 }),
      cell(v25, { width: 1540, align: AlignmentType.RIGHT }),
      cell(p25, { width: 1080, align: AlignmentType.CENTER }),
      cell(v24, { width: 1540, align: AlignmentType.RIGHT }),
      cell(p24, { width: 1000, align: AlignmentType.CENTER }),
    ])),
    row([
      cell("TOTAL CURRENT ASSETS", { width: 3200, bold: true, fill: PALEG }),
      cell("382,291,951", { width: 1540, bold: true, fill: PALEG, align: AlignmentType.RIGHT }),
      cell("52.88%", { width: 1080, bold: true, fill: PALEG, align: AlignmentType.CENTER }),
      cell("277,343,820", { width: 1540, bold: true, fill: PALEG, align: AlignmentType.RIGHT }),
      cell("45.17%", { width: 1000, bold: true, fill: PALEG, align: AlignmentType.CENTER }),
    ]),
    row([
      cell("TOTAL ASSETS", { width: 3200, bold: true, fill: HDR, color: BLUE }),
      cell("722,963,593", { width: 1540, bold: true, fill: HDR, align: AlignmentType.RIGHT }),
      cell("100.00%", { width: 1080, bold: true, fill: HDR, align: AlignmentType.CENTER }),
      cell("614,010,101", { width: 1540, bold: true, fill: HDR, align: AlignmentType.RIGHT }),
      cell("100.00%", { width: 1000, bold: true, fill: HDR, align: AlignmentType.CENTER }),
    ]),
  ],
}));

main.push(spacer(2));
main.push(heading3("Table 3: Common-Size Balance Sheet — Equity & Liabilities (% of Total Assets)"));
main.push(spacer());

main.push(new Table({
  width: { size: 9360, type: WidthType.DXA },
  columnWidths: [3200, 1540, 1080, 1540, 1000],
  rows: [
    row([
      cell("Equity & Liability Item", { width: 3200, bold: true, fill: HDR, color: BLUE, size: 18 }),
      cell("2025 (Rwf '000)", { width: 1540, bold: true, fill: HDR, color: BLUE, size: 18, align: AlignmentType.RIGHT }),
      cell("2025 %", { width: 1080, bold: true, fill: HDR, color: BLUE, size: 18, align: AlignmentType.CENTER }),
      cell("2024 (Rwf '000)", { width: 1540, bold: true, fill: HDR, color: BLUE, size: 18, align: AlignmentType.RIGHT }),
      cell("2024 %", { width: 1000, bold: true, fill: HDR, color: BLUE, size: 18, align: AlignmentType.CENTER }),
    ]),
    row([cell("EQUITY", { width: 3200, bold: true, fill: PALE, color: BLUE }), cell("", { width: 1540, fill: PALE }), cell("", { width: 1080, fill: PALE }), cell("", { width: 1540, fill: PALE }), cell("", { width: 1000, fill: PALE })]),
    ...([
      ["Ordinary Share Capital", "1,350,887", "0.19%", "1,350,887", "0.22%"],
      ["Retained Earnings", "41,278,054", "5.71%", "36,618,319", "5.96%"],
      ["Other Reserves", "15,333,606", "2.12%", "9,215,986", "1.50%"],
    ].map(([label, v25, p25, v24, p24]) => row([
      cell(label, { width: 3200 }),
      cell(v25, { width: 1540, align: AlignmentType.RIGHT }),
      cell(p25, { width: 1080, align: AlignmentType.CENTER }),
      cell(v24, { width: 1540, align: AlignmentType.RIGHT }),
      cell(p24, { width: 1000, align: AlignmentType.CENTER }),
    ]))),
    row([
      cell("TOTAL EQUITY", { width: 3200, bold: true, fill: PALEG }),
      cell("57,962,547", { width: 1540, bold: true, fill: PALEG, align: AlignmentType.RIGHT }),
      cell("8.02%", { width: 1080, bold: true, fill: PALEG, align: AlignmentType.CENTER }),
      cell("47,185,192", { width: 1540, bold: true, fill: PALEG, align: AlignmentType.RIGHT }),
      cell("7.68%", { width: 1000, bold: true, fill: PALEG, align: AlignmentType.CENTER }),
    ]),
    row([cell("NON-CURRENT LIABILITIES", { width: 3200, bold: true, fill: PALE, color: BLUE }), cell("", { width: 1540, fill: PALE }), cell("", { width: 1080, fill: PALE }), cell("", { width: 1540, fill: PALE }), cell("", { width: 1000, fill: PALE })]),
    ...([
      ["Borrowings (NC)", "37,134,545", "5.14%", "55,393,665", "9.02%"],
      ["Lease Liabilities (NC)", "141,454,542", "19.57%", "140,995,036", "22.96%"],
      ["NC Financial Liability (Ericsson)", "7,544,097", "1.04%", "9,636,433", "1.57%"],
      ["NC IRU Liability", "338,870", "0.05%", "374,856", "0.06%"],
    ].map(([label, v25, p25, v24, p24]) => row([
      cell(label, { width: 3200 }),
      cell(v25, { width: 1540, align: AlignmentType.RIGHT }),
      cell(p25, { width: 1080, align: AlignmentType.CENTER }),
      cell(v24, { width: 1540, align: AlignmentType.RIGHT }),
      cell(p24, { width: 1000, align: AlignmentType.CENTER }),
    ]))),
    row([
      cell("TOTAL NON-CURRENT LIABILITIES", { width: 3200, bold: true, fill: PALEG }),
      cell("186,472,054", { width: 1540, bold: true, fill: PALEG, align: AlignmentType.RIGHT }),
      cell("25.79%", { width: 1080, bold: true, fill: PALEG, align: AlignmentType.CENTER }),
      cell("206,399,990", { width: 1540, bold: true, fill: PALEG, align: AlignmentType.RIGHT }),
      cell("33.62%", { width: 1000, bold: true, fill: PALEG, align: AlignmentType.CENTER }),
    ]),
    row([cell("CURRENT LIABILITIES", { width: 3200, bold: true, fill: PALE, color: BLUE }), cell("", { width: 1540, fill: PALE }), cell("", { width: 1080, fill: PALE }), cell("", { width: 1540, fill: PALE }), cell("", { width: 1000, fill: PALE })]),
    ...([
      ["Trade & Other Payables", "142,198,241", "19.67%", "109,733,895", "17.87%"],
      ["Mobile Money Customer Deposits", "286,143,324", "39.58%", "202,142,123", "32.92%"],
      ["Borrowings (Current)", "16,249,760", "2.25%", "15,430,274", "2.51%"],
      ["Lease Liabilities (Current)", "18,773,062", "2.60%", "12,536,657", "2.04%"],
      ["Financial Liability (Current)", "3,636,491", "0.50%", "3,122,723", "0.51%"],
      ["Contract Liabilities & Deferred Income", "3,554,898", "0.49%", "3,998,000", "0.65%"],
      ["Provisions, Tax & Other", "7,937,220", "1.10%", "7,461,506", "1.21%"],
      ["Shareholder Loan", "915,000", "0.13%", "915,000", "0.15%"],
    ].map(([label, v25, p25, v24, p24]) => row([
      cell(label, { width: 3200 }),
      cell(v25, { width: 1540, align: AlignmentType.RIGHT }),
      cell(p25, { width: 1080, align: AlignmentType.CENTER }),
      cell(v24, { width: 1540, align: AlignmentType.RIGHT }),
      cell(p24, { width: 1000, align: AlignmentType.CENTER }),
    ]))),
    row([
      cell("TOTAL CURRENT LIABILITIES", { width: 3200, bold: true, fill: PALEG }),
      cell("478,528,992", { width: 1540, bold: true, fill: PALEG, align: AlignmentType.RIGHT }),
      cell("66.19%", { width: 1080, bold: true, fill: PALEG, align: AlignmentType.CENTER }),
      cell("360,424,919", { width: 1540, bold: true, fill: PALEG, align: AlignmentType.RIGHT }),
      cell("58.70%", { width: 1000, bold: true, fill: PALEG, align: AlignmentType.CENTER }),
    ]),
    row([
      cell("TOTAL LIABILITIES", { width: 3200, bold: true, fill: HDR, color: BLUE }),
      cell("665,001,046", { width: 1540, bold: true, fill: HDR, align: AlignmentType.RIGHT }),
      cell("91.98%", { width: 1080, bold: true, fill: HDR, align: AlignmentType.CENTER }),
      cell("566,824,909", { width: 1540, bold: true, fill: HDR, align: AlignmentType.RIGHT }),
      cell("92.32%", { width: 1000, bold: true, fill: HDR, align: AlignmentType.CENTER }),
    ]),
    row([
      cell("TOTAL EQUITY & LIABILITIES", { width: 3200, bold: true, fill: HDR, color: BLUE }),
      cell("722,963,593", { width: 1540, bold: true, fill: HDR, align: AlignmentType.RIGHT }),
      cell("100.00%", { width: 1080, bold: true, fill: HDR, align: AlignmentType.CENTER }),
      cell("614,010,101", { width: 1540, bold: true, fill: HDR, align: AlignmentType.RIGHT }),
      cell("100.00%", { width: 1000, bold: true, fill: HDR, align: AlignmentType.CENTER }),
    ]),
  ],
}));

main.push(spacer(2));
main.push(heading2("4.1 Asset Structure"));
main.push(body("The most striking shift in asset structure between 2024 and 2025 is the rise of current assets from 45.17% to 52.88% of total assets. This was driven almost entirely by two items: mobile money deposits (up from 32.82% to 39.65%) and cash and cash equivalents (up from 2.36% to 8.14%). Together, these two items alone accounted for 47.79% of total assets at year-end 2025."));
main.push(spacer());
main.push(body("Mobile money deposits represent funds held in trust with commercial banks on behalf of MoMo customers. They are ring-fenced under National Bank of Rwanda regulations and are matched by an equal MoMo customer deposit liability (Rwf 286.1 billion on the liability side). The surge in this balance — from Rwf 201.5 billion to Rwf 286.6 billion — reflects growing MoMo transaction volumes and customer wallet balances. From a structural perspective, these are pass-through items: they inflate both sides of the balance sheet but do not represent productive capital available to the company. Any analysis of MTN's 'true' asset base should exclude this amount and focus on the non-MoMo assets, which total approximately Rwf 436 billion."));
main.push(spacer());
main.push(body("Cash and cash equivalents quadrupled from 2.36% to 8.14% of total assets — an absolute increase of Rwf 44.3 billion. This was partly the result of strong operating cash generation (Rwf 119.7 billion net from operations) and net borrowing activity during the year. High cash balances are generally positive for liquidity but should be examined against upcoming debt repayments."));
main.push(spacer());
main.push(body("Among non-current assets, PPE grew from 19.70% to 18.14% of total assets — a modest proportional decline despite Rwf 30.7 billion in capital additions, because total assets grew faster. ROU assets fell from 18.99% to 15.87% as the balance sheet expanded and some leases ran off. Intangible assets fell from 13.19% to 9.56%, driven by ongoing amortisation of the GSM licence and software. Deferred tax assets doubled from 1.02% to 2.20%, reflecting the additional tax loss carry-forward recognition."));

main.push(spacer());
main.push(heading2("4.2 Liability Structure"));
main.push(body("Total liabilities represented 91.98% of total assets in 2025 — marginally better than 92.32% in 2024 but still extremely high. This is structurally driven by the MoMo business: mobile money customer deposits (Rwf 286.1 billion, 39.58% of total assets) and trade payables (which include large intercompany and regulatory payables) dominate the liability side. If MoMo deposits are excluded, the residual liability-to-asset ratio is more manageable — around 57% — though still elevated."));
main.push(spacer());
main.push(body("Lease liabilities (combining current and non-current portions) stood at 22.16% of total assets — the single largest 'conventional' financial liability. This reflects MTN's strategy of leasing tower space and land rather than owning, as well as vehicle and office leases. The lease portfolio shrank slightly from Rwf 153.5 billion to Rwf 160.2 billion in absolute terms (additions of Rwf 13.4 billion partly offset by repayments), and proportionally fell from 25.00% to 22.16% as the overall balance sheet grew."));
main.push(spacer());
main.push(body("Long-term borrowings fell significantly from 9.02% to 5.14% of total assets as Syndicate Loan I was fully repaid and Syndicate Loan II payments reduced the balance. Current liabilities as a proportion of total assets rose sharply from 58.70% to 66.19%, partly because of the MoMo customer deposit growth and rising trade payables (from 17.87% to 19.67%). This structural shift — more liabilities classified as current — is something to watch, though much of it is MoMo-driven and therefore matched by liquid assets."));

main.push(spacer());
main.push(heading2("4.3 Equity Structure"));
main.push(body("Equity grew in absolute terms from Rwf 47.2 billion to Rwf 58.0 billion — a Rwf 10.8 billion uplift driven entirely by the 2025 net profit (no dividend was declared). As a proportion of total assets, equity moved only marginally from 7.68% to 8.02%. This remains a thin equity base and the primary driver of MTN's high leverage ratios. The company's gearing ratio (as disclosed in Note 3) improved from 81.64% to 72.75%, indicating progress but still pointing to meaningful leverage."));
main.push(body("Other reserves grew from 1.50% to 2.12% of total assets, reflecting the legal reserve fund obligation in Mobile Money Rwanda's articles (minimum 20% of net profit to special reserve). Retained earnings ticked up from 5.96% to 5.71% — slightly lower proportionally because the Rwf 6.1 billion transfer to reserves was netted out."));

main.push(spacer(2));

// Section 5 – Key Findings
main.push(heading1("5. KEY FINDINGS AND IDENTIFICATION OF COST DRIVERS"));
main.push(spacer());
main.push(heading2("5.1 Primary Income Statement Cost Drivers (by percentage weight, 2025)"));
main.push(spacer());
main.push(body("The table below ranks the major cost lines by their share of revenue in 2025 and the change from 2024:"));
main.push(spacer());

main.push(new Table({
  width: { size: 9360, type: WidthType.DXA },
  columnWidths: [3200, 1500, 1500, 1500, 1660],
  rows: [
    row([
      cell("Cost Category", { width: 3200, bold: true, fill: HDR, color: BLUE }),
      cell("2025 % of Rev", { width: 1500, bold: true, fill: HDR, color: BLUE, align: AlignmentType.CENTER }),
      cell("2024 % of Rev", { width: 1500, bold: true, fill: HDR, color: BLUE, align: AlignmentType.CENTER }),
      cell("Change (pp)", { width: 1500, bold: true, fill: HDR, color: BLUE, align: AlignmentType.CENTER }),
      cell("Direction", { width: 1660, bold: true, fill: HDR, color: BLUE, align: AlignmentType.CENTER }),
    ]),
    ...([
      ["Direct Network Operating Costs", "16.40%", "15.76%", "+0.64 pp", "Slight deterioration"],
      ["Sales, Distribution & Marketing", "14.66%", "14.92%", "-0.26 pp", "Minor improvement"],
      ["Other Operating Expenses", "14.64%", "11.89%", "+2.75 pp", "Significant deterioration"],
      ["Finance Costs", "13.02%", "14.94%", "-1.92 pp", "Notable improvement"],
      ["Employee Benefits", "8.84%", "9.09%", "-0.25 pp", "Minor improvement"],
      ["Depreciation — PPE", "6.93%", "9.19%", "-2.26 pp", "Notable improvement"],
      ["Depreciation — ROU Assets", "5.12%", "6.36%", "-1.24 pp", "Notable improvement"],
      ["Amortisation of Intangibles", "5.09%", "5.05%", "+0.04 pp", "Stable"],
      ["Government & Regulatory Costs", "4.71%", "6.21%", "-1.50 pp", "Notable improvement"],
      ["Interconnect & Roaming", "2.24%", "4.26%", "-2.02 pp", "Significant improvement"],
      ["Credit Loss Expense", "1.09%", "0.22%", "+0.87 pp", "Significant deterioration"],
      ["Cost of Handsets", "1.96%", "3.37%", "-1.41 pp", "Notable improvement"],
    ].map(([label, p25, p24, chg, dir], i) => row([
      cell(label, { width: 3200, fill: i % 2 === 0 ? WHITE : GRAY }),
      cell(p25, { width: 1500, align: AlignmentType.CENTER, fill: i % 2 === 0 ? WHITE : GRAY }),
      cell(p24, { width: 1500, align: AlignmentType.CENTER, fill: i % 2 === 0 ? WHITE : GRAY }),
      cell(chg, { width: 1500, align: AlignmentType.CENTER, fill: i % 2 === 0 ? WHITE : GRAY, color: chg.startsWith("+") ? RED : GREEN }),
      cell(dir, { width: 1660, fill: i % 2 === 0 ? WHITE : GRAY }),
    ]))),
  ],
}));

main.push(spacer(2));
main.push(heading2("5.2 Primary Balance Sheet Structural Features"));
main.push(body("Three structural features dominate MTN's balance sheet and need to be understood together to interpret the ratios correctly:"));
main.push(spacer());
main.push(body("First, the mobile money ecosystem. Mobile money deposits (asset) and mobile money customer deposits (liability) form a near-perfect matched pair, each at roughly 39-40% of total assets. They are regulatory pass-through items. The growth from ~33% to ~40% on both sides represents expanding fintech volumes, not balance sheet risk per se — but they distort conventional leverage and liquidity ratios."));
main.push(spacer());
main.push(body("Second, the right-of-use asset and lease liability. At 15.87% of total assets (ROU) and 22.16% (total lease liabilities), this is a material structural feature arising from MTN's tower-lease-rather-than-own model. The lease liability slightly exceeds the ROU asset because the lease liability includes future interest unwinding, while ROU assets are at cost less depreciation."));
main.push(spacer());
main.push(body("Third, the thin equity layer. At 8% of total assets, equity is slender relative to the overall balance sheet size. However, once MoMo pass-through items are excluded, the equity-to-productive-assets ratio improves to approximately 13% — still leveraged, but considerably less alarming than the headline number suggests."));

main.push(spacer(2));

// Section 6 – Scorecard
main.push(heading1("6. FINANCIAL HEALTH SCORECARD"));
main.push(spacer());
main.push(body("The scorecard below rates MTN's financial health across eight dimensions, each scored 1 to 5 (5 = excellent). Ratings are grounded in the vertical analysis findings above."));
main.push(spacer());

main.push(new Table({
  width: { size: 9360, type: WidthType.DXA },
  columnWidths: [2200, 1800, 5360],
  rows: [
    row([
      cell("Indicator", { width: 2200, bold: true, fill: HDR, color: BLUE }),
      cell("Score", { width: 1800, bold: true, fill: HDR, color: BLUE, align: AlignmentType.CENTER }),
      cell("Justification", { width: 5360, bold: true, fill: HDR, color: BLUE }),
    ]),
    scRow("Profitability", 3, "Net margin improved from -2.07% to 3.63%, and operating margin expanded to 18.87%. However, net profitability is still modest and partly supported by a deferred tax asset. Further earnings quality improvement is needed.", WHITE),
    scRow("Liquidity", 3, "Cash position strengthened significantly (8.14% of total assets vs 2.36%). Current assets (52.88%) exceed current liabilities (66.19%), but when MoMo pass-through items are excluded, the liquidity picture tightens. Trade payables rose sharply.", GRAY),
    scRow("Solvency", 2, "Total liabilities at 91.98% of total assets looks severe. Excluding MoMo pass-throughs, residual leverage is still high (~57%) with lease liabilities at 22% of assets. Gearing improved from 81.6% to 72.8% but remains elevated.", WHITE),
    scRow("Efficiency", 3, "Operating leverage improving — revenue grew 14.3% while most cost lines held or fell as a ratio. Other operating expenses deteriorated. Interconnect and regulatory cost efficiencies are positive signals.", GRAY),
    scRow("Growth Potential", 4, "Mobile money commissions grew ~29% year-on-year, and MoMo deposits grew ~42%. Rwanda's digital economy growth trajectory is strong. Revenue composition is shifting toward higher-margin fintech services.", WHITE),
    scRow("Risk Exposure", 2, "High lease obligations (22% of assets), thin equity base, USD-denominated payables with significant FX sensitivity, and credit losses up 5x (1.09% vs 0.22% of revenue). These risks are real and visible.", GRAY),
    scRow("Sustainability", 3, "Operations are improving — EBITDA is ~36% of revenue, and cash flow from operations was Rwf 119.7 billion. However, Rwf 38.6 billion in annual finance costs is a persistent drag, and tax losses providing deferred tax credits will eventually be exhausted.", WHITE),
    scRow("Overall Financial Health", 3, "MTN Rwanda is in recovery mode — moving from loss to profit with improving operational efficiency and strong fintech growth. Structural leverage and cost risks remain, but the trajectory is clearly positive.", GRAY),
  ],
}));

main.push(spacer(2));

// Section 7 – Strategic
main.push(heading1("7. FINAL STRATEGIC QUESTION"));
main.push(spacer());
main.push(new Paragraph({
  children: [new TextRun({ text: '"Based on your analysis, assess the overall financial health of MTN Rwanda Plc and recommend strategic actions that management and investors should take during the 2026 financial year."', font: "Calibri", size: 21, italics: true, color: "333333" })],
  spacing: { before: 80, after: 120 },
  alignment: AlignmentType.LEFT,
}));
main.push(spacer());
main.push(body("MTN Rwanda's overall financial health in 2025 is best described as recovering but structurally constrained. The turnaround from a Rwf 5.4 billion loss to a Rwf 10.8 billion profit is real and operationally meaningful — driven by strong fintech growth, improved interconnect economics, and operating leverage as revenue scaled faster than most fixed costs. The vertical analysis reveals genuine structural improvement: finance costs fell from 14.94% to 13.02% of revenue, depreciation ratios improved across the board, and the operating margin widened by over 4 percentage points to 18.87%."));
main.push(spacer());
main.push(body("That said, the company's balance sheet remains heavily structured around two dominant features: the mobile money pass-through (which is operationally benign but statistically distorts standard ratios) and a large lease liability stack (22% of total assets). Equity at 8% of total assets is thin. Other operating expenses jumped unexpectedly, credit losses surged five-fold in ratio terms, and the company carries significant USD-denominated payables against a Rwandan franc revenue base."));
main.push(spacer());
main.push(heading2("Recommendations for Management"));
main.push(body("Investigate and address the surge in other operating expenses. At 14.64% of revenue and growing, this line needs urgent scrutiny. The CBHI levy is a regulatory cost, but discretionary administrative overhead should be benchmarked and controlled. If the company is scaling its workforce or digital infrastructure to support fintech growth, management should make this explicit in financial guidance."));
main.push(spacer());
main.push(body("Strengthen credit quality in trade receivables. The jump in credit loss expense from 0.22% to 1.09% of revenue — partly from MoMo prefund wallet impairments — reflects systemic quality deterioration. Tighter credit controls on dealer receivables and MoMo agent prefunding arrangements should be prioritised before this erodes profitability further."));
main.push(spacer());
main.push(body("Continue deleveraging on borrowings. Non-current borrowings fell from 9.02% to 5.14% of total assets — that is meaningful progress. Management should prioritise completing the repayment of Syndicate Loan II (final repayment July 2028) and resist the temptation to draw heavily on revolving facilities unless growth capex justifies it."));
main.push(spacer());
main.push(body("Hedge USD exposure. With USD-denominated payables of Rwf 47.2 billion versus USD receivables of Rwf 1.7 billion, the net USD short position is significant. Even a 10% Rwandan franc depreciation would reduce pre-tax profit by approximately Rwf 1.7 billion (the company's own sensitivity disclosure). A partial natural hedge through USD revenue diversification or forward contract coverage is worth evaluating."));
main.push(spacer());
main.push(heading2("Recommendations for Investors"));
main.push(body("The investment case for MTN Rwanda in 2026 is positive but requires patience. The fintech segment — which generated Rwf 39.5 billion in profit in 2025 — is the growth engine, and its trajectory looks strong given Rwanda's digital economy ambitions. Investors should focus on MoMo transaction volume growth, the trajectory of mobile money commissions as a share of revenue, and whether the fintech margin can sustain ~25.7% profit-before-tax margin going forward."));
main.push(spacer());
main.push(body("The absence of a 2025 dividend (the board chose to retain profits to strengthen the balance sheet) is a sensible but short-term disappointment. If profitability continues improving into 2026 and the balance sheet strengthens further, dividend resumption is plausible — Note 10 shows cumulative unpaid dividends of Rwf 12.2 billion still on the books."));
main.push(spacer());
main.push(body("Risk-aware investors should watch the credit loss trend, other operating expense discipline, and any regulatory changes around MoMo that could affect the fee structure or trust account requirements. The current valuation should reflect the discount for thin equity coverage and structural leverage — but the operational improvement trajectory is genuine."));

main.push(spacer(2));

// Section 8 – Conclusion
main.push(heading1("8. CONCLUSION"));
main.push(spacer());
main.push(body("Vertical analysis of MTN Rwandacell Plc's 2025 financial statements reveals a company that has made genuine and visible progress in operational efficiency while managing a business model that is inherently complex — balancing a capital-intensive telecoms infrastructure with a high-volume, lower-margin fintech platform that structurally dominates the balance sheet."));
main.push(spacer());
main.push(body("The income statement shows that revenue growth of 14.3% was partially converted into meaningfully better margins: operating margin grew 4.4 percentage points, and finance costs declined proportionally. The remaining work is in controlling other operating expenses and credit quality. The balance sheet shows a capital-light fintech operation generating large pass-through balances (MoMo), a significant leased network infrastructure, and thin equity — all of which are characteristic of the business model rather than signs of financial distress. The direction of travel is positive."));
main.push(spacer());
main.push(body("For Group BA-05, this analysis demonstrates that vertical (common-size) financial statements are far more informative than raw numbers for understanding cost structure, asset composition, and financial architecture — particularly in a business like MTN Rwanda, where absolute figures are dominated by the scale of the fintech business."));

main.push(spacer(2));
// References
main.push(heading1("9. REFERENCES"));
main.push(spacer());
main.push(body("MTN Rwandacell Plc. (2026). Annual Consolidated Financial Statements for the Year Ended 31 December 2025. Audited by Ernst & Young Rwanda Limited. Kigali: MTN Rwandacell Plc."));
main.push(spacer());
main.push(body("International Financial Reporting Standards (IFRS) as issued by the International Accounting Standards Board (IASB)."));
main.push(spacer());
main.push(body("Law No. 007/2021 of 5/02/2021 Governing Companies in Rwanda, as amended by Law No. 019/2023 of 30 March 2023."));
main.push(spacer());
main.push(body("Brigham, E.F. & Houston, J.F. (2022). Fundamentals of Financial Management (16th ed.). Cengage Learning."));
main.push(spacer());
main.push(body("National Bank of Rwanda. (2023). E-Money Regulations — Reporting and Trust Account Requirements. Kigali: BNR."));

sections.push({
  properties: {
    page: {
      size: { width: 11906, height: 16838 },
      margin: { top: 1440, right: 1440, bottom: 1440, left: 1440 },
    },
  },
  footers: {
    default: new Footer({
      children: [new Paragraph({
        children: [
          new TextRun({ text: "MTN Rwandacell Plc — Vertical Analysis Report  |  Group BA-05  |  ", font: "Calibri", size: 16, color: "888888" }),
          new TextRun({ children: [PageNumber.CURRENT], font: "Calibri", size: 16, color: "888888" }),
        ],
        alignment: AlignmentType.CENTER,
      })],
    }),
  },
  children: main,
});

const doc = new Document({ sections });
Packer.toBuffer(doc).then(buf => {
  fs.writeFileSync("/app/Group5_Vertical_Analysis_MTN_Rwanda_2025.docx", buf);
  console.log("Done");
});
