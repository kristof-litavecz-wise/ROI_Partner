/**
 * ROI Partner — BizOps Investment Model
 * Google Apps Script — Idempotent builder
 * Run: createROIPartnerSheet()
 */

// ── Colour palette ──────────────────────────────────────────────────
const CLR = {
  HEADER_BG:   '#1E2A38',
  HEADER_FG:   '#FFFFFF',
  GREEN:       '#2E7D5B',
  AMBER:       '#C89B3C',
  RED:         '#B94A48',
  BANNER_BG:   '#FFF3CD',
  BANNER_FG:   '#856404',
  WHITE:       '#FFFFFF',
};

// ── Tab definitions ─────────────────────────────────────────────────
const TABS = [
  '00_SYSTEM_RULES',
  '01_INPUT_PROBLEM',
  '02_VALUE_PATH',
  '03_EFFICIENCY_MODEL',
  '04_GROWTH_MODEL',
  '05_COST_MODEL',
  '06_ROI_ENGINE',
  '07_SENSITIVITY',
  '08_REFLECTION',
  '09_FINAL_MEMO',
];

// ── Entry point ─────────────────────────────────────────────────────
function createROIPartnerSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  ss.rename('ROI Partner — BizOps Investment Model');

  ensureTabs_(ss);

  build00SystemRules_(ss);
  build01InputProblem_(ss);
  build02ValuePath_(ss);
  build03EfficiencyModel_(ss);
  build04GrowthModel_(ss);
  build05CostModel_(ss);
  build06ROIEngine_(ss);
  build07Sensitivity_(ss);
  build08Reflection_(ss);
  build09FinalMemo_(ss);

  // Remove default "Sheet1" if it still exists
  const defaultSheet = ss.getSheetByName('Sheet1');
  if (defaultSheet && ss.getSheets().length > 1) {
    ss.deleteSheet(defaultSheet);
  }

  SpreadsheetApp.flush();
  SpreadsheetApp.getUi().alert('✅ ROI Partner build complete.');
}

// ── Helper: ensure all tabs exist in order ──────────────────────────
function ensureTabs_(ss) {
  TABS.forEach((name, i) => {
    let sheet = ss.getSheetByName(name);
    if (!sheet) {
      sheet = ss.insertSheet(name, i);
    } else {
      // Clear existing content for idempotency
      sheet.clear();
      sheet.clearFormats();
      sheet.clearNotes();
      // Remove all existing data validations
      const range = sheet.getDataRange();
      if (range) range.clearDataValidations();
      // Remove existing protections (except sheet-level ones we'll re-add)
      sheet.getProtections(SpreadsheetApp.ProtectionType.RANGE).forEach(p => p.remove());
    }
    sheet.setTabColor(null);
  });
}

// ── Helper: style header row ────────────────────────────────────────
function styleHeader_(sheet, numCols) {
  const hdr = sheet.getRange(1, 1, 1, numCols);
  hdr.setBackground(CLR.HEADER_BG)
     .setFontColor(CLR.HEADER_FG)
     .setFontWeight('bold');
  sheet.setFrozenRows(1);
  for (let c = 1; c <= numCols; c++) {
    sheet.autoResizeColumn(c);
  }
}

// ── Helper: add a dropdown validation ───────────────────────────────
function addDropdown_(sheet, row, col, options) {
  const rule = SpreadsheetApp.newDataValidation()
    .requireValueInList(options, true)
    .setAllowInvalid(false)
    .build();
  sheet.getRange(row, col).setDataValidation(rule);
}

// ── Helper: protect a range ─────────────────────────────────────────
function protectRange_(sheet, a1, description) {
  const protection = sheet.getRange(a1).protect().setDescription(description);
  protection.setWarningOnly(true);
}

// ══════════════════════════════════════════════════════════════════════
// TAB BUILDERS
// ══════════════════════════════════════════════════════════════════════

function build00SystemRules_(ss) {
  const s = ss.getSheetByName('00_SYSTEM_RULES');

  // Banner row
  s.getRange('A1').setValue('⚠️ These assumptions are standardized and cannot be modified. They ensure comparability across all BizOps cases.');
  s.getRange('A1:C1').merge()
    .setBackground(CLR.BANNER_BG)
    .setFontColor(CLR.BANNER_FG)
    .setFontWeight('bold')
    .setWrap(true);

  // Headers
  s.getRange('A2:C2').setValues([['Parameter', 'Value', 'Lock Status']]);
  styleHeader_(s, 3);
  // Unfreeze row 1 (banner), freeze row 2 instead
  s.setFrozenRows(2);

  // Data
  const data = [
    ['Contact Cost',            50,      'LOCK'],
    ['XCCY Take Rate',          0.006,   'LOCK'],
    ['Annualization Factor',    12,      'LOCK'],
    ['Fully Loaded Multiplier', 1.3,     'LOCK'],
    ['Case Cost (Standard)',    'TBD',   'LOCK'],
  ];
  s.getRange(3, 1, data.length, 3).setValues(data);

  // Protect value column (B3:B7)
  protectRange_(s, 'B3:B7', 'System Rules — locked values');

  // Column widths
  s.setColumnWidth(1, 220);
  s.setColumnWidth(2, 120);
  s.setColumnWidth(3, 100);

  s.setTabColor('#1E2A38');
}

function build01InputProblem_(ss) {
  const s = ss.getSheetByName('01_INPUT_PROBLEM');

  s.getRange('A1:B1').setValues([['Field', 'Value']]);
  styleHeader_(s, 2);

  const fields = [
    ['Problem Statement',          ''],
    ['Metric Impacted',            ''],
    ['Baseline Value',             ''],
    ['Target Value',               ''],
    ['Annual Volume',              ''],
    ['Annual Financial Impact (£)',''],
  ];
  s.getRange(2, 1, fields.length, 2).setValues(fields);

  // Note on Problem Statement
  s.getRange('B2').setNote('Must be measurable. No vague phrasing like "efficiency improvement".');

  // Dropdown: Metric Impacted → B3
  addDropdown_(s, 3, 2, ['Time', 'Cost', 'Revenue', 'Risk']);

  s.setColumnWidth(1, 260);
  s.setColumnWidth(2, 300);
}

function build02ValuePath_(ss) {
  const s = ss.getSheetByName('02_VALUE_PATH');

  s.getRange('A1:B1').setValues([['Field', 'Value']]);
  styleHeader_(s, 2);

  s.getRange('A2').setValue('Selected Value Path');
  addDropdown_(s, 2, 2, ['Efficiency', 'Growth', 'Risk']);

  s.getRange('A4').setValue('Only ONE path allowed. No blending models.')
    .setFontStyle('italic')
    .setFontColor('#888888');

  s.setColumnWidth(1, 220);
  s.setColumnWidth(2, 200);
}

function build03EfficiencyModel_(ss) {
  const s = ss.getSheetByName('03_EFFICIENCY_MODEL');

  s.getRange('A1:B1').setValues([['Field', 'Value']]);
  styleHeader_(s, 2);

  // Inputs
  s.getRange('A2').setValue('Contacts Reduced');
  s.getRange('A3').setValue('Cases Reduced');

  // System constants (linked)
  s.getRange('A5').setValue('Contact Cost (system)');
  s.getRange('B5').setFormula("='00_SYSTEM_RULES'!B3");

  s.getRange('A6').setValue('Case Cost (system)');
  s.getRange('B6').setFormula("='00_SYSTEM_RULES'!B7");

  // Calculations
  s.getRange('A8').setValue('Contact Savings');
  s.getRange('B8').setFormula('=B2*B5');

  s.getRange('A9').setValue('Case Savings');
  s.getRange('B9').setFormula('=B3*B6');

  s.getRange('A10').setValue('Total Efficiency Value');
  s.getRange('B10').setFormula('=B8+B9');
  s.getRange('B10').setFontWeight('bold');

  s.setColumnWidth(1, 240);
  s.setColumnWidth(2, 180);
}

function build04GrowthModel_(ss) {
  const s = ss.getSheetByName('04_GROWTH_MODEL');

  s.getRange('A1:B1').setValues([['Field', 'Value']]);
  styleHeader_(s, 2);

  s.getRange('A2').setValue('XCCY Volume');

  s.getRange('A3').setValue('Take Rate (system)');
  s.getRange('B3').setFormula("='00_SYSTEM_RULES'!B4");

  s.getRange('A4').setValue('Revenue Impact');
  s.getRange('B4').setFormula('=B2*B3');

  s.getRange('A5').setValue('Annual Value');
  s.getRange('B5').setFormula("=B4*'00_SYSTEM_RULES'!B5");
  s.getRange('B5').setFontWeight('bold');

  s.setColumnWidth(1, 220);
  s.setColumnWidth(2, 180);
}

function build05CostModel_(ss) {
  const s = ss.getSheetByName('05_COST_MODEL');

  s.getRange('A1:B1').setValues([['Field', 'Value']]);
  styleHeader_(s, 2);

  s.getRange('A2').setValue('Role Level');
  addDropdown_(s, 2, 2, ['Analyst', 'Associate', 'Manager', 'Senior Manager', 'Director']);

  s.getRange('A3').setValue('Base Salary');

  s.getRange('A4').setValue('Fully Loaded Multiplier (system)');
  s.getRange('B4').setFormula("='00_SYSTEM_RULES'!B6");

  s.getRange('A5').setValue('Fully Loaded Cost');
  s.getRange('B5').setFormula('=B3*B4');
  s.getRange('B5').setFontWeight('bold');

  s.setColumnWidth(1, 280);
  s.setColumnWidth(2, 180);
}

function build06ROIEngine_(ss) {
  const s = ss.getSheetByName('06_ROI_ENGINE');

  s.getRange('A1:B1').setValues([['Field', 'Value']]);
  styleHeader_(s, 2);

  // Pull values
  s.getRange('A2').setValue('Selected Path');
  s.getRange('B2').setFormula("='02_VALUE_PATH'!B2");

  s.getRange('A3').setValue('Efficiency Value');
  s.getRange('B3').setFormula("='03_EFFICIENCY_MODEL'!B10");

  s.getRange('A4').setValue('Growth Value');
  s.getRange('B4').setFormula("='04_GROWTH_MODEL'!B5");

  s.getRange('A5').setValue('Fully Loaded Cost');
  s.getRange('B5').setFormula("='05_COST_MODEL'!B5");

  // ROI logic
  s.getRange('A7').setValue('Annual Value');
  s.getRange('B7').setFormula('=IF(B2="Efficiency",B3,B4)');
  s.getRange('B7').setFontWeight('bold');

  s.getRange('A8').setValue('ROI Expected');
  s.getRange('B8').setFormula('=IFERROR(B7/B5,0)');
  s.getRange('B8').setFontWeight('bold').setNumberFormat('0.0x');

  s.getRange('A9').setValue('ROI Conservative');
  s.getRange('B9').setFormula('=B8*0.7');
  s.getRange('B9').setFontWeight('bold').setNumberFormat('0.0x');

  // Conditional formatting — ROI Expected (B8)
  applyROIConditionalFormat_(s, 'B8');
  applyROIConditionalFormat_(s, 'B9');

  // Protect formula cells
  protectRange_(s, 'B2:B9', 'ROI Engine — formulas locked');

  s.setColumnWidth(1, 220);
  s.setColumnWidth(2, 180);

  s.setTabColor('#2E7D5B');
}

function applyROIConditionalFormat_(sheet, cellA1) {
  const range = sheet.getRange(cellA1);
  const rules = sheet.getConditionalFormatRules();

  // Green: >= 3
  rules.push(SpreadsheetApp.newConditionalFormatRule()
    .whenNumberGreaterThanOrEqualTo(3)
    .setBackground(CLR.GREEN)
    .setFontColor(CLR.WHITE)
    .setRanges([range])
    .build());

  // Amber: >= 2
  rules.push(SpreadsheetApp.newConditionalFormatRule()
    .whenNumberBetween(2, 2.999999)
    .setBackground(CLR.AMBER)
    .setFontColor(CLR.WHITE)
    .setRanges([range])
    .build());

  // Red: < 2
  rules.push(SpreadsheetApp.newConditionalFormatRule()
    .whenNumberLessThan(2)
    .setBackground(CLR.RED)
    .setFontColor(CLR.WHITE)
    .setRanges([range])
    .build());

  sheet.setConditionalFormatRules(rules);
}

function build07Sensitivity_(ss) {
  const s = ss.getSheetByName('07_SENSITIVITY');

  s.getRange('A1:B1').setValues([['Field', 'Value']]);
  styleHeader_(s, 2);

  s.getRange('A2').setValue('Primary Driver');
  s.getRange('A3').setValue('Most Sensitive Variable');
  s.getRange('A4').setValue('Risk Level');

  addDropdown_(s, 4, 2, ['Low', 'Medium', 'High']);

  s.getRange('A6').setValue('Identify what changes ROI most significantly.')
    .setFontStyle('italic')
    .setFontColor('#888888');

  s.setColumnWidth(1, 240);
  s.setColumnWidth(2, 200);
}

function build08Reflection_(ss) {
  const s = ss.getSheetByName('08_REFLECTION');

  s.getRange('A1').setValue('Reflection Prompt');
  styleHeader_(s, 1);

  const prompts = [
    ['Why can\'t this be solved by product?'],
    ['What breaks if nothing is done?'],
    ['Is this scaling or local inefficiency?'],
    ['Which assumption is least defensible?'],
  ];
  s.getRange(2, 1, prompts.length, 1).setValues(prompts);

  s.getRange('B1').setValue('Response');
  s.getRange('B1').setBackground(CLR.HEADER_BG)
    .setFontColor(CLR.HEADER_FG)
    .setFontWeight('bold');

  s.setColumnWidth(1, 360);
  s.setColumnWidth(2, 400);
}

function build09FinalMemo_(ss) {
  const s = ss.getSheetByName('09_FINAL_MEMO');

  s.getRange('A1:B1').setValues([['Section', 'Value']]);
  styleHeader_(s, 2);

  s.getRange('A2').setValue('Problem Summary');
  s.getRange('B2').setFormula("='01_INPUT_PROBLEM'!B2");

  s.getRange('A3').setValue('Value Path');
  s.getRange('B3').setFormula("='02_VALUE_PATH'!B2");

  s.getRange('A4').setValue('Annual Value');
  s.getRange('B4').setFormula("='06_ROI_ENGINE'!B7");
  s.getRange('B4').setNumberFormat('£#,##0');

  s.getRange('A5').setValue('ROI Expected');
  s.getRange('B5').setFormula("='06_ROI_ENGINE'!B8");
  s.getRange('B5').setNumberFormat('0.0x');

  s.getRange('A6').setValue('ROI Conservative');
  s.getRange('B6').setFormula("='06_ROI_ENGINE'!B9");
  s.getRange('B6').setNumberFormat('0.0x');

  s.getRange('A7').setValue('Recommendation');
  s.getRange('B7').setFormula('=IF(B5>=3,"Strong Case",IF(B5>=2,"Borderline","Weak Case"))');
  s.getRange('B7').setFontWeight('bold').setFontSize(12);

  // Conditional format on recommendation
  const recRange = s.getRange('B7');
  const rules = [];
  rules.push(SpreadsheetApp.newConditionalFormatRule()
    .whenTextEqualTo('Strong Case')
    .setBackground(CLR.GREEN).setFontColor(CLR.WHITE)
    .setRanges([recRange]).build());
  rules.push(SpreadsheetApp.newConditionalFormatRule()
    .whenTextEqualTo('Borderline')
    .setBackground(CLR.AMBER).setFontColor(CLR.WHITE)
    .setRanges([recRange]).build());
  rules.push(SpreadsheetApp.newConditionalFormatRule()
    .whenTextEqualTo('Weak Case')
    .setBackground(CLR.RED).setFontColor(CLR.WHITE)
    .setRanges([recRange]).build());
  s.setConditionalFormatRules(rules);

  s.setColumnWidth(1, 200);
  s.setColumnWidth(2, 300);

  s.setTabColor('#1E2A38');
}
