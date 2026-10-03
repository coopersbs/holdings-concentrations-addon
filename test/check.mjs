#!/usr/bin/env node
/**
 * The test harness of src/concentration.gs, src/layout.gs, and
 * src/sidebar.html.
 *
 * The harness loads the two script files into one node:vm context with fakes
 * of the Apps Script services. The spreadsheet fake starts with a synthetic
 * Holdings tab alone. It holds each tab as a grid in memory, with the styles,
 * the column widths, the conditional formats, and the data validation that
 * the script sets. No Google service takes part.
 *
 * The harness compares the tabs that the script creates with the accepted
 * layout of the report. The definition files in test/fixtures/ hold that
 * layout.
 *
 * The harness calculates the two run-time formulas B9 and B10 and the
 * coverage label B12 of the report alone, with a small evaluator. It reads
 * the text of each other formula. No formula can hold a column letter of the
 * Holdings tab. No SUM can read a name of a LET that holds TRUE and FALSE,
 * because SUM adds no TRUE in an array. Each range of the
 * sources block must end at the column of position MAX_POSITIONS. Each
 * FILTER must sit in IFNA or IFERROR, or GUARDED_FILTERS must name it.
 *
 * The script computes the direct weights, the unseen block, the stock fund
 * block, the own block, and the group block from each answer. The harness
 * computes the same values from the tab with the steps of the report
 * formulas that computed them before, and compares the two. It also checks
 * that the script writes the two data blocks alone, clears the cells of the
 * last answer outside them, and fits the rows of the grid to the answer.
 *
 * Each good run draws the two bar charts of the report tab again: the
 * company chart, then the holdings chart. The fake of the chart builder
 * keeps the ranges, the colors, the title, the anchor, and the options of a
 * chart. The log of the fake names the block of each chart that the script
 * inserts or removes. The harness reads the formulas of the two chart
 * blocks and of the two anchor cells as text, and it checks that they read
 * the report tab alone. At each SpreadsheetApp.flush, the spreadsheet fake
 * calculates the header row of each chart block and each anchor cell with
 * the rules of those formulas: the series of the company table for the
 * threshold, the share series when the section Holdings holds a
 * holding group, and the row under the header row of each chart in the
 * spill. It also puts the text of each header row into column A of the
 * report tab, in the row above the anchor row. getValues gives the
 * calculated value of such a cell in place of its formula. The harness
 * checks that each chart reads each column whose header holds a text, that
 * it anchors at the calculated row, and that a draw of one chart removes
 * no chart of the other block.
 *
 * The simple trigger onEdit gets fake edit events. The harness checks that an
 * edit of the threshold cell draws the company chart again, that each other
 * edit and each value outside 0 to 1 draws nothing, and that onEdit catches
 * an error of the draw and gives it to console.error. It checks that onEdit
 * with no event object returns and logs nothing, and that a chart header row
 * with a comma as the decimal mark matches. A lag fake gives the old value of
 * the anchor cell and of the chart header row for a count of reads, as a slow
 * spill does right after an edit. The harness checks that the draw waits for
 * the fresh values, and that a header that stays old for the whole wait
 * leaves the chart as it is. A draw of the holdings chart alone gets the same
 * checks of the wait, and a check of an empty share header. The fake of
 * Utilities.sleep records each wait and returns at once.
 *
 * The last runs put tabs of another layout version into the spreadsheet
 * fake. They check that a refresh replaces those tabs and keeps the values
 * that the person typed.
 *
 * The fund mix runs call the server functions of the sidebar: the data of
 * the sidebar, the ticker check, the save, the link, and the delete of a mix.
 * They check the static rules of the sidebar page. Then a refresh sends a
 * position with parts and gets a synthetic answer in the shape of the route
 * with parts.
 *
 * Runs 18 and 19 get synthetic answers in the shape of schema version 1.9.
 * The answer of run 18 holds a held fund, a fund that did not enter, and a
 * line of the class preferred. The answer of run 19 keeps 5 lines and adds
 * the cap line other:lines. Run 16 gets the answer of run 1 in the shape of
 * schema version 1.8, with no entered key and no preferredWeight. Run 20
 * gets a synthetic answer with the line classes trust and cash of a direct
 * position: a trust that the person holds, a trust in a fund mix, and a
 * money market fund with no element of the funds block.
 *
 * Run 1 sends one real request to the concentration route with the key of the
 * environment variable HOLDINGS_API_KEY. The request holds no parts. The other
 * runs send no request. The harness prints no part of the key.
 *
 * The option --live-parts also sends the request of the fund mix run to the
 * route, so the harness sends two real requests. Use it after the route
 * accepts parts.
 *
 * The option --offline skips each real request. Run 1 then gets a synthetic
 * answer, and the harness uses an invented key. Each other assertion runs.
 * Use it when no key is available, such as in a pull request from a fork.
 *
 * Usage, from the repository root:
 *
 *   HOLDINGS_API_KEY=<your key> node test/check.mjs
 *   HOLDINGS_API_KEY=<your key> node test/check.mjs --live-parts
 *   node test/check.mjs --offline
 *
 * Exit codes: 0 = each assertion passed, 1 = an assertion failed, 2 = no key.
 */

import { spawnSync } from "node:child_process";
import { readFileSync } from "node:fs";
import { dirname, resolve } from "node:path";
import { fileURLToPath } from "node:url";
import vm from "node:vm";

const HERE = dirname(fileURLToPath(import.meta.url));
const SRC = resolve(HERE, "..", "src");
const SCRIPT_FILES = ["concentration.gs", "layout.gs"].map((name) => resolve(SRC, name));
const SIDEBAR_FILE = resolve(SRC, "sidebar.html");
const MANIFEST_FILE = resolve(SRC, "appsscript.json");
const FUND_URL = "https://data.coopersbs.com/funds/v1/";
const ROUTE_URL = `${FUND_URL}concentration`;
const EXPOSURE_TAB = "Concentration.Exposure";
const REPORT_TAB = "Concentration";

/**
 * The user property that holds the key.
 */
const KEY_PROPERTY = "HOLDINGS_API_KEY";

/**
 * The fixture of the accepted layout of each of the two tabs.
 */
const ACCEPTED_FILES = {
  [EXPOSURE_TAB]: resolve(HERE, "fixtures", "concentration-exposure.json"),
  [REPORT_TAB]: resolve(HERE, "fixtures", "concentration.json"),
};

/**
 * The start of the name of each document property that holds a fund mix.
 */
const MIX_PREFIX = "FUND_MIX:";

/**
 * The time zone of the spreadsheet fake.
 */
const TIME_ZONE = "America/Chicago";

/**
 * The three scopes of the manifest, and the one URL prefix that the script
 * can fetch. The sidebar needs script.container.ui.
 */
const SCOPES = [
  "https://www.googleapis.com/auth/spreadsheets.currentonly",
  "https://www.googleapis.com/auth/script.external_request",
  "https://www.googleapis.com/auth/script.container.ui",
];
const URL_PREFIXES = ["https://data.coopersbs.com/funds/v1/"];

/**
 * The size of the grid of a new tab, as Apps Script makes it.
 */
const NEW_ROWS = 1000;
const NEW_COLUMNS = 26;

/**
 * True when the harness skips the real request.
 */
const OFFLINE = process.argv.includes("--offline");

/**
 * True when the fund mix run also sends its request to the route. The
 * option --offline wins.
 */
const LIVE_PARTS = !OFFLINE && process.argv.includes("--live-parts");

/**
 * The count of real requests that the harness sends.
 */
const LIVE_LIMIT = OFFLINE ? 0 : LIVE_PARTS ? 2 : 1;

/**
 * The key of the route. The harness reads it from the environment alone. The
 * offline run uses an invented key in the key format of the route, and it
 * does not read the environment.
 */
const API_KEY = OFFLINE ? `csbs_${"0".repeat(43)}` : (process.env.HOLDINGS_API_KEY ?? "").trim();

/**
 * The synthetic Holdings tab. Each account name, description, and value is
 * invented. The tickers are public symbols that the route knows: two index
 * funds that hold one stock, the stock itself, and a money market fund. The
 * bond identifier is invented, so the route does not know it. The column
 * order differs from a Tiller tab. The values are round, so a reader can
 * check each weight by hand. The two rows of fund A add to one position.
 */
const HOLDINGS_HEADER = ["Symbol", "Account", "Value", "Description"];
const FUND_A = "VOO";
const FUND_B = "IVV";
const STOCK = "AAPL";
const MONEY = "SPAXX";
const PLAN_FUND = "Example Plan Collective Trust";
const BOND = "912810ZZ1";
const TRUST = "SPY";
const MIX_TRUST = "GLD";
const HOLDINGS_ROWS = [
  [FUND_A, "Example brokerage", 3000, "Example index fund A"],
  [FUND_B, "Example brokerage", 2000, "Example index fund B"],
  [STOCK, "Example brokerage", 1000, "Example company stock"],
  [FUND_A, "Example retirement", 1000, "Example index fund A"],
  [BOND, "Example brokerage", 1000, "Example Treasury bond"],
  [MONEY, "Example brokerage", 1000, "Example money market fund"],
  ["", "Example plan", 1000, PLAN_FUND],
  ["", "Example brokerage", 500, ""],
  ["XYZ", "Example brokerage", "n/a", "A row with a value that is not a number"],
];

/**
 * The hand calculation: the sum of each key and its share of the total of
 * 10,000. The empty key and the row with no number add nothing.
 */
const EXPECTED = [
  { id: FUND_A, ticker: FUND_A, sum: 3000 + 1000 },
  { id: FUND_B, ticker: FUND_B, sum: 2000 },
  { id: STOCK, ticker: STOCK, sum: 1000 },
  { id: BOND, ticker: BOND, sum: 1000 },
  { id: MONEY, ticker: MONEY, sum: 1000 },
  { id: PLAN_FUND, ticker: null, sum: 1000 },
];
const EXPECTED_TOTAL = 10000;

const MEASURES = [
  "lineCount",
  "top10Weight",
  "hhi",
  "effectiveCount",
  "lookedThroughWeight",
  "notLookedThroughWeight",
  "weightSum",
  "weightDifference",
  "unknownWeight",
  "preferredWeight",
];
const EQUITY = ["weight", "lineCount", "top10Weight", "hhi", "effectiveCount"];
const FUND_FIELDS = [
  "id",
  "ticker",
  "reportDate",
  "accessionNumber",
  "holdingCount",
  "weight",
  "coveredWeight",
  "mergedByTicker",
  "mergedByLei",
  "mergedByName",
  "partWeight",
];
const MIX_FIELDS = ["mixId", "mixWeight", "entered", "partTicker", "partWeight", "substitute"];
const LINE_FIELDS = ["key", "name", "ticker", "lei", "class", "weight", "stockWeight"];

/**
 * The first column of each block of the Concentration.Exposure tab: the
 * funds in D, the overlaps in P, the mixes in U, the unseen block in AB, the
 * stock fund block in AE, the own block in AG, the group block in AM, the
 * lines in AR, the part name in AY, the direct weight in AZ, and the sources
 * in BA. Column 1 is A.
 */
const FUND_AT = 4;
const OVERLAP_AT = 16;
const MIX_AT = 21;
const UNSEEN_AT = 28;
const STOCK_FUND_AT = 31;
const OWN_AT = 33;
const GROUP_AT = 39;
const LINE_AT = 44;
const PART_AT = LINE_AT + LINE_FIELDS.length;
const DIRECT_AT = PART_AT + 1;
const SOURCE_AT = DIRECT_AT + 1;

/**
 * The count of position columns of the sources block, and the count of rows
 * of a new tab.
 */
const MAX_POSITIONS = 200;
const TAB_ROWS = 1000;

/**
 * The chart block of the Concentration.Exposure tab: its first column IT,
 * one empty column after the column of position MAX_POSITIONS, the count of
 * data rows under the header in row 4, the count of fund columns, and the
 * count of columns. The columns are the company name, Direct, the fund
 * columns, and Other funds.
 */
const CHART_AT = SOURCE_AT + MAX_POSITIONS + 1;
const CHART_ROWS = 100;
const CHART_FUNDS = 5;
const CHART_WIDTH = CHART_FUNDS + 3;

/**
 * The column of the anchor cell JC5, one empty column after the chart block,
 * and the count of blank rows of the band under the chart header row of the
 * report spill.
 */
const ANCHOR_AT = CHART_AT + CHART_WIDTH + 1;
const BAND_ROWS = 26;

/**
 * The holdings chart block of the Concentration.Exposure tab: its first
 * column JE, one empty column after the anchor cell of the company chart,
 * the count of data rows under the header in row 4, and the count of
 * columns. The columns are the label of each holding group and its share of
 * the portfolio. The holdings anchor cell JH5 is one empty column after the
 * block. The text of the header row of the holdings chart in the spill is
 * fixed.
 */
const HOLDINGS_AT = ANCHOR_AT + 2;
const HOLDINGS_CHART_ROWS = MAX_POSITIONS;
const HOLDINGS_WIDTH = 2;
const HOLDINGS_ANCHOR_AT = HOLDINGS_AT + HOLDINGS_WIDTH + 1;
const HOLDINGS_HEADING = "Holdings by share of the portfolio";

/**
 * The fixed colors of the series of the bar chart: Direct, the five funds in
 * rank order, and Other funds.
 */
const DIRECT_COLOR = "#2a78d6";
const FUND_COLORS = ["#f26e3b", "#037952", "#eba007", "#c75d87", "#036503"];
const OTHER_COLOR = "#a9a7a0";

/**
 * The options that the script sets on the bar chart with setOption. Apps
 * Script checks no option name, so the harness checks each name and value.
 * The width is the width of the columns A to D of the report tab.
 */
const CHART_OPTIONS = {
  "annotations.total.enabled": true,
  hAxis: { format: "0%" },
  height: 520,
  legend: { position: "top" },
  useFirstColumnAsDomain: true,
  width: 250 + 300 + 150 + 175,
};

/**
 * The options that the script sets on the holdings chart: no legend, and
 * the value of the one series as a data label at the end of each bar.
 */
const HOLDINGS_OPTIONS = {
  hAxis: { format: "0%" },
  height: 520,
  legend: { position: "none" },
  series: { 0: { dataLabel: "value", dataLabelPlacement: "outsideEnd", hasAnnotations: true } },
  useFirstColumnAsDomain: true,
  width: 250 + 300 + 150 + 175,
};

/**
 * The header row and the first row of the run-time block of the
 * Concentration.Exposure tab: one row under the last measure.
 */
const RUN_HEADER = 5 + MEASURES.length;
const RUN_ROW = RUN_HEADER + 1;

/**
 * The header row and the first row of the equity block of the
 * Concentration.Exposure tab: one empty row under the run-time block.
 */
const EQUITY_HEADER = RUN_HEADER + 10 + 2;
const EQUITY_ROW = EQUITY_HEADER + 1;

/**
 * The cell of the Concentration.Exposure tab that holds the layout version,
 * and the two cells of the Concentration tab that a person types in.
 */
const VERSION_CELL = "B3";
const THRESHOLD_CELL = "B23";
const MINIMUM_CELL = "B24";

/**
 * The cells of the Concentration tab: the disclaimer, the subtitle, the total
 * value, the note under the status cell, the coverage label, and the report
 * spill.
 */
const DISCLAIMER_CELL = "B1";
const SUBTITLE_CELL = "B2";
const TOTAL_CELL = "B6";
const NOTE_CELL = "B4";
const COVERAGE_CELL = "B12";
const SPILL_CELL = "A26";
const REPORT_ROW = 26;

/**
 * The texts of the report that a person reads about the funds not looked
 * through.
 */
const UNSEEN_TITLE = "Funds not looked through";
const ADD_MIX_NOTE = "Add its fund mix: Concentration › Describe a fund";

/**
 * The text of the sidebar for a ticker that the fund route answers with
 * not_a_fund: a company stock or a trust.
 */
const NOT_A_FUND = "A company stock or a trust, not a fund. It counts as one holding.";

/**
 * The error of a failed assertion. The message names the assertion.
 */
class AssertionFailure extends Error {}

let passed = 0;

/**
 * Stop the harness with the name of the assertion when the condition is
 * false.
 */
function check(condition, name) {
  if (!condition) throw new AssertionFailure(name);
  passed += 1;
}

/**
 * True when two numbers differ by less than 1e-12.
 */
function near(a, b) {
  return typeof a === "number" && Math.abs(a - b) < 1e-12;
}

/**
 * True when the value is a Date of any realm. The script makes its Date in
 * the vm context, so `instanceof Date` of this realm is false.
 */
function isDate(value) {
  return Object.prototype.toString.call(value) === "[object Date]";
}

/**
 * The column number of a column letter, such as 1 for A and 27 for AA.
 */
function columnNumber(letters) {
  return [...letters].reduce((n, c) => n * 26 + c.charCodeAt(0) - 64, 0);
}

/**
 * The bounds of an A1 range inside a grid of the given size. An open end,
 * such as E18:E or U:AZ, runs to the edge of the grid.
 */
function parseA1(a1, maxRows, maxColumns) {
  const m = /^([A-Z]+)?(\d+)?(?::([A-Z]+)?(\d+)?)?$/.exec(a1);
  if (!m || (!m[1] && !m[2])) throw new Error(`The range "${a1}" is not an A1 range.`);
  const single = !a1.includes(":");
  const row = m[2] ? Number(m[2]) : 1;
  const column = m[1] ? columnNumber(m[1]) : 1;
  const endRowText = single ? m[2] : m[4];
  const endColumnText = single ? m[1] : m[3];
  const endRow = endRowText ? Number(endRowText) : maxRows;
  const endColumn = endColumnText ? columnNumber(endColumnText) : maxColumns;
  return { row, column, rows: endRow - row + 1, columns: endColumn - column + 1 };
}

/**
 * The key of one cell in a style map or a validation map.
 */
function cellKey(row, column) {
  return `${row},${column}`;
}

/**
 * A grid of rows and columns in memory with the methods of an Apps Script
 * Sheet that the script calls. Each change goes into the log of the fake.
 */
class FakeSheet {
  constructor(name, rows, columns, log) {
    this.name = name;
    this.log = log;
    this.grid = Array.from({ length: rows }, () => Array(columns).fill(""));
    this.styles = new Map();
    this.validation = new Map();
    this.widths = new Map();
    this.rules = [];
    this.hidden = false;
    this.frozenRows = 0;
    this.charts = [];
    this.calculated = new Map();
  }

  record(op, extra = {}) {
    this.log.push({ sheet: this.name, op, ...extra });
  }

  getName() {
    return this.name;
  }

  getMaxRows() {
    return this.grid.length;
  }

  getMaxColumns() {
    return this.grid[0].length;
  }

  insertRowsAfter(after, count) {
    if (after !== this.getMaxRows()) throw new Error("The fake inserts rows at the end of the grid alone.");
    for (let i = 0; i < count; i += 1) this.grid.push(Array(this.getMaxColumns()).fill(""));
    this.record("insertRowsAfter", { count });
  }

  insertColumnsAfter(after, count) {
    if (after !== this.getMaxColumns()) throw new Error("The fake inserts columns at the end of the grid alone.");
    for (const row of this.grid) row.push(...Array(count).fill(""));
    this.record("insertColumnsAfter", { count });
  }

  deleteRows(start, count) {
    if (start + count - 1 !== this.getMaxRows()) throw new Error("The fake deletes rows at the end of the grid alone.");
    this.grid.splice(start - 1, count);
    this.record("deleteRows", { count });
  }

  deleteColumns(start, count) {
    if (start + count - 1 !== this.getMaxColumns()) {
      throw new Error("The fake deletes columns at the end of the grid alone.");
    }
    for (const row of this.grid) row.splice(start - 1, count);
    this.record("deleteColumns", { count });
  }

  getIndex() {
    return book.sheets.indexOf(this) + 1;
  }

  hideSheet() {
    this.hidden = true;
    this.record("hideSheet");
  }

  setFrozenRows(rows) {
    this.frozenRows = rows;
    this.record("setFrozenRows", { rows });
  }

  setColumnWidths(start, count, width) {
    for (let c = start; c < start + count; c += 1) this.widths.set(c, width);
    this.record("setColumnWidths", { start, count, width });
  }

  setConditionalFormatRules(rules) {
    this.rules = rules.map((rule) => ({ ...rule }));
    this.record("setConditionalFormatRules", { count: rules.length });
  }

  getCharts() {
    return [...this.charts];
  }

  newChart() {
    return fakeChartBuilder();
  }

  insertChart(chart) {
    this.charts.push(chart);
    this.record("insertChart", { block: chartBlock(chart) });
  }

  removeChart(chart) {
    const index = this.charts.indexOf(chart);
    if (index < 0) throw new Error(`The sheet "${this.name}" holds no such chart.`);
    this.charts.splice(index, 1);
    this.record("removeChart", { block: chartBlock(chart) });
  }

  getRange(first, column, rows = 1, columns = 1) {
    let bounds = { row: first, column, rows, columns };
    if (typeof first === "string") bounds = parseA1(first, this.getMaxRows(), this.getMaxColumns());
    const b = bounds;
    if (b.row < 1 || b.column < 1 || b.rows < 1 || b.columns < 1) {
      throw new Error(`The range ${b.row},${b.column},${b.rows},${b.columns} has a bad start or a bad size.`);
    }
    if (b.row + b.rows - 1 > this.getMaxRows() || b.column + b.columns - 1 > this.getMaxColumns()) {
      throw new Error("The coordinates of the range are outside the dimensions of the sheet.");
    }
    return new FakeRange(this, b.row, b.column, b.rows, b.columns);
  }

  /**
   * The last row that holds a value, or 0 for an empty grid.
   */
  getLastRow() {
    let last = 0;
    this.grid.forEach((cells, r) => {
      if (cells.some((value) => value !== "")) last = r + 1;
    });
    return last;
  }

  /**
   * The last column that holds a value, or 0 for an empty grid.
   */
  getLastColumn() {
    let last = 0;
    for (const cells of this.grid) {
      cells.forEach((value, c) => {
        if (value !== "") last = Math.max(last, c + 1);
      });
    }
    return last;
  }

  getDataRange() {
    let lastRow = 1;
    let lastColumn = 1;
    this.grid.forEach((cells, r) => {
      cells.forEach((value, c) => {
        if (value !== "") {
          lastRow = Math.max(lastRow, r + 1);
          lastColumn = Math.max(lastColumn, c + 1);
        }
      });
    });
    return this.getRange(1, 1, lastRow, lastColumn);
  }

  /**
   * The value that a cell shows, by its row and its column: the calculated
   * value of a formula when the fake calculated one, else the content of the
   * cell.
   */
  shown(row, column) {
    const key = cellKey(row, column);
    return this.calculated.has(key) ? this.calculated.get(key) : this.grid[row - 1][column - 1];
  }

  /**
   * The value that the script reads from a cell, by its row and its column.
   * A cell of state.lagging gives its old value while it has reads left,
   * and each read uses one. Each other cell gives the value of shown.
   */
  read(row, column) {
    const lag = state.lagging.size === 0 ? undefined : state.lagging.get(`${this.name}!${cellKey(row, column)}`);
    if (lag !== undefined && lag.reads > 0) {
      lag.reads -= 1;
      return lag.value;
    }
    return this.shown(row, column);
  }

  /**
   * The value of one cell by its A1 name, such as B1.
   */
  cell(a1) {
    const b = parseA1(a1, this.getMaxRows(), this.getMaxColumns());
    return this.grid[b.row - 1]?.[b.column - 1];
  }

  /**
   * The values of a block of cells, with row 1 and column 1 as the start.
   */
  block(row, column, rows, columns) {
    return this.grid.slice(row - 1, row - 1 + rows).map((cells) => cells.slice(column - 1, column - 1 + columns));
  }

  /**
   * A copy of the grid, with each Date as its time in milliseconds.
   */
  snapshot() {
    return this.grid.map((cells) => cells.map((v) => (isDate(v) ? v.getTime() : v)));
  }

  /**
   * A copy of the whole state of the tab as JSON text: the grid, the styles,
   * the validation, the column widths, the conditional formats, the hidden
   * flag, and the frozen rows.
   */
  state() {
    const sorted = (map) => [...map.entries()].sort(([a], [b]) => (a < b ? -1 : a > b ? 1 : 0));
    return JSON.stringify({
      grid: this.snapshot(),
      styles: sorted(this.styles).map(([key, style]) => [key, Object.fromEntries(Object.entries(style).sort())]),
      validation: sorted(this.validation),
      widths: sorted(this.widths),
      rules: this.rules.map((rule) => ({ ...rule, ranges: rule.ranges.map((r) => r.bounds()) })),
      hidden: this.hidden,
      frozenRows: this.frozenRows,
    });
  }
}

/**
 * A range of a FakeSheet with the methods of an Apps Script Range that the
 * script calls.
 */
class FakeRange {
  constructor(sheet, row, column, rows, columns) {
    Object.assign(this, { sheet, row, column, rows, columns });
  }

  bounds() {
    const { row, column, rows, columns } = this;
    return { row, column, rows, columns };
  }

  getSheet() {
    return this.sheet;
  }

  getRow() {
    return this.row;
  }

  getColumn() {
    return this.column;
  }

  getNumRows() {
    return this.rows;
  }

  getNumColumns() {
    return this.columns;
  }

  record(op, extra = {}) {
    this.sheet.record(op, { ...this.bounds(), ...extra });
  }

  /**
   * Call the function for the key of each cell of the range.
   */
  eachCell(fn) {
    for (let r = this.row; r < this.row + this.rows; r += 1) {
      for (let c = this.column; c < this.column + this.columns; c += 1) fn(cellKey(r, c));
    }
  }

  /**
   * Set one style property on each cell of the range and log the call.
   */
  style(property, value, op) {
    this.eachCell((key) => {
      const style = this.sheet.styles.get(key) ?? {};
      style[property] = value;
      this.sheet.styles.set(key, style);
    });
    this.record(op, { value });
    return this;
  }

  getValues() {
    return Array.from({ length: this.rows }, (_, r) =>
      Array.from({ length: this.columns }, (_, c) => this.sheet.read(this.row + r, this.column + c)),
    );
  }

  getValue() {
    return this.sheet.read(this.row, this.column);
  }

  setValues(values) {
    if (values.length !== this.rows || values.some((cells) => cells.length !== this.columns)) {
      throw new Error(`The data has ${values.length} rows, and the range has ${this.rows} rows.`);
    }
    values.forEach((cells, r) => {
      cells.forEach((value, c) => {
        this.sheet.grid[this.row - 1 + r][this.column - 1 + c] = value;
      });
    });
    this.eachCell((key) => this.sheet.calculated.delete(key));
    this.record("setValues");
    return this;
  }

  clearContent() {
    for (let r = this.row; r < this.row + this.rows; r += 1) {
      for (let c = this.column; c < this.column + this.columns; c += 1) this.sheet.grid[r - 1][c - 1] = "";
    }
    this.eachCell((key) => this.sheet.calculated.delete(key));
    this.record("clearContent");
    return this;
  }

  setNumberFormat(format) {
    return this.style("numberFormat", format, "setNumberFormat");
  }

  setFontWeight(value) {
    return this.style("fontWeight", value, "setFontWeight");
  }

  setFontStyle(value) {
    return this.style("fontStyle", value, "setFontStyle");
  }

  setFontSize(value) {
    return this.style("fontSize", value, "setFontSize");
  }

  setFontColor(value) {
    return this.style("fontColor", value, "setFontColor");
  }

  setBackground(value) {
    return this.style("background", value, "setBackground");
  }

  setHorizontalAlignment(value) {
    return this.style("horizontalAlignment", value, "setHorizontalAlignment");
  }

  setWrap(value) {
    return this.style("wrap", value, "setWrap");
  }

  setDataValidation(rule) {
    this.eachCell((key) => this.sheet.validation.set(key, { ...rule }));
    this.record("setDataValidation");
    return this;
  }
}

/**
 * A spreadsheet in memory with the methods of an Apps Script Spreadsheet that
 * the script calls. The tabs keep their order.
 */
class FakeBook {
  constructor(log) {
    this.log = log;
    this.sheets = [];
  }

  getSheetByName(name) {
    state.tabsAsked.add(name);
    return this.sheets.find((sheet) => sheet.name === name) ?? null;
  }

  getNumSheets() {
    return this.sheets.length;
  }

  getSpreadsheetTimeZone() {
    return TIME_ZONE;
  }

  insertSheet(name, index) {
    if (this.sheets.some((sheet) => sheet.name === name)) throw new Error(`A sheet named "${name}" exists.`);
    const sheet = new FakeSheet(name, NEW_ROWS, NEW_COLUMNS, this.log);
    this.sheets.splice(index, 0, sheet);
    this.log.push({ sheet: name, op: "insertSheet", index });
    return sheet;
  }

  deleteSheet(sheet) {
    const index = this.sheets.indexOf(sheet);
    if (index < 0) throw new Error(`The spreadsheet holds no sheet "${sheet.name}".`);
    if (this.sheets.filter((other) => other !== sheet && !other.hidden).length === 0) {
      throw new Error("A spreadsheet must keep one visible sheet.");
    }
    this.sheets.splice(index, 1);
    this.log.push({ sheet: sheet.name, op: "deleteSheet", index });
  }

  /**
   * The names of the tabs in their order.
   */
  names() {
    return this.sheets.map((sheet) => sheet.name);
  }

  /**
   * A tab by its name, with no entry in the list of the tabs that the script
   * asked for.
   */
  tab(name) {
    return this.sheets.find((sheet) => sheet.name === name);
  }
}

/**
 * A builder of a rule with the chain methods of Apps Script. Each method
 * sets one field of the rule. The method `build` returns a copy of the rule.
 */
function fakeBuilder(methods) {
  const rule = {};
  const builder = { build: () => ({ ...rule }) };
  for (const [method, set] of Object.entries(methods)) {
    builder[method] = (...args) => {
      set(rule, ...args);
      return builder;
    };
  }
  return builder;
}

/**
 * A builder of an embedded bar chart with the methods of an Apps Script
 * EmbeddedChartBuilder and EmbeddedBarChartBuilder that the script calls.
 * Each method records its arguments in the chart. A call of another method
 * fails the harness, because the builder has no such member. setOption
 * keeps each option by its name, as Apps Script does, and checks no name.
 * The method `build` returns the chart with getRanges.
 */
function fakeChartBuilder() {
  const chart = {
    type: null,
    ranges: [],
    numHeaders: null,
    stacked: false,
    title: null,
    colors: null,
    position: null,
    options: {},
  };
  const builder = {
    asBarChart: () => {
      chart.type = "BAR";
      return builder;
    },
    addRange: (range) => {
      chart.ranges.push(range);
      return builder;
    },
    setNumHeaders: (count) => {
      chart.numHeaders = count;
      return builder;
    },
    setStacked: () => {
      chart.stacked = true;
      return builder;
    },
    setTitle: (title) => {
      chart.title = title;
      return builder;
    },
    setColors: (colors) => {
      chart.colors = [...colors];
      return builder;
    },
    setPosition: (row, column, offsetX, offsetY) => {
      chart.position = { row, column, offsetX, offsetY };
      return builder;
    },
    setOption: (name, value) => {
      chart.options[name] = value;
      return builder;
    },
    build: () => {
      const built = { ...chart, ranges: [...chart.ranges], options: { ...chart.options } };
      built.getRanges = () => [...built.ranges];
      return built;
    },
  };
  return builder;
}

/**
 * The block of the Concentration.Exposure tab that a chart reads: company
 * when each range of the chart is in the company chart block, holdings when
 * each range is in the holdings chart block, else other.
 */
function chartBlock(chart) {
  const inside = (first, width) =>
    chart.ranges.every(
      (range) => range.sheet.name === EXPOSURE_TAB && range.column >= first && range.column < first + width,
    );
  if (inside(CHART_AT, CHART_WIDTH)) return "company";
  if (inside(HOLDINGS_AT, HOLDINGS_WIDTH)) return "holdings";
  return "other";
}

/**
 * The company charts of a tab.
 */
function companyCharts(sheet) {
  return sheet.charts.filter((chart) => chartBlock(chart) === "company");
}

/**
 * The holdings charts of a tab.
 */
function holdingsCharts(sheet) {
  return sheet.charts.filter((chart) => chartBlock(chart) === "holdings");
}

/**
 * The state of the fakes that the runs change: the log of changes, the user
 * properties, the document properties, the answers of the key dialog, the
 * messages of the alerts, the sidebars, the lock, the fetch handler, the
 * fetch calls, the menus, the tabs that the script asked for, the state of
 * the tabs at the time of each request, the flag that stops the
 * calculation of the chart cells at a flush, the cells that give an old
 * value for a count of reads, the waits of Utilities.sleep, and the calls of
 * console.warn and console.error.
 */
const state = {
  log: [],
  userProperties: new Map(),
  documentProperties: new Map(),
  sidebars: [],
  promptAnswers: [],
  prompts: [],
  alerts: [],
  lockHeldByOther: false,
  lockTaken: 0,
  lockReleased: 0,
  fetchHandler: null,
  fetchCalls: [],
  menus: [],
  tabsAsked: new Set(),
  atFetch: [],
  frozenValues: false,
  lagging: new Map(),
  sleeps: [],
  logged: [],
};

/**
 * A Holdings fake with the synthetic header and the given rows.
 */
function makeHoldings(rows) {
  const holdings = new FakeSheet("Holdings", rows.length + 1, HOLDINGS_HEADER.length, state.log);
  holdings.grid[0] = [...HOLDINGS_HEADER];
  rows.forEach((cells, r) => {
    holdings.grid[r + 1] = [...cells];
  });
  return holdings;
}

/**
 * Put a new Holdings fake in the place of the old one.
 */
function replaceHoldings(rows) {
  book.sheets[book.names().indexOf("Holdings")] = makeHoldings(rows);
}

const book = new FakeBook(state.log);
book.sheets.push(makeHoldings(HOLDINGS_ROWS));

/**
 * The ambient services of Apps Script, as fakes over the state above.
 */
const BUTTON = { OK: "OK", CANCEL: "CANCEL", CLOSE: "CLOSE" };
const BUTTON_SET = { OK: "OK", OK_CANCEL: "OK_CANCEL" };

/**
 * The Ui fake. It has the add-on menu builder, the input dialog, the alert,
 * and the sidebar. The input dialog takes its answer from the queue
 * state.promptAnswers. showSidebar records the page in state.sidebars. The
 * other dialogs, such as showModalDialog, are absent, so a call to one of
 * them fails the harness. A published add-on puts its menu under Extensions,
 * so createMenu fails the harness too.
 */
function fakeUi() {
  const menu = { addon: true, items: [] };
  const builder = {
    addItem: (label, handler) => {
      menu.items.push({ label, handler });
      return builder;
    },
    addToUi: () => {
      state.menus.push(menu);
    },
  };
  return {
    Button: BUTTON,
    ButtonSet: BUTTON_SET,
    createAddonMenu: () => builder,
    createMenu: () => {
      throw new Error("The script calls createMenu. A published add-on uses createAddonMenu.");
    },
    prompt: (title, text, buttons) => {
      state.prompts.push({ title, text, buttons });
      const answer = state.promptAnswers.shift();
      if (answer === undefined) throw new Error("The harness has no answer for the input dialog.");
      return { getSelectedButton: () => answer.button, getResponseText: () => answer.text };
    },
    alert: (text) => {
      state.alerts.push(String(text));
      return BUTTON.OK;
    },
    showSidebar: (page) => {
      state.sidebars.push(page);
    },
  };
}

/**
 * A fake of a properties store over a Map.
 */
function fakeStore(map) {
  return {
    getProperty: (name) => map.get(name) ?? null,
    getProperties: () => Object.fromEntries(map),
    setProperty(name, value) {
      map.set(name, String(value));
      return this;
    },
    deleteProperty(name) {
      map.delete(name);
      return this;
    },
  };
}

const services = {
  SpreadsheetApp: {
    getActiveSpreadsheet: () => book,
    getUi: fakeUi,
    flush: () => {
      state.log.push({ op: "flush" });
      if (!state.frozenValues) calculateChart();
    },
    newConditionalFormatRule: () =>
      fakeBuilder({
        whenFormulaSatisfied: (rule, formula) => (rule.formula = formula),
        setRanges: (rule, ranges) => (rule.ranges = ranges),
        setBackground: (rule, value) => (rule.background = value),
        setFontColor: (rule, value) => (rule.color = value),
        setBold: (rule, value) => (rule.bold = value),
        setItalic: (rule, value) => (rule.italic = value),
      }),
    newDataValidation: () =>
      fakeBuilder({
        requireNumberBetween: (rule, min, max) => Object.assign(rule, { min, max }),
        setAllowInvalid: (rule, value) => (rule.allowInvalid = value),
        setHelpText: (rule, text) => (rule.helpText = text),
      }),
  },
  PropertiesService: {
    getScriptProperties: () => {
      throw new Error("The script reads the script properties, which all users of an add-on share.");
    },
    getDocumentProperties: () => fakeStore(state.documentProperties),
    getUserProperties: () => fakeStore(state.userProperties),
  },
  HtmlService: {
    createHtmlOutputFromFile: (name) => {
      const source = readFileSync(resolve(SRC, `${name}.html`), "utf8");
      const page = { file: name, source, title: "" };
      page.setTitle = (title) => {
        page.title = title;
        return page;
      };
      return page;
    },
  },
  Utilities: {
    formatDate: (date, zone, format) => {
      if (format !== "yyyy-MM-dd") throw new Error(`The fake formats yyyy-MM-dd alone, not ${format}.`);
      return new Intl.DateTimeFormat("en-CA", {
        timeZone: zone,
        year: "numeric",
        month: "2-digit",
        day: "2-digit",
      }).format(date);
    },
    sleep: (ms) => {
      state.sleeps.push(ms);
    },
  },
  console: {
    warn: (...args) => {
      state.logged.push({ level: "warn", args });
    },
    error: (...args) => {
      state.logged.push({ level: "error", args });
    },
  },
  LockService: {
    getScriptLock: () => {
      throw new Error("The script takes the script lock, which blocks all users of an add-on.");
    },
    getDocumentLock: () => ({
      tryLock: (ms) => {
        if (ms !== 0) throw new Error(`The script asked for a lock wait of ${ms} ms, not 0.`);
        if (state.lockHeldByOther) return false;
        state.lockTaken += 1;
        return true;
      },
      releaseLock: () => {
        state.lockReleased += 1;
      },
    }),
  },
  UrlFetchApp: {
    fetch: (url, options) => {
      state.fetchCalls.push({ url, options });
      state.log.push({ op: "fetch" });
      state.atFetch.push({ names: book.names(), tabs: book.sheets.map((sheet) => [sheet.name, sheet.state()]) });
      if (state.fetchHandler === null) throw new Error("The harness allows no fetch in this run.");
      return state.fetchHandler(url, options);
    },
  },
};

/**
 * The response fake of UrlFetchApp.
 */
function fakeResponse(status, text) {
  return { getResponseCode: () => status, getContentText: () => text };
}

/**
 * The source of the child process that sends the real request. It reads the
 * request from its standard input, so the key never appears in an argument
 * of a process. It prints the status and the body as JSON.
 */
const CHILD = `
let input = "";
for await (const chunk of process.stdin) input += chunk;
const req = JSON.parse(input);
try {
  const res = await fetch(req.url, { method: req.method, headers: req.headers, body: req.body });
  process.stdout.write(JSON.stringify({ status: res.status, text: await res.text() }));
} catch (e) {
  process.stdout.write(JSON.stringify({ error: String(e && e.message) }));
}
`;

let liveRequests = 0;

/**
 * The body text of the answer of run 1. The harness compares the cells with
 * it, and later runs send it again.
 */
let liveText = null;

/**
 * The body text of the last answer of liveFetch or offlineFetch.
 */
let answerText = null;

/**
 * The fetch handler of a run with a real request. It sends the request of
 * the script to the route and blocks until the answer arrives, as
 * UrlFetchApp.fetch does. The harness allows LIVE_LIMIT such requests.
 */
function liveFetch(url, options) {
  liveRequests += 1;
  if (liveRequests > LIVE_LIMIT) throw new Error(`The harness sends ${LIVE_LIMIT} real requests at most.`);
  const request = {
    url,
    method: String(options.method).toUpperCase(),
    headers: { ...options.headers, "Content-Type": options.contentType },
    body: options.payload,
  };
  const child = spawnSync(process.execPath, ["--input-type=module", "-e", CHILD], {
    input: JSON.stringify(request),
    encoding: "utf8",
    maxBuffer: 256 * 1024 * 1024,
    timeout: 120000,
  });
  if (child.status !== 0) throw new Error(`The fetch process stopped with status ${child.status}.`);
  const out = JSON.parse(child.stdout);
  if (out.error) throw new Error(`The fetch failed: ${out.error}`);
  answerText = out.text;
  return fakeResponse(out.status, out.text);
}

let offlineRequests = 0;

/**
 * The invented holdings of the two index funds of the synthetic answer. Each
 * holding has a merge key, a name, a ticker, a class, and a percent of the
 * fund. Fund A holds stocks, a preferred stock, and cash. Fund B holds the
 * stock and a bond of company B, a bond of company C, and a held fund with a
 * negative percent that did not enter. The percents of each fund add up to
 * 100.
 */
const OFFLINE_FUNDS = {
  [FUND_A]: [
    { key: "name:EXAMPLE COMPANY B", name: "Example Company B", ticker: null, class: "stock", pct: 85 },
    { key: "name:EXAMPLE COMPANY C", name: "Example Company C", ticker: null, class: "stock", pct: 5 },
    { key: `ticker:${STOCK}`, name: "Example company", ticker: STOCK, class: "stock", pct: 5 },
    { key: "name:CASH", name: "Cash", ticker: null, class: "cash", pct: 4 },
    { key: "name:EXAMPLE PREFERRED D", name: "Example Preferred D", ticker: null, class: "preferred", pct: 1 },
  ],
  [FUND_B]: [
    { key: "name:EXAMPLE COMPANY B", name: "Example Company B", ticker: null, class: "stock", pct: 80 },
    { key: "name:EXAMPLE COMPANY B", name: "Example Company B", ticker: null, class: "other", pct: 5 },
    { key: "name:EXAMPLE COMPANY C", name: "Example Company C", ticker: null, class: "other", pct: 5 },
    { key: `ticker:${STOCK}`, name: "Example company", ticker: STOCK, class: "stock", pct: 5 },
    { key: "name:CASH", name: "Cash", ticker: null, class: "cash", pct: 6 },
    { key: "name:EXAMPLE HELD FUND", name: "Example Held Fund", ticker: null, class: "fund", pct: -1 },
  ],
};

/**
 * The invented series of each fund of the synthetic answer.
 */
const OFFLINE_SERIES = { [FUND_A]: "S000000901", [FUND_B]: "S000000902", [MONEY]: "S000000903" };

/**
 * The elements of the funds block of an answer that the funds block of the
 * tab holds: each element with no heldBy whose entered value is not false.
 * An element with no entered key entered.
 */
function topFundsOf(funds) {
  return funds.filter((fund) => (fund.heldBy ?? null) === null && fund.entered !== false);
}

/**
 * The cause of each fund of the synthetic answer that is in the store and
 * did not enter through its holdings.
 */
const OFFLINE_NOT_ENTERED = { [MONEY]: "no_report" };

/**
 * The line of each direct position of the synthetic answer. The bond and the
 * plan fund get the class unknown, and the money market fund gets the class
 * fund, as the route gives them.
 */
const OFFLINE_DIRECT = {
  [STOCK]: { key: `ticker:${STOCK}`, name: "Example company", ticker: STOCK, class: "stock" },
  [BOND]: { key: `ticker:${BOND}`, name: BOND, ticker: BOND, class: "unknown" },
  [MONEY]: { key: `ticker:${MONEY}`, name: "Example money market fund", ticker: MONEY, class: "fund" },
  [PLAN_FUND]: { key: `id:${PLAN_FUND}`, name: PLAN_FUND, ticker: null, class: "unknown" },
};

/**
 * The direct lines and the funds that did not enter of a synthetic answer
 * with the line classes trust and cash of a direct position. A trust gets
 * the class trust and no stock weight. The money market fund gets the class
 * cash and no element of the funds block.
 */
const TRUST_KINDS = {
  direct: {
    ...OFFLINE_DIRECT,
    [MONEY]: { key: `ticker:${MONEY}`, name: "Example money market fund", ticker: MONEY, class: "cash" },
    [TRUST]: { key: `ticker:${TRUST}`, name: "Example index trust", ticker: TRUST, class: "trust" },
    [MIX_TRUST]: { key: `ticker:${MIX_TRUST}`, name: "Example metal trust", ticker: MIX_TRUST, class: "trust" },
  },
  notEntered: {},
};

/**
 * The top 10 sum, the HHI, and the effective count of a list of weights,
 * from the absolute share of each weight. Each is 0 when the absolute values
 * sum to 0.
 */
function shareMeasures(weights) {
  const total = weights.reduce((a, w) => a + Math.abs(w), 0);
  if (total === 0) return { top10Weight: 0, hhi: 0, effectiveCount: 0 };
  const shares = weights.map((w) => Math.abs(w) / total);
  const squares = shares.reduce((a, share) => a + share * share, 0);
  const top = [...shares].sort((a, b) => b - a).slice(0, 10);
  return { top10Weight: top.reduce((a, b) => a + b, 0), hhi: squares * 10000, effectiveCount: 1 / squares };
}

/**
 * The equity block of a list of stock weights, or null when their sum is not
 * above 0.
 */
function equityOf(stockWeights) {
  const weight = stockWeights.reduce((a, b) => a + b, 0);
  if (!(weight > 0)) return null;
  const counted = stockWeights.filter((w) => w !== 0);
  return { weight, lineCount: counted.length, ...shareMeasures(counted) };
}

/**
 * The overlaps block of a list of positions and lines: one element for each
 * pair of positions of the funds block that hold stock in one or more of the
 * same lines, largest overlap first. Each position holds its id and its
 * weight.
 */
function overlapsOf(funds, lines) {
  const overlaps = [];
  funds.forEach((first, i) => {
    for (const second of funds.slice(i + 1)) {
      let overlap = 0;
      let sharedLineCount = 0;
      for (const line of lines) {
        const a = (line.stockSources[first.id] ?? 0) / first.weight;
        const b = (line.stockSources[second.id] ?? 0) / second.weight;
        if (a > 0 && b > 0) {
          overlap += Math.min(a, b);
          sharedLineCount += 1;
        }
      }
      if (sharedLineCount > 0) overlaps.push({ ids: [first.id, second.id], overlap, sharedLineCount });
    }
  });
  return overlaps.sort((a, b) => b.overlap - a.overlap);
}

/**
 * A synthetic answer of the route for a list of positions, in the shape of a
 * real answer of schema version 1.9. Each name and each number is invented.
 * A position of OFFLINE_FUNDS enters through its holdings and gets a
 * residual line of the class cash. Each other position enters as one line of
 * OFFLINE_DIRECT. Holdings with one key share one line. The line of company
 * B holds a stock from both funds and a bond from fund B. The line of
 * company C holds a stock from fund A and a bond from fund B. Each line
 * holds classWeights.
 *
 * The funds block holds an element for each fund that entered, an element
 * with heldBy set for each held fund of the class fund, and an element with
 * entered false for each fund of OFFLINE_NOT_ENTERED.
 *
 * A position with parts enters through each part at the weight of the
 * position times the weight of the part. A fund part gets an element of the
 * funds block with its partWeight. The rest of a mix under 100% enters as one
 * line of the class unknown. Each contribution goes under the position id.
 * The overlaps block counts each position one time.
 *
 * When the answer holds more than `maxLines` lines, the lines block keeps
 * the `maxLines` lines with the largest absolute weight, and one cap line
 * with the key other:lines holds the sums of the other lines. The route
 * keeps 2,000 lines.
 *
 * `kinds` holds the direct line of each ticker and the cause of each fund
 * that did not enter. TRUST_KINDS gives the classes trust and cash of a
 * direct position.
 */
function offlineAnswer(
  positions,
  maxLines = 2000,
  kinds = { direct: OFFLINE_DIRECT, notEntered: OFFLINE_NOT_ENTERED },
) {
  const byKey = new Map();
  let lookedThrough = 0;
  const add = (holding, id, weight) => {
    if (!byKey.has(holding.key)) {
      const { key, name, ticker } = holding;
      byKey.set(key, {
        key,
        name,
        ticker,
        weight: 0,
        sources: {},
        stockWeight: 0,
        stockSources: {},
        classWeights: {},
        sizes: {},
      });
    }
    const line = byKey.get(holding.key);
    line.weight += weight;
    line.sources[id] = (line.sources[id] ?? 0) + weight;
    if (holding.class === "stock") {
      line.stockWeight += weight;
      line.stockSources[id] = (line.stockSources[id] ?? 0) + weight;
    }
    line.classWeights[holding.class] = (line.classWeights[holding.class] ?? 0) + weight;
    line.sizes[holding.class] = (line.sizes[holding.class] ?? 0) + Math.abs(weight);
  };
  const funds = [];
  const fundPositions = [];
  /**
   * The fields of an element of the funds block that adds nothing to the
   * counts: a held fund or a fund that did not enter.
   */
  const notEntered = (reason) => ({
    entered: false,
    notEnteredReason: reason,
    reportDate: null,
    accessionNumber: null,
    holdingCount: 0,
    coveredWeight: 0,
    mergedByTicker: 0,
    mergedByLei: 0,
    mergedByName: 0,
  });
  /**
   * Enter one ticker at a weight under a position id. The result is true
   * when the ticker is a fund of OFFLINE_FUNDS.
   */
  const enter = (id, ticker, weight, partWeight) => {
    const holdings = OFFLINE_FUNDS[ticker];
    if (holdings === undefined) {
      const direct = kinds.direct[ticker] ?? { key: `ticker:${ticker}`, name: ticker, ticker, class: "unknown" };
      add(direct, id, weight);
      if (ticker in kinds.notEntered) {
        funds.push({
          id,
          ticker,
          partWeight,
          seriesId: OFFLINE_SERIES[ticker],
          name: direct.name,
          heldBy: null,
          weight,
          ...notEntered(kinds.notEntered[ticker]),
        });
      }
      return false;
    }
    for (const holding of holdings) {
      add(holding, id, (weight * holding.pct) / 100);
      if (holding.class !== "fund") lookedThrough += (weight * holding.pct) / 100;
    }
    const covered = holdings.reduce((a, holding) => a + holding.pct, 0) / 100;
    const residual = {
      key: `residual:${ticker}`,
      name: "Cash and other net assets",
      ticker: null,
      class: "cash",
    };
    add(residual, id, weight * Math.max(0, 1 - covered));
    funds.push({
      id,
      ticker,
      partWeight,
      seriesId: OFFLINE_SERIES[ticker],
      name: `Example index fund ${ticker}`,
      heldBy: null,
      entered: true,
      notEnteredReason: null,
      reportDate: "2026-06-30",
      accessionNumber: `0000000000-26-00000${funds.length + 1}`,
      holdingCount: holdings.length,
      weight,
      coveredWeight: weight * covered,
      mergedByTicker: 1,
      mergedByLei: 0,
      mergedByName: holdings.length - 1,
    });
    for (const holding of holdings.filter((h) => h.class === "fund")) {
      funds.push({
        id,
        ticker: holding.ticker,
        partWeight,
        seriesId: null,
        name: holding.name,
        heldBy: OFFLINE_SERIES[ticker],
        weight: (weight * holding.pct) / 100,
        ...notEntered("unknown_series"),
      });
    }
    return true;
  };
  for (const position of positions) {
    const { id, weight } = position;
    if (!Array.isArray(position.parts)) {
      if (enter(id, position.ticker ?? id, weight, null)) fundPositions.push({ id, weight });
      continue;
    }
    let described = 0;
    let looked = false;
    for (const part of position.parts) {
      described += part.weight;
      looked = enter(id, part.ticker, weight * part.weight, part.weight) || looked;
    }
    if (looked) fundPositions.push({ id, weight });
    const rest = weight * (1 - described);
    if (rest > 1e-12) add({ key: `id:${id}`, name: `${id} (not described)`, ticker: null, class: "unknown" }, id, rest);
  }
  const lines = [...byKey.values()].map((line) => ({
    key: line.key,
    name: line.name,
    ticker: line.ticker,
    lei: null,
    class: Object.entries(line.sizes).sort((a, b) => b[1] - a[1])[0][0],
    weight: line.weight,
    sources: line.sources,
    stockWeight: line.stockWeight,
    stockSources: line.stockSources,
    classWeights: line.classWeights,
  }));
  lines.sort((a, b) => b.weight - a.weight || (a.key < b.key ? -1 : 1));
  const weights = lines.map((line) => line.weight);
  const sum = weights.reduce((a, b) => a + b, 0);
  const seen = lines.filter((line) => line.class !== "unknown").map((line) => line.weight);
  const measures = {
    lineCount: lines.length,
    ...shareMeasures(seen),
    lookedThroughWeight: lookedThrough,
    notLookedThroughWeight: sum - lookedThrough,
    weightSum: sum,
    weightDifference: sum - 1,
    unknownWeight: lines.filter((line) => line.class === "unknown").reduce((a, line) => a + line.weight, 0),
    preferredWeight: lines.reduce((a, line) => a + (line.classWeights.preferred ?? 0), 0),
    equity: equityOf(lines.map((line) => line.stockWeight)),
  };
  const overlaps = overlapsOf(fundPositions, lines);
  let kept = lines;
  if (lines.length > maxLines) {
    const ranked = [...lines].sort((a, b) => Math.abs(b.weight) - Math.abs(a.weight) || (a.key < b.key ? -1 : 1));
    const keep = new Set(ranked.slice(0, maxLines));
    const rest = lines.filter((line) => !keep.has(line));
    const total = (pick) => {
      const map = {};
      for (const line of rest) {
        for (const [name, value] of Object.entries(pick(line))) map[name] = (map[name] ?? 0) + value;
      }
      return map;
    };
    kept = [
      ...lines.filter((line) => keep.has(line)),
      {
        key: "other:lines",
        name: "Other lines",
        ticker: null,
        lei: null,
        class: "other",
        weight: rest.reduce((a, line) => a + line.weight, 0),
        sources: total((line) => line.sources),
        stockWeight: rest.reduce((a, line) => a + line.stockWeight, 0),
        stockSources: total((line) => line.stockSources),
        classWeights: total((line) => line.classWeights),
      },
    ];
  }
  return {
    measures,
    funds,
    overlaps,
    lines: kept,
    meta: {
      schemaVersion: "1.9",
      source: "Invented data of the offline run",
      pctValueUnit: "percent of net assets",
      disclaimer: "Invented data. Not investment advice.",
    },
  };
}

/**
 * The fetch handler of a run with a synthetic answer. It answers with the
 * synthetic answer of the positions of the request and sends no request.
 */
function offlineFetch(url, options) {
  offlineRequests += 1;
  answerText = JSON.stringify(offlineAnswer(JSON.parse(options.payload).positions));
  return fakeResponse(200, answerText);
}

/**
 * The fixture of the accepted layout of one tab.
 */
function acceptedDefinition(tab) {
  return JSON.parse(readFileSync(ACCEPTED_FILES[tab], "utf8"));
}

/**
 * The expected state of a new tab from its accepted definition file, in the
 * form of FakeSheet.state. The function applies the Sheets API meaning of
 * each key of the file: a style entry sets only the properties that it names,
 * and a later entry wins. A cell entry with `eachRow` puts its one row of
 * values into each row of its range.
 */
function expectedState(definition) {
  const format = definition.format ?? {};
  const rows = NEW_ROWS;
  const columns = Math.max(NEW_COLUMNS, format.columns ?? 0);
  const sheet = new FakeSheet(definition.tab, rows, columns, []);
  for (const entry of definition.cells) {
    const b = parseA1(entry.range, rows, columns);
    const values = entry.eachRow ? Array.from({ length: b.rows }, () => entry.eachRow) : entry.values;
    values.forEach((cells, r) => {
      cells.forEach((value, c) => {
        sheet.grid[b.row - 1 + r][b.column - 1 + c] = value;
      });
    });
  }
  const names = {
    bold: ["fontWeight", (v) => (v ? "bold" : "normal")],
    italic: ["fontStyle", (v) => (v ? "italic" : "normal")],
    fontSize: ["fontSize", (v) => v],
    color: ["fontColor", (v) => v],
    background: ["background", (v) => v],
    numberFormat: ["numberFormat", (v) => v],
    align: ["horizontalAlignment", (v) => v.toLowerCase()],
    wrap: ["wrap", (v) => v],
  };
  for (const entry of format.cells ?? []) {
    const range = sheet.getRange(entry.range);
    for (const [key, value] of Object.entries(entry)) {
      if (key === "range") continue;
      const [property, convert] = names[key];
      range.style(property, convert(value), "expected");
    }
  }
  for (const [spec, width] of Object.entries(format.columnWidths ?? {})) {
    const [first, last] = spec.split(":");
    for (let c = columnNumber(first); c <= columnNumber(last ?? first); c += 1) sheet.widths.set(c, width);
  }
  sheet.rules = (format.conditional ?? []).map((rule) => {
    const out = { formula: rule.formula, ranges: [sheet.getRange(rule.range)] };
    for (const key of ["background", "color", "bold", "italic"]) if (rule[key] !== undefined) out[key] = rule[key];
    return out;
  });
  for (const v of format.validation ?? []) {
    sheet
      .getRange(v.range)
      .setDataValidation({ min: v.between[0], max: v.between[1], allowInvalid: false, helpText: v.message });
  }
  sheet.hidden = Boolean(format.hidden);
  sheet.frozenRows = format.frozenRows ?? 0;
  return JSON.parse(sheet.state());
}

/**
 * Compare the state of a tab with the expected state, part by part. The
 * assertion names the first cell or the first entry that differs.
 */
function checkLayout(tab, actualText, expected) {
  const actual = JSON.parse(actualText);
  check(
    actual.grid.length === expected.grid.length && actual.grid[0].length === expected.grid[0].length,
    `${tab}: the grid is ${expected.grid.length} rows by ${expected.grid[0].length} columns ` +
      `(${actual.grid.length} by ${actual.grid[0].length})`,
  );
  const cells = [];
  expected.grid.forEach((row, r) => {
    row.forEach((value, c) => {
      if (actual.grid[r][c] !== value) cells.push(`row ${r + 1} column ${c + 1}`);
    });
  });
  check(cells.length === 0, `${tab}: each formula and label is at its cell (differs: ${cells.slice(0, 3).join(", ")})`);
  const differ = (part) => {
    const a = new Map(actual[part].map(([k, v]) => [String(k), JSON.stringify(v)]));
    const e = new Map(expected[part].map(([k, v]) => [String(k), JSON.stringify(v)]));
    const keys = [...new Set([...a.keys(), ...e.keys()])].filter((k) => a.get(k) !== e.get(k));
    return keys.slice(0, 3).map((k) => `${k}: ${a.get(k) ?? "none"} for ${e.get(k) ?? "none"}`);
  };
  check(differ("styles").length === 0, `${tab}: each cell style matches (${differ("styles").join("; ")})`);
  check(differ("widths").length === 0, `${tab}: each column width matches (${differ("widths").join("; ")})`);
  check(differ("validation").length === 0, `${tab}: the data validation matches (${differ("validation").join("; ")})`);
  check(
    JSON.stringify(actual.rules) === JSON.stringify(expected.rules),
    `${tab}: the ${expected.rules.length} conditional formats match in their order`,
  );
  check(actual.hidden === expected.hidden, `${tab}: the hidden flag is ${expected.hidden}`);
  check(actual.frozenRows === expected.frozenRows, `${tab}: ${expected.frozenRows} rows are frozen`);
}

/**
 * Wait the given count of milliseconds, so the next Date differs.
 */
function pause(ms) {
  Atomics.wait(new Int32Array(new SharedArrayBuffer(4)), 0, 0, ms);
}

/**
 * Print rows as a table with a header, with each number to 6 decimals.
 */
function table(header, rows) {
  const text = (v) => (typeof v === "number" && !Number.isInteger(v) ? v.toFixed(6) : String(v));
  const all = [header, ...rows].map((cells) => cells.map(text));
  const widths = header.map((_, c) => Math.min(48, Math.max(...all.map((cells) => cells[c].length))));
  for (const cells of all) {
    console.log(`  ${cells.map((v, c) => v.slice(0, 48).padEnd(widths[c])).join("  ")}`);
  }
}

/**
 * Each number inside a JSON value.
 */
function numbersOf(value) {
  if (typeof value === "number") return [value];
  if (Array.isArray(value)) return value.flatMap(numbersOf);
  if (value !== null && typeof value === "object") return Object.values(value).flatMap(numbersOf);
  return [];
}

/**
 * Check that every cell of the Exposure grid outside B1:B2 matches the
 * snapshot.
 */
function checkKept(before, name) {
  const after = book.tab(EXPOSURE_TAB).snapshot();
  check(after.length === before.length && after[0].length === before[0].length, `${name}: the grid size stays`);
  const changed = [];
  after.forEach((cells, r) => {
    cells.forEach((value, c) => {
      if (c === 1 && r < 2) return;
      if (value !== before[r][c]) changed.push(`row ${r + 1} column ${c + 1}`);
    });
  });
  check(changed.length === 0, `${name}: the lines and the funds stay (changed: ${changed.slice(0, 3).join(", ")})`);
}

/**
 * Check that a run with tabs of the current layout creates no tab, deletes
 * no tab, changes no cell and no format of the report tab, and changes no
 * cell of the Holdings tab. A good run, with OK in B1, draws the company
 * chart, then the holdings chart. Each draw removes the chart of its own
 * block from the last draw when one exists, then inserts one chart when its
 * block gives a series. A fault changes no chart. `names` holds the tabs of
 * the spreadsheet in their order.
 */
function checkNoTabChange(reportBefore, logFrom, name, names = ["Holdings", EXPOSURE_TAB, REPORT_TAB]) {
  const entries = state.log.slice(logFrom);
  const chartEntries = entries.filter((e) => e.op === "insertChart" || e.op === "removeChart");
  const chartOps = chartEntries.map((e) => `${e.op} ${e.block}`);
  const good = book.tab(EXPOSURE_TAB).cell("B1") === "OK";
  const report = book.tab(REPORT_TAB);
  const count = (op) => chartOps.filter((o) => o === op).length;
  const before = (block, now) => now - count(`insertChart ${block}`) + count(`removeChart ${block}`);
  const want = expectedSeries(book.tab(EXPOSURE_TAB), report.cell(THRESHOLD_CELL));
  const drawn = want.direct || want.funds > 0 || want.other;
  const ops = (block, had, draws) => [
    ...(had > 0 ? [`removeChart ${block}`] : []),
    ...(draws ? [`insertChart ${block}`] : []),
  ];
  const expected = good
    ? [
        ...ops("company", before("company", companyCharts(report).length), drawn),
        ...ops("holdings", before("holdings", holdingsCharts(report).length), holdingGroups() > 0),
      ]
    : [];
  check(
    JSON.stringify(chartOps) === JSON.stringify(expected),
    `${name}: ${good ? "a good run draws the company chart, then the holdings chart: each draw removes the chart of its block from the last draw, then inserts one chart when its block gives a series" : "a fault changes no chart"} ` +
      `(${chartOps.join(", ") || "no chart change"})`,
  );
  check(
    !entries.some((e) => e.op === "insertSheet" || e.op === "deleteSheet"),
    `${name}: the script creates no tab and deletes no tab`,
  );
  check(
    blockWritesSince(logFrom).length === 0,
    `${name}: both tabs hold the current layout, so the script writes no cell of the chart block`,
  );
  check(
    JSON.stringify(book.names()) === JSON.stringify(names),
    `${name}: the spreadsheet holds the same three tabs in the same order`,
  );
  check(book.tab(REPORT_TAB).state() === reportBefore, `${name}: no cell and no format of the report tab changes`);
  check(
    entries.every(
      (e) =>
        e.sheet === undefined ||
        e.sheet === EXPOSURE_TAB ||
        (e.sheet === REPORT_TAB && (e.op === "insertChart" || e.op === "removeChart")),
    ),
    `${name}: the script changes the Concentration.Exposure tab and the chart of the report tab alone`,
  );
}

/**
 * The series that the bar chart must show for the Concentration.Exposure tab
 * and a threshold. The companies are the rows of the stock part with a stock
 * weight at or above the threshold, largest first, CHART_ROWS at most, as in
 * the company table. Direct is present when the direct weights of those
 * rows, each rounded to 12 decimals, add up to a number other than 0. A fund
 * of the stock fund block counts when its source cells of those rows add up
 * to a number other than 0. The first CHART_FUNDS funds that count get a
 * series each, and Other funds is present when more funds count.
 */
function expectedSeries(sheet, threshold) {
  const ids = sheet.block(4, SOURCE_AT, 1, MAX_POSITIONS)[0];
  const at = (column) => column - LINE_AT;
  const stockAt = LINE_FIELDS.indexOf("stockWeight");
  const listed = sheet
    .block(5, LINE_AT, sheet.getMaxRows() - 4, SOURCE_AT - LINE_AT + MAX_POSITIONS)
    .filter(
      (cells) => cells[at(PART_AT)] === "stock" && typeof cells[stockAt] === "number" && cells[stockAt] >= threshold,
    )
    .sort((a, b) => b[stockAt] - a[stockAt]);
  const rows = listed.slice(0, CHART_ROWS);
  const direct = rows.reduce((sum, cells) => sum + Math.round(cells[at(DIRECT_AT)] * 1e12) / 1e12, 0);
  const counted = blockRows(sheet, STOCK_FUND_AT, 1).filter(
    ([id]) => sumCells(rows.map((cells) => cells[at(SOURCE_AT) + ids.indexOf(id)])) !== 0,
  ).length;
  return {
    direct: direct !== 0,
    funds: Math.min(counted, CHART_FUNDS),
    other: counted > CHART_FUNDS,
    companies: rows.length,
    listed: listed.length,
  };
}

/**
 * The count of holding groups of the section Holdings. A holding group
 * is a key of the Holdings tab with a number in Value: its Symbol, or its
 * Description when the Symbol is empty, cut to 64 characters. The count is
 * 0 when the Holdings tab holds no Symbol, Description, or Value column.
 */
function holdingGroups() {
  const [header, ...rows] = book.tab("Holdings").grid;
  const at = Object.fromEntries(
    ["Symbol", "Description", "Value"].map((name) => [name, header.findIndex((v) => String(v).trim() === name)]),
  );
  if (!Object.values(at).every((i) => i >= 0)) return 0;
  const keys = rows
    .filter((cells) => typeof cells[at.Value] === "number")
    .map((cells) => (String(cells[at.Symbol]).trim() || String(cells[at.Description]).trim()).slice(0, 64));
  return new Set(keys).size;
}

/**
 * The count of rows of the table of the section Holdings under its
 * header: one row for each holding group and the Total row, or one note row.
 */
function yoursRows() {
  const groups = holdingGroups();
  return groups > 0 ? groups + 1 : 1;
}

/**
 * The row of the report tab under the header row of the holdings chart in
 * the report spill. The spill holds these rows from row 26: the title
 * Holdings, the header of the section, the rows of yoursRows, a blank row,
 * and the header row of the holdings chart.
 */
function holdingsAnchorRow() {
  return REPORT_ROW + 4 + yoursRows();
}

/**
 * The row of the report tab under the header row of the company chart in
 * the report spill, for a count of companies at or above the threshold.
 * After the header row of the holdings chart, the spill holds the band of
 * BAND_ROWS blank rows, the trust note, the header of the company table,
 * the company rows or one note row, the row of the companies under the
 * threshold, a blank row, and the header row of the company chart.
 */
function anchorRow(companies) {
  return holdingsAnchorRow() + BAND_ROWS + 5 + Math.max(1, companies);
}

/**
 * Calculate the header row of the chart block and the anchor cell of the
 * Concentration.Exposure tab, as Google Sheets calculates their formulas,
 * and keep the values in the calculated map of the tab. The header holds
 * Company, then Direct when the Direct series is present, the name of each
 * fund series, and Other funds when that series is present, by
 * expectedSeries. Each other header cell is empty. The fund names are
 * placeholders, because the draw reads only whether a header cell holds a
 * text. The anchor cell holds anchorRow for the count of the companies at
 * or above the threshold. Column A of the report tab holds the text of the
 * header row of the company chart, by spillHeading, in the row above the
 * anchor row.
 *
 * The function also calculates the header row of the holdings chart block
 * and the holdings anchor cell: Holding, then % of portfolio when the
 * section Holdings holds a holding group, else an empty string. The
 * holdings anchor cell holds holdingsAnchorRow, and column A of the report
 * tab holds HOLDINGS_HEADING in the row above it. Column A of the report
 * tab holds no other calculated value. The function calculates nothing
 * for a block while a header cell or the anchor cell of the block holds no
 * formula, such as before ensureTabs writes them.
 */
function calculateChart() {
  const hidden = book.tab(EXPOSURE_TAB);
  const report = book.tab(REPORT_TAB);
  if (hidden === undefined || report === undefined) return;
  const formulas = (first, width, anchor) =>
    hidden.getMaxColumns() >= anchor &&
    [...hidden.grid[3].slice(first - 1, first - 1 + width), hidden.grid[4][anchor - 1]].every(
      (value) => typeof value === "string" && value.startsWith("="),
    );
  const company = formulas(CHART_AT, CHART_WIDTH, ANCHOR_AT);
  const holdings = formulas(HOLDINGS_AT, HOLDINGS_WIDTH, HOLDINGS_ANCHOR_AT);
  if (!company && !holdings) return;
  report.calculated.clear();
  if (company) {
    const want = expectedSeries(hidden, report.cell(THRESHOLD_CELL));
    const names = [
      "Company",
      want.direct ? "Direct" : "",
      ...Array.from({ length: CHART_FUNDS }, (_, k) => (k < want.funds ? `Fund ${k + 1}` : "")),
      want.other ? "Other funds" : "",
    ];
    names.forEach((name, i) => hidden.calculated.set(cellKey(4, CHART_AT + i), name));
    hidden.calculated.set(cellKey(5, ANCHOR_AT), anchorRow(want.listed));
    report.calculated.set(cellKey(anchorRow(want.listed) - 1, 1), spillHeading(report.cell(THRESHOLD_CELL)));
  }
  if (holdings) {
    hidden.calculated.set(cellKey(4, HOLDINGS_AT), "Holding");
    hidden.calculated.set(cellKey(4, HOLDINGS_AT + 1), holdingGroups() > 0 ? "% of portfolio" : "");
    hidden.calculated.set(cellKey(5, HOLDINGS_ANCHOR_AT), holdingsAnchorRow());
    report.calculated.set(cellKey(holdingsAnchorRow() - 1, 1), HOLDINGS_HEADING);
  }
}

/**
 * The text of the chart header row of the report spill for a threshold, as
 * TEXT(threshold,"0.00%") gives it in the spill formula.
 */
function spillHeading(threshold) {
  return `Companies at or over ${(threshold * 100).toFixed(2)}% of the portfolio, by source`;
}

/**
 * The text of an options object with its keys in order, for a comparison.
 */
function sortedJson(value) {
  if (Array.isArray(value)) return `[${value.map(sortedJson).join(",")}]`;
  if (value === null || typeof value !== "object") return JSON.stringify(value);
  return `{${Object.keys(value)
    .sort()
    .map((key) => `${JSON.stringify(key)}:${sortedJson(value[key])}`)
    .join(",")}}`;
}

/**
 * Check the bar chart of the report tab after a draw, and give the series of
 * expectedSeries and the anchor row. The anchor cell JC5 holds the row under
 * the chart header row of the spill, by anchorRow. The report tab holds one
 * chart: a stacked bar chart with one header row, anchored at column A of
 * the row of the anchor cell with no offset. Each range of the chart is one
 * column of the chart block from row 4 to row 4 + CHART_ROWS, so the range
 * never moves. The columns are the company names, then each column whose
 * header holds a text: Direct, the fund columns, and Other funds, in this
 * order. The colors and the options follow. The chart has no title, and its
 * width is the width of the columns A to D of the report tab.
 */
function checkChart(name) {
  const report = book.tab(REPORT_TAB);
  const hidden = book.tab(EXPOSURE_TAB);
  const want = expectedSeries(hidden, report.cell(THRESHOLD_CELL));
  const anchor = hidden.shown(5, ANCHOR_AT);
  check(
    anchor === anchorRow(want.listed),
    `${name}: the anchor cell JC5 holds row ${anchorRow(want.listed)}, under the chart header row (${anchor})`,
  );
  check(
    report.charts.every((chart) => chartBlock(chart) !== "other"),
    `${name}: each chart of the report tab reads the company chart block or the holdings chart block alone`,
  );
  const charts = companyCharts(report);
  const header = Array.from({ length: CHART_WIDTH }, (_, i) => hidden.shown(4, CHART_AT + i));
  const offsets = [0, ...header.map((_, o) => o).filter((o) => o > 0 && header[o] !== "")];
  if (offsets.length === 1) {
    check(
      charts.length === 0,
      `${name}: no header after Company holds a text, so the report tab holds no company chart (${charts.length})`,
    );
    return { ...want, series: 0, anchor };
  }
  check(charts.length === 1, `${name}: the report tab holds one company chart (${charts.length})`);
  const [chart] = charts;
  check(chart.type === "BAR" && chart.stacked === true, `${name}: the chart is a stacked bar chart`);
  check(
    sortedJson(chart.position) === sortedJson({ row: anchor, column: 1, offsetX: 0, offsetY: 0 }),
    `${name}: the chart anchors at A${anchor}, the first row of the band, with no offset ` +
      `(${JSON.stringify(chart.position)})`,
  );
  const bounds = chart.ranges.map((range) => range.bounds());
  check(
    chart.ranges.every((range) => range.sheet === hidden) &&
      bounds.every((b) => b.row === 4 && b.rows === CHART_ROWS + 1 && b.columns === 1) &&
      JSON.stringify(bounds.map((b) => b.column - CHART_AT)) === JSON.stringify(offsets),
    `${name}: the chart reads the columns ${offsets.join(", ")} of the chart block, each from row 4 to row ` +
      `${4 + CHART_ROWS} (${bounds.map((b) => `${b.row},${b.column},${b.rows},${b.columns}`).join("; ")})`,
  );
  check(chart.numHeaders === 1, `${name}: row 4 of the block gives the series names`);
  const colors = offsets
    .slice(1)
    .map((o) => (o === 1 ? DIRECT_COLOR : o <= 1 + CHART_FUNDS ? FUND_COLORS[o - 2] : OTHER_COLOR));
  check(
    JSON.stringify(chart.colors) === JSON.stringify(colors),
    `${name}: each series has its fixed color (${JSON.stringify(chart.colors)})`,
  );
  check(chart.title === null && !("title" in chart.options), `${name}: the chart has no title (${chart.title})`);
  const span = [1, 2, 3, 4].reduce((sum, column) => sum + report.widths.get(column), 0);
  check(
    chart.options.width === span,
    `${name}: the chart spans the columns A to D of the report tab (${chart.options.width} px for ${span} px)`,
  );
  check(
    sortedJson(chart.options) === sortedJson(CHART_OPTIONS),
    `${name}: the chart options are ${sortedJson(CHART_OPTIONS)} (${sortedJson(chart.options)})`,
  );
  return { ...want, series: chart.ranges.length - 1, anchor };
}

/**
 * Check the holdings chart of the report tab after a draw, and give the count
 * of holding groups, the anchor row, and the chart. The holdings anchor cell
 * JH5 holds the row under the header row of the holdings chart in the spill,
 * by holdingsAnchorRow. When the section Holdings holds no holding
 * group, the report tab holds no holdings chart. Else it holds one: a bar
 * chart with no stacking and one header row, anchored at column A of the row
 * of the holdings anchor cell with no offset. Its two ranges are the label
 * column and the share column of the holdings chart block, each from row 4 to
 * row 4 + HOLDINGS_CHART_ROWS. The one series has the Direct color. The chart
 * has no title and no legend, and the data label of the series shows the
 * value at the end of each bar.
 */
function checkHoldingsChart(name) {
  const report = book.tab(REPORT_TAB);
  const hidden = book.tab(EXPOSURE_TAB);
  const groups = holdingGroups();
  const anchor = hidden.shown(5, HOLDINGS_ANCHOR_AT);
  check(
    anchor === holdingsAnchorRow(),
    `${name}: the holdings anchor cell JH5 holds row ${holdingsAnchorRow()}, under the header row of the ` +
      `holdings chart (${anchor})`,
  );
  check(
    report.shown(anchor - 1, 1) === HOLDINGS_HEADING,
    `${name}: the row above the holdings anchor row shows "${HOLDINGS_HEADING}" in column A`,
  );
  const charts = holdingsCharts(report);
  if (groups === 0) {
    check(
      charts.length === 0,
      `${name}: the section Holdings holds no holding group, so the report tab holds no holdings chart ` +
        `(${charts.length})`,
    );
    return { groups, anchor, chart: null };
  }
  check(charts.length === 1, `${name}: the report tab holds one holdings chart (${charts.length})`);
  const [chart] = charts;
  check(chart.type === "BAR" && chart.stacked === false, `${name}: the holdings chart is a bar chart with no stacking`);
  check(
    sortedJson(chart.position) === sortedJson({ row: anchor, column: 1, offsetX: 0, offsetY: 0 }),
    `${name}: the holdings chart anchors at A${anchor}, the first row of its band, with no offset ` +
      `(${JSON.stringify(chart.position)})`,
  );
  const bounds = chart.ranges.map((range) => range.bounds());
  check(
    chart.ranges.every((range) => range.sheet === hidden) &&
      JSON.stringify(bounds) ===
        JSON.stringify(
          [0, 1].map((o) => ({ row: 4, column: HOLDINGS_AT + o, rows: HOLDINGS_CHART_ROWS + 1, columns: 1 })),
        ),
    `${name}: the holdings chart reads the label column and the share column of the holdings chart block, each ` +
      `from row 4 to row ${4 + HOLDINGS_CHART_ROWS} (${bounds.map((b) => `${b.row},${b.column},${b.rows},${b.columns}`).join("; ")})`,
  );
  check(chart.numHeaders === 1, `${name}: row 4 of the holdings chart block gives the series name`);
  check(
    JSON.stringify(chart.colors) === JSON.stringify([DIRECT_COLOR]),
    `${name}: the one series of the holdings chart has the Direct color (${JSON.stringify(chart.colors)})`,
  );
  check(
    chart.title === null && !("title" in chart.options),
    `${name}: the holdings chart has no title (${chart.title})`,
  );
  check(
    sortedJson(chart.options) === sortedJson(HOLDINGS_OPTIONS),
    `${name}: the holdings chart options are ${sortedJson(HOLDINGS_OPTIONS)} (${sortedJson(chart.options)})`,
  );
  return { groups, anchor, chart };
}

/**
 * The rows of the run-time block A16:B25 of the Exposure tab, with each
 * start time in milliseconds.
 */
function runRows() {
  return book
    .tab(EXPOSURE_TAB)
    .block(RUN_ROW, 1, 10, 2)
    .map(([start, seconds]) => [isDate(start) ? start.getTime() : start, seconds]);
}

/**
 * The value of a formula of the report that reads the run-time block or the
 * unknown weight. The evaluator knows the parts that the two run-time
 * formulas and the coverage label use alone: a number, a text in quotes, a
 * reference to a cell or a range of the tab Concentration.Exposure, the
 * operators `&` and `-` from left to right, the comparison `=`, and the
 * functions IF, ISNUMBER, COUNT, AVERAGE, and TEXT. COUNT and AVERAGE read
 * the numbers of a range and skip each other cell, as Google Sheets does.
 * AVERAGE of no number gives the error #DIV/0!. TEXT knows the formats "0%"
 * and "0.0%". Another part stops the harness.
 */
function calculate(formula, sheet) {
  const body = formula.replace(/^=/, "");
  const tokens =
    body.match(/'[^']*'!\$?[A-Z]+\$?\d+(?::\$?[A-Z]+\$?\d+)?|"[^"]*"|[A-Z]+\(|\d+(?:\.\d+)?|[(),=&-]/g) ?? [];
  if (tokens.join("") !== body) throw new Error(`The evaluator cannot read the formula ${formula}.`);
  let i = 0;
  const next = () => tokens[i++];
  const expect = (token) => {
    if (next() !== token) throw new Error(`The evaluator expected "${token}" in ${formula}.`);
  };
  const primary = () => {
    const token = next();
    if (token.endsWith("(")) {
      const args = [];
      if (tokens[i] !== ")") {
        args.push(expression());
        while (tokens[i] === ",") {
          next();
          args.push(expression());
        }
      }
      expect(")");
      return { call: token.slice(0, -1), args };
    }
    if (token.startsWith("'")) {
      const [tab, a1] = token.split("!");
      if (tab !== `'${EXPOSURE_TAB}'`) throw new Error(`The evaluator reads ${EXPOSURE_TAB} alone, not ${tab}.`);
      return {
        ref: parseA1(a1.replaceAll("$", ""), sheet.getMaxRows(), sheet.getMaxColumns()),
        single: !a1.includes(":"),
      };
    }
    if (token.startsWith('"')) return { value: token.slice(1, -1) };
    if (/^\d/.test(token)) return { value: Number(token) };
    throw new Error(`The evaluator cannot read "${token}" in ${formula}.`);
  };
  const chain = () => {
    let left = primary();
    while (tokens[i] === "&" || tokens[i] === "-") {
      const op = next();
      left = { op, args: [left, primary()] };
    }
    return left;
  };
  const expression = () => {
    const left = chain();
    if (tokens[i] !== "=") return left;
    next();
    return { equal: [left, chain()] };
  };
  const numbers = (node) => {
    if (!node.ref) throw new Error("COUNT and AVERAGE take a range in the run-time formulas.");
    const { row, column, rows, columns } = node.ref;
    return sheet
      .block(row, column, rows, columns)
      .flat()
      .filter((v) => typeof v === "number");
  };
  const value = (node) => {
    if ("value" in node) return node.value;
    if (node.ref) {
      if (!node.single) throw new Error("A range stands alone in the run-time formulas.");
      return sheet.block(node.ref.row, node.ref.column, 1, 1)[0][0];
    }
    if (node.equal) return value(node.equal[0]) === value(node.equal[1]);
    if (node.op === "&") return `${value(node.args[0])}${value(node.args[1])}`;
    if (node.op === "-") return value(node.args[0]) - value(node.args[1]);
    const { call, args } = node;
    if (call === "IF") return value(args[0]) ? value(args[1]) : value(args[2]);
    if (call === "ISNUMBER") return typeof value(args[0]) === "number";
    if (call === "COUNT") return numbers(args[0]).length;
    if (call === "AVERAGE") {
      const list = numbers(args[0]);
      return list.length === 0 ? "#DIV/0!" : list.reduce((a, b) => a + b, 0) / list.length;
    }
    if (call === "TEXT") {
      const format = value(args[1]);
      const places = { "0%": 0, "0.0%": 1 }[format];
      if (places === undefined) throw new Error(`The evaluator does not know the TEXT format ${format}.`);
      return `${(value(args[0]) * 100).toFixed(places)}%`;
    }
    throw new Error(`The evaluator does not know the function ${call}.`);
  };
  const tree = expression();
  if (i !== tokens.length) throw new Error(`The evaluator did not read the whole formula ${formula}.`);
  return value(tree);
}

/**
 * The count of setValues calls in the log from the index `from`.
 */
function writesSince(from) {
  return state.log.slice(from).filter((entry) => entry.op === "setValues").length;
}

/**
 * The index of the first log entry from the index `from` that meets the
 * test, or -1.
 */
function indexSince(from, test) {
  const index = state.log.slice(from).findIndex(test);
  return index < 0 ? -1 : from + index;
}

/**
 * The indexes of the setValues entries of the log from the index `from` that
 * write a cell of the company chart block IT4:JA104, the anchor cell JC5,
 * the holdings chart block JE4:JF204, or the holdings anchor cell JH5 of the
 * Concentration.Exposure tab.
 */
function blockWritesSince(from) {
  const meets = (e, row, column, rows, columns) =>
    e.row <= row + rows - 1 &&
    e.row + e.rows - 1 >= row &&
    e.column <= column + columns - 1 &&
    e.column + e.columns - 1 >= column;
  const writes = [];
  state.log.slice(from).forEach((e, i) => {
    if (e.op !== "setValues" || e.sheet !== EXPOSURE_TAB) return;
    if (
      meets(e, 4, CHART_AT, CHART_ROWS + 1, CHART_WIDTH) ||
      meets(e, 5, ANCHOR_AT, 1, 1) ||
      meets(e, 4, HOLDINGS_AT, HOLDINGS_CHART_ROWS + 1, HOLDINGS_WIDTH) ||
      meets(e, 5, HOLDINGS_ANCHOR_AT, 1, 1)
    ) {
      writes.push(from + i);
    }
  });
  return writes;
}

/**
 * Check, by the order of the log from the index `from`, that a run which
 * replaced a tab writes the chart blocks and the anchor cells as the last
 * step of ensureTabs: one setValues call for IT4:JA104, one for JC5, one for
 * JE4:JF204, then one for JH5, after the creation of the report tab and
 * after each other tab operation of the run, and before the request.
 */
function checkBlockWrite(name, from) {
  const writes = blockWritesSince(from);
  const bounds = writes.map((at) => {
    const e = state.log[at];
    return `${e.row},${e.column},${e.rows},${e.columns}`;
  });
  const want = [
    `4,${CHART_AT},${CHART_ROWS + 1},${CHART_WIDTH}`,
    `5,${ANCHOR_AT},1,1`,
    `4,${HOLDINGS_AT},${HOLDINGS_CHART_ROWS + 1},${HOLDINGS_WIDTH}`,
    `5,${HOLDINGS_ANCHOR_AT},1,1`,
  ];
  check(
    JSON.stringify(bounds) === JSON.stringify(want),
    `${name}: ensureTabs writes the chart block IT4:JA${4 + CHART_ROWS}, the anchor cell JC5, the holdings chart ` +
      `block JE4:JF${4 + HOLDINGS_CHART_ROWS}, and the holdings anchor cell JH5, with one call each, in this order ` +
      `(${bounds.join("; ")})`,
  );
  const reportAt = indexSince(from, (e) => e.op === "insertSheet" && e.sheet === REPORT_TAB);
  const fetchAt = indexSince(from, (e) => e.op === "fetch");
  const tabOpsAt = state.log
    .slice(from, fetchAt)
    .map((e, i) => (["insertSheet", "deleteSheet", "hideSheet"].includes(e.op) ? from + i : -1));
  check(
    reportAt >= 0 && writes[0] > reportAt && writes[0] > Math.max(...tabOpsAt) && writes.at(-1) < fetchAt,
    `${name}: the writes of the blocks and the anchor cells follow the creation of the report tab and each other ` +
      "tab operation, and precede the request",
  );
}

/**
 * The rows that the script must write for the lines of an answer, by the
 * part rule: one row for the stock part of a line, one row for its other
 * part, or both. A row holds the fields of LINE_FIELDS, the part name, and
 * one cell for each position id. A stock cell holds the entry of
 * stockSources. An other cell holds the entry of sources minus the entry of
 * stockSources. A part of 0 gives an empty cell.
 */
function expectedParts(lines, ids) {
  const entry = (map, id) => (Object.hasOwn(map ?? {}, id) ? map[id] : null);
  const rows = [];
  for (const line of lines) {
    const fields = LINE_FIELDS.map((name) => line[name] ?? "");
    const stock = ids.map((id) => entry(line.stockSources, id) ?? "");
    const other = ids.map((id) => {
      const part = (entry(line.sources, id) ?? 0) - (entry(line.stockSources, id) ?? 0);
      return part === 0 ? "" : part;
    });
    const hasStock = line.stockWeight !== 0 || stock.some((value) => value !== "");
    if (hasStock) rows.push([...fields, "stock", ...stock]);
    if (!hasStock || other.some((value) => value !== "")) rows.push([...fields, "other", ...other]);
  }
  return rows;
}

/**
 * The part rows of the lines block of a tab, in the form of expectedParts:
 * the fields, the part name, and the source cells, with no direct weight.
 * `count` is the count of position ids.
 */
function partRowsOf(sheet, rows, count) {
  return sheet
    .block(5, LINE_AT, rows, SOURCE_AT - LINE_AT + count)
    .map((cells) => [...cells.slice(0, PART_AT - LINE_AT + 1), ...cells.slice(SOURCE_AT - LINE_AT)]);
}

/**
 * The values that the report formulas of layout version 2 computed from the
 * tab: the direct weight of each part row, the rows of their own, the
 * groups, the weight of each position in the lines of the class unknown, and
 * the funds that hold a stock. The function reads the lines block, the
 * sources block with MAX_POSITIONS columns, the id row, the funds block, and
 * the mix block, and it follows the steps of those formulas: dw, keep, came,
 * cls, gw, gn, gf, uid, uu, fid, and f. The other part of the cap line,
 * with the key other:lines, also gets a row of its own. The own row of a
 * residual line holds the fund ticker of its key as the name. The script
 * writes these values into the tab, and the harness compares the two.
 */
function formulaReference(sheet) {
  const all = sheet.getMaxRows() - 4;
  const header = sheet.block(4, SOURCE_AT, 1, MAX_POSITIONS)[0];
  const rows = sheet
    .block(5, LINE_AT, all, SOURCE_AT - LINE_AT + MAX_POSITIONS)
    .filter((cells) => cells[PART_AT - LINE_AT] !== "");
  const field = (cells, name) => cells[LINE_FIELDS.indexOf(name)];
  const part = (cells) => cells[PART_AT - LINE_AT];
  const num = (value) => (typeof value === "number" ? value : 0);
  const source = (cells, j) => cells[SOURCE_AT - LINE_AT + j];
  const fundColumn = sheet.block(5, FUND_AT, all, 1).map((cells) => cells[0]);
  const fid = [...new Set(fundColumn.filter((id) => id !== ""))];
  const mixIds = new Set(sheet.block(5, MIX_AT, all, 1).map((cells) => cells[0]));
  const isf = header.map((h) => (h !== "" && fid.includes(h) ? 1 : 0));
  const pw = rows.map((cells) =>
    part(cells) === "stock"
      ? num(field(cells, "stockWeight"))
      : part(cells) === "other"
        ? num(field(cells, "weight")) - num(field(cells, "stockWeight"))
        : 0,
  );
  const dw = rows.map((cells, r) => pw[r] - header.reduce((sum, _, j) => sum + num(source(cells, j)) * isf[j], 0));
  const residual = (cells) => String(field(cells, "key")).startsWith("residual:");
  const cap = (cells) => field(cells, "key") === "other:lines";
  const own = rows.map(
    (cells, r) =>
      part(cells) === "other" &&
      field(cells, "class") !== "stock" &&
      (residual(cells) || cap(cells) || Math.abs(dw[r]) > 1e-12 || field(cells, "class") === "unknown"),
  );
  const grp = rows.map((cells, r) => part(cells) === "other" && !own[r]);
  const came = (cells) => header.filter((h, j) => h !== "" && source(cells, j) !== "" && source(cells, j) !== 0);
  const ownRows = rows
    .map((cells, r) => [cells, r])
    .filter(([cells, r]) => own[r] && !(residual(cells) && pw[r] === 0))
    .map(([cells, r]) => [
      field(cells, "key"),
      residual(cells) ? String(field(cells, "key")).slice("residual:".length) : field(cells, "name"),
      field(cells, "class"),
      pw[r],
      came(cells).join(", "),
    ]);
  const cls = [...new Set(rows.filter((_, r) => grp[r]).map((cells) => field(cells, "class")))];
  const groupRows = cls.map((kind) => {
    const members = rows.map((cells, r) => [cells, r]).filter(([cells, r]) => grp[r] && field(cells, "class") === kind);
    const ids = header.filter(
      (h, j) => h !== "" && members.reduce((sum, [cells]) => sum + num(source(cells, j)), 0) !== 0,
    );
    return [kind, members.reduce((sum, [, r]) => sum + pw[r], 0), members.length, ids.join(", ")];
  });
  const unseenRows = header
    .filter((h) => h !== "")
    .map((h) => {
      if (mixIds.has(h)) return [h, 0];
      const j = header.indexOf(h);
      const weight = rows
        .filter((cells) => field(cells, "class") === "unknown" && part(cells) === "other")
        .reduce((sum, cells) => sum + num(source(cells, j)), 0);
      return [h, weight];
    });
  const stockFunds = fid
    .filter((id) => {
      const j = header.indexOf(id);
      return (
        j >= 0 &&
        rows.filter((cells) => part(cells) === "stock").reduce((sum, cells) => sum + num(source(cells, j)), 0) !== 0
      );
    })
    .map((id) => [id]);
  return { direct: dw, own: ownRows, groups: groupRows, unseen: unseenRows, stockFunds };
}

/**
 * The rows of a block of the tab from row 5 down to the last row whose first
 * cell holds a value.
 */
function blockRows(sheet, column, width) {
  const rows = sheet.block(5, column, sheet.getMaxRows() - 4, width);
  let count = 0;
  rows.forEach((cells, r) => {
    if (cells[0] !== "") count = r + 1;
  });
  return rows.slice(0, count);
}

/**
 * True when two rows of cells are equal: each text is the same, and each
 * number is the same to 1e-12.
 */
function sameRows(a, b) {
  return (
    a.length === b.length &&
    a.every(
      (cells, r) =>
        cells.length === b[r].length &&
        cells.every((value, c) =>
          typeof value === "number" && typeof b[r][c] === "number"
            ? Math.abs(value - b[r][c]) < 1e-12
            : value === b[r][c],
        ),
    )
  );
}

/**
 * Check that the computed blocks and the direct weights of the tab hold the
 * values of formulaReference.
 */
function checkComputed(sheet, name, partCount) {
  const want = formulaReference(sheet);
  const direct = sheet.block(5, DIRECT_AT, partCount, 1).map((cells) => cells[0]);
  check(
    direct.length === want.direct.length && direct.every((v, r) => Math.abs(v - want.direct[r]) < 1e-12),
    `${name}: column AZ holds the direct weight dw of each of the ${partCount} part rows`,
  );
  check(
    sameRows(blockRows(sheet, OWN_AT, 5), want.own),
    `${name}: the own block AG:AK holds the ${want.own.length} rows of their own, with keep and came of the formula`,
  );
  check(
    sameRows(blockRows(sheet, GROUP_AT, 4), want.groups),
    `${name}: the group block AM:AP holds cls, gw, gn, and gf of the formula (${want.groups.length} classes)`,
  );
  check(
    sameRows(blockRows(sheet, UNSEEN_AT, 2), want.unseen),
    `${name}: the unseen block AB:AC holds uid and uu of the formula (${want.unseen.length} positions)`,
  );
  check(
    sameRows(blockRows(sheet, STOCK_FUND_AT, 1), want.stockFunds),
    `${name}: the stock fund block AE holds f of the formula (${want.stockFunds.length} funds)`,
  );
  return want;
}

/**
 * The last row of the blocks left of the lines block: the equity block, or
 * the last row of the longest of the funds, the overlaps, the mixes, and the
 * blocks that the script computes.
 */
function topLastOf(sheet, computed) {
  return Math.max(
    EQUITY_ROW + EQUITY.length - 1,
    ...[FUND_AT, OVERLAP_AT, MIX_AT].map((column) => 4 + blockRows(sheet, column, 1).length),
    ...[computed.unseen, computed.stockFunds, computed.own, computed.groups].map((rows) => 4 + rows.length),
  );
}

/**
 * Check the two data writes of a good run in the log entries after the
 * request: the blocks left of the lines block from B4 to AQ down to
 * `topLast`, then the lines block from AR4 down to the last part row and to
 * the column of the last position. No setValues call reaches a column after
 * the last position. The function prints the count of cells of the two
 * writes and the count of a write of the whole grid, and gives the first.
 */
function checkDataWrites(entries, name, topLast, partCount, idCount, sheet) {
  const writes = entries.filter((e) => e.op === "setValues");
  const [top, lines] = writes;
  check(
    top.row === 4 && top.column === 2 && top.rows === topLast - 3 && top.columns === LINE_AT - 2,
    `${name}: the first data write covers B4:AQ${topLast}, the blocks left of the lines block ` +
      `(${top.row},${top.column},${top.rows},${top.columns})`,
  );
  const lastColumn = SOURCE_AT - 1 + idCount;
  check(
    lines.row === 4 &&
      lines.column === LINE_AT &&
      lines.rows === partCount + 1 &&
      lines.columns === lastColumn - LINE_AT + 1,
    `${name}: the second data write covers row 4 to the last part row, from AR to the column of position ${idCount} ` +
      `(${lines.row},${lines.column},${lines.rows},${lines.columns})`,
  );
  check(
    writes.every((e) => e.column + e.columns - 1 <= lastColumn),
    `${name}: no setValues call reaches a column after the column of the last position`,
  );
  const written = top.rows * top.columns + lines.rows * lines.columns;
  const grid = (sheet.getMaxRows() - 3) * (sheet.getMaxColumns() - 1);
  check(written < grid, `${name}: the two writes hold fewer cells than a write of the whole grid`);
  console.log(`  ${name}: the data writes hold ${written} cells; a write of the whole grid holds ${grid}.`);
  return written;
}

/**
 * The sum of the numbers of a list of cells. A cell with no number adds 0.
 */
function sumCells(cells) {
  return cells.reduce((sum, value) => sum + (typeof value === "number" ? value : 0), 0);
}

/**
 * A synthetic answer of the route for the given position ids: one fund, no
 * pair of funds, no equity block, and the given count of lines. Line i has
 * one source, the position i modulo the count of ids, and no stock.
 */
function bigAnswer(ids, lineCount) {
  const weight = 1 / lineCount;
  return {
    measures: { ...Object.fromEntries(MEASURES.map((name, i) => [name, i + 0.5])), equity: null },
    funds: [
      {
        id: ids[0],
        ticker: ids[0],
        reportDate: "2026-06-30",
        accessionNumber: "0000000000-26-000001",
        holdingCount: 100,
        weight: 0.025,
        coveredWeight: 0.025,
        mergedByTicker: 0,
        mergedByLei: 0,
        mergedByName: 0,
      },
    ],
    overlaps: [],
    lines: Array.from({ length: lineCount }, (_, i) => ({
      key: `line:${i}`,
      name: `Line ${i}`,
      ticker: null,
      lei: null,
      class: "other",
      weight,
      sources: { [ids[i % ids.length]]: weight },
      stockWeight: 0,
      stockSources: {},
    })),
  };
}

/**
 * The FILTER calls of the report formulas that need no IFNA and no IFERROR,
 * by the text of their condition arguments. The guards are texts of the
 * same formula. Together they make a call with no match unused. The reason
 * states why.
 */
const GUARDED_FILTERS = {
  yok: {
    guards: ["IF(SUM(yok)=0,"],
    reason: "the section Holdings uses the Holdings rows only when a row holds a number in Value",
  },
  sel: {
    guards: ["top,IF(SUM(sel)=0,"],
    reason: "top uses the stock rows only when a stock is at or above the threshold",
  },
  keep: {
    guards: ["main,IF(nown=0,FILTER(grpAll,big),", "nis,IF(nown+ng=0,"],
    reason: "main uses ownRows only when nown, the count of the rows of the own block, is above 0",
  },
  grp: {
    guards: ["ng,IF(SUM(grp)=0,0,SUM(big)),", "nsm,IF(SUM(grp)=0,0,SUM(sm)),"],
    reason: "ng and nsm are 0 when the group block is empty, so no group row is used",
  },
  big: {
    guards: ["ng,IF(SUM(grp)=0,0,SUM(big)),", "main,IF(nown=0,FILTER(grpAll,big),IF(ng=0,", "nis,IF(nown+ng=0,"],
    reason: "main uses the big group rows only when ng, the count of the big groups, is above 0",
  },
  sm: {
    guards: ["nsm,IF(SUM(grp)=0,0,SUM(sm)),", "nis,IF(nown+ng=0,IF(nsm=0,", "IF(nsm=0,SORT(main,4,FALSE),"],
    reason: "nis uses smallRow only when nsm, the count of the small groups, is above 0",
  },
  psel: {
    guards: ["plist,IF(SUM(psel)=0,"],
    reason: "plist uses the pair rows only when a pair is at or above the overlap minimum",
  },
};

/**
 * Each formula of a tab state: each cell text that starts with "=", and the
 * formula of each conditional format.
 */
function formulasOf(stateText) {
  const tab = JSON.parse(stateText);
  const cells = [];
  tab.grid.forEach((row, r) => {
    row.forEach((value, c) => {
      if (typeof value === "string" && value.startsWith("=")) cells.push({ at: `row ${r + 1} column ${c + 1}`, value });
    });
  });
  return [...cells, ...tab.rules.map((rule, i) => ({ at: `conditional format ${i + 1}`, value: rule.formula }))];
}

/**
 * The string literals of a formula, and the formula with each literal
 * replaced by an empty literal.
 */
function splitLiterals(formula) {
  const literals = [];
  const code = formula.replace(/"(?:[^"]|"")*"/g, (text) => {
    literals.push(text);
    return '""';
  });
  return { code, literals };
}

/**
 * Each FILTER call of a formula. A call holds its text, the text of its
 * condition arguments, and the enclosing calls from the outside in, each
 * with the index of the argument that holds the FILTER. The parser skips
 * string literals, and it treats an array literal in braces as one call.
 */
function filterCalls(formula) {
  const calls = [];
  const stack = [];
  let i = 0;
  while (i < formula.length) {
    const ch = formula[i];
    if (ch === '"') {
      i += 1;
      while (i < formula.length && !(formula[i] === '"' && formula[i + 1] !== '"')) i += formula[i] === '"' ? 2 : 1;
    } else if (ch === "(" || ch === "{") {
      const name = ch === "{" ? "{" : (/[A-Za-z_][A-Za-z0-9_.]*$/.exec(formula.slice(0, i))?.[0] ?? "");
      const outer = stack.map((f) => ({ name: f.name, arg: f.arg }));
      stack.push({ name: name.toUpperCase(), arg: 0, start: i - name.length, commas: [], outer });
    } else if (ch === "," || ch === ";") {
      const top = stack.at(-1);
      if (top) {
        top.arg += 1;
        top.commas.push(i);
      }
    } else if (ch === ")" || ch === "}") {
      const frame = stack.pop();
      if (frame?.name === "FILTER") {
        const conditionAt = frame.commas[0] ?? i;
        calls.push({
          text: formula.slice(frame.start, i + 1),
          condition: formula.slice(conditionAt + 1, i),
          outer: frame.outer,
        });
      }
    }
    i += 1;
  }
  return calls;
}

/**
 * The treatment of one FILTER call: the IFNA or IFERROR that holds it in its
 * first argument, the guards of GUARDED_FILTERS, or null.
 */
function filterTreatment(call, formula) {
  const wrap = call.outer.findLast((f) => (f.name === "IFNA" || f.name === "IFERROR") && f.arg === 0);
  if (wrap) return wrap.name;
  const guarded = GUARDED_FILTERS[call.condition];
  if (guarded && guarded.guards.every((text) => formula.includes(text))) return `guard ${call.condition}`;
  return null;
}

/**
 * The arguments of the call whose open parenthesis is at `open`, as texts.
 * The function skips string literals, and it counts a brace as a
 * parenthesis.
 */
function callArgs(formula, open) {
  const args = [];
  let depth = 0;
  let start = open + 1;
  for (let i = open; i < formula.length; i += 1) {
    const ch = formula[i];
    if (ch === '"') {
      i += 1;
      while (i < formula.length && !(formula[i] === '"' && formula[i + 1] !== '"')) i += formula[i] === '"' ? 2 : 1;
    } else if (ch === "(" || ch === "{") {
      depth += 1;
    } else if (ch === ")" || ch === "}") {
      depth -= 1;
      if (depth === 0) {
        args.push(formula.slice(start, i));
        return args;
      }
    } else if (ch === "," && depth === 1) {
      args.push(formula.slice(start, i));
      start = i + 1;
    }
  }
  throw new Error(`The call at ${open} of the formula has no end.`);
}

/**
 * The name and the value text of each pair of each LET of a formula with no
 * string literals.
 */
function letDefinitions(code) {
  const names = new Map();
  for (const m of code.matchAll(/\bLET\(/g)) {
    const args = callArgs(code, m.index + m[0].length - 1);
    for (let k = 0; k + 1 < args.length; k += 2) names.set(args[k].trim(), args[k + 1].trim());
  }
  return names;
}

/**
 * True when a value text of a LET gives TRUE and FALSE: a comparison, or a
 * call of ISNUMBER, EXACT, or a function of that kind, with no arithmetic
 * outside the parentheses. One ARRAYFORMULA around the whole text does not
 * change the result.
 */
function isLogical(text) {
  let value = text.trim();
  if (/^ARRAYFORMULA\(/i.test(value)) {
    const args = callArgs(value, "ARRAYFORMULA".length);
    if (args.length === 1 && args[0].length + "ARRAYFORMULA()".length === value.length) value = args[0].trim();
  }
  let outside = "";
  let depth = 0;
  for (const ch of value) {
    if (ch === "(" || ch === "{") depth += 1;
    else if (ch === ")" || ch === "}") depth -= 1;
    else if (depth === 0) outside += ch;
  }
  if (/[*+\-/&]/.test(outside)) return false;
  if (/[=<>]/.test(outside)) return true;
  const call = /^(ISNUMBER|ISTEXT|ISBLANK|ISNA|ISERROR|ISERR|ISLOGICAL|EXACT|AND|OR|NOT)\(/i.exec(value);
  return call !== null && callArgs(value, call[0].length - 1).join(",").length + call[0].length + 1 === value.length;
}

function main() {
  if (OFFLINE) {
    console.log("OFFLINE: the harness skips each real request to the concentration route.\n");
  } else if (API_KEY === "") {
    console.error("Set HOLDINGS_API_KEY in the environment. The usage line in the header of check.mjs shows how.");
    process.exit(2);
  } else if (!LIVE_PARTS) {
    console.log("The fund mix run gets a synthetic answer. Add --live-parts to send it to the route.\n");
  }
  const context = vm.createContext({ ...services });
  for (const file of SCRIPT_FILES) {
    const source = readFileSync(file, "utf8");
    const name = file.slice(SRC.length + 1);
    for (const text of ["ScriptApp", "newTrigger", "getScriptProperties", "Logger"]) {
      check(!source.includes(text), `${name} holds no "${text}" text`);
    }
    check(
      [...source.matchAll(/console\.(\w+)/g)].every(([, method]) => method === "warn" || method === "error"),
      `${name} calls console.warn and console.error alone`,
    );
    vm.runInContext(source, context, { filename: file });
  }
  for (const name of [
    "onOpen",
    "onInstall",
    "onEdit",
    "setApiKey",
    "refreshConcentration",
    "ensureTabs",
    "writeChartBlock",
    "drawChart",
    "readInputs",
    "showMixSidebar",
    "mixSidebarData",
    "lookupFund",
    "saveMix",
    "deleteMix",
    "linkMix",
  ]) {
    check(typeof context[name] === "function", `the script defines ${name}`);
  }
  const accepted = {
    [EXPOSURE_TAB]: expectedState(acceptedDefinition(EXPOSURE_TAB)),
    [REPORT_TAB]: expectedState(acceptedDefinition(REPORT_TAB)),
  };

  /**
   * Run Refresh once, and give the time before and after the call in
   * milliseconds.
   */
  const timedRun = () => {
    const before = Date.now();
    context.refreshConcentration();
    return { before, after: Date.now() };
  };

  /**
   * The start time of each good run, newest first, as the harness expects
   * the run-time block to hold it.
   */
  const runStarts = [];

  /**
   * Check that a good run put its start time and its seconds at the top of
   * the run-time block, that the block holds the 10 newest good runs alone,
   * newest first, and that the run drew the bar chart.
   */
  const checkRecorded = (name, times) => {
    const rows = runRows();
    const [start, seconds] = rows[0];
    check(start >= times.before && start <= times.after, `${name}: A${RUN_ROW} holds the start time of the run`);
    check(
      typeof seconds === "number" && seconds >= 0 && seconds <= (times.after - times.before) / 1000 + 1e-9,
      `${name}: B${RUN_ROW} holds the seconds of the run (${seconds})`,
    );
    runStarts.unshift(start);
    checkBlock(name);
    checkChart(name);
    checkHoldingsChart(name);
  };

  /**
   * Check that the run-time block holds the 10 newest good runs alone,
   * newest first.
   */
  const checkBlock = (name) => {
    const kept = runStarts.slice(0, 10);
    check(
      runRows().every(([s, sec], n) =>
        n < kept.length ? s === kept[n] && typeof sec === "number" : s === "" && sec === "",
      ),
      `${name}: the run-time block holds the ${kept.length} newest good runs, newest first`,
    );
  };

  /**
   * Check that a run with a fault, or a run that cannot get the lock, adds
   * no row to the run-time block.
   */
  const checkNoRecord = (name) => checkBlock(`${name} adds no run time`);

  console.log("== Manifest");
  const manifest = JSON.parse(readFileSync(MANIFEST_FILE, "utf8"));
  check(JSON.stringify(manifest.oauthScopes) === JSON.stringify(SCOPES), "the manifest holds the three scopes alone");
  check(
    JSON.stringify(manifest.urlFetchWhitelist) === JSON.stringify(URL_PREFIXES),
    "the manifest allows fetches under https://data.coopersbs.com/funds/v1/ alone",
  );
  check(ROUTE_URL.startsWith(URL_PREFIXES[0]), "the route is under the allowed prefix");
  check(manifest.runtimeVersion === "V8", "the manifest sets the V8 runtime");
  console.log("  pass");

  console.log("\n== Pure functions");
  check(context.faultText(502, "<html>") === "FAULT: 502", "a body that is not JSON gives the status alone");
  check(
    context.faultText(401, '{"error":{"code":"invalid_key","status":401}}') === "FAULT: 401 invalid_key",
    "a JSON error body gives the status and the code",
  );
  const long = "D".repeat(70);
  const cut = context.buildPositions(
    [
      ["", 1, long],
      ["", 1, `${long}x`],
    ],
    { symbol: 0, value: 1, description: 2 },
  );
  check(cut.length === 1 && cut[0].id.length === 64, "two keys with one first 64 code units give one position");
  check(context.tickerOf("BRK.B") === "BRK.B" && context.tickerOf("$abc") === "$abc", "a valid symbol is a ticker");
  check(context.tickerOf("CASH SWEEP") === null, "a symbol that fails the ticker pattern gives no ticker");
  check(
    context.normalTicker(" $vtiax ") === "VTIAX" && context.normalTicker("A B") === null,
    "normalTicker trims, removes one $, and changes to upper case",
  );
  check(
    context.daySerial("1899-12-30") === 0 && context.daySerial("2026-09-30") === 46295,
    "daySerial gives the date serial number of Google Sheets",
  );
  check(
    ["10%", "0.1", " 10 % ", ".1", 0.1].every((value) => context.typedShare(value) === 0.1) &&
      context.typedShare("100%") === 1 &&
      context.typedShare("0") === 0,
    "typedShare reads a typed 10% and a typed 0.1 as the share 0.1",
  );
  check(
    ["150%", "1.5", "-1%", "ten", "1,5", "", undefined, null, 2, -0.5].every(
      (value) => context.typedShare(value) === null,
    ),
    "typedShare gives no share for a value outside 0 to 1, an empty value, or a text that is not a number",
  );
  check(
    [0.1, 0.01, 0.025, 1, 0].every((share) => context.headingText(share) === spillHeading(share)) &&
      context.headingText(0.01005) === spillHeading(0.0101),
    "headingText gives the text of the chart header row with the threshold as a percent with two decimals, and " +
      "1.005% rounds up",
  );
  check(
    context.sameHeading(spillHeading(0.1).replace(".", ","), context.headingText(0.1)) &&
      context.sameHeading(spillHeading(0.1), context.headingText(0.1)) &&
      !context.sameHeading(spillHeading(0.05).replace(".", ","), context.headingText(0.1)),
    "sameHeading accepts the chart header row with a comma or a dot as the decimal mark, and refuses another " +
      "threshold",
  );
  const part = (ticker, percent, substitute = false) => ({ ticker, percent, substitute });
  check(context.mixProblem([part("VOO", 60), part("IVV", 40)]) === "", "a mix of 100% is valid");
  check(
    context.mixProblem([part("VOO", 33.33), part("IVV", 33.33), part("AAPL", 33.34)]) === "",
    "33.33 + 33.33 + 33.34 is valid",
  );
  const problems = [
    [[], "Add at least one fund."],
    [Array.from({ length: 21 }, (_, n) => part(`T${n}`, 1)), "A mix can hold 20 funds at most."],
    [[part("VOO", 60), part("IVV", 40.01)], "The total is over 100%. Lower a percent to save."],
    [[part("VOO", 10), part("VOO", 20)], "VOO is in the mix two times."],
    [[part("voo", 10)], 'Check the ticker "voo".'],
    [[part("VOO", 0)], "Give VOO a percent above 0 and up to 100."],
    [[{ ticker: "VOO", percent: 10 }], "Each row needs a substitute box."],
  ];
  for (const [parts, text] of problems) {
    check(context.mixProblem(parts) === text, `mixProblem gives "${text}"`);
  }
  check(
    JSON.stringify(
      context.parseMix('{"entered":"2026-02-30","parts":[{"ticker":"VOO","percent":1,"substitute":false}]}'),
    ) === "null" &&
      context.parseMix("{not json") === null &&
      context.parseMix('{"entered":"2026-01-05","parts":[]}') === null,
    "parseMix refuses a day that does not exist, a text that is not JSON, and a mix with no fund",
  );
  const mixes = new Map([["PLAN", { entered: "2026-01-05", parts: [part("VOO", 50, true), part("AAPL", 25)] }]]);
  const mixed = context.buildPositions(
    [
      ["PLAN", 3, "Example plan"],
      ["VOO", 1, "Example fund"],
    ],
    { symbol: 0, value: 1, description: 2 },
    mixes,
  );
  check(
    JSON.stringify(mixed) ===
      JSON.stringify([
        {
          id: "PLAN",
          weight: 0.75,
          parts: [
            { ticker: "VOO", weight: 0.5 },
            { ticker: "AAPL", weight: 0.25 },
          ],
        },
        { id: "VOO", ticker: "VOO", weight: 0.25 },
      ]),
    `a holding with a mix sends id, weight, and parts, and no ticker (${JSON.stringify(mixed)})`,
  );
  check(
    JSON.stringify(context.mixRows(mixed, mixes)) ===
      JSON.stringify([
        ["PLAN", 0.75, context.daySerial("2026-01-05"), "VOO", 0.5, true],
        ["PLAN", 0.75, context.daySerial("2026-01-05"), "AAPL", 0.25, false],
      ]),
    "mixRows gives one row for each fund of each mix, with the substitute flag",
  );
  check(context.columnNumber("A") === 1 && context.columnNumber("AZ") === 52, "columnNumber reads A and AZ");
  check(
    [1, 26, 27, 52, 220, 702, 703].every((n) => context.columnNumber(context.columnLetter(n)) === n),
    "columnLetter is the inverse of columnNumber",
  );
  const small = new FakeSheet("small", 2, 2, []);
  context.fitGrid(small, 3, 5);
  check(
    small.getMaxRows() === TAB_ROWS && small.getMaxColumns() === 5,
    `fitGrid adds rows up to ${TAB_ROWS} at least, and adds columns`,
  );
  context.fitGrid(small, TAB_ROWS + 50, 1);
  check(
    small.getMaxRows() === TAB_ROWS + 50 && small.getMaxColumns() === 5,
    "fitGrid adds rows for a longer answer and deletes no column",
  );
  context.fitGrid(small, 10, 1);
  check(small.getMaxRows() === TAB_ROWS, `fitGrid deletes the rows after ${TAB_ROWS} for a shorter answer`);

  console.log("\n== Values that the script computes from an answer");
  const computeIds = ["F", "S", "P", "M"];
  /**
   * The funds block of the hand answer. F and M entered through their
   * holdings. F also holds a held fund that entered and one that did not. S
   * names a fund that did not enter, so S is a direct position. An element
   * with no entered key and no heldBy key entered.
   */
  const computeFunds = [
    { id: "F", ticker: "F", heldBy: null, entered: true, accessionNumber: "0000000000-26-000001" },
    { id: "M" },
    { id: "F", ticker: "F2", heldBy: null, entered: true, accessionNumber: "0000000000-26-000002" },
    { id: "F", ticker: null, heldBy: "S000000001", entered: true, accessionNumber: "0000000000-26-000003" },
    { id: "F", ticker: null, heldBy: "S000000001", entered: false, notEnteredReason: "unknown_series" },
    { id: "S", ticker: "S", heldBy: null, entered: false, notEnteredReason: "no_report", accessionNumber: null },
  ];
  const computeRows = context.fundRows(computeFunds);
  check(
    JSON.stringify(computeRows.map((cells) => cells.slice(0, 2))) ===
      JSON.stringify([
        ["F", "F"],
        ["M", ""],
        ["F", "F2"],
      ]) && JSON.stringify(context.fundIds(computeFunds)) === JSON.stringify(["F", "M"]),
    "fundRows and fundIds keep each element with no heldBy that entered, and an element with no entered key " +
      `(${JSON.stringify(computeRows.map((cells) => cells[0]))})`,
  );
  const nullRow = context.fundRows([{ id: "N", ticker: null, accessionNumber: null, reportDate: null }])[0];
  check(
    nullRow[FUND_FIELDS.indexOf("ticker")] === "" &&
      nullRow[FUND_FIELDS.indexOf("accessionNumber")] === "" &&
      nullRow[FUND_FIELDS.indexOf("reportDate")] === "",
    "fundRows writes an empty cell for a null ticker, a null accessionNumber, and a null reportDate",
  );
  const lineOf = (key, kind, weight, sources, stockWeight = 0, stockSources = {}) => ({
    key,
    name: `Name of ${key}`,
    ticker: null,
    lei: null,
    class: kind,
    weight,
    sources,
    stockWeight,
    stockSources,
  });
  const computeLines = [
    lineOf("lei:X", "stock", 0.4, { F: 0.25, S: 0.15 }, 0.4, { F: 0.25, S: 0.15 }),
    lineOf("name:BOND", "other", 0.1, { F: 0.06, S: 0.04 }),
    lineOf("name:CASH", "cash", 0.05, { F: 0.03, M: 0.02 }),
    lineOf("name:TBILL", "treasury", 0.02, { M: 0.02 }),
    lineOf("name:CASH2", "cash", 0.01, { F: 0.01 }),
    lineOf("id:P", "unknown", 0.3, { P: 0.3 }),
    lineOf("residual:F", "cash", 0, { F: 0 }),
    lineOf("residual:M", "cash", 0.01, { M: 0.01 }),
    lineOf("id:M", "unknown", 0.11, { M: 0.11 }),
    lineOf("name:Q", "unknown", 0.02, { F: 0.02 }),
    lineOf("lei:Y", "stock", 0.03, { F: 0.03 }, 0.02, { F: 0.02 }),
    lineOf("lei:PREF", "preferred", 0.004, { F: 0.004 }),
    lineOf("other:lines", "other", 0.006, { F: 0.006 }, 0.004, { F: 0.004 }),
  ];
  const computeParts = context.partRows(computeLines, computeIds);
  const computeDirect = context.directWeights(computeParts, computeIds, computeFunds);
  const wantDirect = [0.15, 0.04, 0, 0, 0, 0.3, 0, 0, 0, 0, 0, 0, 0, 0, 0];
  check(
    computeDirect.length === wantDirect.length && computeDirect.every((v, r) => Math.abs(v - wantDirect[r]) < 1e-12),
    "directWeights gives the part weight minus the cells of the funds F and M, and keeps the cells of S, " +
      `a fund that did not enter (${JSON.stringify(computeDirect)})`,
  );
  const computeOwn = context.ownRows(computeParts, computeIds, computeDirect);
  check(
    sameRows(computeOwn, [
      ["name:BOND", "Name of name:BOND", "other", 0.1, "F, S"],
      ["id:P", "Name of id:P", "unknown", 0.3, "P"],
      ["residual:M", "M", "cash", 0.01, "M"],
      ["id:M", "Name of id:M", "unknown", 0.11, "M"],
      ["name:Q", "Name of name:Q", "unknown", 0.02, "F"],
      ["other:lines", "Name of other:lines", "other", 0.006 - 0.004, "F"],
    ]),
    "ownRows gives a direct line, each line of the class unknown, a residual line above 0 with the fund ticker as " +
      `the name, and the other part of the cap line, in their order (${JSON.stringify(computeOwn)})`,
  );
  const computeGroups = context.groupRows(computeParts, computeIds, computeDirect);
  check(
    sameRows(computeGroups, [
      ["cash", 0.05 + 0.01, 2, "F, M"],
      ["treasury", 0.02, 1, "M"],
      ["stock", 0.03 - 0.02, 1, "F"],
      ["preferred", 0.004, 1, "F"],
    ]),
    "groupRows gives each class of the other rows with no row of their own, with the weight, the count, " +
      `and the positions, a group of the class preferred, and skips a residual line of 0 (${JSON.stringify(computeGroups)})`,
  );
  const computeMixes = [["M", 0.2, 46000, "VOO", 1, false]];
  const computeUnseen = context.unseenRows(computeParts, computeIds, computeMixes);
  check(
    sameRows(computeUnseen, [
      ["F", 0.02],
      ["S", 0],
      ["P", 0.3],
      ["M", 0],
    ]),
    `unseenRows gives each position its weight in the lines of the class unknown, and 0 with a mix (${JSON.stringify(computeUnseen)})`,
  );
  const computeStock = context.stockFundRows(computeParts, computeIds, computeFunds);
  check(
    JSON.stringify(computeStock) === JSON.stringify([["F"]]),
    "stockFundRows gives each fund one time, no fund whose stock cells add up to 0, and no fund that did not " +
      `enter (${JSON.stringify(computeStock)})`,
  );
  const computeSheet = new FakeSheet(EXPOSURE_TAB, 40, SOURCE_AT + MAX_POSITIONS - 1, []);
  computeSheet.grid[3].splice(SOURCE_AT - 1, computeIds.length, ...computeIds);
  computeRows.forEach((cells, r) => {
    computeSheet.grid[4 + r][FUND_AT - 1] = cells[0];
  });
  computeSheet.grid[4][MIX_AT - 1] = "M";
  computeParts.forEach((cells, r) => {
    const row = [...cells.slice(0, LINE_FIELDS.length + 1), computeDirect[r], ...cells.slice(LINE_FIELDS.length + 1)];
    computeSheet.grid[4 + r].splice(LINE_AT - 1, row.length, ...row);
  });
  const placeRows = (column, rows) => {
    rows.forEach((cells, r) => computeSheet.grid[4 + r].splice(column - 1, cells.length, ...cells));
  };
  placeRows(OWN_AT, computeOwn);
  placeRows(GROUP_AT, computeGroups);
  placeRows(UNSEEN_AT, computeUnseen);
  placeRows(STOCK_FUND_AT, computeStock);
  checkComputed(computeSheet, "the hand answer", computeParts.length);
  const handOwn = blockRows(computeSheet, OWN_AT, 5);
  const handGroups = blockRows(computeSheet, GROUP_AT, 4);
  const handOther = computeParts
    .filter((cells) => cells[LINE_FIELDS.length] === "other")
    .reduce((sum, cells) => sum + cells[LINE_FIELDS.indexOf("weight")] - cells[LINE_FIELDS.indexOf("stockWeight")], 0);
  const handSection = sumCells(handOwn.map((cells) => cells[3])) + sumCells(handGroups.map((cells) => cells[1]));
  check(
    handOwn.some((cells) => cells[2] === "unknown") &&
      handOwn.some((cells) => String(cells[0]).startsWith("residual:")) &&
      ["cash", "treasury"].every((kind) => handGroups.some((cells) => cells[0] === kind)) &&
      Math.abs(handSection - handOther) < 1e-12,
    "the own rows and the group rows of the hand answer, with lines of the class unknown, a residual line, a cash " +
      `group, and a treasury group, add up to the Other holdings row of the composition (${handSection} for ${handOther})`,
  );

  const handLines = [
    {
      key: "name:BOTH",
      class: "stock",
      weight: 0.5,
      sources: { A: 0.25, B: 0.25 },
      stockWeight: 0.375,
      stockSources: { A: 0.25, B: 0.125 },
    },
    { key: "name:BOND", class: "other", weight: 0.125, sources: { A: 0.125 }, stockWeight: 0, stockSources: {} },
    {
      key: "name:STOCK",
      class: "stock",
      weight: 0.0625,
      sources: { B: 0.0625 },
      stockWeight: 0.0625,
      stockSources: { B: 0.0625 },
    },
    { key: "residual:A", class: "other", weight: 0, sources: { A: 0 }, stockWeight: 0, stockSources: {} },
    {
      key: "name:SHORT",
      class: "stock",
      weight: 0,
      sources: { A: 0.5, B: -0.5 },
      stockWeight: 0,
      stockSources: { A: 0.5, B: -0.5 },
    },
  ];
  const handRows = context
    .partRows(handLines, ["A", "B"])
    .map((cells) => [cells[0], ...cells.slice(LINE_FIELDS.length)]);
  check(
    JSON.stringify(handRows) ===
      JSON.stringify([
        ["name:BOTH", "stock", 0.25, 0.125],
        ["name:BOTH", "other", "", 0.125],
        ["name:BOND", "other", 0.125, ""],
        ["name:STOCK", "stock", "", 0.0625],
        ["residual:A", "other", "", ""],
        ["name:SHORT", "stock", 0.5, -0.5],
      ]),
    `partRows gives the stock part and the other part of each source (${JSON.stringify(handRows)})`,
  );

  const defaults = JSON.stringify({ threshold: 0.01, overlapMinimum: 0.1 });
  const typed = new FakeSheet("typed", 30, 3, []);
  check(JSON.stringify(context.readInputs(null)) === defaults, "readInputs gives the defaults for an absent tab");
  check(
    JSON.stringify(context.readInputs(typed)) === defaults,
    "readInputs gives the defaults for a tab with no label",
  );
  typed.grid[4] = ["Threshold", 1.5, ""];
  typed.grid[5] = [" Overlap minimum ", "10%", ""];
  check(
    JSON.stringify(context.readInputs(typed)) === defaults,
    "readInputs gives the default for a number above 1 and for a text",
  );
  typed.grid[4][1] = 0;
  typed.grid[5][1] = 1;
  check(
    JSON.stringify(context.readInputs(typed)) === JSON.stringify({ threshold: 0, overlapMinimum: 1 }),
    "readInputs reads 0 and 1 from the cells next to the two labels",
  );
  console.log("  pass");

  console.log("\n== onOpen and onInstall");
  const menuItems = [
    { label: "Refresh", handler: "refreshConcentration" },
    { label: "Describe a fund", handler: "showMixSidebar" },
    { label: "Set API key", handler: "setApiKey" },
  ];
  context.onOpen({ authMode: "NONE" });
  context.onInstall({ authMode: "FULL" });
  check(state.menus.length === 2, "onOpen and onInstall each add one menu");
  for (const menu of state.menus) {
    check(
      menu.addon === true && JSON.stringify(menu.items) === JSON.stringify(menuItems),
      "the add-on menu holds Refresh, Describe a fund, and Set API key",
    );
  }
  check(state.log.length === 0, "onOpen and onInstall change no tab");
  check(
    state.userProperties.size === 0 && state.documentProperties.size === 0,
    "onOpen and onInstall write no property",
  );
  check(
    state.sidebars.length === 0 && state.prompts.length === 0,
    "onOpen and onInstall open no sidebar and no dialog",
  );
  console.log("  pass");

  console.log("\n== Set API key");
  state.promptAnswers.push({ button: BUTTON.CANCEL, text: API_KEY });
  context.setApiKey();
  check(state.userProperties.size === 0 && state.alerts.length === 0, "Cancel saves nothing and shows no message");
  state.promptAnswers.push({ button: BUTTON.CLOSE, text: API_KEY });
  context.setApiKey();
  check(state.userProperties.size === 0 && state.alerts.length === 0, "the close button saves nothing");
  state.promptAnswers.push({ button: BUTTON.OK, text: "not a key\nsecond line" });
  context.setApiKey();
  check(state.userProperties.size === 0, "a text with a space or a line break is not saved");
  check(
    state.alerts.at(-1) === "This text is not an API key. Nothing changed.",
    "the person sees that nothing changed",
  );
  state.promptAnswers.push({ button: BUTTON.OK, text: `  ${API_KEY}\n` });
  context.setApiKey();
  check(state.userProperties.get(KEY_PROPERTY) === API_KEY, "OK saves the key without the white space");
  check(state.userProperties.size === 1, "the key is the one user property");
  check(state.alerts.at(-1) === "The API key is saved. Use Refresh in the add-on menu.", "the person sees the save");
  check(
    state.prompts.every((p) => p.buttons === BUTTON_SET.OK_CANCEL),
    "the input dialog has the buttons OK and Cancel",
  );
  check(
    [...state.alerts, ...state.prompts.flatMap((p) => [p.title, p.text])].every((text) => !text.includes(API_KEY)),
    "no dialog text and no message holds the key",
  );
  check(state.log.length === 0, "Set API key changes no tab");
  console.log("  pass");

  console.log("\n== Run 1: the Holdings tab alone, then one real request to the route");
  check(JSON.stringify(book.names()) === '["Holdings"]', "the spreadsheet fake starts with the Holdings tab alone");
  state.fetchHandler = OFFLINE ? offlineFetch : liveFetch;
  const logStart = state.log.length;
  const runOne = timedRun();
  liveText = answerText;
  check(state.fetchCalls.length === 1 && liveRequests + offlineRequests === 1, "run 1 sends one request");
  check(state.lockTaken === 1 && state.lockReleased === 1, "run 1 takes the lock and releases it");

  const inserts = state.log.slice(logStart).filter((e) => e.op === "insertSheet");
  check(
    inserts.map((e) => e.sheet).join() === `${EXPOSURE_TAB},${REPORT_TAB}`,
    "run 1 creates Concentration.Exposure, then Concentration",
  );
  check(
    JSON.stringify(book.names()) === JSON.stringify(["Holdings", EXPOSURE_TAB, REPORT_TAB]),
    "the two new tabs follow the Holdings tab",
  );
  const fetchAt = indexSince(logStart, (e) => e.op === "fetch");
  const createCalls = state.log.slice(logStart, fetchAt);
  check(
    fetchAt > logStart && createCalls.every((e) => e.sheet === EXPOSURE_TAB || e.sheet === REPORT_TAB),
    `the script creates and formats the two tabs before the request (${createCalls.length} calls)`,
  );
  const [snap] = state.atFetch;
  check(snap.names.length === 3, "both tabs exist at the time of the request");
  checkBlockWrite("run 1", logStart);
  const created = Object.fromEntries(snap.tabs);
  for (const tab of [EXPOSURE_TAB, REPORT_TAB]) checkLayout(tab, created[tab], accepted[tab]);
  const report = book.tab(REPORT_TAB);
  const layoutVersion = vm.runInContext("LAYOUT_VERSION", context);
  check(
    Number.isInteger(layoutVersion) && JSON.parse(created[EXPOSURE_TAB]).grid[2][1] === layoutVersion,
    `${VERSION_CELL} of the hidden tab holds the layout version of layout.gs (${layoutVersion})`,
  );
  check(report.cell(THRESHOLD_CELL) === 0.01, `the threshold cell ${THRESHOLD_CELL} holds its default 0.01`);
  check(report.cell(MINIMUM_CELL) === 0.1, `the overlap minimum cell ${MINIMUM_CELL} holds its default 0.1`);
  check(
    report.cell(SPILL_CELL).startsWith("=LET(") && report.cell(SPILL_CELL).includes("\n"),
    `${SPILL_CELL} holds the report formula`,
  );

  console.log("\n== Report formulas");
  const formulas = [EXPOSURE_TAB, REPORT_TAB].flatMap((tab) =>
    formulasOf(created[tab]).map((f) => ({ ...f, at: `${tab} ${f.at}` })),
  );
  check(formulas.length > 10, `the two tabs hold the report formulas (${formulas.length})`);

  const lettered = formulas.filter(({ value }) => {
    const { code, literals } = splitLiterals(value);
    return /'?Holdings'?!\$?[A-Z]/i.test(code) || literals.some((text) => /Holdings'?!\$?[A-Z]+\$?[\d:]/i.test(text));
  });
  check(
    lettered.length === 0,
    `no formula holds a reference of the form Holdings!<letter> (${lettered.map((f) => f.at).join(", ")})`,
  );
  const holdingsColumns = vm.runInContext("HOLDINGS_COLUMNS", context);
  check(
    report.cell(TOTAL_CELL).includes('EXACT(TRIM(Holdings!$1:$1),"Value")'),
    `${TOTAL_CELL} finds the Value column by the header text in row 1`,
  );
  check(
    holdingsColumns.every(
      (name) => report.cell(TOTAL_CELL).includes(`"${name}"`) || report.cell(SPILL_CELL).includes(`"${name}"`),
    ),
    `${TOTAL_CELL} and ${SPILL_CELL} find each column of findColumns by its header text (${holdingsColumns.join(", ")})`,
  );
  check(
    report.cell(TOTAL_CELL).includes('"No Holdings column Value"'),
    `${TOTAL_CELL} shows a text when no Value column exists`,
  );

  const lastSource = vm.runInContext("SOURCE_COLUMN + MAX_POSITIONS - 1", context);
  const firstSource = vm.runInContext("SOURCE_COLUMN", context);
  check(firstSource === SOURCE_AT, `the sources block starts at column ${SOURCE_AT}`);
  const exposureRange = /'Concentration\.Exposure'!\$?([A-Z]+)\$?(\d+):\$?([A-Z]+)\$?(\d*)/g;
  /**
   * The ranges of a formula that start in the first column of the sources
   * block.
   */
  const sourceRangesOf = (cell) =>
    [...report.cell(cell).matchAll(exposureRange)].filter((m) => columnNumber(m[1]) === firstSource);
  check(
    sourceRangesOf(NOTE_CELL).length === 0,
    `${NOTE_CELL} reads a block that the script computes, and no range of the sources block`,
  );
  const spillRanges = sourceRangesOf(SPILL_CELL);
  check(
    spillRanges.length === 2,
    `${SPILL_CELL} reads the sources block and the id row (${spillRanges.length} ranges)`,
  );
  check(
    spillRanges.every((m) => columnNumber(m[3]) === lastSource),
    `the source ranges of ${SPILL_CELL} end at column ${lastSource}, the column of position MAX_POSITIONS ` +
      `(${spillRanges.map((m) => m[0]).join(", ")})`,
  );
  check(
    [NOTE_CELL, SPILL_CELL].every((cell) => !/\b(MMULT|BYROW)\(|\bnum,|\bisf,|\bfid,/.test(report.cell(cell))),
    `${NOTE_CELL} and ${SPILL_CELL} hold no MMULT, no BYROW, and no num, isf, or fid`,
  );
  for (const text of [
    "thr,$B$23,",
    'sel,ARRAYFORMULA((LP="stock")*ISNUMBER(LX)*(LX>=thr)*(LK<>"other:lines")),',
    'rsel,ARRAYFORMULA((LP="stock")*ISNUMBER(LX)*((LX<thr)+(LK="other:lines")>0)),',
    'rc,COUNTIFS(LP,"stock",LX,"<"&thr,LX,"<>0",LK,"<>other:lines"),',
    "rf,MAP(f,LAMBDA(x,IFERROR(SUM(FILTER(INDEX(LS,0,XMATCH(x,LH)),rsel)),0))),",
    "psel,ARRAYFORMULA(ISNUMBER(po)*(po>=$B$24)),",
    "IF(ABS(gw*tot)>=100,1,0)",
  ]) {
    check(
      report.cell(SPILL_CELL).includes(text),
      `${SPILL_CELL} computes ${text.slice(0, 40)} in the formula, so a change of the threshold, the overlap ` +
        "minimum, or the Holdings tab shows with no refresh",
    );
  }
  const exposureState = JSON.parse(created[EXPOSURE_TAB]);
  const exposureStyles = new Map(exposureState.styles);
  check(exposureState.grid[0].length >= lastSource, `the Exposure grid holds column ${lastSource}`);
  check(
    new Map(exposureState.widths).get(lastSource) === 110,
    `the column width of the sources block reaches column ${lastSource}`,
  );
  check(
    exposureStyles.get(cellKey(4, lastSource))?.fontWeight === "bold" &&
      exposureStyles.get(cellKey(5, lastSource))?.numberFormat === "0.00000",
    `the header format and the number format of the sources block reach column ${lastSource}`,
  );

  check(
    vm.runInContext("CHART_COLUMN", context) === CHART_AT && CHART_AT === lastSource + 2,
    `the chart block starts at column ${CHART_AT}, one empty column after the sources block`,
  );
  check(
    vm.runInContext("ANCHOR_COLUMN", context) === ANCHOR_AT,
    `the anchor cell is in column ${ANCHOR_AT}, one empty column after the chart block`,
  );
  const chartBlock = exposureState.grid
    .slice(3, 4 + CHART_ROWS)
    .map((cells) => cells.slice(CHART_AT - 1, CHART_AT - 1 + CHART_WIDTH));
  const [chartHeader, chartRow] = chartBlock;
  check(
    chartBlock.flat().every((value) => typeof value === "string" && value.startsWith("=")),
    `each of the ${chartBlock.flat().length} cells of the chart block IT4:JA${4 + CHART_ROWS} is a formula`,
  );
  check(
    exposureState.grid[0].length === HOLDINGS_ANCHOR_AT &&
      exposureState.grid.every((cells) =>
        [CHART_AT, ANCHOR_AT, HOLDINGS_AT, HOLDINGS_ANCHOR_AT].every((column) => cells[column - 2] === ""),
      ) &&
      exposureState.grid
        .slice(4 + CHART_ROWS)
        .every((cells) => cells.slice(CHART_AT - 1, ANCHOR_AT).every((v) => v === "")) &&
      exposureState.grid
        .slice(4 + HOLDINGS_CHART_ROWS)
        .every((cells) => cells.slice(HOLDINGS_AT - 1).every((v) => v === "")) &&
      exposureState.grid.every(
        (cells, r) => r === 3 || r === 4 || (cells[ANCHOR_AT - 1] === "" && cells[HOLDINGS_ANCHOR_AT - 1] === ""),
      ),
    "the grid ends at the holdings anchor column, the four gap columns are empty, and no cell under a block or " +
      "around an anchor cell holds a value",
  );
  /**
   * The text of the chart header row of the spill, as an expression with
   * the threshold.
   */
  const heading = (threshold) => `"Companies at or over "&TEXT(${threshold},"0.00%")&" of the portfolio, by source"`;
  check(
    report
      .cell(SPILL_CELL)
      .includes(` top,rest,"",\n ${heading("thr")},\n MAKEARRAY(${BAND_ROWS},1,LAMBDA(i,j,"")),\n "Fund overlap",`),
    `${SPILL_CELL} puts the chart header row with the threshold and a band of ${BAND_ROWS} blank rows between the ` +
      "company table and the fund overlap list",
  );
  check(
    (BAND_ROWS - 1) * 21 >= CHART_OPTIONS.height && (BAND_ROWS - 2) * 21 < CHART_OPTIONS.height,
    `the band of ${BAND_ROWS} rows holds the ${CHART_OPTIONS.height} px of the chart at 21 px a row, and one row of margin`,
  );
  check(
    exposureState.grid[3][ANCHOR_AT - 1] === "chartRow" &&
      exposureState.grid[4][ANCHOR_AT - 1] ===
        `=IFERROR(XMATCH(${heading("'Concentration'!$B$23")},'Concentration'!$A:$A)+1,"")`,
    "the anchor cell JC5 finds the text of the chart header row in column A of the report tab, with the threshold " +
      "cell, and adds 1",
  );
  check(
    chartBlock.slice(1).every((cells) => JSON.stringify(cells) === JSON.stringify(chartRow)) &&
      chartRow.every((formula) => formula.includes("r,ROW()-4,")),
    `each of the ${CHART_ROWS} data rows holds the formulas of row 5, which find the company of rank ROW()-4`,
  );
  check(
    chartHeader[0] === '="Company"' &&
      chartHeader[1].includes(
        `IF(m=0,"",IF(SUM(CHOOSEROWS('Concentration'!$G$26:$G,SEQUENCE(m,1,h+1)))=0,"","Direct"))`,
      ) &&
      chartHeader.at(-1).includes(`IF(m=0,"",LET(g,`) &&
      chartHeader.at(-1).includes(`IF(nz<=${CHART_FUNDS},"","Other funds")`),
    "the header row names the company column; the Direct header is empty when the direct weights of the companies " +
      `add up to 0, and the Other funds header is empty with ${CHART_FUNDS} funds or fewer`,
  );
  check(
    [1, 2, 3, 4, 5].every(
      (k) =>
        chartHeader[1 + k].includes(`LEFT(INDEX('Concentration'!$H$26:$Z,h,INDEX(o,${k})),24)`) &&
        chartRow[1 + k].includes(`INDEX(g,r,INDEX(o,${k}))`),
    ),
    "fund column k shows the fund at place k of the order o, and its header is the id of the fund, cut to 24 characters",
  );
  check(
    chartRow.every((formula) => formula.includes('IF(r>m,"",')),
    "a data row with no company of its rank is blank",
  );
  check(
    chartRow.at(-1).includes(`IF(nz<=${CHART_FUNDS},"",`) && chartRow.at(-1).includes(`XMATCH(j,o)>${CHART_FUNDS}`),
    `the Other funds column adds each fund after place ${CHART_FUNDS}, and it is blank with ${CHART_FUNDS} funds or fewer`,
  );
  const blockReferences = chartBlock
    .flat()
    .flatMap((formula) => [...formula.matchAll(/'([^']+)'!\$?([A-Z]+)\$?(\d+)/g)]);
  check(
    blockReferences.length > 0 &&
      blockReferences.every((m) => m[1] === REPORT_TAB && m[3] === "26" && ["A", "B", "G", "H"].includes(m[2])),
    "each formula of the chart block reads the company table of the report spill alone: the columns A, B, G, and " +
      "H:Z from row 26, and no value that the script computes",
  );
  const blockRanges = new Set(blockReferences.map((m) => `${m[2]}${m[3]}`));
  check(
    JSON.stringify([...blockRanges].sort()) === JSON.stringify(["A26", "B26", "G26", "H26"]),
    `the chart block reads the rank, the company, the direct, and the fund columns of the spill (${[...blockRanges].join(", ")})`,
  );
  const exposureReads = formulas
    .filter((f) => f.at.startsWith(`${REPORT_TAB} `))
    .flatMap(({ value }) => [...value.matchAll(/'Concentration\.Exposure'!\$?([A-Z]+)\$?\d+(?::\$?([A-Z]+))?/g)])
    .flatMap((m) => [m[1], m[2]].filter(Boolean));
  check(
    exposureReads.length > 0 && exposureReads.every((letters) => columnNumber(letters) <= lastSource),
    `no formula of the report tab reads a column after column ${lastSource}, so the chart block makes no circular reference`,
  );
  check(
    report.cell(SPILL_CELL).includes('SPARKLINE(x,{"charttype","bar";"max",mx;"color1","#2a78d6"})'),
    "the company table keeps its column of SPARKLINE bars",
  );
  check(
    report.cell(SPILL_CELL).includes('rest,HSTACK(SUM(sel)+1,"Securities under "'),
    "the row of the securities under the threshold holds the next rank, the count of the company rows plus 1",
  );
  const reportRules = JSON.parse(created[REPORT_TAB]).rules;
  const restRules = reportRules.filter((rule) => rule.formula === '=LEFT($B26,17)="Securities under "');
  check(
    restRules.length === 1 &&
      restRules[0].background === undefined &&
      restRules[0].italic === true &&
      restRules[0].color === "#3c3b37",
    "the row of the securities under the threshold is italic, with no background",
  );
  check(
    reportRules.some(
      (rule) =>
        rule.bold === true &&
        rule.formula === '=AND(ISNUMBER($A26),LEFT($B26,17)<>"Securities under ",N($G26)>0,SUM($H26:$Z26)>0)',
    ),
    "the bold rule of column B holds no row of the securities under the threshold",
  );
  check(
    [...chartHeader.slice(1), ...chartRow].every((formula) =>
      formula.includes(
        "wb,CHOOSEROWS('Concentration'!$B$26:$B,SEQUENCE(ROWS(w),1,h+1)),\n" +
          'm,SUMPRODUCT((w=SEQUENCE(ROWS(w)))*(LEFT(wb,17)<>"Securities under ")*1),',
      ),
    ),
    "m of each formula of the chart block counts no row of the securities under the threshold",
  );

  check(
    vm.runInContext("HOLDINGS_CHART_COLUMN", context) === HOLDINGS_AT &&
      vm.runInContext("HOLDINGS_ANCHOR_COLUMN", context) === HOLDINGS_ANCHOR_AT &&
      vm.runInContext("HOLDINGS_CHART_ROWS", context) === HOLDINGS_CHART_ROWS,
    `the holdings chart block starts at column ${HOLDINGS_AT}, one empty column after the anchor cell of the ` +
      `company chart, with ${HOLDINGS_CHART_ROWS} data rows, and the holdings anchor cell is in column ${HOLDINGS_ANCHOR_AT}`,
  );
  const holdingsBlock = exposureState.grid
    .slice(3, 4 + HOLDINGS_CHART_ROWS)
    .map((cells) => cells.slice(HOLDINGS_AT - 1, HOLDINGS_AT - 1 + HOLDINGS_WIDTH));
  const [holdingsHeader, holdingsRow] = holdingsBlock;
  check(
    holdingsBlock.flat().every((value) => typeof value === "string" && value.startsWith("=")),
    `each of the ${holdingsBlock.flat().length} cells of the holdings chart block JE4:JF${4 + HOLDINGS_CHART_ROWS} is a formula`,
  );
  check(
    report
      .cell(SPILL_CELL)
      .includes(
        ` yours,\n "",\n "${HOLDINGS_HEADING}",\n MAKEARRAY(${BAND_ROWS},1,LAMBDA(i,j,"")),\n "A trust or a closed-end fund, such as SPY, GLD, or IBIT, is not a company. Other holdings shows it as one line.",`,
      ),
    `${SPILL_CELL} puts the header row of the holdings chart and a band of ${BAND_ROWS} blank rows between the ` +
      "Total row of the section Holdings and the trust note",
  );
  check(
    exposureState.grid[3][HOLDINGS_ANCHOR_AT - 1] === "holdingsChartRow" &&
      exposureState.grid[4][HOLDINGS_ANCHOR_AT - 1] ===
        `=IFERROR(XMATCH("${HOLDINGS_HEADING}",'Concentration'!$A:$A)+1,"")`,
    "the holdings anchor cell JH5 finds the header row of the holdings chart in column A of the report tab, and adds 1",
  );
  check(
    holdingsBlock.slice(1).every((cells) => JSON.stringify(cells) === JSON.stringify(holdingsRow)) &&
      holdingsRow.every((formula) => formula.includes("r,ROW()-4,") && formula.includes('IF(r>m,"",')),
    `each of the ${HOLDINGS_CHART_ROWS} data rows holds the formulas of row 5, which show the holding group at place ` +
      "ROW()-4 of the table, and a row with no group at its place is blank",
  );
  check(
    holdingsBlock
      .flat()
      .every((formula) =>
        formula === '="Holding"'
          ? true
          : formula.includes(`h,XMATCH(1,ARRAYFORMULA((sa="Holding")*('Concentration'!$B$26:$B="Ticker")))`) &&
            formula.includes(
              `w,CHOOSEROWS('Concentration'!$C$26:$C,SEQUENCE(MIN(${HOLDINGS_CHART_ROWS},ROWS(sa)-h),1,h+1))`,
            ) &&
            formula.includes("m,IFNA(XMATCH(0,ARRAYFORMULA(ISNUMBER(w)*1)),ROWS(w)+1)-1,"),
      ),
    "the holdings chart block finds the header of the table of the section Holdings, and counts the group " +
      "rows under it up to the first row with no number in the Accounts column, so the Total row and a note row " +
      "are no group",
  );
  check(
    holdingsHeader[0] === '="Holding"' &&
      holdingsHeader[1].includes(
        `IF(m=0,"",IF(COUNT(CHOOSEROWS('Concentration'!$E$26:$E,SEQUENCE(m,1,h+1)))=0,"","% of portfolio"))`,
      ),
    "the header row names the label column; the share header is empty when the table holds no group or no share",
  );
  check(
    holdingsRow[0].includes(
      `IF(r>m,"",IF(INDEX('Concentration'!$B$26:$B,h+r)<>"",INDEX('Concentration'!$B$26:$B,h+r),LEFT(INDEX(sa,h+r),24)))`,
    ) && holdingsRow[1].includes(`IF(r>m,"",INDEX('Concentration'!$E$26:$E,h+r))`),
    "the label is the Ticker of the group, or the Holding cut to 24 characters, and the share is its % of portfolio",
  );
  /**
   * True when the parentheses of a formula outside its quoted texts balance,
   * and no ")" comes before its "(".
   */
  const balanced = (formula) => {
    let depth = 0;
    for (const part of formula.split('"').filter((_, i) => i % 2 === 0)) {
      for (const c of part) {
        depth += c === "(" ? 1 : c === ")" ? -1 : 0;
        if (depth < 0) return false;
      }
    }
    return depth === 0;
  };
  const chartFormulaCells = [
    ...chartBlock.flat(),
    ...holdingsBlock.flat(),
    exposureState.grid[4][ANCHOR_AT - 1],
    exposureState.grid[4][HOLDINGS_ANCHOR_AT - 1],
  ];
  check(
    chartFormulaCells.every(balanced),
    "each formula of the two chart blocks and the two anchor cells has balanced parentheses",
  );
  const holdingsReferences = holdingsBlock
    .flat()
    .flatMap((formula) => [...formula.matchAll(/'([^']+)'!\$?([A-Z]+)\$?(\d+)/g)]);
  check(
    holdingsReferences.length > 0 &&
      holdingsReferences.every((m) => m[1] === REPORT_TAB && m[3] === "26") &&
      JSON.stringify([...new Set(holdingsReferences.map((m) => m[2]))].sort()) === JSON.stringify(["A", "B", "C", "E"]),
    "each formula of the holdings chart block reads the Holding, Ticker, Accounts, and % of portfolio columns of " +
      "the report spill from row 26 alone, and no value that the script computes",
  );

  const reportFormulas = formulas.filter((f) => f.at.startsWith(`${REPORT_TAB} `));
  const allLines = reportFormulas.filter((f) => /'Concentration\.Exposure'!\$?B\$?[678]\b/.test(f.value));
  check(
    allLines.length === 0,
    `no formula reads the top 10 weight, the HHI, or the effective count of all the lines (${allLines.map((f) => f.at).join(", ")})`,
  );
  const reportCells = JSON.parse(created[REPORT_TAB]).grid.flat();
  check(
    !reportCells.includes("Sum check") && reportFormulas.every((f) => !/"pass"|"fail"/.test(f.value)),
    "the report holds no Sum check cell and no pass or fail text",
  );
  check(
    reportCells.every((value) => typeof value !== "string" || !value.includes("fund inside funds")) &&
      report.cell(SPILL_CELL).includes('"fund","Funds held by the funds, not looked through"'),
    "the class fund has a label of its own, so no row shows the text fund inside funds",
  );

  const filterRows = [];
  const usedGuards = new Set();
  for (const { at, value } of formulas) {
    for (const call of filterCalls(value)) {
      const treatment = filterTreatment(call, value);
      check(treatment !== null, `${at}: ${call.text} sits in IFNA or IFERROR, or GUARDED_FILTERS names it`);
      if (treatment.startsWith("guard")) usedGuards.add(call.condition);
      filterRows.push([at.replace(/^Concentration /, ""), call.text, treatment]);
    }
  }
  check(filterRows.length > 0, "the report formulas hold FILTER calls");
  check(
    Object.keys(GUARDED_FILTERS).every((condition) => usedGuards.has(condition)),
    "each entry of GUARDED_FILTERS guards a FILTER call",
  );
  check(
    isLogical("ARRAYFORMULA(ISNUMBER('Concentration.Exposure'!$AJ$5:$AJ))") &&
      isLogical('ARRAYFORMULA(LP="stock")') &&
      !isLogical("ARRAYFORMULA(ISNUMBER('Concentration.Exposure'!$AJ$5:$AJ)*1)") &&
      !isLogical('ARRAYFORMULA((LP="stock")*ISNUMBER(LX))'),
    "isLogical finds an array of TRUE and FALSE, and passes an array of 1 and 0",
  );
  const logicalSums = [];
  for (const { at, value } of formulas) {
    const { code } = splitLiterals(value);
    const names = letDefinitions(code);
    for (const m of code.matchAll(/\bSUM\(([A-Za-z_][A-Za-z0-9_]*)\)/g)) {
      if (names.has(m[1]) && isLogical(names.get(m[1]))) logicalSums.push(`${at}: SUM(${m[1]})`);
    }
  }
  check(
    logicalSums.length === 0,
    "no SUM reads a LET name that holds TRUE and FALSE, so the counts nown and ng of the section Other holdings " +
      `are not 0 (${logicalSums.join(", ")})`,
  );
  console.log("  pass");
  console.log("\nFILTER calls and their treatment:");
  table(["cell", "call", "treatment"], filterRows);
  console.log("\nGuards of GUARDED_FILTERS:");
  for (const [condition, { reason }] of Object.entries(GUARDED_FILTERS)) console.log(`  ${condition}: ${reason}`);

  const call = state.fetchCalls[0];
  check(call.url === ROUTE_URL, "the request goes to the concentration route");
  check(call.options.method === "post", "the method is post");
  check(call.options.contentType === "application/json", "the content type is application/json");
  check(call.options.muteHttpExceptions === true, "muteHttpExceptions is true");
  check(
    Object.keys(call.options.headers).join() === "Authorization" &&
      call.options.headers.Authorization === `Bearer ${API_KEY}`,
    "the Authorization header holds the key of the user property alone",
  );
  const body = JSON.parse(call.options.payload);
  check(Object.keys(body).join() === "positions", "the body holds only positions");
  const sent = body.positions;
  check(sent.length === EXPECTED.length, `the body holds ${EXPECTED.length} positions`);
  sent.forEach((p, i) => {
    const want = EXPECTED[i];
    check(
      Object.keys(p).every((k) => ["id", "ticker", "weight"].includes(k)),
      `position ${i} holds id, ticker, and weight alone`,
    );
    check(p.id === want.id, `position ${i} has the id of the hand calculation`);
    check(
      want.ticker === null ? !("ticker" in p) : p.ticker === want.ticker,
      `position ${i} has the ticker of the hand calculation`,
    );
    check(near(p.weight, want.sum / EXPECTED_TOTAL), `position ${i} has the weight of the hand calculation`);
  });
  check(!sent.some((p) => p.id === "" || p.id === "XYZ"), "the two skipped rows send nothing");
  check(!("ticker" in sent.find((p) => p.id === PLAN_FUND)), "the plan fund sends no ticker");
  check(
    numbersOf(body).every((n) => n > 0 && n <= 1),
    "each number of the body is a weight, and no dollar value appears",
  );
  check(
    !/Test (brokerage|IRA|plan)/.test(call.options.payload) &&
      !/\b(3000|2000|1000|500|10000)\b/.test(call.options.payload),
    "no account name and no value of the Holdings fake appears in the body",
  );

  console.log("\nPositions sent, with the hand calculation (total of the kept rows: 10,000):");
  table(
    ["id", "ticker", "sum of the key", "sum / 10,000", "weight sent"],
    sent.map((p, i) => [p.id, p.ticker ?? "(none)", EXPECTED[i].sum, EXPECTED[i].sum / EXPECTED_TOTAL, p.weight]),
  );

  const x = book.tab(EXPOSURE_TAB);
  check(x.cell("B1") === "OK", `B1 is OK (B1 holds "${x.cell("B1")}")`);
  check(isDate(x.cell("B2")), "B2 is a Date");
  const firstTime = x.cell("B2").getTime();

  const answer = JSON.parse(liveText);
  const measures = x.block(5, 2, MEASURES.length, 1).map((cells) => cells[0]);
  MEASURES.forEach((name, i) =>
    check(measures[i] === (answer.measures[name] ?? ""), `B${5 + i} holds ${name}, or is empty with no ${name}`),
  );

  /**
   * The values that B28:B32 must hold for an equity block. A null block
   * gives five empty cells.
   */
  const equityWant = (equity) => EQUITY.map((name) => equity?.[name] ?? "");
  const equityCells = (sheet) => sheet.block(EQUITY_ROW, 2, EQUITY.length, 1).map((cells) => cells[0]);
  check(
    JSON.stringify(equityCells(x)) === JSON.stringify(equityWant(answer.measures.equity)),
    `B${EQUITY_ROW}:B${EQUITY_ROW + EQUITY.length - 1} holds the equity measures, in the order ${EQUITY.join(", ")}`,
  );

  const funds = topFundsOf(answer.funds);
  const fundCells = x.block(5, FUND_AT, funds.length, FUND_FIELDS.length);
  funds.forEach((fund, i) => {
    FUND_FIELDS.forEach((name, c) => {
      check(fundCells[i][c] === (fund[name] ?? ""), `D${5 + i}:N holds ${name} of fund ${i}`);
    });
  });
  check(
    blockRows(x, FUND_AT, 1).length === funds.length,
    `D5:N holds the ${funds.length} funds that entered, and none of the ` +
      `${answer.funds.length - funds.length} held funds and funds that did not enter`,
  );
  console.log(`  The funds block of the answer holds ${answer.funds.length} elements, and D5:N holds ${funds.length}.`);
  check(
    x.block(5, MIX_AT, x.getMaxRows() - 4, MIX_FIELDS.length).every((c) => c.every((v) => v === "")),
    "the mix block U5:Z is empty when no holding has a mix",
  );

  const pairs = answer.overlaps;
  /**
   * The count of cells of the overlaps block that differ from the answer.
   */
  const pairFaults = (sheet) => {
    const now = sheet.block(5, OVERLAP_AT, pairs.length, 4);
    return pairs.filter((pair, i) => {
      const want = [pair.ids[0], pair.ids[1], pair.overlap, pair.sharedLineCount];
      return want.some((value, c) => now[i][c] !== value);
    }).length;
  };
  check(
    pairFaults(x) === 0,
    `P5:S holds the two ids, overlap, and sharedLineCount of each of the ${pairs.length} pairs`,
  );
  check(
    x.block(5 + pairs.length, OVERLAP_AT, x.getMaxRows() - 4 - pairs.length, 4).every((c) => c.every((v) => v === "")),
    "the rows under the last pair are empty",
  );

  const ids = sent.map((p) => p.id);
  const idCells = x.block(4, SOURCE_AT, 1, MAX_POSITIONS)[0];
  check(
    ids.every((id, i) => idCells[i] === id),
    "AJ4 onward holds the ids in the body order",
  );
  check(
    idCells.slice(ids.length).every((v) => v === ""),
    "no id follows the last position",
  );

  const lines = answer.lines;
  const parts = expectedParts(lines, ids);
  const partWidth = LINE_FIELDS.length + 1 + ids.length;
  /**
   * The count of cells of the lines block and of the sources block that
   * differ from the part rows of the answer.
   */
  const partFaults = (sheet) => {
    let faults = 0;
    const now = partRowsOf(sheet, parts.length, ids.length);
    parts.forEach((cells, r) => {
      cells.forEach((value, c) => {
        if (now[r]?.[c] !== value) faults += 1;
      });
    });
    return faults;
  };
  check(
    partFaults(x) === 0,
    `AR5:AY and the sources from BA5 hold the ${parts.length} part rows of the ${lines.length} lines, in the answer order`,
  );
  check(
    x
      .block(5 + parts.length, LINE_AT, x.getMaxRows() - 4 - parts.length, partWidth + 1)
      .every((c) => c.every((v) => v === "")),
    "the rows under the last part row are empty",
  );
  const partCells = partRowsOf(x, parts.length, ids.length);
  const stockIndex = LINE_FIELDS.indexOf("stockWeight");
  const weightIndex = LINE_FIELDS.indexOf("weight");
  const partIndex = LINE_FIELDS.length;
  const stockRows = partCells.filter((cells) => cells[partIndex] === "stock");
  const otherRows = partCells.filter((cells) => cells[partIndex] === "other");
  check(stockRows.length + otherRows.length === partCells.length, "column AY of each row holds stock or other");
  check(
    stockRows.every((cells) => Math.abs(sumCells(cells.slice(partIndex + 1)) - cells[stockIndex]) < 1e-9),
    "the source cells of each stock row add up to the stock weight of its line",
  );
  check(
    otherRows.every(
      (cells) => Math.abs(sumCells(cells.slice(partIndex + 1)) - (cells[weightIndex] - cells[stockIndex])) < 1e-9,
    ),
    "the source cells of each other row add up to the weight minus the stock weight of its line",
  );
  const stockSum = sumCells(stockRows.map((cells) => cells[stockIndex]));
  const otherSum = otherRows.reduce((sum, cells) => sum + cells[weightIndex] - cells[stockIndex], 0);
  check(
    Math.abs(stockSum + otherSum - answer.measures.weightSum) < 1e-9,
    `the stock rows and the other rows add up to weightSum (${stockSum + otherSum} for ${answer.measures.weightSum})`,
  );
  if (answer.measures.equity !== null) {
    check(
      Math.abs(stockSum - answer.measures.equity.weight) < 1e-9,
      `the stock weights of the stock rows add up to equity.weight (${stockSum} for ${answer.measures.equity.weight})`,
    );
  }
  const lostLabels = [];
  accepted[EXPOSURE_TAB].grid.forEach((row, r) => {
    row.forEach((value, c) => {
      if (value !== "" && x.grid[r][c] !== value) lostLabels.push(`row ${r + 1} column ${c + 1}`);
    });
  });
  check(
    lostLabels.length === 0,
    `each label and formula of Concentration.Exposure stays after the write (${lostLabels.slice(0, 3).join(", ")})`,
  );

  const afterFetch = state.log.slice(fetchAt + 1);
  check(
    afterFetch.every(
      (e) => e.sheet === EXPOSURE_TAB || e.op === "flush" || (e.sheet === REPORT_TAB && e.op === "insertChart"),
    ),
    "after the request, run 1 changes the Concentration.Exposure tab and inserts the charts of the report tab alone",
  );
  const chartInserts = afterFetch.map((e, n) => (e.op === "insertChart" ? n : -1)).filter((n) => n >= 0);
  const insertAt = chartInserts[0];
  check(
    JSON.stringify(chartInserts.map((n) => afterFetch[n].block)) === JSON.stringify(["company", "holdings"]) &&
      afterFetch.filter((e) => e.op === "removeChart").length === 0,
    "run 1 inserts the company chart, then the holdings chart, and removes none, because the new report tab " +
      "holds no chart",
  );
  check(
    [...state.tabsAsked].every((name) => ["Holdings", EXPOSURE_TAB, REPORT_TAB].includes(name)),
    "the script asks for the Holdings tab and the two report tabs alone",
  );
  check(
    writesSince(fetchAt) === 4,
    "after the request, run 1 makes 2 setValues calls for the data, 1 for the status, and 1 for the run time " +
      `(${writesSince(fetchAt)})`,
  );
  const writeRanges = afterFetch
    .filter((e) => e.op === "setValues")
    .map((e) => `${e.row},${e.column},${e.rows},${e.columns}`);
  const flushes = afterFetch.map((e, n) => (e.op === "flush" ? n : -1)).filter((n) => n >= 0);
  const setAt = afterFetch.map((e, n) => (e.op === "setValues" ? n : -1)).filter((n) => n >= 0);
  check(
    writeRanges[2] === "1,2,2,1" && writeRanges[3] === `${RUN_ROW},1,10,2`,
    `the status write B1:B2 comes before the run-time write A${RUN_ROW}:B${RUN_ROW + 9} ` +
      `(${writeRanges.slice(2).join("; ")})`,
  );
  check(
    flushes.length === 3 &&
      flushes[0] > setAt[2] &&
      flushes[0] < chartInserts[0] &&
      flushes[1] > chartInserts[0] &&
      flushes[1] < chartInserts[1],
    "run 1 calls SpreadsheetApp.flush after the status write, before each draw reads its chart block",
  );
  check(
    insertAt > setAt[2] && flushes[2] > chartInserts[1] && flushes[2] < setAt[3],
    "run 1 draws the two charts, then flushes again before the run-time write, so the run time includes the charts",
  );
  const runOneComputed = checkComputed(x, "run 1", parts.length);
  checkDataWrites(afterFetch, "run 1", topLastOf(x, runOneComputed), parts.length, ids.length, x);
  const textRanges = afterFetch
    .filter((e) => e.op === "setNumberFormat" && e.value === "@")
    .map((e) => `${e.row},${e.column},${e.columns}`);
  check(
    [
      `5,${FUND_AT},2`,
      `5,${FUND_AT + 3},1`,
      `5,${OVERLAP_AT},2`,
      `5,${MIX_AT},1`,
      `5,${MIX_AT + 3},1`,
      `5,${UNSEEN_AT},1`,
      `5,${STOCK_FUND_AT},1`,
      `5,${OWN_AT},3`,
      `5,${OWN_AT + 4},1`,
      `5,${GROUP_AT},1`,
      `5,${GROUP_AT + 3},1`,
      `5,${LINE_AT},4`,
      `4,${SOURCE_AT},${MAX_POSITIONS}`,
    ].every((r) => textRanges.includes(r)),
    "the script sets the text format on D:E, G, P:Q, U, X, AB, AE, AG:AI, AK, AM, AP, AR:AU, and BA4:IR4",
  );
  check(
    afterFetch.every((e) => e.op !== "setNumberFormat" || e.column + e.columns - 1 < CHART_AT),
    "no number format of a refresh reaches the chart block",
  );
  const firstFormat = afterFetch.findIndex((e) => e.op === "setNumberFormat");
  const firstWrite = afterFetch.findIndex((e) => e.op === "setValues");
  check(firstFormat >= 0 && firstFormat < firstWrite, "the text format comes before the data write");

  const stock = lines.find((l) => l.sources && STOCK in l.sources);
  check(stock !== undefined, "a line holds a source from the direct stock");
  check(
    [STOCK, FUND_A, FUND_B].every((id) => typeof stock.sources[id] === "number" && stock.sources[id] > 0),
    "the direct stock line has a source from the stock and from each of the two index funds",
  );
  const plan = lines.find((l) => l.sources && PLAN_FUND in l.sources);
  check(plan?.class === "unknown", `the plan fund line has the class unknown (${plan?.class})`);
  const money = lines.find((l) => l.sources && MONEY in l.sources);
  check(
    money?.class === "fund" || money?.class === "cash",
    `the money market fund line has the class fund or the class cash (${money?.class})`,
  );

  console.log("\n== Run time and chart of run 1");
  checkRecorded("run 1", runOne);
  const oneChart = checkChart("run 1");
  check(
    oneChart.direct && oneChart.funds === 2 && !oneChart.other && oneChart.companies > 0,
    `run 1: the companies at or above 1% come from the direct stock and from the two index funds (${oneChart.companies} companies)`,
  );
  check(
    oneChart.series === 1 + oneChart.funds + Number(oneChart.other),
    `run 1: the chart has 1 + ${oneChart.funds} + ${Number(oneChart.other)} series: Direct, each fund, and Other funds ` +
      `(${oneChart.series})`,
  );
  const drawn = companyCharts(report)[0];
  console.log(
    `  chart at A${oneChart.anchor}; ranges ${drawn.ranges.map((r) => `${r.row},${r.column},${r.rows},${r.columns}`).join("; ")}`,
  );
  console.log(`  colors ${drawn.colors.join(", ")}`);
  check(
    x.cell(`A${RUN_HEADER}`) === "runStart" && x.cell(`B${RUN_HEADER}`) === "seconds",
    "the header of the run-time block stays",
  );
  console.log(`  A${RUN_ROW}: ${new Date(runRows()[0][0]).toISOString()}; B${RUN_ROW}: ${runRows()[0][1]} seconds`);

  const partHeader = [...LINE_FIELDS, "part", ...ids.map((id) => id.slice(0, 12))];
  console.log("\nGrid of Concentration.Exposure after run 1:", `${x.getMaxRows()} rows, ${x.getMaxColumns()} columns.`);
  console.log(`The answer holds ${lines.length} lines, and the tab holds ${parts.length} part rows.`);
  console.log("\nFunds block, D5:N:");
  table(FUND_FIELDS, fundCells);
  console.log("\nOverlaps block, P5:S:");
  table(["firstId", "secondId", "overlap", "sharedLineCount"], x.block(5, OVERLAP_AT, pairs.length, 4));
  console.log("\nFirst 10 part rows, AR5:AY and the sources from BA5:");
  table(partHeader, partCells.slice(0, 10));
  console.log("\nPart rows of the lines of the direct positions:");
  const directKeys = new Set(
    lines.filter((l) => [STOCK, BOND, MONEY, PLAN_FUND].some((id) => l.sources && id in l.sources)).map((l) => l.key),
  );
  table(
    partHeader,
    partCells.filter((cells) => directKeys.has(cells[0])),
  );
  console.log(`\nMeasures, B5:B${4 + MEASURES.length}:`);
  table(
    ["measure", "value"],
    MEASURES.map((name, i) => [name, measures[i]]),
  );
  console.log(`\nEquity measures, B${EQUITY_ROW}:B${EQUITY_ROW + EQUITY.length - 1}:`);
  table(
    ["equity", "value"],
    EQUITY.map((name, i) => [name, equityCells(x)[i]]),
  );

  console.log("\n== Edit of the threshold cell: the simple trigger onEdit");
  /**
   * A fake edit event of one range, in AuthMode.LIMITED, with the value of
   * the edit when the call gives one.
   */
  const edit = (sheet, a1, value) => ({
    authMode: "LIMITED",
    source: book,
    range: sheet.getRange(a1),
    ...(value === undefined ? {} : { value }),
  });
  /**
   * Call onEdit with an event, check that it throws no error, sends no
   * request, takes no lock, opens no dialog, and writes no cell, and give the
   * operations of the log of the call.
   */
  const runEdit = (name, event) => {
    const from = state.log.length;
    const counts = () => [
      state.fetchCalls.length,
      state.lockTaken,
      state.prompts.length,
      state.alerts.length,
      state.sidebars.length,
    ];
    const before = counts();
    let thrown = null;
    try {
      context.onEdit(event);
    } catch (e) {
      thrown = e;
    }
    check(thrown === null, `${name}: onEdit throws no error (${thrown?.message ?? "no error"})`);
    check(
      JSON.stringify(counts()) === JSON.stringify(before),
      `${name}: onEdit sends no request, takes no lock, and opens no dialog`,
    );
    const ops = state.log.slice(from).map((e) => e.op);
    check(
      ops.every((op) => ["flush", "removeChart", "insertChart"].includes(op)),
      `${name}: onEdit writes no cell (${ops.join(", ") || "no change"})`,
    );
    return ops;
  };
  /**
   * Make the anchor cell and the chart header row above it give their
   * present values for a count of reads, as a slow spill does right after
   * an edit. Call it before the edit.
   */
  const lag = (reads) => {
    const anchor = x.shown(5, ANCHOR_AT);
    state.lagging.set(`${EXPOSURE_TAB}!${cellKey(5, ANCHOR_AT)}`, { value: anchor, reads });
    state.lagging.set(`${REPORT_TAB}!${cellKey(anchor - 1, 1)}`, { value: report.shown(anchor - 1, 1), reads });
    return anchor;
  };
  /**
   * The calls of console.warn or of console.error since an index of
   * state.logged.
   */
  const loggedSince = (from, level) => state.logged.slice(from).filter((entry) => entry.level === level);

  const refreshChart = companyCharts(report)[0];
  const refreshHoldings = holdingsCharts(report)[0];
  const editLogFrom = state.log.length;
  const bareFrom = state.logged.length;
  const bare = runEdit("a run from the script editor with no event object", undefined);
  check(
    bare.length === 0 && companyCharts(report)[0] === refreshChart && state.logged.length === bareFrom,
    "onEdit() with no event object returns at once, draws nothing, and logs nothing",
  );
  const skipFrom = state.logged.length;
  for (const [name, event] of [
    ["an edit of the overlap minimum B24", edit(report, MINIMUM_CELL, "20%")],
    ["an edit of C23 in the threshold row", edit(report, "C23", "5%")],
    ["an edit of the label A23", edit(report, "A23", "5%")],
    ["an edit of B22 in column B", edit(report, "B22", "5%")],
    ["a paste into B23:B24", edit(report, "B23:B24")],
    ["an edit of B30 in the spill", edit(report, "B30", "5%")],
    ["an edit of B2 on the Holdings tab", edit(book.tab("Holdings"), "B2", "5%")],
    ["an edit of B23 on the hidden tab", edit(x, THRESHOLD_CELL, "5%")],
    ["an edit event with no range", {}],
    ["an edit of the threshold cell with no value", edit(report, THRESHOLD_CELL)],
    ["an edit of the threshold to 150%", edit(report, THRESHOLD_CELL, "150%")],
    ["an edit of the threshold to -1%", edit(report, THRESHOLD_CELL, "-1%")],
    ["an edit of the threshold to a text", edit(report, THRESHOLD_CELL, "ten")],
  ]) {
    const ops = runEdit(name, event);
    check(
      ops.length === 0 && companyCharts(report).length === 1 && companyCharts(report)[0] === refreshChart,
      `${name}: onEdit draws nothing`,
    );
  }
  check(
    state.logged.length === skipFrom,
    "each edit that draws nothing, and the edit event with no range, logs nothing",
  );

  const atOne = checkChart("threshold 1% after the refresh");
  report.grid[22][1] = 0.05;
  const raised = runEdit("an edit of the threshold to 5%", edit(report, THRESHOLD_CELL, "5%"));
  check(
    JSON.stringify(raised) === JSON.stringify(["flush", "removeChart", "insertChart"]),
    `an edit of the threshold cell flushes, removes the old chart, and inserts one chart (${raised.join(", ")})`,
  );
  const atFive = checkChart("threshold 5%");
  check(
    atFive.listed < atOne.listed && atFive.anchor < atOne.anchor,
    `threshold 5%: ${atFive.listed} companies are at or above 5% and ${atOne.listed} at or above 1%, so the ` +
      `chart moves up from row ${atOne.anchor} to row ${atFive.anchor}`,
  );

  report.grid[22][1] = 1;
  const none = runEdit("an edit of the threshold to 100%", edit(report, THRESHOLD_CELL, "1"));
  check(
    JSON.stringify(none) === JSON.stringify(["flush", "removeChart"]) && companyCharts(report).length === 0,
    "threshold 100%: no company is at or above it, so no header after Company holds a text, and onEdit removes " +
      `the chart and inserts none (${none.join(", ")})`,
  );

  report.grid[22][1] = 0.01;
  const back = runEdit("an edit of the threshold back to 1%", edit(report, THRESHOLD_CELL, "0.01"));
  check(
    JSON.stringify(back) === JSON.stringify(["flush", "insertChart"]),
    `an edit of the threshold back to 1% inserts the chart again (${back.join(", ")})`,
  );
  const again = checkChart("threshold 1% again");
  check(
    again.anchor === atOne.anchor && again.series === atOne.series,
    "threshold 1% again: the chart is back at the row and with the series of the refresh",
  );
  check(
    state.sleeps.length === 0 && state.logged.filter((entry) => entry.level === "warn").length === 0,
    `each draw so far reads the fresh chart header row at the first read and waits no time (${state.sleeps.length} waits)`,
  );

  state.frozenValues = true;
  ["Company", "", "Fund 1", "Fund 2", "", "", "", ""].forEach((text, i) =>
    x.calculated.set(cellKey(4, CHART_AT + i), text),
  );
  x.calculated.set(cellKey(5, ANCHOR_AT), 40);
  report.calculated.clear();
  report.calculated.set(cellKey(39, 1), spillHeading(0.01));
  context.drawChart(book, 0.01);
  const planted = companyCharts(report)[0];
  check(
    companyCharts(report).length === 1 &&
      JSON.stringify(planted.ranges.map((range) => range.column - CHART_AT)) === "[0,2,3]" &&
      JSON.stringify(planted.colors) === JSON.stringify(FUND_COLORS.slice(0, 2)) &&
      planted.position.row === 40,
    "with an empty Direct header and two fund headers, the chart reads the company column and the two fund " +
      "columns, with the first two fund colors, at the row of the anchor cell",
  );
  report.calculated.set(cellKey(39, 1), spillHeading(0.01).replace(".", ","));
  context.drawChart(book, 0.01);
  check(
    companyCharts(report).length === 1 &&
      companyCharts(report)[0] !== planted &&
      companyCharts(report)[0].position.row === 40 &&
      state.sleeps.length === 0,
    "a chart header row with a comma decimal mark, as a comma locale shows it, matches at the first read, and " +
      "the chart anchors at the row of the anchor cell",
  );
  const commaChart = companyCharts(report)[0];
  x.calculated.set(cellKey(5, ANCHOR_AT), "");
  let warnFrom = state.logged.length;
  context.drawChart(book, 0.01);
  check(
    companyCharts(report).length === 1 &&
      companyCharts(report)[0] === commaChart &&
      loggedSince(warnFrom, "warn").length === 1,
    "an anchor cell with no row number for the whole wait leaves the chart as it is and logs one warning",
  );
  state.sleeps.length = 0;
  state.frozenValues = false;
  runEdit("an edit of the threshold cell after the planted values", edit(report, THRESHOLD_CELL, "0.01"));
  checkChart("the draw after the planted values");

  const lastChart = companyCharts(report)[0];
  const failure = new Error("The chart service failed.");
  report.newChart = () => {
    throw failure;
  };
  const errorFrom = state.logged.length;
  const failed = runEdit(
    "an edit of the threshold cell with a failing chart builder",
    edit(report, THRESHOLD_CELL, "0.01"),
  );
  delete report.newChart;
  check(
    JSON.stringify(failed) === JSON.stringify(["flush"]) &&
      companyCharts(report).length === 1 &&
      companyCharts(report)[0] === lastChart,
    `onEdit catches the error of the draw and returns, and the chart of the last draw stays (${failed.join(", ")})`,
  );
  const errors = loggedSince(errorFrom, "error");
  check(
    errors.length === 1 &&
      [failure.message, failure.stack].every((part) => errors[0].args.some((arg) => String(arg).includes(part))),
    `onEdit gives the message and the stack of the error to console.error (${errors.length} calls)`,
  );

  const staleFrom = lag(2);
  report.grid[22][1] = 0.05;
  const lagged = runEdit(
    "an edit of the threshold to 5% while the spill lags two reads",
    edit(report, THRESHOLD_CELL, "5%"),
  );
  state.lagging.clear();
  const fresh = checkChart("threshold 5% after two stale reads");
  check(
    JSON.stringify(lagged) === JSON.stringify(["flush", "removeChart", "insertChart"]) &&
      fresh.anchor === atFive.anchor &&
      fresh.anchor !== staleFrom &&
      companyCharts(report)[0].position.row === fresh.anchor,
    `the draw waits for the fresh anchor row ${fresh.anchor} in place of the stale row ${staleFrom} ` +
      `(${lagged.join(", ")})`,
  );
  check(
    JSON.stringify(state.sleeps) === "[500,500]",
    `the draw waits 500 ms after each of the two stale reads (${JSON.stringify(state.sleeps)})`,
  );

  const keptChart = companyCharts(report)[0];
  state.sleeps.length = 0;
  warnFrom = state.logged.length;
  lag(Infinity);
  report.grid[22][1] = 0.01;
  const stale = runEdit(
    "an edit of the threshold to 1% while the spill stays stale",
    edit(report, THRESHOLD_CELL, "0.01"),
  );
  state.lagging.clear();
  check(
    JSON.stringify(stale) === JSON.stringify(["flush"]) &&
      companyCharts(report).length === 1 &&
      companyCharts(report)[0] === keptChart,
    `a chart header row that stays stale for the whole wait leaves the old chart in place (${stale.join(", ")})`,
  );
  check(
    state.sleeps.every((ms) => ms === 500) &&
      state.sleeps.reduce((sum, ms) => sum + ms, 0) === 20000 &&
      loggedSince(warnFrom, "warn").length === 1,
    `the draw waits 500 ms between reads, 20 s in total, then logs one warning (${state.sleeps.length} waits)`,
  );
  state.sleeps.length = 0;
  runEdit("an edit of the threshold to 1% after the stale spill", edit(report, THRESHOLD_CELL, "0.01"));
  checkChart("threshold 1% after the stale spill");
  check(
    refreshHoldings !== undefined &&
      holdingsCharts(report).length === 1 &&
      holdingsCharts(report)[0] === refreshHoldings &&
      state.log
        .slice(editLogFrom)
        .every((e) => (e.op !== "insertChart" && e.op !== "removeChart") || e.block === "company"),
    "each edit and each direct draw of the company chart removes and inserts company charts alone, so the " +
      "holdings chart of the refresh stays",
  );

  console.log("\n== Holdings chart: a draw of its block alone");
  const companyKept = companyCharts(report)[0];
  const holdingsFrom = state.log.length;
  const holdingsAnchor = x.shown(5, HOLDINGS_ANCHOR_AT);
  const holdingsKept = holdingsCharts(report)[0];
  state.frozenValues = true;
  report.calculated.delete(cellKey(holdingsAnchor - 1, 1));
  state.sleeps.length = 0;
  warnFrom = state.logged.length;
  context.drawHoldingsChart(book);
  check(
    holdingsKept !== undefined &&
      holdingsCharts(report).length === 1 &&
      holdingsCharts(report)[0] === holdingsKept &&
      state.sleeps.every((ms) => ms === 500) &&
      state.sleeps.reduce((sum, ms) => sum + ms, 0) === 20000 &&
      loggedSince(warnFrom, "warn").length === 1,
    "when the row above the holdings anchor row never shows the header row of the holdings chart, the holdings " +
      "draw waits 500 ms between reads, 20 s in total, leaves the holdings chart in place, and logs one warning",
  );
  report.calculated.set(cellKey(holdingsAnchor - 1, 1), HOLDINGS_HEADING);
  x.calculated.set(cellKey(4, HOLDINGS_AT + 1), "");
  context.drawHoldingsChart(book);
  check(
    holdingsCharts(report).length === 0 &&
      companyCharts(report).length === 1 &&
      companyCharts(report)[0] === companyKept,
    "with an empty share header, as when the section Holdings shows a note, the holdings draw removes the " +
      "holdings chart, inserts none, raises no error, and keeps the company chart",
  );
  state.frozenValues = false;
  state.sleeps.length = 0;
  state.lagging.set(`${EXPOSURE_TAB}!${cellKey(5, HOLDINGS_ANCHOR_AT)}`, { value: holdingsAnchor + 3, reads: 2 });
  context.drawHoldingsChart(book);
  state.lagging.clear();
  const holdingsDrawn = checkHoldingsChart("the holdings draw after two stale reads of its anchor cell");
  check(
    holdingsDrawn.chart !== null &&
      holdingsDrawn.chart.position.row === holdingsAnchor &&
      JSON.stringify(state.sleeps) === "[500,500]",
    `the holdings draw waits 500 ms after each of the two stale reads, then anchors at the fresh row ${holdingsAnchor}`,
  );
  check(
    companyCharts(report).length === 1 &&
      companyCharts(report)[0] === companyKept &&
      state.log
        .slice(holdingsFrom)
        .every((e) => (e.op !== "insertChart" && e.op !== "removeChart") || e.block === "holdings"),
    "each holdings draw removes and inserts holdings charts alone, so the company chart stays",
  );
  state.sleeps.length = 0;

  const reportState = report.state();

  const bigLines = x.getMaxRows() - 4 + 200;
  const bigRows = bigLines + 4;
  console.log(`\n== Run 2: status 200 with 40 positions, ${bigLines} lines, no pair, and no equity block`);
  const forty = Array.from({ length: 40 }, (_, i) => [`T${String(i).padStart(3, "0")}`, "Example brokerage", 1, "A"]);
  replaceHoldings(forty);
  const fortyIds = forty.map((row) => row[0]);
  const big = bigAnswer(fortyIds, bigLines);
  const at = (row, column) => x.block(row, column, 1, 1)[0][0];
  let logFrom = state.log.length;
  let fetches = state.fetchCalls.length;
  state.fetchHandler = () => fakeResponse(200, JSON.stringify(big));
  pause(2);
  const runTwo = timedRun();
  check(state.fetchCalls.length === fetches + 1, "run 2 calls the fake once");
  checkRecorded("run 2", runTwo);
  checkNoTabChange(reportState, logFrom, "run 2");
  check(x.cell("B1") === "OK", `B1 is OK (B1 holds "${x.cell("B1")}")`);
  const exposureColumns = accepted[EXPOSURE_TAB].grid[0].length;
  check(
    x.getMaxRows() === bigRows && x.getMaxColumns() === exposureColumns,
    `the grid grew to ${bigRows} rows and keeps ${exposureColumns} columns (${x.getMaxRows()} by ${x.getMaxColumns()})`,
  );
  check(
    at(5, LINE_AT + 1) === "Line 0" && at(bigRows, LINE_AT + 1) === `Line ${bigLines - 1}`,
    `column U holds the ${bigLines} lines`,
  );
  check(
    at(5, PART_AT) === "other" && at(bigRows, PART_AT) === "other",
    "a line with no stock gives one row of the other part",
  );
  const lastId = fortyIds[(bigLines - 1) % 40];
  check(
    at(4, SOURCE_AT + fortyIds.indexOf(lastId)) === lastId &&
      near(at(bigRows, SOURCE_AT + fortyIds.indexOf(lastId)), 1 / bigLines),
    `the source column of ${lastId} holds the weight of the last line`,
  );
  check(at(4, SOURCE_AT + 39) === "T039", "the column of position 40 holds the id T039");
  check(
    equityCells(x).every((value) => value === ""),
    `B${EQUITY_ROW}:B${EQUITY_ROW + EQUITY.length - 1} is empty when the answer holds no equity block`,
  );
  check(
    x.cell(`A${EQUITY_HEADER}`) === "equity" && x.cell(`B${EQUITY_HEADER}`) === "value",
    "the header of the equity block stays",
  );
  check(
    x.block(5, OVERLAP_AT, x.getMaxRows() - 4, 4).every((c) => c.every((v) => v === "")),
    "P5:S is empty when the answer holds no pair",
  );
  const runTwoAfter = state.log.slice(indexSince(logFrom, (e) => e.op === "fetch") + 1);
  const runTwoComputed = checkComputed(x, "run 2", bigLines);
  checkDataWrites(runTwoAfter, "run 2", topLastOf(x, runTwoComputed), bigLines, fortyIds.length, x);
  console.log(`  B1: ${x.cell("B1")}; grid ${x.getMaxRows()} rows, ${x.getMaxColumns()} columns`);

  const shorterLines = bigLines - 100;
  const thirty = fortyIds.slice(0, 30);
  console.log(`\n== Run 2b: status 200 with 30 positions and ${shorterLines} lines, so the answer is shorter`);
  replaceHoldings(forty.slice(0, 30));
  logFrom = state.log.length;
  state.fetchHandler = () => fakeResponse(200, JSON.stringify(bigAnswer(thirty, shorterLines)));
  pause(2);
  const runTwoB = timedRun();
  checkRecorded("run 2b", runTwoB);
  checkNoTabChange(reportState, logFrom, "run 2b");
  check(
    x.getMaxRows() === 4 + shorterLines && x.getMaxColumns() === exposureColumns,
    `the grid shrinks to the ${4 + shorterLines} rows of the answer and keeps ${exposureColumns} columns ` +
      `(${x.getMaxRows()} by ${x.getMaxColumns()})`,
  );
  const runTwoBEntries = state.log.slice(logFrom);
  check(
    runTwoBEntries.some((e) => e.op === "deleteRows" && e.count === bigLines - shorterLines),
    `run 2b deletes the ${bigLines - shorterLines} rows after the answer at the end of the grid`,
  );
  const clears = runTwoBEntries.filter((e) => e.op === "clearContent");
  check(
    clears.length > 0 && clears.every((e) => e.column >= 2),
    "run 2b clears the cells of run 2 that the new answer does not write, and no cell of column A",
  );
  check(
    clears.some((e) => e.column === SOURCE_AT + 30 && e.row === 4),
    "run 2b clears the source columns of positions 31 to 40 from row 4",
  );
  check(
    x.block(4, SOURCE_AT + 30, x.getMaxRows() - 3, MAX_POSITIONS - 30).every((c) => c.every((v) => v === "")),
    "no id and no cell of positions 31 to 40 of run 2 stays",
  );
  const runTwoBComputed = checkComputed(x, "run 2b", shorterLines);
  checkDataWrites(
    state.log.slice(indexSince(logFrom, (e) => e.op === "fetch") + 1),
    "run 2b",
    topLastOf(x, runTwoBComputed),
    shorterLines,
    thirty.length,
    x,
  );

  console.log("\n== Run 3: status 200 with the answer of run 1 again, and tabs of the current layout version");
  replaceHoldings(HOLDINGS_ROWS);
  logFrom = state.log.length;
  state.fetchHandler = () => fakeResponse(200, liveText);
  pause(2);
  const runThree = timedRun();
  checkNoTabChange(reportState, logFrom, "run 3");
  check(x.cell(VERSION_CELL) === layoutVersion, `${VERSION_CELL} holds the current layout version, so both tabs stay`);
  checkRecorded("run 3", runThree);
  check(
    x.getMaxRows() === TAB_ROWS && x.getMaxColumns() === exposureColumns,
    `the grid shrinks to ${TAB_ROWS} rows, the rows of a new tab, and keeps ${exposureColumns} columns`,
  );
  check(partFaults(x) === 0, "the part rows of run 1 are back");
  check(pairFaults(x) === 0, "the pairs of run 1 are back");
  check(
    JSON.stringify(equityCells(x)) === JSON.stringify(equityWant(answer.measures.equity)),
    "the equity measures of run 1 are back",
  );
  check(
    x
      .block(5 + funds.length, FUND_AT, x.getMaxRows() - 4 - funds.length, FUND_FIELDS.length)
      .every((c) => c.every((v) => v === "")),
    "the rows under the last fund are empty",
  );
  check(
    x
      .block(5 + parts.length, LINE_AT, x.getMaxRows() - 4 - parts.length, SOURCE_AT + MAX_POSITIONS - LINE_AT)
      .every((c) => c.every((v) => v === "")),
    "the rows under the last part row are empty, so no line of run 2 stays",
  );
  check(
    x.block(4, SOURCE_AT + ids.length, 1, MAX_POSITIONS - ids.length)[0].every((v) => v === ""),
    "no id of run 2 follows the last position",
  );
  console.log(`  B1: ${x.cell("B1")}`);

  console.log("\n== Run 4: status 429 with an error body of the route");
  let before = x.snapshot();
  const thirdTime = x.cell("B2").getTime();
  check(thirdTime >= firstTime, "B2 holds the time of run 3");
  pause(5);
  logFrom = state.log.length;
  fetches = state.fetchCalls.length;
  state.fetchHandler = () =>
    fakeResponse(
      429,
      JSON.stringify({ error: { code: "rate_limited", message: "The request is over the limit.", status: 429 } }),
    );
  context.refreshConcentration();
  check(state.fetchCalls.length === fetches + 1, "run 4 calls the fake once");
  check(x.cell("B1") === "FAULT: 429 rate_limited", `B1 is FAULT: 429 rate_limited (B1 holds "${x.cell("B1")}")`);
  check(isDate(x.cell("B2")) && x.cell("B2").getTime() > thirdTime, "B2 holds a later time");
  checkKept(before, "run 4");
  checkNoRecord("run 4");
  checkNoTabChange(reportState, logFrom, "run 4");
  console.log(`  B1: ${x.cell("B1")}`);

  console.log("\n== Run 5: status 200 with an answer that the report cannot show");
  const noOverlaps = JSON.parse(liveText);
  delete noOverlaps.overlaps;
  const nullStock = JSON.parse(liveText);
  nullStock.lines[0].stockWeight = null;
  const noStock = JSON.parse(liveText);
  delete noStock.lines.at(-1).stockWeight;
  const textStock = JSON.parse(liveText);
  textStock.lines[0].stockWeight = "0.1";
  const badAnswers = [
    ["an answer with no overlaps list", noOverlaps],
    ["a line with null in stockWeight", nullStock],
    ["a line with no stockWeight field", noStock],
    ["a line with a text in stockWeight", textStock],
  ];
  for (const [name, bad] of badAnswers) {
    check(context.parseAnswer(JSON.stringify(bad)) === null, `parseAnswer refuses ${name}`);
    before = x.snapshot();
    const lastTime = x.cell("B2").getTime();
    pause(5);
    logFrom = state.log.length;
    fetches = state.fetchCalls.length;
    state.fetchHandler = () => fakeResponse(200, JSON.stringify(bad));
    context.refreshConcentration();
    check(state.fetchCalls.length === fetches + 1, `${name}: run 5 calls the fake once`);
    check(
      x.cell("B1") === "FAULT: 200 bad_answer",
      `${name}: B1 is FAULT: 200 bad_answer (B1 holds "${x.cell("B1")}")`,
    );
    check(x.cell("B2").getTime() > lastTime, `${name}: B2 holds a later time`);
    checkKept(before, name);
    checkNoRecord(name);
    checkNoTabChange(reportState, logFrom, name);
  }
  check(context.parseAnswer(liveText) !== null, "parseAnswer accepts the answer of run 1");
  console.log(`  B1: ${x.cell("B1")}`);

  console.log("\n== Run 6: Set API key with an empty box removes the key");
  before = x.snapshot();
  const fifthTime = x.cell("B2").getTime();
  pause(5);
  logFrom = state.log.length;
  fetches = state.fetchCalls.length;
  state.promptAnswers.push({ button: BUTTON.OK, text: "  " });
  context.setApiKey();
  check(!state.userProperties.has(KEY_PROPERTY), "an empty box removes the saved key");
  check(state.alerts.at(-1) === "The saved API key is removed.", "the person sees the removal");
  state.fetchHandler = null;
  context.refreshConcentration();
  check(state.fetchCalls.length === fetches, "run 6 calls no fetch");
  check(
    x.cell("B1") === "FAULT: no API key. Use Set API key in the add-on menu.",
    `B1 names the missing key and the menu item (B1 holds "${x.cell("B1")}")`,
  );
  check(x.cell("B2").getTime() > fifthTime, "B2 holds a later time");
  checkKept(before, "run 6");
  checkNoRecord("run 6");
  checkNoTabChange(reportState, logFrom, "run 6");
  console.log(`  B1: ${x.cell("B1")}`);

  console.log("\n== Run 7: the lock is held");
  state.userProperties.set(KEY_PROPERTY, API_KEY);
  state.lockHeldByOther = true;
  before = x.snapshot();
  const logSeven = state.log.length;
  pause(5);
  context.refreshConcentration();
  check(state.fetchCalls.length === fetches, "run 7 calls no fetch");
  check(state.log.length === logSeven, "run 7 changes no cell");
  check(JSON.stringify(x.snapshot()) === JSON.stringify(before), "run 7 leaves each cell, B1 and B2 too");
  checkNoRecord("run 7");
  state.lockHeldByOther = false;
  console.log(`  B1: ${x.cell("B1")}`);

  console.log("\n== Run 8: 201 keys");
  const many = Array.from({ length: 201 }, (_, i) => [`T${String(i).padStart(3, "0")}`, "Example brokerage", 1, "A"]);
  replaceHoldings(many);
  before = x.snapshot();
  logFrom = state.log.length;
  context.refreshConcentration();
  check(state.fetchCalls.length === fetches, "run 8 calls no fetch");
  check(x.cell("B1") === "FAULT: too many positions", `B1 is FAULT: too many positions (B1 holds "${x.cell("B1")}")`);
  checkKept(before, "run 8");
  checkNoRecord("run 8");
  checkNoTabChange(reportState, logFrom, "run 8");
  console.log(`  B1: ${x.cell("B1")}`);

  console.log("\n== Run 9: seven good runs, so eleven good runs in all");
  replaceHoldings(HOLDINGS_ROWS);
  state.fetchHandler = () => fakeResponse(200, liveText);
  const firstStart = runStarts.at(-1);
  for (let n = 5; n <= 11; n += 1) {
    pause(2);
    const times = timedRun();
    check(x.cell("B1") === "OK", `good run ${n}: B1 is OK`);
    checkRecorded(`good run ${n}`, times);
  }
  check(runStarts.length === 11, "the harness counted eleven good runs");
  const rowsNow = runRows();
  check(
    rowsNow.every(([start]) => start !== firstStart) && rowsNow.at(-1)[0] === runStarts[9],
    "the eleventh good run keeps 10 rows and drops the run of run 1, the oldest",
  );
  table(
    ["row", "runStart", "seconds"],
    rowsNow.map(([start, seconds], n) => [`${15 + n}`, new Date(start).toISOString(), seconds]),
  );

  console.log("\n== Run-time formulas B9 and B10 of the report");
  const lastFormula = report.cell("B9");
  const averageFormula = report.cell("B10");
  check(
    report.cell("A9") === "Last run time" && report.cell("A10") === "Average (last 10)",
    "A9 and A10 hold the labels of the run time",
  );
  const known = (seconds) => {
    const sheet = new FakeSheet(EXPOSURE_TAB, 30, 4, []);
    sheet.grid[RUN_HEADER - 1][0] = "runStart";
    sheet.grid[RUN_HEADER - 1][1] = "seconds";
    seconds.forEach((value, n) => {
      sheet.grid[RUN_HEADER + n][0] = new Date(Date.UTC(2026, 0, 10 - n));
      sheet.grid[RUN_HEADER + n][1] = value;
    });
    return sheet;
  };
  const cases = [
    { name: "no recorded run", seconds: [], last: "", average: "" },
    { name: "one run", seconds: [4.2], last: 4.2, average: 4.2 },
    { name: "three runs", seconds: [3, 1.5, 6], last: 3, average: 3.5 },
    { name: "ten runs", seconds: [10, 9, 8, 7, 6, 5, 4, 3, 2, 1], last: 10, average: 5.5 },
  ];
  const formulaRows = [];
  for (const c of cases) {
    const sheet = known(c.seconds);
    const last = calculate(lastFormula, sheet);
    const average = calculate(averageFormula, sheet);
    check(c.last === "" ? last === "" : near(last, c.last), `${c.name}: B9 is ${JSON.stringify(c.last)} (${last})`);
    check(
      c.average === "" ? average === "" : near(average, c.average),
      `${c.name}: B10 is ${JSON.stringify(c.average)} (${average})`,
    );
    formulaRows.push([c.name, c.seconds.join(", ") || "(none)", JSON.stringify(last), JSON.stringify(average)]);
  }
  const recorded = runRows().map(([, seconds]) => seconds);
  check(
    near(calculate(lastFormula, x), recorded[0]),
    `B9 gives B${RUN_ROW} of the Exposure tab after eleven good runs`,
  );
  check(
    near(calculate(averageFormula, x), recorded.reduce((a, b) => a + b, 0) / 10),
    "B10 gives the average of the 10 kept runs of the Exposure tab",
  );
  table(["block", "seconds, newest first", "B9", "B10"], formulaRows);

  console.log("\n== Describe a fund: the sidebar page and the report text");
  logFrom = state.log.length;
  context.showMixSidebar();
  check(state.sidebars.length === 1, "Describe a fund opens one sidebar");
  check(
    state.sidebars[0].file === "sidebar" && state.sidebars[0].title === "Describe a fund",
    "the sidebar shows sidebar.html with the title Describe a fund",
  );
  check(
    state.log.slice(logFrom).every((e) => e.sheet === undefined),
    "the sidebar opens with no change to a tab",
  );
  check(state.documentProperties.size === 0, "the sidebar opens with no property write");
  const page = readFileSync(SIDEBAR_FILE, "utf8");
  const pageScript = [...page.matchAll(/<script>([\s\S]*?)<\/script>/g)].map((m) => m[1]).join("\n");
  const pageText = page
    .replace(/<script>[\s\S]*?<\/script>/g, "")
    .replace(/<style>[\s\S]*?<\/style>/g, "")
    .replace(/<[^>]+>/g, " ");
  check(!/<script[^>]*\ssrc=|<link\b|https?:\/\//i.test(page), "the sidebar loads no file and names no address");
  check(
    !/innerHTML|outerHTML|insertAdjacentHTML|document\.write/.test(page),
    "the sidebar puts each text into the page as text, never as HTML",
  );
  const serverCalls = [...new Set([...pageScript.matchAll(/call\("(\w+)"/g)].map((m) => m[1]))];
  check(
    serverCalls.length === 5 && serverCalls.every((name) => typeof context[name] === "function"),
    `each server function that the sidebar calls exists (${serverCalls.join(", ")})`,
  );
  const jargon = /\bAPI\b|\broute\b|payload|schema|\bparts\b|unknown class|position id/i;
  check(!jargon.test(pageText), "the text of the sidebar holds no API words");
  check(
    !/\bAPI\b|\broute\b|payload|schema|position id/i.test(pageScript),
    "the messages of the sidebar script hold no API words",
  );
  const reportTexts = JSON.parse(created[REPORT_TAB])
    .grid.flat()
    .filter((value) => typeof value === "string")
    .flatMap((value) => (value.startsWith("=") ? splitLiterals(value).literals : [value]));
  const wordy = reportTexts.filter((text) => jargon.test(text));
  check(wordy.length === 0, `no text of the report holds an API word (${wordy.join(" | ")})`);
  const spill = report.cell(SPILL_CELL);
  const stockWords = reportTexts.filter(
    (text) =>
      /\bstocks?\b/i.test(text) &&
      text !== '"stock"' &&
      text !== '"Bonds and other securities of companies whose stock you hold"',
  );
  check(
    stockWords.length === 0,
    `each label says security or securities, not stock, except the group of bonds (${stockWords.join(" | ")})`,
  );
  check(
    report.cell(DISCLAIMER_CELL) ===
      "This report is intended for informational purposes only. It does not constitute investment advice. The presented data might be inaccurate, incomplete, or out of date." &&
      report.cell(SUBTITLE_CELL) === "Securities by company, with a look inside each fund.",
    `${DISCLAIMER_CELL} holds the disclaimer, and ${SUBTITLE_CELL} holds the subtitle`,
  );
  const order = [
    '"Holdings",',
    '{"Holding","Ticker","Accounts","Value","% of portfolio"},',
    "\n yours,\n",
    `"${HOLDINGS_HEADING}",`,
    '"A trust or a closed-end fund, such as SPY, GLD, or IBIT, is not a company. Other holdings shows it as one line.",',
    'HSTACK({"Rank","Company","Ticker","Value","% of portfolio","","Direct"},TRANSPOSE(f)),',
    "top,rest,",
    '"Fund overlap",',
    `"${UNSEEN_TITLE}",`,
    '"Other holdings",',
    'HSTACK("Total of all lines",',
  ];
  const orderAt = order.map((text) => spill.lastIndexOf(text));
  check(
    orderAt.every((at, n) => at > 0 && (n === 0 || at > orderAt[n - 1])),
    "the spill holds Holdings, the note on trusts, the company table, Fund overlap, " +
      `${UNSEEN_TITLE}, Other holdings, and the total, in this order (${orderAt.join(", ")})`,
  );
  for (const text of [
    'hv,XMATCH(TRUE,ARRAYFORMULA(EXACT(TRIM(hh),"Value"))),',
    'fk,FILTER(ARRAYFORMULA(LEFT(IF(ys<>"",ys,yd),64)),yok)',
    "yi,MAP(yu,LAMBDA(k,XMATCH(TRUE,ARRAYFORMULA(EXACT(fk,k))))),",
    "yc,MAP(yu,LAMBDA(k,SUMPRODUCT(EXACT(fk,k)*1))),",
    "yw,MAP(yu,LAMBDA(k,SUMPRODUCT(EXACT(fk,k)*fv))),",
    'SORT(HSTACK(yn,yt,yc,yw,ARRAYFORMULA(IFERROR(yw/tot,""))),4,FALSE)',
    'HSTACK("Total","","",tot,IFERROR(tot/tot,""))',
  ]) {
    check(
      spill.includes(text),
      `the section Holdings computes ${text.slice(0, 48)} from the Holdings tab, so it changes with no refresh`,
    );
  }
  for (const text of [
    `"${ADD_MIX_NOTE}"`,
    "uu>=0.01",
    '"Not described"',
    '" (substitute)"',
    '"Mix entered "&TEXT(d,"mmmm yyyy")',
    'IF(TODAY()-d>182," — check the fact sheet","")',
  ]) {
    check(spill.includes(text), `${SPILL_CELL} holds ${text}`);
  }
  const note = report.cell(NOTE_CELL);
  check(
    report.cell("A3") === "Status" &&
      note.includes('n&IF(n=1," holding ("," holdings (")') &&
      note.includes('" of your portfolio) "&IF(n=1,"is a fund","are funds")&" not looked through."'),
    `${NOTE_CELL}, under the status cell, counts the funds not looked through and their share`,
  );

  console.log("\n== Describe a fund: the sidebar data and the ticker check");
  logFrom = state.log.length;
  state.fetchHandler = null;
  let data = context.mixSidebarData();
  check(
    JSON.stringify(data.holdings.map((h) => h.key)) === JSON.stringify([BOND, PLAN_FUND]),
    `the drop-down lists the holdings that are not looked through (${data.holdings.map((h) => h.key).join(", ")})`,
  );
  check(
    data.holdings.every((h) => h.mix === null) && data.holdings[0].label === `${BOND} – Example Treasury bond`,
    "no holding has a mix yet, and a holding with a symbol shows the symbol and the description",
  );
  check(
    JSON.stringify(data.others.map((h) => h.key)) === JSON.stringify(EXPECTED.map((e) => e.id)) &&
      data.orphans.length === 0,
    "each holding can take a saved mix, and no saved mix is without a holding",
  );
  fetches = state.fetchCalls.length;
  state.fetchHandler = (url) => {
    const ticker = url.slice(FUND_URL.length);
    const error = (code) => fakeResponse(404, JSON.stringify({ error: { code, message: "Invented.", status: 404 } }));
    if (ticker === FUND_A) {
      return fakeResponse(200, JSON.stringify({ fund: { ticker: FUND_A, seriesName: "EXAMPLE INDEX FUND A" } }));
    }
    if (ticker === STOCK || ticker === TRUST) return error("not_a_fund");
    if (ticker === "QQQQX") return error("fund_not_found");
    return fakeResponse(503, "busy");
  };
  const found = context.lookupFund(" $voo ");
  check(
    JSON.stringify(found) === JSON.stringify({ ok: true, ticker: FUND_A, text: "EXAMPLE INDEX FUND A" }),
    `a fund ticker shows the fund name (${JSON.stringify(found)})`,
  );
  const ask = state.fetchCalls.at(-1);
  check(
    ask.url === `${FUND_URL}${FUND_A}` && ask.options.method === "get" && ask.options.payload === undefined,
    "the check reads the fund route of the normalized ticker, with no body",
  );
  check(
    Object.keys(ask.options.headers).join() === "Authorization" &&
      ask.options.headers.Authorization === `Bearer ${API_KEY}`,
    "the check sends the key of the person alone",
  );
  check(
    JSON.stringify(context.lookupFund("aapl")) === JSON.stringify({ ok: true, ticker: STOCK, text: NOT_A_FUND }),
    "a company ticker shows that it counts as one holding",
  );
  check(
    JSON.stringify(context.lookupFund("spy")) === JSON.stringify({ ok: true, ticker: TRUST, text: NOT_A_FUND }),
    "a trust ticker shows that it counts as one holding",
  );
  check(
    JSON.stringify(context.lookupFund("qqqqx")) ===
      JSON.stringify({ ok: false, ticker: "QQQQX", text: "We don't know this ticker." }),
    "a ticker that the service does not know shows that we don't know it",
  );
  check(
    context.lookupFund("ZZZZ").text === "We couldn't check this ticker right now. You can still save it.",
    "another answer shows that the check failed",
  );
  const checked = state.fetchCalls.length;
  check(
    context.lookupFund("not a ticker").ok === false && state.fetchCalls.length === checked,
    "a text that is not a ticker sends no request",
  );
  state.userProperties.delete(KEY_PROPERTY);
  const noKey = context.lookupFund(FUND_A);
  check(
    !noKey.ok && noKey.text.includes("Set API key") && state.fetchCalls.length === checked,
    "with no key, the check names Set API key and sends no request",
  );
  state.userProperties.set(KEY_PROPERTY, API_KEY);
  check(state.fetchCalls.length === fetches + 5, "the ticker checks send five requests");
  check(
    state.log.slice(logFrom).every((e) => e.sheet === undefined),
    "the sidebar data and the ticker check change no tab",
  );

  console.log("\n== Describe a fund: save, link, and delete a mix");
  state.fetchHandler = null;
  const refusals = [
    ["a holding that is not on the Holdings tab", "Example Old Plan Trust", [{ ticker: FUND_A, percent: 10 }]],
    [
      "a total over 100%",
      PLAN_FUND,
      [
        { ticker: FUND_A, percent: 60 },
        { ticker: FUND_B, percent: 50 },
      ],
    ],
    [
      "a ticker two times",
      PLAN_FUND,
      [
        { ticker: "voo", percent: 10 },
        { ticker: FUND_A, percent: 20 },
      ],
    ],
    ["21 funds", PLAN_FUND, Array.from({ length: 21 }, (_, n) => ({ ticker: `T${n}`, percent: 1 }))],
    ["no fund", PLAN_FUND, []],
    ["a text that is not a ticker", PLAN_FUND, [{ ticker: "not a ticker", percent: 10 }]],
    ["an empty ticker", PLAN_FUND, [{ ticker: "", percent: 5 }]],
    ["a percent of 0", PLAN_FUND, [{ ticker: FUND_A, percent: "0" }]],
    ["a percent that is not a number", PLAN_FUND, [{ ticker: FUND_A, percent: "abc" }]],
  ];
  const texts = [
    "This holding is not on your Holdings tab.",
    "The total is over 100%. Lower a percent to save.",
    "VOO is in the mix two times.",
    "A mix can hold 20 funds at most.",
    "Add at least one fund.",
    'Check the ticker "not a ticker".',
    "Type a ticker in each row.",
    "Give VOO a percent above 0 and up to 100.",
    "Give VOO a percent above 0 and up to 100.",
  ];
  refusals.forEach(([name, key, rows], n) => {
    const result = context.saveMix(key, rows);
    check(!result.ok && result.text === texts[n], `the sidebar refuses ${name} (${result.text})`);
  });
  check(state.documentProperties.size === 0, "a refused mix writes no property");
  const todayText = new Intl.DateTimeFormat("en-CA", {
    timeZone: TIME_ZONE,
    year: "numeric",
    month: "2-digit",
    day: "2-digit",
  }).format(new Date());
  const saved = context.saveMix(PLAN_FUND, [
    { ticker: " voo ", percent: "60", substitute: false },
    { ticker: "$ivv", percent: 25, substitute: true },
    { ticker: STOCK, percent: 10 },
  ]);
  check(
    saved.ok && saved.text === "Saved. Choose Refresh to update your report.",
    "a mix with a total of 95% is saved",
  );
  check(
    [...state.documentProperties.keys()].join() === `${MIX_PREFIX}${PLAN_FUND}`,
    `the mix is the document property ${MIX_PREFIX}${PLAN_FUND}, the prefix and the key of the holding`,
  );
  const stored = state.documentProperties.get(`${MIX_PREFIX}${PLAN_FUND}`);
  const wantStored = {
    entered: todayText,
    parts: [
      { ticker: FUND_A, percent: 60, substitute: false },
      { ticker: FUND_B, percent: 25, substitute: true },
      { ticker: STOCK, percent: 10, substitute: false },
    ],
  };
  check(
    stored === JSON.stringify(wantStored),
    `the property holds the entry date of today and the normalized funds (${stored})`,
  );
  check(
    saved.data.holdings.find((h) => h.key === PLAN_FUND)?.mix?.entered === todayText &&
      !saved.data.others.some((h) => h.key === PLAN_FUND),
    "the sidebar data holds the new mix, and the holding cannot take another saved mix",
  );
  const ORPHAN = "Example Old Plan Trust";
  const orphanMix = JSON.stringify({
    entered: "2025-01-15",
    parts: [{ ticker: FUND_A, percent: 100, substitute: false }],
  });
  state.documentProperties.set(`${MIX_PREFIX}${ORPHAN}`, orphanMix);
  state.documentProperties.set(`${MIX_PREFIX}Example broken mix`, "{not json");
  data = context.mixSidebarData();
  check(
    JSON.stringify(data.orphans) === JSON.stringify([{ key: ORPHAN, mix: JSON.parse(orphanMix) }]),
    "the sidebar lists the saved mix that matches no holding, and skips a value that is not a mix",
  );
  check(
    context.linkMix(ORPHAN, PLAN_FUND).text === "That holding has a mix already." &&
      context.linkMix(ORPHAN, "Not a holding").text === "Pick a holding from your Holdings tab." &&
      context.linkMix("Example broken mix", BOND).text === "We can't find this mix.",
    "a link to a holding with a mix, to a key with no holding, and of a broken value is refused",
  );
  const linked = context.linkMix(ORPHAN, BOND);
  check(
    linked.ok &&
      state.documentProperties.get(`${MIX_PREFIX}${BOND}`) === orphanMix &&
      !state.documentProperties.has(`${MIX_PREFIX}${ORPHAN}`),
    "a link moves the mix to the holding and keeps its entry date",
  );
  check(
    linked.data.holdings.find((h) => h.key === BOND)?.mix?.entered === "2025-01-15" && linked.data.orphans.length === 0,
    "after the link, the holding shows the mix and no mix is without a holding",
  );
  const removed = context.deleteMix(BOND);
  check(removed.ok && !state.documentProperties.has(`${MIX_PREFIX}${BOND}`), "a delete removes the mix");
  state.documentProperties.set(`${MIX_PREFIX}${ORPHAN}`, orphanMix);
  check(
    state.log.slice(logFrom).every((e) => e.sheet === undefined),
    "the save, the link, and the delete change no tab",
  );

  console.log(
    `\n== Run 15: a holding with a mix sends parts, with ${LIVE_PARTS ? "a real request" : "a synthetic answer"}`,
  );
  state.fetchHandler = LIVE_PARTS ? liveFetch : offlineFetch;
  fetches = state.fetchCalls.length;
  logFrom = state.log.length;
  pause(2);
  const runMix = timedRun();
  check(state.fetchCalls.length === fetches + 1, "run 15 sends one request");
  checkNoTabChange(reportState, logFrom, "run 15");
  const mixBody = JSON.parse(state.fetchCalls.at(-1).options.payload);
  const mixIds = mixBody.positions.map((p) => p.id);
  const planWeight = EXPECTED.find((e) => e.id === PLAN_FUND).sum / EXPECTED_TOTAL;
  const planSent = mixBody.positions.find((p) => p.id === PLAN_FUND);
  check(
    JSON.stringify(Object.keys(planSent)) === '["id","weight","parts"]' && near(planSent.weight, planWeight),
    "the holding with a mix sends id, weight, and parts, and no ticker",
  );
  const wantParts = [
    [FUND_A, 0.6],
    [FUND_B, 0.25],
    [STOCK, 0.1],
  ];
  check(
    planSent.parts.length === 3 &&
      planSent.parts.every(
        (p, n) =>
          Object.keys(p).join() === "ticker,weight" && p.ticker === wantParts[n][0] && p.weight === wantParts[n][1],
      ),
    "each part holds the ticker and the share of the holding: the percent divided by 100",
  );
  check(
    JSON.stringify(mixIds) === JSON.stringify(EXPECTED.map((e) => e.id)) &&
      mixBody.positions.every((p) => p.id === PLAN_FUND || !("parts" in p)),
    "the body holds the same holdings in the same order, no other parts, and no saved mix without a holding",
  );
  check(
    numbersOf(mixBody).every((n) => n > 0 && n <= 1),
    "each number of the body is a weight, and no dollar value appears",
  );
  check(x.cell("B1") === "OK", `run 15: B1 is OK (B1 holds "${x.cell("B1")}")`);
  checkRecorded("run 15", runMix);
  const mixAnswer = JSON.parse(answerText);
  const unknownLines = mixAnswer.lines.filter((l) => l.class === "unknown");
  check(
    typeof mixAnswer.measures.unknownWeight === "number" &&
      Math.abs(mixAnswer.measures.unknownWeight - unknownLines.reduce((a, l) => a + l.weight, 0)) < 1e-9,
    "unknownWeight is the sum of the weights of the lines of the class unknown",
  );
  const mixTop = topFundsOf(mixAnswer.funds);
  const planFunds = mixTop.filter((f) => f.id === PLAN_FUND);
  check(
    JSON.stringify(planFunds.map((f) => f.ticker).sort()) === JSON.stringify([FUND_A, FUND_B].sort()),
    "the funds block holds one element for each fund of the mix, under the holding id, and none for the stock",
  );
  check(
    planFunds.every(
      (f) =>
        f.partWeight === (f.ticker === FUND_A ? 0.6 : 0.25) && Math.abs(f.weight - planWeight * f.partWeight) < 1e-9,
    ),
    "each element of a fund of the mix holds its share and the weight of the holding times the share",
  );
  check(
    mixTop.filter((f) => f.id !== PLAN_FUND).every((f) => f.partWeight === null),
    "each fund that is not in a mix holds partWeight null",
  );
  const restLine = mixAnswer.lines.find((l) => l.key === `id:${PLAN_FUND}`);
  check(
    restLine?.name === `${PLAN_FUND} (not described)` &&
      restLine.class === "unknown" &&
      Math.abs(restLine.weight - planWeight * 0.05) < 1e-9 &&
      restLine.stockWeight === 0 &&
      JSON.stringify(Object.keys(restLine.sources)) === JSON.stringify([PLAN_FUND]),
    "the 5% that the mix does not describe enters as one line of the class unknown, named (not described)",
  );
  const stockLine = mixAnswer.lines.find((l) => l.ticker === STOCK);
  check(
    (stockLine?.stockSources[PLAN_FUND] ?? 0) > planWeight * 0.1 - 1e-9,
    "the stock of the mix and the stock inside its funds come through the holding id",
  );
  check(
    mixAnswer.lines.every(
      (l) =>
        Object.keys(l.sources).every((id) => mixIds.includes(id)) &&
        Object.keys(l.stockSources).every((id) => mixIds.includes(id)),
    ),
    "each source of each line is a holding id",
  );
  const pairKeys = mixAnswer.overlaps.map((o) => o.ids.join("|"));
  check(
    mixAnswer.overlaps.every((o) => o.ids.every((id) => mixIds.includes(id))) &&
      new Set(pairKeys).size === pairKeys.length &&
      pairKeys.includes(`${FUND_A}|${PLAN_FUND}`),
    "the overlaps block gives one pair for each pair of holdings, the holding with a mix too",
  );
  check(x.cell("B13") === mixAnswer.measures.unknownWeight, "B13 holds unknownWeight");
  const fundsNow = x.block(5, FUND_AT, mixTop.length, FUND_FIELDS.length);
  check(
    mixTop.every((f, n) => fundsNow[n][FUND_FIELDS.length - 1] === (f.partWeight ?? "")),
    "column N holds partWeight, and it is empty for a fund that is not in a mix",
  );
  const serial = context.daySerial(todayText);
  const wantMix = [
    [PLAN_FUND, planSent.weight, serial, FUND_A, 0.6, false],
    [PLAN_FUND, planSent.weight, serial, FUND_B, 0.25, true],
    [PLAN_FUND, planSent.weight, serial, STOCK, 0.1, false],
  ];
  const mixCells = x.block(5, MIX_AT, 4, MIX_FIELDS.length);
  check(
    JSON.stringify(mixCells.slice(0, 3)) === JSON.stringify(wantMix) && mixCells[3].every((v) => v === ""),
    `U5:Z holds one row for each fund of the mix, with the entry date and the substitute flag (${JSON.stringify(mixCells.slice(0, 3))})`,
  );
  const mixParts = expectedParts(mixAnswer.lines, mixIds);
  const mixPartCells = partRowsOf(x, mixParts.length, mixIds.length);
  check(
    JSON.stringify(mixPartCells) === JSON.stringify(mixParts),
    "the part rows hold the lines of the answer with parts",
  );
  const mixComputed = checkComputed(x, "run 15", mixParts.length);
  check(
    mixComputed.unseen.find(([id]) => id === PLAN_FUND)?.[1] === 0 &&
      blockRows(x, UNSEEN_AT, 2).find(([id]) => id === PLAN_FUND)?.[1] === 0,
    "the unseen block gives 0 to the holding with a mix",
  );
  const coverage = calculate(report.cell(COVERAGE_CELL), x);
  const wantCoverage = LIVE_PARTS
    ? `These measures cover ${((1 - mixAnswer.measures.unknownWeight) * 100).toFixed(1)}% of the portfolio.`
    : "These measures cover 89.5% of the portfolio.";
  check(coverage === wantCoverage, `${COVERAGE_CELL} shows the coverage label "${wantCoverage}" (${coverage})`);
  console.log(`  ${COVERAGE_CELL}: ${coverage}`);
  console.log("\nMix block, U5:Z:");
  table(MIX_FIELDS, mixCells.slice(0, 3));
  console.log("\nFunds of the mix in the funds block:");
  table(
    ["id", "ticker", "partWeight", "weight"],
    planFunds.map((f) => [f.id, f.ticker, f.partWeight, f.weight]),
  );

  console.log("\n== Run 16: an answer with no unknownWeight, no preferredWeight, no partWeight, and no entered");
  state.documentProperties.clear();
  const oldShape = JSON.parse(liveText);
  delete oldShape.measures.unknownWeight;
  delete oldShape.measures.preferredWeight;
  oldShape.funds = topFundsOf(oldShape.funds);
  for (const fund of oldShape.funds) {
    for (const name of ["partWeight", "heldBy", "entered", "notEnteredReason"]) delete fund[name];
  }
  state.fetchHandler = () => fakeResponse(200, JSON.stringify(oldShape));
  logFrom = state.log.length;
  pause(2);
  const runOld = timedRun();
  check(x.cell("B1") === "OK", `run 16: B1 is OK (B1 holds "${x.cell("B1")}")`);
  checkRecorded("run 16", runOld);
  checkNoTabChange(reportState, logFrom, "run 16");
  check(x.cell("B13") === "", "B13 is empty when the answer holds no unknownWeight");
  check(x.cell("B14") === "", "B14 is empty when the answer holds no preferredWeight");
  check(
    blockRows(x, FUND_AT, 1).length === oldShape.funds.length,
    `the funds block holds each of the ${oldShape.funds.length} funds with no entered key and no heldBy key`,
  );
  check(
    calculate(report.cell(COVERAGE_CELL), x) === "",
    `${COVERAGE_CELL} shows no coverage label when the answer holds no unknownWeight`,
  );
  check(
    partFaults(x) === 0 &&
      x.block(5, MIX_AT, x.getMaxRows() - 4, MIX_FIELDS.length).every((c) => c.every((v) => v === "")),
    "the part rows of run 1 are back, and the mix block is empty",
  );

  console.log("\n== Run 17: seven funds and a threshold of 2% give five fund series and Other funds");
  const sevenFunds = ["FA", "FB", "FC", "FD", "FE", "FF", "FG"];
  replaceHoldings(
    [...sevenFunds, "SX"].map((symbol) => [symbol, "Example brokerage", 1000, `Example holding ${symbol}`]),
  );
  /**
   * A line of one invented company of the class stock, with its sources.
   */
  const companyLine = (key, sources) => {
    const weight = Object.values(sources).reduce((a, b) => a + b, 0);
    return {
      key: `name:${key}`,
      name: `Example Company ${key}`,
      ticker: null,
      lei: null,
      class: "stock",
      weight,
      sources,
      stockWeight: weight,
      stockSources: sources,
    };
  };
  const sevenAnswer = {
    measures: { ...Object.fromEntries(MEASURES.map((name, i) => [name, i + 0.5])), equity: null },
    funds: sevenFunds.map((id) => ({
      id,
      ticker: id,
      reportDate: "2026-06-30",
      accessionNumber: "0000000000-26-000001",
      holdingCount: 10,
      weight: 0.125,
      coveredWeight: 0.125,
      mergedByTicker: 0,
      mergedByLei: 0,
      mergedByName: 0,
    })),
    overlaps: [],
    lines: [
      companyLine("A", { FA: 0.1, FB: 0.05, FC: 0.04, FD: 0.03, FE: 0.02, FF: 0.01, SX: 0.05 }),
      companyLine("B", { FA: 0.02, FB: 0.02, FF: 0.005 }),
      companyLine("C", { FC: 0.03 }),
      companyLine("D", { FG: 0.004 }),
    ],
  };
  book.tab(REPORT_TAB).grid[22][1] = 0.02;
  const sevenState = book.tab(REPORT_TAB).state();
  state.fetchHandler = () => fakeResponse(200, JSON.stringify(sevenAnswer));
  logFrom = state.log.length;
  pause(2);
  const runSeven = timedRun();
  check(x.cell("B1") === "OK", `run 17: B1 is OK (B1 holds "${x.cell("B1")}")`);
  checkNoTabChange(sevenState, logFrom, "run 17");
  check(
    JSON.stringify(blockRows(x, STOCK_FUND_AT, 1).map(([id]) => id)) === JSON.stringify(sevenFunds),
    "the stock fund block holds the seven funds",
  );
  checkRecorded("run 17", runSeven);
  const sevenChart = checkChart("run 17");
  check(
    sevenChart.companies === 3 && sevenChart.direct && sevenChart.funds === 5 && sevenChart.other,
    "run 17: three companies are at or above 2%, the fund FG has no stock in them, and six funds have, " +
      "so the chart has Direct, five funds, and Other funds",
  );
  check(
    sevenChart.series === 1 + sevenChart.funds + Number(sevenChart.other) && sevenChart.series === 7,
    `run 17: the chart has 1 + 5 + 1 series (${sevenChart.series})`,
  );
  check(
    sevenChart.anchor === REPORT_ROW + 8 + 9 + 1 + BAND_ROWS + 3,
    "run 17: the chart anchors under the chart header row, after 8 holding rows, the Total row, the header row " +
      `and the band of the holdings chart, and 3 companies (row ${sevenChart.anchor})`,
  );
  const sevenHoldings = checkHoldingsChart("run 17");
  check(
    sevenHoldings.groups === 8 && sevenHoldings.anchor === REPORT_ROW + 4 + 9,
    "run 17: the holdings chart anchors under its header row, after 8 holding rows and the Total row " +
      `(row ${sevenHoldings.anchor})`,
  );

  console.log("\n== Run 18: a synthetic answer with a held fund, a fund that did not enter, and a preferred line");
  replaceHoldings(HOLDINGS_ROWS);
  book.tab(REPORT_TAB).grid[22][1] = 0.01;
  /**
   * A fetch handler that answers with the synthetic answer of the positions
   * of the request, with a cap of `maxLines` lines, and sends no request.
   */
  const syntheticFetch = (maxLines) => (url, options) => {
    answerText = JSON.stringify(offlineAnswer(JSON.parse(options.payload).positions, maxLines));
    return fakeResponse(200, answerText);
  };
  /**
   * The position ids of the id row of the hidden tab.
   */
  const idRow = () => x.block(4, SOURCE_AT, 1, MAX_POSITIONS)[0].filter((id) => id !== "");
  state.fetchHandler = syntheticFetch(2000);
  const fullState = book.tab(REPORT_TAB).state();
  logFrom = state.log.length;
  pause(2);
  const runFull = timedRun();
  check(x.cell("B1") === "OK", `run 18: B1 is OK (B1 holds "${x.cell("B1")}")`);
  checkNoTabChange(fullState, logFrom, "run 18");
  checkRecorded("run 18", runFull);
  const fullAnswer = JSON.parse(answerText);
  const fullTop = topFundsOf(fullAnswer.funds);
  check(
    fullAnswer.funds.some((f) => f.heldBy !== null) &&
      fullAnswer.funds.some((f) => f.heldBy === null && f.entered === false && f.accessionNumber === null) &&
      fullAnswer.lines.some((l) => l.class === "preferred"),
    "run 18: the answer holds a held fund, a fund that did not enter with no accession number, and a preferred line",
  );
  check(
    JSON.stringify(blockRows(x, FUND_AT, 1).map(([id]) => id)) === JSON.stringify(fullTop.map((f) => f.id)) &&
      JSON.stringify(fullTop.map((f) => f.id)) === JSON.stringify([FUND_A, FUND_B]),
    `run 18: the funds block holds ${FUND_A} and ${FUND_B} alone, and not the held fund or ${MONEY}`,
  );
  const fullParts = expectedParts(fullAnswer.lines, idRow());
  check(
    JSON.stringify(partRowsOf(x, fullParts.length, idRow().length)) === JSON.stringify(fullParts),
    `run 18: the lines block holds the ${fullParts.length} part rows of the answer`,
  );
  checkComputed(x, "run 18", fullParts.length);
  const fullOwn = blockRows(x, OWN_AT, 5);
  const moneyRow = fullOwn.find((cells) => cells[0] === `ticker:${MONEY}`);
  check(
    moneyRow !== undefined && moneyRow[2] === "fund" && near(moneyRow[3], 0.1),
    `run 18: ${MONEY}, a fund that did not enter, is a direct holding with a row of its own (${JSON.stringify(moneyRow)})`,
  );
  const preferredGroup = blockRows(x, GROUP_AT, 4).find((cells) => cells[0] === "preferred");
  check(
    preferredGroup !== undefined && near(preferredGroup[1], 0.004) && preferredGroup[3] === FUND_A,
    `run 18: the group block holds the class preferred (${JSON.stringify(preferredGroup)})`,
  );
  check(
    near(x.cell("B14"), fullAnswer.measures.preferredWeight) && near(x.cell("B14"), 0.004),
    `run 18: B14 holds preferredWeight (${x.cell("B14")})`,
  );
  check(
    JSON.stringify(blockRows(x, STOCK_FUND_AT, 1).map(([id]) => id)) === JSON.stringify([FUND_A, FUND_B]),
    `run 18: the stock fund block holds ${FUND_A} and ${FUND_B}, and not ${MONEY}`,
  );

  console.log("\n== Run 19: a synthetic answer with a cap of 5 lines and the cap line");
  state.fetchHandler = syntheticFetch(5);
  const capState = book.tab(REPORT_TAB).state();
  logFrom = state.log.length;
  pause(2);
  const runCap = timedRun();
  check(x.cell("B1") === "OK", `run 19: B1 is OK (B1 holds "${x.cell("B1")}")`);
  checkNoTabChange(capState, logFrom, "run 19");
  checkRecorded("run 19", runCap);
  const capAnswer = JSON.parse(answerText);
  const capLine = capAnswer.lines.at(-1);
  check(
    capAnswer.lines.length === 6 && capLine.key === "other:lines" && capLine.stockWeight > 0,
    `run 19: the answer holds 5 lines and the cap line, which holds stock (${capLine.stockWeight})`,
  );
  const capParts = expectedParts(capAnswer.lines, idRow());
  const capRows = partRowsOf(x, capParts.length, idRow().length);
  check(
    JSON.stringify(capRows) === JSON.stringify(capParts) &&
      JSON.stringify(capRows.slice(-2).map((cells) => [cells[0], cells[LINE_FIELDS.length]])) ===
        JSON.stringify([
          ["other:lines", "stock"],
          ["other:lines", "other"],
        ]),
    "run 19: the lines block holds the stock part and the other part of the cap line last",
  );
  checkComputed(x, "run 19", capParts.length);
  const capOwn = blockRows(x, OWN_AT, 5).find((cells) => cells[0] === "other:lines");
  check(
    capOwn !== undefined &&
      capOwn[1] === "Other lines" &&
      near(capOwn[3], capLine.weight - capLine.stockWeight) &&
      capOwn[4] === `${FUND_A}, ${FUND_B}`,
    `run 19: the other part of the cap line has a row of its own (${JSON.stringify(capOwn)})`,
  );
  const spillText = report.cell(SPILL_CELL);
  check(
    spillText.includes('IF(k="other:lines","many small lines, combined",') &&
      spillText.includes('IF(rk>0," and more","")') &&
      spillText.includes('"preferred","Preferred shares inside funds"') &&
      spillText.includes('"preferred","preferred shares"'),
    "the report names the cap line and the class preferred, and the row of the securities under the threshold " +
      "says and more when the cap line holds stock",
  );
  check(
    report.cell("C19").endsWith(",'Concentration.Exposure'!$AR$5:$AR,\"<>other:lines\")") &&
      report
        .cell("C20")
        .endsWith(
          "+SUMIFS('Concentration.Exposure'!$AX$5:$AX,'Concentration.Exposure'!$AY$5:$AY,\"stock\",'Concentration.Exposure'!$AR$5:$AR,\"other:lines\")",
        ),
    "the composition puts the stock part of the cap line under the threshold, whatever its weight",
  );

  console.log("\n== Run 20: a synthetic answer with the line classes trust and cash of a direct position");
  replaceHoldings([...HOLDINGS_ROWS, [TRUST, "Example brokerage", 500, "Example index trust"]]);
  state.documentProperties.set(
    `${MIX_PREFIX}${PLAN_FUND}`,
    JSON.stringify({
      entered: "2026-09-01",
      parts: [
        { ticker: FUND_A, percent: 70, substitute: false },
        { ticker: MIX_TRUST, percent: 30, substitute: false },
      ],
    }),
  );
  state.fetchHandler = (url, options) => {
    answerText = JSON.stringify(offlineAnswer(JSON.parse(options.payload).positions, 2000, TRUST_KINDS));
    return fakeResponse(200, answerText);
  };
  const trustState = book.tab(REPORT_TAB).state();
  logFrom = state.log.length;
  pause(2);
  const runTrust = timedRun();
  check(x.cell("B1") === "OK", `run 20: B1 is OK (B1 holds "${x.cell("B1")}")`);
  checkNoTabChange(trustState, logFrom, "run 20");
  checkRecorded("run 20", runTrust);
  const trustAnswer = JSON.parse(answerText);
  const trustLine = trustAnswer.lines.find((l) => l.key === `ticker:${TRUST}`);
  const mixTrustLine = trustAnswer.lines.find((l) => l.key === `ticker:${MIX_TRUST}`);
  const moneyLine = trustAnswer.lines.find((l) => l.key === `ticker:${MONEY}`);
  check(
    trustLine?.class === "trust" &&
      trustLine.stockWeight === 0 &&
      mixTrustLine?.class === "trust" &&
      moneyLine?.class === "cash" &&
      !trustAnswer.funds.some((f) => f.ticker === MONEY),
    "run 20: the answer holds two trust lines with no stock weight, and a cash line of the money market fund " +
      "with no element of the funds block",
  );
  const trustParts = expectedParts(trustAnswer.lines, idRow());
  check(
    JSON.stringify(partRowsOf(x, trustParts.length, idRow().length)) === JSON.stringify(trustParts),
    `run 20: the lines block holds the ${trustParts.length} part rows of the answer`,
  );
  check(
    trustParts
      .filter((cells) => cells[0] === `ticker:${TRUST}` || cells[0] === `ticker:${MIX_TRUST}`)
      .every((cells) => cells[LINE_FIELDS.length] === "other"),
    "run 20: each part row of a trust line is of the other part, so no trust is a company row",
  );
  checkComputed(x, "run 20", trustParts.length);
  const trustOwn = blockRows(x, OWN_AT, 5);
  const trustRow = trustOwn.find((cells) => cells[0] === `ticker:${TRUST}`);
  check(
    trustRow !== undefined && trustRow[2] === "trust" && near(trustRow[3], trustLine.weight) && trustRow[4] === TRUST,
    `run 20: the trust that the person holds has a row of its own at its full weight (${JSON.stringify(trustRow)})`,
  );
  const cashRow = trustOwn.find((cells) => cells[0] === `ticker:${MONEY}`);
  check(
    cashRow !== undefined && cashRow[2] === "cash" && near(cashRow[3], moneyLine.weight) && cashRow[4] === MONEY,
    `run 20: the money market fund that the person holds has a row of its own (${JSON.stringify(cashRow)})`,
  );
  const trustGroups = blockRows(x, GROUP_AT, 4);
  const trustGroup = trustGroups.find((cells) => cells[0] === "trust");
  check(
    !trustOwn.some((cells) => cells[0] === `ticker:${MIX_TRUST}`) &&
      trustGroup !== undefined &&
      near(trustGroup[1], mixTrustLine.weight) &&
      trustGroup[2] === 1 &&
      trustGroup[3] === PLAN_FUND,
    `run 20: the trust in the fund mix goes into the group trust (${JSON.stringify(trustGroup)})`,
  );
  check(
    JSON.stringify(blockRows(x, FUND_AT, 1).map(([id]) => id)) ===
      JSON.stringify(topFundsOf(trustAnswer.funds).map((f) => f.id)) &&
      !blockRows(x, FUND_AT, 1).some(([id]) => id === MONEY || id === TRUST),
    `run 20: the funds block holds neither ${MONEY} nor ${TRUST}`,
  );
  const trustSpill = report.cell(SPILL_CELL);
  check(
    trustSpill.includes('"trust","trust or closed-end fund, no holdings data","cash","money market fund"') &&
      trustSpill.includes('"trust","Trusts and closed-end funds inside funds"') &&
      trustSpill.includes('"trust","trust or closed-end fund",x'),
    "the report names a direct trust line, a direct cash line, and the group trust",
  );
  state.documentProperties.clear();
  replaceHoldings(HOLDINGS_ROWS);

  /**
   * The accepted state of the report tab with other values in the two cells
   * that a person types in.
   */
  const withInputs = (threshold, minimum) => {
    const expected = structuredClone(accepted[REPORT_TAB]);
    const at = (a1) => parseA1(a1, NEW_ROWS, NEW_COLUMNS);
    expected.grid[at(THRESHOLD_CELL).row - 1][at(THRESHOLD_CELL).column - 1] = threshold;
    expected.grid[at(MINIMUM_CELL).row - 1][at(MINIMUM_CELL).column - 1] = minimum;
    return expected;
  };

  /**
   * The tab operations of the log from the index `from`: each delete and
   * each insert, with the name of the tab.
   */
  const tabOps = (from) =>
    state.log
      .slice(from)
      .filter((e) => e.op === "deleteSheet" || e.op === "insertSheet")
      .map((e) => `${e.op} ${e.sheet}`);

  /**
   * Check the two tabs after a refresh that replaced them: the layout at the
   * time of the request, the two typed values, the layout version, the
   * hidden flags, and the answer.
   */
  const checkReplaced = (name, threshold, minimum) => {
    const tabs = Object.fromEntries(state.atFetch.at(-1).tabs);
    checkLayout(`${name}, ${EXPOSURE_TAB}`, tabs[EXPOSURE_TAB], accepted[EXPOSURE_TAB]);
    checkLayout(`${name}, ${REPORT_TAB}`, tabs[REPORT_TAB], withInputs(threshold, minimum));
    const hidden = book.tab(EXPOSURE_TAB);
    check(hidden.cell(VERSION_CELL) === layoutVersion, `${name}: the new hidden tab holds the current layout version`);
    check(
      hidden.hidden && !book.tab(REPORT_TAB).hidden,
      `${name}: the hidden tab is hidden, and the report tab is not`,
    );
    check(hidden.cell("B1") === "OK", `${name}: B1 is OK (B1 holds "${hidden.cell("B1")}")`);
    check(partFaults(hidden) === 0 && pairFaults(hidden) === 0, `${name}: the new hidden tab holds the answer`);
  };

  console.log("\n== Run 10: tabs with no layout version are replaced and keep a threshold of 2.5%");
  const oldExposure = new FakeSheet(EXPOSURE_TAB, 1200, 220, state.log);
  oldExposure.grid[0][0] = "Status";
  oldExposure.grid[0][1] = "OK";
  oldExposure.grid[4][14] = "ticker:OLD";
  oldExposure.grid[4][19] = 0.5;
  oldExposure.hidden = true;
  const oldReport = new FakeSheet(REPORT_TAB, NEW_ROWS, NEW_COLUMNS, state.log);
  oldReport.grid[14][0] = "Threshold";
  oldReport.grid[14][1] = 0.025;
  oldReport.grid[17][0] = "=LET(old,1,old)";
  book.sheets.splice(0, book.sheets.length, oldReport, makeHoldings(HOLDINGS_ROWS), oldExposure);
  state.fetchHandler = () => fakeResponse(200, liveText);
  runStarts.length = 0;
  logFrom = state.log.length;
  pause(2);
  const runTen = timedRun();
  check(
    JSON.stringify(tabOps(logFrom)) ===
      JSON.stringify([
        `deleteSheet ${EXPOSURE_TAB}`,
        `insertSheet ${EXPOSURE_TAB}`,
        `deleteSheet ${REPORT_TAB}`,
        `insertSheet ${REPORT_TAB}`,
      ]),
    `run 10 deletes each old tab and creates it again, the hidden tab first (${tabOps(logFrom).join("; ")})`,
  );
  checkBlockWrite("run 10", logFrom);
  check(
    indexSince(logFrom, (e) => e.op === "fetch") > indexSince(logFrom, (e) => e.op === "hideSheet"),
    "run 10 replaces the tabs before the request",
  );
  check(!book.sheets.includes(oldExposure) && !book.sheets.includes(oldReport), "run 10: no old tab stays");
  check(
    JSON.stringify(book.names()) === JSON.stringify([REPORT_TAB, "Holdings", EXPOSURE_TAB]),
    "run 10: each new tab takes the position of the old tab",
  );
  check(book.tab(REPORT_TAB).cell(THRESHOLD_CELL) === 0.025, "run 10: the new report tab keeps the threshold of 2.5%");
  check(
    book.tab(REPORT_TAB).cell(MINIMUM_CELL) === 0.1,
    "run 10: the overlap minimum gets its default, because the old tab holds none",
  );
  checkReplaced("run 10", 0.025, 0.1);
  checkRecorded("run 10", runTen);

  console.log("\n== Run 11: tabs with the current layout version stay");
  const keptOrder = [REPORT_TAB, "Holdings", EXPOSURE_TAB];
  const keptExposure = book.tab(EXPOSURE_TAB);
  const keptReport = book.tab(REPORT_TAB);
  const keptState = keptReport.state();
  logFrom = state.log.length;
  pause(2);
  const runEleven = timedRun();
  checkNoTabChange(keptState, logFrom, "run 11", keptOrder);
  check(
    book.tab(EXPOSURE_TAB) === keptExposure && book.tab(REPORT_TAB) === keptReport,
    "run 11: both tabs are the tabs of run 10",
  );
  check(keptReport.cell(THRESHOLD_CELL) === 0.025, "run 11: the threshold of 2.5% stays");
  checkRecorded("run 11", runEleven);

  console.log("\n== Run 12: tabs with an older layout version are replaced and keep both typed values");
  keptExposure.grid[2][1] = layoutVersion - 1;
  keptReport.grid[22][1] = 0.03;
  keptReport.grid[23][1] = 0.25;
  runStarts.length = 0;
  logFrom = state.log.length;
  pause(2);
  const runTwelve = timedRun();
  check(tabOps(logFrom).length === 4, `run 12 deletes and creates both tabs (${tabOps(logFrom).join("; ")})`);
  checkBlockWrite("run 12", logFrom);
  check(
    JSON.stringify(book.names()) === JSON.stringify(keptOrder),
    "run 12: each new tab takes the position of the old tab",
  );
  check(
    book.tab(REPORT_TAB).cell(THRESHOLD_CELL) === 0.03 && book.tab(REPORT_TAB).cell(MINIMUM_CELL) === 0.25,
    "run 12: the new report tab keeps the threshold of 3% and the overlap minimum of 25%",
  );
  checkReplaced("run 12", 0.03, 0.25);
  checkRecorded("run 12", runTwelve);

  console.log("\n== Run 13: the report tab is absent, and the hidden tab holds the current layout version");
  const stayed = book.tab(EXPOSURE_TAB);
  /**
   * The cells of the chart block IT4:JA104 and of the holdings chart block
   * JE4:JF204 of a grid.
   */
  const blockOf = (grid) => [
    grid.slice(3, 4 + CHART_ROWS).map((cells) => cells.slice(CHART_AT - 1, CHART_AT - 1 + CHART_WIDTH)),
    grid
      .slice(3, 4 + HOLDINGS_CHART_ROWS)
      .map((cells) => cells.slice(HOLDINGS_AT - 1, HOLDINGS_AT - 1 + HOLDINGS_WIDTH)),
  ];
  for (const cells of stayed.grid.slice(3, 4 + CHART_ROWS)) {
    cells.fill("stale", CHART_AT - 1, CHART_AT - 1 + CHART_WIDTH);
  }
  for (const cells of stayed.grid.slice(3, 4 + HOLDINGS_CHART_ROWS)) {
    cells.fill("stale", HOLDINGS_AT - 1, HOLDINGS_AT - 1 + HOLDINGS_WIDTH);
  }
  stayed.grid[4][ANCHOR_AT - 1] = "stale";
  stayed.grid[4][HOLDINGS_ANCHOR_AT - 1] = "stale";
  book.sheets.splice(book.sheets.indexOf(book.tab(REPORT_TAB)), 1);
  logFrom = state.log.length;
  pause(2);
  const runThirteen = timedRun();
  check(
    JSON.stringify(tabOps(logFrom)) === JSON.stringify([`insertSheet ${REPORT_TAB}`]),
    `run 13 creates the report tab alone (${tabOps(logFrom).join("; ")})`,
  );
  check(book.tab(EXPOSURE_TAB) === stayed, "run 13: the hidden tab stays");
  checkBlockWrite("run 13", logFrom);
  check(
    JSON.stringify(blockOf(JSON.parse(Object.fromEntries(state.atFetch.at(-1).tabs)[EXPOSURE_TAB]).grid)) ===
      JSON.stringify(blockOf(accepted[EXPOSURE_TAB].grid)),
    "run 13: at the request, the two chart blocks of the hidden tab hold the formulas of the fixture again, " +
      "in place of the stale cells",
  );
  check(
    [ANCHOR_AT, HOLDINGS_ANCHOR_AT].every(
      (column) =>
        JSON.parse(Object.fromEntries(state.atFetch.at(-1).tabs)[EXPOSURE_TAB]).grid[4][column - 1] ===
        accepted[EXPOSURE_TAB].grid[4][column - 1],
    ),
    "run 13: at the request, the two anchor cells hold the formulas of the fixture again",
  );
  check(
    JSON.stringify(book.names()) === JSON.stringify(["Holdings", EXPOSURE_TAB, REPORT_TAB]),
    "run 13: the new report tab is the last tab",
  );
  checkLayout(`run 13, ${REPORT_TAB}`, Object.fromEntries(state.atFetch.at(-1).tabs)[REPORT_TAB], accepted[REPORT_TAB]);
  checkRecorded("run 13", runThirteen);

  console.log("\n== Run 14: the hidden tab is absent, so the script replaces the report tab too");
  book.tab(REPORT_TAB).grid[22][1] = 0.04;
  book.sheets.splice(book.sheets.indexOf(book.tab(EXPOSURE_TAB)), 1);
  runStarts.length = 0;
  logFrom = state.log.length;
  pause(2);
  const runFourteen = timedRun();
  check(
    JSON.stringify(tabOps(logFrom)) ===
      JSON.stringify([`insertSheet ${EXPOSURE_TAB}`, `deleteSheet ${REPORT_TAB}`, `insertSheet ${REPORT_TAB}`]),
    `run 14 creates the hidden tab, then deletes and creates the report tab (${tabOps(logFrom).join("; ")})`,
  );
  checkBlockWrite("run 14", logFrom);
  checkReplaced("run 14", 0.04, 0.1);
  checkRecorded("run 14", runFourteen);

  const keyCells = book.sheets.flatMap((sheet) =>
    sheet.grid.flatMap((cells) => cells.filter((v) => typeof v === "string" && v.includes(API_KEY))),
  );
  check(keyCells.length === 0, "no cell of any tab holds the key");
  check(
    state.alerts.every((text) => !text.includes(API_KEY)),
    "no message holds the key",
  );
  check(
    [...state.documentProperties].every(([name, value]) => name.startsWith(MIX_PREFIX) && !value.includes(API_KEY)),
    "each document property is a mix, and no document property holds the key",
  );
  check(liveRequests === LIVE_LIMIT, `the harness sent ${LIVE_LIMIT} real requests (${liveRequests})`);
  const sentText = ["no real request", "one real request", "two real requests"][LIVE_LIMIT];
  const partsNote = LIVE_PARTS ? "" : " The request with parts got a synthetic answer.";
  console.log(`\nAll ${passed} assertions passed. The harness sent ${sentText}.${partsNote}`);
}

try {
  main();
} catch (e) {
  if (e instanceof AssertionFailure) {
    console.error(`\nFAILED: ${e.message}`);
    process.exit(1);
  }
  throw e;
}
