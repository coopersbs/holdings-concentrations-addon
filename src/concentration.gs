/**
 * The concentration script of a Tiller spreadsheet.
 *
 * The script reads the Holdings tab. It sends the weight of each position to
 * the concentration route of the holdings API. It writes the answer into the
 * hidden tab Concentration.Exposure. The report tab Concentration reads that
 * answer. layout.gs holds the layout of the two tabs. The script creates each
 * of them that is absent, and it replaces both when the hidden tab holds
 * another layout version. It reads no other tab, and it writes no other tab.
 * The menu item Refresh of the add-on menu is the one way to run the
 * report. The add-on menu is under Extensions, with the name of the add-on.
 * Each good refresh also draws the two bar charts of the report tab again,
 * the company chart and the holdings chart, and records its time in the
 * hidden tab. The simple trigger onEdit draws the company chart again after
 * an edit of the threshold cell. It sends no request and writes no cell.
 *
 * The user properties of each person hold the API key of that person. The
 * menu item Set API key writes it. The script does not use the script
 * properties, because all users of an add-on share them.
 *
 * The menu item Describe a fund opens a sidebar. In the sidebar, a person
 * describes a holding that the route cannot look through as a mix of funds.
 * The document properties of the spreadsheet hold each mix, so each editor
 * of the spreadsheet can read it. The sidebar writes the document
 * properties alone. The next refresh sends the mix as the parts of the
 * position and writes it into the hidden tab.
 */

/**
 * The address of the fund routes. The fund route is this address and a
 * ticker.
 */
const FUND_URL = "https://data.coopersbs.com/funds/v1/";

/**
 * The address of the concentration route.
 */
const ROUTE_URL = `${FUND_URL}concentration`;

/**
 * The tab that the script reads. Tiller fills it.
 */
const HOLDINGS_TAB = "Holdings";

/**
 * The hidden tab that holds the answer of the route. Each refresh writes it.
 */
const EXPOSURE_TAB = "Concentration.Exposure";

/**
 * The user property that holds the API key.
 */
const KEY_PROPERTY = "HOLDINGS_API_KEY";

/**
 * The function that the menu item Refresh runs.
 */
const HANDLER = "refreshConcentration";

/**
 * The function that the menu item Set API key runs.
 */
const KEY_HANDLER = "setApiKey";

/**
 * The function that the menu item Describe a fund runs.
 */
const MIX_HANDLER = "showMixSidebar";

/**
 * The HTML file of the sidebar, with no extension.
 */
const SIDEBAR_FILE = "sidebar";

/**
 * The start of the name of each document property that holds a fund mix.
 * The rest of the name is the key of the holding: its Symbol, or its
 * Description when the Symbol is empty, as buildPositions makes it. The
 * value is JSON: the entry date `entered` as yyyy-mm-dd, and `parts`, a list
 * of 1 to MAX_PARTS elements with a `ticker`, a `percent` of the holding,
 * and a `substitute` flag.
 */
const MIX_PREFIX = "FUND_MIX:";

/**
 * The largest count of funds in one mix. The route accepts 20 parts for each
 * position.
 */
const MAX_PARTS = 20;

/**
 * The amount that the percents of a mix can add above 100 and stay valid.
 * The route accepts a sum of part weights up to 1 + 1e-9.
 */
const PERCENT_TOLERANCE = 1e-7;

/**
 * The characters that an API key can hold. A key with another character,
 * such as a space or a line break, cannot go into the Authorization header.
 */
const KEY_PATTERN = /^[A-Za-z0-9_-]{1,256}$/;

/**
 * The largest count of positions that the route accepts in one request.
 */
const MAX_POSITIONS = 200;

/**
 * The largest length of a position id, in UTF-16 code units.
 */
const MAX_ID_LENGTH = 64;

/**
 * The ticker pattern of the concentration route. The pattern applies after
 * the route removes one leading `$` and changes each letter to upper case.
 */
const TICKER_PATTERN = /^[A-Z0-9][A-Z0-9.-]{0,11}$/;

/**
 * The header text of the three columns of the Holdings tab that the script
 * reads.
 */
const HOLDINGS_COLUMNS = ["Description", "Symbol", "Value"];

/**
 * The measures of the answer, in the order of the rows B5:B14. An answer
 * with no unknownWeight leaves B13 empty, and an answer with no
 * preferredWeight leaves B14 empty.
 */
const MEASURE_NAMES = [
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

/**
 * The fields of the equity block of the measures, in the order of the rows
 * B28:B32. The equity block holds the measures of the stock part alone.
 */
const EQUITY_NAMES = ["weight", "lineCount", "top10Weight", "hhi", "effectiveCount"];

/**
 * The fields of each element of the funds block, in the order of the columns
 * D:N. An element of a fund in a mix holds the position id, the ticker of the
 * fund, and its share of the position in partWeight. Each other element
 * holds no partWeight.
 */
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

/**
 * The fields of the mix block, in the order of the columns U:Z. Each fund of
 * a mix that the request sent gets one row: the position id, the weight of
 * the position, the entry date of the mix as a date serial number, the
 * ticker of the fund, its share of the position, and the substitute flag.
 */
const MIX_FIELDS = ["mixId", "mixWeight", "entered", "partTicker", "partWeight", "substitute"];

/**
 * The fields of each line, in the order of the columns AR:AX. Column AY
 * holds the part of the line that the row gives: STOCK_PART or OTHER_PART.
 * Column AZ holds the direct weight of the row.
 */
const LINE_FIELDS = ["key", "name", "ticker", "lei", "class", "weight", "stockWeight"];

/**
 * The fields of the unseen block, in the order of the columns AB:AC. Each
 * position gets one row, in the order of the request: the position id and
 * the weight of the lines of the class unknown that came through the
 * position. A position with a mix gets 0.
 */
const UNSEEN_FIELDS = ["unseenId", "unseenWeight"];

/**
 * The field of the stock fund block, in column AE. Each position that the
 * route looked through and that gave weight to the stock part gets one row,
 * in the order of the funds block, one time each.
 */
const STOCK_FUND_FIELDS = ["stockFundId"];

/**
 * The fields of the own block, in the order of the columns AG:AK. Each row
 * of the lines block that the report shows on a row of its own gets one row:
 * the key, the name, and the class of its line, the part weight, and the ids
 * of the positions that gave weight to the row.
 */
const OWN_FIELDS = ["ownKey", "ownName", "ownClass", "ownWeight", "ownSources"];

/**
 * The fields of the group block, in the order of the columns AM:AP. Each
 * class of the rows of the lines block that go into a group gets one row:
 * the class, the sum of the part weights, the count of rows, and the ids of
 * the positions that gave weight to the group.
 */
const GROUP_FIELDS = ["groupClass", "groupWeight", "groupCount", "groupSources"];

/**
 * The part name of a row that holds the stock part of a line.
 */
const STOCK_PART = "stock";

/**
 * The part name of a row that holds the part of a line that is not stock.
 */
const OTHER_PART = "other";

/**
 * The row of the labels and of the position ids.
 */
const HEADER_ROW = 4;

/**
 * The first row of each data block.
 */
const FIRST_DATA_ROW = 5;

/**
 * The first column of the funds block, D. Column 1 is A.
 */
const FUND_COLUMN = 4;

/**
 * The count of columns of the overlaps block.
 */
const OVERLAP_WIDTH = 4;

/**
 * The first column of the overlaps block, P. The block holds the two position
 * ids, the overlap, and the count of shared lines of each pair, in P:S. One
 * empty column separates each block from the next.
 */
const OVERLAP_COLUMN = FUND_COLUMN + FUND_FIELDS.length + 1;

/**
 * The first column of the mix block, U.
 */
const MIX_COLUMN = OVERLAP_COLUMN + OVERLAP_WIDTH + 1;

/**
 * The first column of the unseen block, AB.
 */
const UNSEEN_COLUMN = MIX_COLUMN + MIX_FIELDS.length + 1;

/**
 * The column of the stock fund block, AE.
 */
const STOCK_FUND_COLUMN = UNSEEN_COLUMN + UNSEEN_FIELDS.length + 1;

/**
 * The first column of the own block, AG.
 */
const OWN_COLUMN = STOCK_FUND_COLUMN + STOCK_FUND_FIELDS.length + 1;

/**
 * The first column of the group block, AM.
 */
const GROUP_COLUMN = OWN_COLUMN + OWN_FIELDS.length + 1;

/**
 * The first column of the lines block, AR.
 */
const LINE_COLUMN = GROUP_COLUMN + GROUP_FIELDS.length + 1;

/**
 * The column of the part name of each row of the lines block, AY.
 */
const PART_COLUMN = LINE_COLUMN + LINE_FIELDS.length;

/**
 * The column of the direct weight of each row of the lines block, AZ: the
 * part weight of the row minus the cells of the positions that the route
 * looked through.
 */
const DIRECT_COLUMN = PART_COLUMN + 1;

/**
 * The first column of the sources block, BA. Each position gets one column.
 */
const SOURCE_COLUMN = DIRECT_COLUMN + 1;

/**
 * The first column of the company chart block, IT. One empty column separates
 * it from the column of position MAX_POSITIONS. The block holds a header in
 * row HEADER_ROW and CHART_ROWS data rows. Its columns are the company name,
 * Direct, CHART_FUNDS fund columns, and Other funds. ensureTabs writes the
 * block after both tabs exist. The write of the answer writes no cell of the
 * block.
 */
const CHART_COLUMN = SOURCE_COLUMN + MAX_POSITIONS + 1;

/**
 * The count of companies that the company chart can show, the count of funds
 * that get a series of their own, and the count of columns of the chart
 * block.
 */
const CHART_ROWS = 100;
const CHART_FUNDS = 5;
const CHART_BLOCK_WIDTH = CHART_FUNDS + 3;

/**
 * The column of the anchor cell of the company chart, JC. One empty column
 * separates it from the company chart block. Row HEADER_ROW holds its label,
 * and row FIRST_DATA_ROW holds the anchor cell: the row of the report tab
 * where the company chart starts. ensureTabs writes the anchor cell after
 * both tabs exist.
 */
const ANCHOR_COLUMN = CHART_COLUMN + CHART_BLOCK_WIDTH + 1;

/**
 * The first column of the holdings chart block, JE. One empty column
 * separates it from the anchor cell. The block holds a header in row
 * HEADER_ROW and HOLDINGS_CHART_ROWS data rows. Its columns are the label of
 * each holding group and the share of the portfolio of the group. ensureTabs
 * writes the block after both tabs exist. The write of the answer writes no
 * cell of the block.
 */
const HOLDINGS_CHART_COLUMN = ANCHOR_COLUMN + 2;

/**
 * The count of holding groups that the holdings chart can show, and the
 * count of columns of the holdings chart block. A refresh sends one position
 * for each holding group, and MAX_POSITIONS positions at most.
 */
const HOLDINGS_CHART_ROWS = MAX_POSITIONS;
const HOLDINGS_BLOCK_WIDTH = 2;

/**
 * The column of the holdings anchor cell, JH. One empty column separates it
 * from the holdings chart block. Row HEADER_ROW holds its label, and row
 * FIRST_DATA_ROW holds the anchor cell: the row of the report tab where the
 * holdings chart starts. ensureTabs writes the anchor cell after both tabs
 * exist.
 */
const HOLDINGS_ANCHOR_COLUMN = HOLDINGS_CHART_COLUMN + HOLDINGS_BLOCK_WIDTH + 1;

/**
 * The row of the header of the run-time block, A15:B15. The block holds one
 * row for each good refresh from row 16: the start time in column A and the
 * seconds in column B, newest first.
 */
const RUN_HEADER_ROW = FIRST_DATA_ROW + MEASURE_NAMES.length;

/**
 * The largest count of refreshes that the run-time block keeps. The rows are
 * A16:B25.
 */
const RUN_LIMIT = 10;

/**
 * The row of the header of the equity block, A27:B27. One empty row
 * separates it from the run-time block. The block holds one row for each
 * field of EQUITY_NAMES from row 28.
 */
const EQUITY_HEADER_ROW = RUN_HEADER_ROW + RUN_LIMIT + 2;

/**
 * Add the items Refresh, Describe a fund, and Set API key to the add-on menu.
 * The add-on menu is under Extensions, with the name of the add-on. The
 * spreadsheet runs this function when a person opens it. The function reads
 * no property and no tab, and it opens no dialog, so it also works before
 * the person gives access.
 */
function onOpen() {
  SpreadsheetApp.getUi()
    .createAddonMenu()
    .addItem("Refresh", HANDLER)
    .addItem("Describe a fund", MIX_HANDLER)
    .addItem("Set API key", KEY_HANDLER)
    .addToUi();
}

/**
 * Add the menu when a person installs the add-on, so the person can use it
 * before the next open of the spreadsheet.
 */
function onInstall() {
  onOpen();
}

/**
 * Draw the company chart of the report tab again after an edit of the
 * threshold cell. The spreadsheet runs this simple trigger after each edit by
 * a person, in AuthMode.LIMITED, which gives it the spreadsheet and no
 * network. It sends no request and writes no cell.
 *
 * The function returns at once, with no log line, when no event object
 * exists or the event has no range, as for a run from the script editor.
 * It also returns at once unless the edited range is one cell in
 * column B of the report tab, in the row whose label in column A is the
 * label of the threshold, as readInputs finds it. It also returns when
 * typedShare gives no threshold for the value of the edit. Then it calls
 * drawChart with that threshold. It catches each error, logs the message
 * and the stack of the error with console.error, and returns, so a failed
 * draw shows no error on the edit.
 */
function onEdit(e) {
  if (!e || !e.range) return;
  try {
    const range = e.range;
    if (range.getNumRows() !== 1 || range.getNumColumns() !== 1 || range.getColumn() !== 2) return;
    const sheet = range.getSheet();
    if (sheet.getName() !== REPORT_TAB || range.getRow() > INPUT_ROWS) return;
    const labels = sheet.getRange(1, 1, Math.min(INPUT_ROWS, sheet.getMaxRows()), 1).getValues();
    if (range.getRow() !== labelIndex(labels, INPUTS.threshold.label) + 1) return;
    const threshold = typedShare(e.value);
    if (threshold === null) return;
    drawChart(SpreadsheetApp.getActiveSpreadsheet(), threshold);
  } catch (error) {
    const { message = String(error), stack = "" } = Object(error);
    console.error(`The chart draw after an edit of the threshold failed: ${message}`, stack);
  }
}

/**
 * The share from 0 to 1 that the value of an edit event gives, as the
 * threshold cell holds it, or null. A number stays as it is. A text is a
 * number with a dot as the decimal mark and an optional percent sign at the
 * end, so 10% and 0.1 both give 0.1. A value outside 0 to 1, an empty
 * value, and each other text give null.
 */
function typedShare(value) {
  if (typeof value === "number") return value >= 0 && value <= 1 ? value : null;
  const match = /^(\d+(?:\.\d*)?|\.\d+)\s*(%?)$/.exec(cellText(value));
  if (match === null) return null;
  const share = Number(match[1]) / (match[2] === "%" ? 100 : 1);
  return share <= 1 ? share : null;
}

/**
 * Ask for the API key and write it into the user properties of the person.
 * An empty answer removes the saved key. Cancel changes nothing. The key does
 * not go into a cell, a log line, or a message.
 */
function setApiKey() {
  const ui = SpreadsheetApp.getUi();
  const response = ui.prompt(
    "Holdings API key",
    "Paste your API key. Leave the box empty to remove the saved key.",
    ui.ButtonSet.OK_CANCEL,
  );
  if (response.getSelectedButton() !== ui.Button.OK) return;
  const key = String(response.getResponseText() || "").trim();
  const store = PropertiesService.getUserProperties();
  if (key === "") {
    store.deleteProperty(KEY_PROPERTY);
    ui.alert("The saved API key is removed.");
    return;
  }
  if (!KEY_PATTERN.test(key)) {
    ui.alert("This text is not an API key. Nothing changed.");
    return;
  }
  store.setProperty(KEY_PROPERTY, key);
  ui.alert("The API key is saved. Use Refresh in the add-on menu.");
}

/**
 * Send the positions of the Holdings tab to the concentration route and write
 * the answer into the tab Concentration.Exposure. The function takes the
 * document lock first, so two runs in one spreadsheet cannot overlap. Runs in
 * different spreadsheets do not wait for each other. When another run holds
 * the lock, the function stops and changes no cell.
 */
function refreshConcentration() {
  const lock = LockService.getDocumentLock();
  if (!lock.tryLock(0)) return;
  try {
    runRefresh();
  } finally {
    lock.releaseLock();
  }
}

/**
 * Do the steps of one refresh. The caller holds the document lock. The first
 * step makes sure that the two tabs of the report hold the layout of
 * layout.gs. A fault writes the status cell B1 and the time cell B2 alone, so
 * the last good answer stays in the other cells.
 *
 * A good refresh draws the two bar charts of the report tab again after the
 * status: the company chart with drawChart and the threshold of readInputs,
 * then the holdings chart with drawHoldingsChart. A fault draws no chart, so
 * the charts of the last good refresh stay.
 *
 * A good refresh records its time in the run-time block. The time starts at
 * the start of this function. It ends after the write of the answer, the
 * status, and the charts. SpreadsheetApp.flush applies the pending writes
 * before the end, so the time includes them. A fault records no time.
 */
function runRefresh() {
  const started = new Date();
  const book = SpreadsheetApp.getActiveSpreadsheet();
  ensureTabs(book);
  const out = book.getSheetByName(EXPOSURE_TAB);
  const holdings = book.getSheetByName(HOLDINGS_TAB);
  if (holdings === null) {
    writeStatus(out, `FAULT: no ${HOLDINGS_TAB} tab`);
    return;
  }
  const rows = holdings.getDataRange().getValues();
  const columns = findColumns(rows.length > 0 ? rows[0] : []);
  if (columns === null) {
    writeStatus(out, `FAULT: no ${HOLDINGS_TAB} column ${missingColumns(rows.length > 0 ? rows[0] : []).join(", ")}`);
    return;
  }
  const mixes = readMixes();
  const positions = buildPositions(rows.slice(1), columns, mixes);
  if (positions.length === 0) {
    writeStatus(out, "FAULT: no positions");
    return;
  }
  if (positions.length > MAX_POSITIONS) {
    writeStatus(out, "FAULT: too many positions");
    return;
  }
  const key = String(PropertiesService.getUserProperties().getProperty(KEY_PROPERTY) || "").trim();
  if (key === "") {
    writeStatus(out, "FAULT: no API key. Use Set API key in the add-on menu.");
    return;
  }
  let response;
  try {
    response = UrlFetchApp.fetch(ROUTE_URL, {
      method: "post",
      contentType: "application/json",
      headers: { Authorization: "Bearer " + key },
      payload: JSON.stringify({ positions }),
      muteHttpExceptions: true,
    });
  } catch (e) {
    writeStatus(out, "FAULT: no answer");
    throw e;
  }
  const status = response.getResponseCode();
  const text = response.getContentText();
  if (status !== 200) {
    writeStatus(out, faultText(status, text));
    return;
  }
  const answer = parseAnswer(text);
  if (answer === null) {
    writeStatus(out, "FAULT: 200 bad_answer");
    return;
  }
  const ids = positions.map((position) => position.id);
  writeAnswer(out, ids, answer, mixRows(positions, mixes));
  writeStatus(out, "OK");
  drawChart(book, readInputs(book.getSheetByName(REPORT_TAB)).threshold);
  drawHoldingsChart(book);
  SpreadsheetApp.flush();
  recordRun(out, started, (Date.now() - started.getTime()) / 1000);
}

/**
 * The text of a cell value with the leading and the trailing white space
 * removed. An empty cell, null, and undefined give an empty string.
 */
function cellText(value) {
  if (value === null || value === undefined) return "";
  return String(value).trim();
}

/**
 * The zero-based index of each column that the script reads, found by the
 * header text of row 1. The result is null when the header holds no
 * Description, Symbol, or Value column.
 */
function findColumns(header) {
  const names = header.map(cellText);
  const found = {
    description: names.indexOf("Description"),
    symbol: names.indexOf("Symbol"),
    value: names.indexOf("Value"),
  };
  if (found.description < 0 || found.symbol < 0 || found.value < 0) return null;
  return found;
}

/**
 * The header text of each column that the script reads and that the header
 * row does not hold.
 */
function missingColumns(header) {
  const names = header.map(cellText);
  return HOLDINGS_COLUMNS.filter((name) => names.indexOf(name) < 0);
}

/**
 * The ticker that the request sends for a symbol, or null. The function
 * applies the ticker normalization of the route to a copy of the symbol. A
 * symbol that fails the pattern gives null, so the position goes with no
 * ticker and the route does not refuse the body.
 */
function tickerOf(symbol) {
  if (symbol === "") return null;
  return normalTicker(symbol) === null ? null : symbol;
}

/**
 * The ticker of a value in the normalized form of the route, or null. The
 * function removes the white space at the two ends and one leading `$`, and
 * it changes each letter to upper case. A result that fails the ticker
 * pattern gives null.
 */
function normalTicker(value) {
  const normal = cellText(value).replace(/^\$/, "").toUpperCase();
  return TICKER_PATTERN.test(normal) ? normal : null;
}

/**
 * The holdings of the rows of the Holdings tab, one for each key, in the
 * order of the first row of each key.
 *
 * The key of a row is its Symbol, or its Description when the Symbol is
 * empty, cut to 64 UTF-16 code units. The function skips a row with an empty
 * key or a Value that is not a finite number. It adds the values of each key.
 * Each holding holds the key in `id`, the Symbol and the Description of the
 * first row of the key, and the sum of the values.
 */
function holdingGroups(rows, columns) {
  const groups = new Map();
  for (const row of rows) {
    const symbol = cellText(row[columns.symbol]);
    const description = cellText(row[columns.description]);
    const key = (symbol !== "" ? symbol : description).slice(0, MAX_ID_LENGTH);
    const value = row[columns.value];
    if (key === "" || typeof value !== "number" || !isFinite(value)) continue;
    if (!groups.has(key)) groups.set(key, { id: key, symbol, description, sum: 0 });
    groups.get(key).sum += value;
  }
  return [...groups.values()];
}

/**
 * The positions of the request, in the order of the first row of each key.
 *
 * holdingGroups gives the keys and their sums. A key with a sum that is not
 * above 0 gives no position, because the route accepts a weight above 0
 * alone. Each weight is the sum of the key divided by the total of the kept
 * keys. A position holds `id`, `ticker` when the key has one, and `weight`.
 * The first row of a key decides whether the position has a ticker.
 *
 * A key with a mix in `mixes` gives a position with `id`, `weight`, and
 * `parts`, and no ticker. Each part holds the ticker of a fund of the mix
 * and its share of the position: the percent divided by 100.
 *
 * A position holds no value, no share count, no account, and no total. The
 * result is empty when no key has a sum above 0.
 */
function buildPositions(rows, columns, mixes = new Map()) {
  const kept = holdingGroups(rows, columns).filter((group) => group.sum > 0);
  const total = kept.reduce((sum, group) => sum + group.sum, 0);
  if (!(total > 0)) return [];
  return kept.map((group) => {
    const position = { id: group.id };
    const mix = mixes.get(group.id);
    if (mix !== undefined) {
      position.weight = group.sum / total;
      position.parts = mix.parts.map((part) => ({ ticker: part.ticker, weight: part.percent / 100 }));
      return position;
    }
    const ticker = tickerOf(group.symbol);
    if (ticker !== null) position.ticker = ticker;
    position.weight = group.sum / total;
    return position;
  });
}

/**
 * The text of a problem of the funds of a mix, or an empty string when the
 * mix is valid. Each fund holds a ticker in the normalized form, a percent
 * above 0 and up to 100, and a substitute flag. A mix holds 1 to MAX_PARTS
 * funds, no ticker two times, and a total of 100 or less. The sidebar shows
 * the text to the person.
 */
function mixProblem(parts) {
  if (!Array.isArray(parts) || parts.length === 0) return "Add at least one fund.";
  if (parts.length > MAX_PARTS) return `A mix can hold ${MAX_PARTS} funds at most.`;
  const seen = new Set();
  let total = 0;
  for (const part of parts) {
    if (part === null || typeof part !== "object") return "Each row needs a ticker and a percent.";
    if (typeof part.ticker !== "string" || part.ticker === "") return "Type a ticker in each row.";
    if (normalTicker(part.ticker) !== part.ticker) return `Check the ticker ${JSON.stringify(part.ticker)}.`;
    if (seen.has(part.ticker)) return `${part.ticker} is in the mix two times.`;
    seen.add(part.ticker);
    if (!isNumber(part.percent) || part.percent <= 0 || part.percent > 100) {
      return `Give ${part.ticker} a percent above 0 and up to 100.`;
    }
    if (typeof part.substitute !== "boolean") return "Each row needs a substitute box.";
    total += part.percent;
  }
  if (total > 100 + PERCENT_TOLERANCE) return "The total is over 100%. Lower a percent to save.";
  return "";
}

/**
 * True when the value is a date text of the form yyyy-mm-dd that names a
 * real day.
 */
function isDay(value) {
  if (typeof value !== "string" || !/^\d{4}-\d{2}-\d{2}$/.test(value)) return false;
  const [year, month, day] = value.split("-").map(Number);
  const date = new Date(Date.UTC(year, month - 1, day));
  return date.getUTCFullYear() === year && date.getUTCMonth() === month - 1 && date.getUTCDate() === day;
}

/**
 * The date serial number of a day of the form yyyy-mm-dd: the count of days
 * after 1899-12-30, the day 0 of Google Sheets.
 */
function daySerial(day) {
  const [year, month, date] = day.split("-").map(Number);
  return Date.UTC(year, month - 1, date) / 86400000 + 25569;
}

/**
 * The mix of the JSON text of a document property, or null. The result
 * holds the entry date and the funds. The result is null when the text is
 * not JSON, when the entry date is not a day, or when mixProblem finds a
 * problem.
 */
function parseMix(text) {
  let value;
  try {
    value = JSON.parse(text);
  } catch {
    return null;
  }
  if (value === null || typeof value !== "object" || !isDay(value.entered)) return null;
  if (mixProblem(value.parts) !== "") return null;
  const parts = value.parts.map((part) => ({
    ticker: part.ticker,
    percent: part.percent,
    substitute: part.substitute,
  }));
  return { entered: value.entered, parts };
}

/**
 * The valid mixes of the document properties, by the key of the holding. A
 * property with a name that does not start with MIX_PREFIX is not a mix. The
 * function skips a value that parseMix refuses.
 */
function readMixes() {
  const mixes = new Map();
  const properties = PropertiesService.getDocumentProperties().getProperties();
  for (const [name, text] of Object.entries(properties)) {
    if (!name.startsWith(MIX_PREFIX)) continue;
    const mix = parseMix(text);
    if (mix !== null) mixes.set(name.slice(MIX_PREFIX.length), mix);
  }
  return mixes;
}

/**
 * The rows of the mix block for the positions of a request: one row for
 * each part of each position that holds parts, in the order of the request.
 * A row holds the fields of MIX_FIELDS.
 */
function mixRows(positions, mixes) {
  const rows = [];
  for (const position of positions) {
    if (!Array.isArray(position.parts)) continue;
    const mix = mixes.get(position.id);
    const entered = daySerial(mix.entered);
    position.parts.forEach((part, i) => {
      rows.push([position.id, position.weight, entered, part.ticker, part.weight, mix.parts[i].substitute]);
    });
  }
  return rows;
}

/**
 * The error code of a body in the error format of the route, or an empty
 * string.
 */
function errorCode(text) {
  try {
    const body = JSON.parse(text);
    if (body && body.error && typeof body.error.code === "string") return body.error.code;
  } catch {
    return "";
  }
  return "";
}

/**
 * The status text of an answer with a status other than 200. A JSON body in
 * the error format of the route adds its error code, such as
 * `FAULT: 429 rate_limited`. Another body gives the status alone, such as
 * `FAULT: 502`.
 */
function faultText(status, text) {
  const code = errorCode(text);
  return code === "" ? `FAULT: ${status}` : `FAULT: ${status} ${code}`;
}

/**
 * The parsed answer of the route, or null. The result is null when the text
 * is not JSON, when it does not hold the blocks measures, funds, overlaps,
 * and lines, or when a line holds no number in stockWeight. The report cannot
 * show the stock part of an answer with no stock weights.
 */
function parseAnswer(text) {
  let body;
  try {
    body = JSON.parse(text);
  } catch {
    return null;
  }
  if (!body || typeof body.measures !== "object" || body.measures === null) return null;
  if (!Array.isArray(body.funds) || !Array.isArray(body.overlaps) || !Array.isArray(body.lines)) return null;
  if (!body.lines.every((line) => line !== null && typeof line === "object" && isNumber(line.stockWeight))) return null;
  return body;
}

/**
 * True when the value is a finite number.
 */
function isNumber(value) {
  return typeof value === "number" && isFinite(value);
}

/**
 * The value that a cell gets for a field. Null and an absent field give an
 * empty string.
 */
function cellValue(value) {
  return value === null || value === undefined ? "" : value;
}

/**
 * The rows B5:B14: one row for each measure, in the order of MEASURE_NAMES.
 */
function measureRows(measures) {
  return MEASURE_NAMES.map((name) => [cellValue(measures[name])]);
}

/**
 * The rows B28:B32: one row for each field of the equity block, in the order
 * of EQUITY_NAMES. The route gives null for the block when the portfolio
 * holds no stock, and each row is then an empty string.
 */
function equityRows(equity) {
  const block = equity !== null && typeof equity === "object" ? equity : {};
  return EQUITY_NAMES.map((name) => [cellValue(block[name])]);
}

/**
 * True when an element of the funds block names a position or a part that
 * entered through the holdings of its fund. The block also names each held
 * fund, with heldBy set, and each fund that did not enter, with entered
 * false. An element with no entered key entered, as in an answer of schema
 * version 1.8.
 */
function isTopFund(fund) {
  return fund !== null && typeof fund === "object" && cellValue(fund.heldBy) === "" && fund.entered !== false;
}

/**
 * The elements of the funds block that isTopFund keeps, in the order of the
 * block.
 */
function topFunds(funds) {
  return funds.filter(isTopFund);
}

/**
 * The rows D5:N: one row for each element of the funds block that isTopFund
 * keeps, with the fields of FUND_FIELDS. A null field gives an empty cell.
 */
function fundRows(funds) {
  return topFunds(funds).map((fund) => FUND_FIELDS.map((name) => cellValue(fund[name])));
}

/**
 * The rows P5:S: one row for each element of the overlaps block, in the order
 * of the answer. A row holds the two position ids, the overlap, and the count
 * of shared lines.
 */
function overlapRows(overlaps) {
  return overlaps.map((pair) => {
    const ids = Array.isArray(pair.ids) ? pair.ids : [];
    return [cellValue(ids[0]), cellValue(ids[1]), cellValue(pair.overlap), cellValue(pair.sharedLineCount)];
  });
}

/**
 * The number that a map of the answer holds for a position id, or null when
 * the map holds no number for it.
 */
function sourceOf(map, id) {
  if (map === null || typeof map !== "object" || !Object.prototype.hasOwnProperty.call(map, id)) return null;
  return isNumber(map[id]) ? map[id] : null;
}

/**
 * The part rows of the lines of an answer. A line gives one row for its
 * stock part, one row for its other part, or both, in the order of the
 * answer. A row holds the fields of LINE_FIELDS, then the part name, then
 * one cell for each position id in the order of the request. writeAnswer
 * puts the direct weight of each row between the part name and the cells,
 * and writes the rows from AR5.
 *
 * A cell of a stock row holds the stock weight that came through the
 * position: the entry of stockSources. A cell of an other row holds the
 * weight that came through the position minus that stock weight: the entry
 * of sources minus the entry of stockSources. A cell is an empty string when
 * the position gave nothing to the part.
 *
 * A line gives a stock row when its stock weight is not 0 or when a position
 * has a stock cell. A line gives an other row when a position has an other
 * cell, or when the line gives no stock row. The two rows of a line split
 * each source between them, so the block needs no second column for each
 * position.
 */
function partRows(lines, ids) {
  const rows = [];
  for (const line of lines) {
    const fields = LINE_FIELDS.map((name) => cellValue(line[name]));
    const stockCells = ids.map((id) => cellValue(sourceOf(line.stockSources, id)));
    const otherCells = ids.map((id) => {
      const whole = sourceOf(line.sources, id);
      const stock = sourceOf(line.stockSources, id);
      if (whole === null && stock === null) return "";
      const other = (whole === null ? 0 : whole) - (stock === null ? 0 : stock);
      return other === 0 ? "" : other;
    });
    const hasStock = line.stockWeight !== 0 || stockCells.some((value) => value !== "");
    if (hasStock) rows.push([...fields, STOCK_PART, ...stockCells]);
    if (!hasStock || otherCells.some((value) => value !== "")) rows.push([...fields, OTHER_PART, ...otherCells]);
  }
  return rows;
}

/**
 * The class of a line of stock.
 */
const STOCK_CLASS = "stock";

/**
 * The class of a line that the route cannot name a kind for: a holding that
 * the route cannot look through.
 */
const UNKNOWN_CLASS = "unknown";

/**
 * The start of the key of a residual line: the part of a fund that its
 * report does not list. The rest of the key is the ticker of the fund.
 */
const RESIDUAL_PREFIX = "residual:";

/**
 * The key of the cap line: the sum of the lines after the largest 2,000
 * lines of an answer. The line holds many securities, so it never is a
 * company row of the report.
 */
const CAP_KEY = "other:lines";

/**
 * The smallest direct weight that gives a row of the other part a row of its
 * own in the report.
 */
const DIRECT_FLOOR = 1e-12;

/**
 * The number of a cell, or 0 when the cell holds no number.
 */
function cellNumber(value) {
  return isNumber(value) ? value : 0;
}

/**
 * The cell of a part row for the position at index `i` of the request.
 */
function sourceCell(row, i) {
  return row[LINE_FIELDS.length + 1 + i];
}

/**
 * The part weight of a part row: the stock weight on a row of the stock
 * part, and the weight minus the stock weight on a row of the other part.
 */
function partWeight(row) {
  const stock = cellNumber(row[LINE_FIELDS.indexOf("stockWeight")]);
  const part = row[LINE_FIELDS.length];
  if (part === STOCK_PART) return stock;
  if (part === OTHER_PART) return cellNumber(row[LINE_FIELDS.indexOf("weight")]) - stock;
  return 0;
}

/**
 * The ids of the positions of the request, in the request order, that gave
 * a number other than 0 to a part row, joined with a comma and a space.
 */
function sourceIds(row, ids) {
  return ids.filter((_, i) => cellNumber(sourceCell(row, i)) !== 0).join(", ");
}

/**
 * The position ids of the elements of the funds block that isTopFund keeps,
 * one time each, in the order of the block. An element with no id adds
 * nothing.
 */
function fundIds(funds) {
  const ids = new Set();
  for (const fund of topFunds(funds)) {
    const id = cellValue(fund.id);
    if (id !== "") ids.add(id);
  }
  return [...ids];
}

/**
 * The direct weight of each part row: the part weight minus the sum of the
 * cells of the positions that the funds block names. The sum adds the cells
 * in the order of the request.
 */
function directWeights(parts, ids, funds) {
  const looked = new Set(fundIds(funds));
  const at = [];
  ids.forEach((id, i) => {
    if (looked.has(id)) at.push(i);
  });
  return parts.map((row) => partWeight(row) - at.reduce((sum, i) => sum + cellNumber(sourceCell(row, i)), 0));
}

/**
 * True when the report gives a part row a row of its own: a row of the
 * other part whose line is not of the class stock, and whose line is a
 * residual line, is the cap line, has a direct weight, or has the class
 * unknown.
 */
function isOwnRow(row, direct) {
  const key = String(row[0]);
  const kind = row[LINE_FIELDS.indexOf("class")];
  return (
    row[LINE_FIELDS.length] === OTHER_PART &&
    kind !== STOCK_CLASS &&
    (key.startsWith(RESIDUAL_PREFIX) || key === CAP_KEY || Math.abs(direct) > DIRECT_FLOOR || kind === UNKNOWN_CLASS)
  );
}

/**
 * The rows of the own block: one row for each part row with a row of its
 * own, in the order of the part rows. A residual row with a part weight of
 * 0 gets no row. The name of a residual row is the ticker of its fund, from
 * the key, because each residual line has the same name. `direct` holds the
 * direct weight of each part row.
 */
function ownRows(parts, ids, direct) {
  const rows = [];
  parts.forEach((row, r) => {
    if (!isOwnRow(row, direct[r])) return;
    const weight = partWeight(row);
    const key = String(row[0]);
    const residual = key.startsWith(RESIDUAL_PREFIX);
    if (residual && weight === 0) return;
    const name = residual ? key.slice(RESIDUAL_PREFIX.length) : row[LINE_FIELDS.indexOf("name")];
    rows.push([row[0], name, row[LINE_FIELDS.indexOf("class")], weight, sourceIds(row, ids)]);
  });
  return rows;
}

/**
 * The rows of the group block: one row for each class of the rows of the
 * other part with no row of their own, in the order of the first row of each
 * class. A row holds the class, the sum of the part weights, the count of
 * rows, and the ids of the positions whose cells add up to a number other
 * than 0. The sums add the rows in their order.
 */
function groupRows(parts, ids, direct) {
  const groups = new Map();
  parts.forEach((row, r) => {
    if (row[LINE_FIELDS.length] !== OTHER_PART || isOwnRow(row, direct[r])) return;
    const kind = row[LINE_FIELDS.indexOf("class")];
    if (!groups.has(kind)) groups.set(kind, { weight: 0, count: 0, sums: ids.map(() => 0) });
    const group = groups.get(kind);
    group.weight += partWeight(row);
    group.count += 1;
    ids.forEach((_, i) => {
      group.sums[i] += cellNumber(sourceCell(row, i));
    });
  });
  return [...groups].map(([kind, group]) => [
    kind,
    group.weight,
    group.count,
    ids.filter((_, i) => group.sums[i] !== 0).join(", "),
  ]);
}

/**
 * The rows of the unseen block: one row for each position, in the order of
 * the request. A row holds the id and the sum of the cells of the position on
 * the rows of the other part of the lines of the class unknown. A position
 * with a mix gets 0. `mixes` holds the rows of the mix block.
 */
function unseenRows(parts, ids, mixes) {
  const mixed = new Set(mixes.map((row) => row[0]));
  const sums = ids.map(() => 0);
  for (const row of parts) {
    if (row[LINE_FIELDS.indexOf("class")] !== UNKNOWN_CLASS || row[LINE_FIELDS.length] !== OTHER_PART) continue;
    ids.forEach((_, i) => {
      sums[i] += cellNumber(sourceCell(row, i));
    });
  }
  return ids.map((id, i) => [id, mixed.has(id) ? 0 : sums[i]]);
}

/**
 * The rows of the stock fund block: each position id of the funds block, one
 * time each and in the order of the block, whose cells on the rows of the
 * stock part add up to a number other than 0.
 */
function stockFundRows(parts, ids, funds) {
  return fundIds(funds)
    .filter((id) => {
      const i = ids.indexOf(id);
      if (i < 0) return false;
      const sum = parts.reduce(
        (total, row) => total + (row[LINE_FIELDS.length] === STOCK_PART ? cellNumber(sourceCell(row, i)) : 0),
        0,
      );
      return sum !== 0;
    })
    .map((id) => [id]);
}

/**
 * Write the status text into B1 and the time of the run into B2 with one
 * call.
 */
function writeStatus(sheet, text) {
  sheet.getRange(1, 2, 2, 1).setValues([[text], [new Date()]]);
}

/**
 * Put one good refresh at the top of the run-time block A16:B25 with one
 * setValues call: the start time in column A and the seconds in column B. The
 * rows of the earlier runs move down one row. The function keeps RUN_LIMIT
 * rows, so the oldest row goes when the block is full. A row with no number
 * of seconds is not a run, and the function drops it.
 */
function recordRun(sheet, started, seconds) {
  const block = sheet.getRange(RUN_HEADER_ROW + 1, 1, RUN_LIMIT, 2);
  const earlier = block.getValues().filter((row) => typeof row[1] === "number");
  const rows = [[started, seconds], ...earlier].slice(0, RUN_LIMIT);
  while (rows.length < RUN_LIMIT) rows.push(["", ""]);
  block.setValues(rows);
}

/**
 * Make the grid hold the given count of rows, and TAB_ROWS rows at least.
 * The function adds rows at the end of the grid or deletes the rows after
 * that count. It adds columns at the end until the grid holds the given
 * count of columns, and it deletes no column, because the report formulas
 * read the sources block up to the column of position MAX_POSITIONS.
 */
function fitGrid(sheet, rows, columns) {
  const want = Math.max(rows, TAB_ROWS);
  const haveRows = sheet.getMaxRows();
  if (want > haveRows) sheet.insertRowsAfter(haveRows, want - haveRows);
  if (want < haveRows) sheet.deleteRows(want + 1, haveRows - want);
  const haveColumns = sheet.getMaxColumns();
  if (columns > haveColumns) sheet.insertColumnsAfter(haveColumns, columns - haveColumns);
}

/**
 * Clear the cells of the last answer that the new answer does not write.
 * `topLast` is the last row of the blocks left of the lines block,
 * `lineLast` is the last row of the lines block, and `lastColumn` is the
 * column of the last position. The function clears three areas: the columns
 * after lastColumn from row 4, the rows after topLast from column B to the
 * column before the lines block, and the rows after lineLast from the lines
 * block to lastColumn. Each area ends at the last row that holds a value.
 * Each area ends at the last column that holds a value, and at the column
 * of position MAX_POSITIONS at most, so the chart blocks and the anchor cells
 * keep their formulas. Column A and the cells that the new answer writes
 * keep their values.
 */
function clearStale(sheet, topLast, lineLast, lastColumn) {
  const lastRow = sheet.getLastRow();
  const oldColumn = Math.min(sheet.getLastColumn(), SOURCE_COLUMN + MAX_POSITIONS - 1);
  if (lastRow < HEADER_ROW) return;
  if (oldColumn > lastColumn) {
    sheet.getRange(HEADER_ROW, lastColumn + 1, lastRow - HEADER_ROW + 1, oldColumn - lastColumn).clearContent();
  }
  if (lastRow > topLast) sheet.getRange(topLast + 1, 2, lastRow - topLast, LINE_COLUMN - 2).clearContent();
  const right = Math.min(oldColumn, lastColumn);
  if (lastRow > lineLast && right >= LINE_COLUMN) {
    sheet.getRange(lineLast + 1, LINE_COLUMN, lastRow - lineLast, right - LINE_COLUMN + 1).clearContent();
  }
}

/**
 * Set the number format `@` on the text columns from row 5: D:E, G, P:Q, U,
 * X, AB, AE, AG:AI, AK, AM, AP, and AR:AU. Also set it on row 4 of the
 * sources block, BA4:IR4. A name that starts with `=`, `+`, `-`, or `@` then
 * stays text, and a ticker such as 0700 stays text.
 */
function setTextFormat(sheet) {
  const dataRows = sheet.getMaxRows() - HEADER_ROW;
  const text = (column, count) => sheet.getRange(FIRST_DATA_ROW, column, dataRows, count).setNumberFormat("@");
  text(FUND_COLUMN, 2);
  text(FUND_COLUMN + 3, 1);
  text(OVERLAP_COLUMN, 2);
  text(MIX_COLUMN, 1);
  text(MIX_COLUMN + 3, 1);
  text(UNSEEN_COLUMN, 1);
  text(STOCK_FUND_COLUMN, 1);
  text(OWN_COLUMN, 3);
  text(OWN_COLUMN + OWN_FIELDS.indexOf("ownSources"), 1);
  text(GROUP_COLUMN, 1);
  text(GROUP_COLUMN + GROUP_FIELDS.indexOf("groupSources"), 1);
  text(LINE_COLUMN, 4);
  sheet.getRange(HEADER_ROW, SOURCE_COLUMN, 1, MAX_POSITIONS).setNumberFormat("@");
}

/**
 * Copy a block of rows into a grid. The row and the column are sheet
 * positions. The grid starts at row HEADER_ROW and at the column `first`.
 */
function placeBlock(grid, first, row, column, values) {
  values.forEach((cells, r) => {
    cells.forEach((value, c) => {
      grid[row - HEADER_ROW + r][column - first + c] = value;
    });
  });
}

/**
 * Write a good answer into the tab with two setValues calls.
 *
 * The first call writes the blocks left of the lines block, from B4 to the
 * column before the lines block and down to the last row of the longest of
 * those blocks: the labels of row 4, the measures, the equity measures, the
 * funds, the overlaps, the mix block, the unseen block, the stock fund block,
 * the own block, and the group block. It also holds column B of the run-time
 * block and of the header of the equity block, and the function copies those
 * cells as they are. The second call writes the lines block, the direct
 * weights, and the sources block, from row 4 down to the last part row and
 * from AR to the column of the last position. Row 4 of that block holds the
 * labels and the position ids.
 *
 * Before the two calls, the function fits the grid to the rows of the answer
 * and clears each cell of the last answer outside the two blocks. Each cell
 * inside a block that the answer does not fill gets an empty string.
 * `mixes` holds the rows of the mix block that mixRows gives.
 */
function writeAnswer(sheet, ids, answer, mixes = []) {
  const funds = fundRows(answer.funds);
  const overlaps = overlapRows(answer.overlaps);
  const parts = partRows(answer.lines, ids);
  const direct = directWeights(parts, ids, answer.funds);
  const unseen = unseenRows(parts, ids, mixes);
  const stockFunds = stockFundRows(parts, ids, answer.funds);
  const own = ownRows(parts, ids, direct);
  const groups = groupRows(parts, ids, direct);
  const topLast = Math.max(
    EQUITY_HEADER_ROW + EQUITY_NAMES.length,
    ...[funds, overlaps, mixes, unseen, stockFunds, own, groups].map((rows) => HEADER_ROW + rows.length),
  );
  const lineLast = HEADER_ROW + parts.length;
  const lastColumn = SOURCE_COLUMN - 1 + ids.length;
  fitGrid(sheet, Math.max(topLast, lineLast), lastColumn);
  clearStale(sheet, topLast, lineLast, lastColumn);
  setTextFormat(sheet);

  const topWidth = LINE_COLUMN - 2;
  const top = Array.from({ length: topLast - HEADER_ROW + 1 }, () => new Array(topWidth).fill(""));
  top[0] = sheet.getRange(HEADER_ROW, 2, 1, topWidth).getValues()[0];
  const kept = sheet.getRange(RUN_HEADER_ROW, 2, EQUITY_HEADER_ROW - RUN_HEADER_ROW + 1, 1).getValues();
  placeBlock(top, 2, FIRST_DATA_ROW, 2, measureRows(answer.measures));
  placeBlock(top, 2, RUN_HEADER_ROW, 2, kept);
  placeBlock(top, 2, EQUITY_HEADER_ROW + 1, 2, equityRows(answer.measures.equity));
  placeBlock(top, 2, FIRST_DATA_ROW, FUND_COLUMN, funds);
  placeBlock(top, 2, FIRST_DATA_ROW, OVERLAP_COLUMN, overlaps);
  placeBlock(top, 2, FIRST_DATA_ROW, MIX_COLUMN, mixes);
  placeBlock(top, 2, FIRST_DATA_ROW, UNSEEN_COLUMN, unseen);
  placeBlock(top, 2, FIRST_DATA_ROW, STOCK_FUND_COLUMN, stockFunds);
  placeBlock(top, 2, FIRST_DATA_ROW, OWN_COLUMN, own);
  placeBlock(top, 2, FIRST_DATA_ROW, GROUP_COLUMN, groups);
  sheet.getRange(HEADER_ROW, 2, top.length, topWidth).setValues(top);

  const labels = sheet.getRange(HEADER_ROW, LINE_COLUMN, 1, SOURCE_COLUMN - LINE_COLUMN).getValues()[0];
  const lines = parts.map((row, r) => [
    ...row.slice(0, LINE_FIELDS.length + 1),
    direct[r],
    ...row.slice(LINE_FIELDS.length + 1),
  ]);
  sheet
    .getRange(HEADER_ROW, LINE_COLUMN, lines.length + 1, lastColumn - LINE_COLUMN + 1)
    .setValues([[...labels, ...ids], ...lines]);
}

/**
 * Open the sidebar of the menu item Describe a fund. The sidebar reads and
 * writes the mixes through the functions below. It writes no cell.
 */
function showMixSidebar() {
  const page = HtmlService.createHtmlOutputFromFile(SIDEBAR_FILE).setTitle("Describe a fund");
  SpreadsheetApp.getUi().showSidebar(page);
}

/**
 * The holdings of the Holdings tab with a sum above 0, as holdingGroups gives
 * them. The result is empty when the tab or one of its three columns is
 * absent.
 */
function readHoldingGroups(book) {
  const holdings = book.getSheetByName(HOLDINGS_TAB);
  if (holdings === null) return [];
  const rows = holdings.getDataRange().getValues();
  const columns = findColumns(rows.length > 0 ? rows[0] : []);
  if (columns === null) return [];
  return holdingGroups(rows.slice(1), columns).filter((group) => group.sum > 0);
}

/**
 * The ids of the positions of the last good answer that gave weight to a line
 * of the class unknown. The function reads the hidden tab. The result is
 * empty when the tab is absent or holds another layout version.
 */
function unknownIds(book) {
  const found = new Set();
  const sheet = book.getSheetByName(EXPOSURE_TAB);
  if (sheet === null || sheet.getRange(VERSION_CELL).getValue() !== LAYOUT_VERSION) return found;
  const width = Math.min(MAX_POSITIONS, sheet.getMaxColumns() - SOURCE_COLUMN + 1);
  const ids = sheet.getRange(HEADER_ROW, SOURCE_COLUMN, 1, width).getValues()[0];
  const rows = sheet
    .getRange(FIRST_DATA_ROW, LINE_COLUMN, sheet.getMaxRows() - HEADER_ROW, SOURCE_COLUMN - LINE_COLUMN + width)
    .getValues();
  const classAt = LINE_FIELDS.indexOf("class");
  const partAt = LINE_FIELDS.length;
  for (const row of rows) {
    if (row[classAt] !== "unknown" || row[partAt] !== OTHER_PART) continue;
    ids.forEach((id, i) => {
      const value = row[SOURCE_COLUMN - LINE_COLUMN + i];
      if (id !== "" && isNumber(value) && value > 0) found.add(id);
    });
  }
  return found;
}

/**
 * The text that the sidebar shows for a holding: the Symbol and the
 * Description, or the key alone.
 */
function holdingLabel(group) {
  if (group.symbol === "" || group.description === "" || group.description === group.symbol) return group.id;
  return `${group.symbol} – ${group.description}`;
}

/**
 * The data of the sidebar. `holdings` holds each holding of the Holdings tab
 * that the last good answer could not look through, and each holding with a
 * mix, with its mix or null. `others` holds each holding with no mix, so a
 * person can link a saved mix to it. `orphans` holds each saved mix whose
 * key matches no holding of the Holdings tab.
 */
function mixSidebarData() {
  const book = SpreadsheetApp.getActiveSpreadsheet();
  const groups = readHoldingGroups(book);
  const mixes = readMixes();
  const unknown = unknownIds(book);
  const keys = new Set(groups.map((group) => group.id));
  return {
    holdings: groups
      .filter((group) => mixes.has(group.id) || unknown.has(group.id))
      .map((group) => ({ key: group.id, label: holdingLabel(group), mix: mixes.get(group.id) || null })),
    others: groups
      .filter((group) => !mixes.has(group.id))
      .map((group) => ({ key: group.id, label: holdingLabel(group) })),
    orphans: [...mixes.entries()].filter(([key]) => !keys.has(key)).map(([key, mix]) => ({ key, mix })),
  };
}

/**
 * The day of today in the time zone of the spreadsheet, as yyyy-mm-dd.
 */
function today(book) {
  return Utilities.formatDate(new Date(), book.getSpreadsheetTimeZone(), "yyyy-MM-dd");
}

/**
 * Save the mix of a holding in the document properties, with the entry date
 * of today. `rows` holds a ticker, a percent, and a substitute flag for each
 * fund. The function normalizes each ticker. It refuses a key that names no
 * holding of the Holdings tab, and a mix that mixProblem refuses. The result
 * holds `ok`, the text for the person, and the new sidebar data.
 */
function saveMix(key, rows) {
  const book = SpreadsheetApp.getActiveSpreadsheet();
  const id = cellText(key);
  if (!readHoldingGroups(book).some((group) => group.id === id)) {
    return { ok: false, text: "This holding is not on your Holdings tab." };
  }
  const parts = (Array.isArray(rows) ? rows : []).map((row) => {
    const cells = row !== null && typeof row === "object" ? row : {};
    const ticker = normalTicker(cells.ticker);
    const percent = typeof cells.percent === "number" ? cells.percent : Number(cellText(cells.percent));
    return {
      ticker: ticker === null ? cellText(cells.ticker) : ticker,
      percent,
      substitute: cells.substitute === true,
    };
  });
  const problem = mixProblem(parts);
  if (problem !== "") return { ok: false, text: problem };
  const value = JSON.stringify({ entered: today(book), parts });
  PropertiesService.getDocumentProperties().setProperty(MIX_PREFIX + id, value);
  return { ok: true, text: "Saved. Choose Refresh to update your report.", data: mixSidebarData() };
}

/**
 * Delete the saved mix of a key. The result holds the text for the person
 * and the new sidebar data.
 */
function deleteMix(key) {
  PropertiesService.getDocumentProperties().deleteProperty(MIX_PREFIX + cellText(key));
  return { ok: true, text: "The mix is deleted. Choose Refresh to update your report.", data: mixSidebarData() };
}

/**
 * Move a saved mix from one key to a holding of the Holdings tab. The mix
 * keeps its entry date. The function refuses a key with no valid mix, a
 * target that is not a holding, and a target with a mix.
 */
function linkMix(fromKey, toKey) {
  const book = SpreadsheetApp.getActiveSpreadsheet();
  const store = PropertiesService.getDocumentProperties();
  const from = cellText(fromKey);
  const to = cellText(toKey);
  const text = store.getProperty(MIX_PREFIX + from);
  if (text === null || parseMix(text) === null) return { ok: false, text: "We can't find this mix." };
  if (!readHoldingGroups(book).some((group) => group.id === to)) {
    return { ok: false, text: "Pick a holding from your Holdings tab." };
  }
  if (store.getProperty(MIX_PREFIX + to) !== null) return { ok: false, text: "That holding has a mix already." };
  store.setProperty(MIX_PREFIX + to, text);
  store.deleteProperty(MIX_PREFIX + from);
  return { ok: true, text: "Linked. Choose Refresh to update your report.", data: mixSidebarData() };
}

/**
 * The name of the fund of a ticker, for the sidebar. The function sends one
 * request to the fund route with the key of the person. The result holds
 * `ok`, the normalized ticker, and the text for the person: the fund name, a
 * note that the ticker is a company stock or a trust, or a note that the
 * service does not know the ticker. A text that is not a ticker sends no
 * request.
 */
function lookupFund(ticker) {
  const normal = normalTicker(ticker);
  if (normal === null) return { ok: false, ticker: cellText(ticker), text: "This is not a ticker." };
  const key = String(PropertiesService.getUserProperties().getProperty(KEY_PROPERTY) || "").trim();
  const later = "We couldn't check this ticker right now. You can still save it.";
  if (key === "") {
    return {
      ok: false,
      ticker: normal,
      text: "Set your key first: Extensions › Holdings Concentration for Tiller › Set API key.",
    };
  }
  let response;
  try {
    response = UrlFetchApp.fetch(FUND_URL + normal, {
      method: "get",
      headers: { Authorization: "Bearer " + key },
      muteHttpExceptions: true,
    });
  } catch {
    return { ok: false, ticker: normal, text: later };
  }
  const status = response.getResponseCode();
  const text = response.getContentText();
  if (status === 200) {
    let name = "";
    try {
      const body = JSON.parse(text);
      if (body && body.fund && typeof body.fund.seriesName === "string") name = body.fund.seriesName.trim();
    } catch {
      name = "";
    }
    return { ok: true, ticker: normal, text: name === "" ? normal : name };
  }
  const code = errorCode(text);
  if (status === 404 && code === "not_a_fund") {
    return { ok: true, ticker: normal, text: "A company stock or a trust, not a fund. It counts as one holding." };
  }
  if (status === 404 && code === "fund_not_found")
    return { ok: false, ticker: normal, text: "We don't know this ticker." };
  return { ok: false, ticker: normal, text: later };
}
