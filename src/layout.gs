/**
 * The layout of the two tabs of the concentration report, and the functions
 * that create a tab from its layout.
 *
 * The tab Concentration.Exposure is hidden. It holds the layout version, the
 * answer of the concentration route, the fund mixes that the request sent,
 * the values that the script computes from the answer, and the run times of
 * the 10 newest good refreshes. concentration.gs writes the answer, the
 * mixes, the computed values, and the run times. No computed value reads
 * the threshold, the overlap minimum, or the total value of the Holdings
 * tab. The report formulas read those cells, so a change of one of them
 * shows in the report with no refresh. The tab
 * Concentration is the report. Each of its cells is a label, a formula, or
 * one of the two cells that the person types in: the threshold and the
 * overlap minimum. The formulas read Concentration.Exposure and the Holdings
 * tab.
 *
 * The hidden tab also holds two chart blocks after the sources block, each
 * with an anchor cell after it. Each cell of the company chart block is a
 * formula that reads the company table of the report spill. Each cell of
 * the holdings chart block is a formula that reads the section
 * Holdings of the report spill. An anchor cell holds the row of the report
 * tab where its bar chart starts: the first row of the band of blank rows
 * under the chart header row of that chart in the report spill.
 * drawBlockChart draws a bar chart from its block at that row. Each good
 * refresh draws both charts. onEdit draws the company chart alone, after an
 * edit of the threshold cell.
 *
 * A refresh calls ensureTabs before the request. The function compares the
 * layout version of the hidden tab with LAYOUT_VERSION. When the two are
 * equal, it creates the report tab if that tab is absent. When they differ,
 * or when the hidden tab is absent, it replaces both tabs and keeps the two
 * values that the person typed. Each time it replaces a tab, its last step
 * writes the chart blocks and the anchor cells, after both tabs exist.
 * Delete the hidden tab to get the layout of both tabs again on the next
 * refresh.
 */

/**
 * The name of the report tab.
 */
const REPORT_TAB = "Concentration";

/**
 * The version of the layout of the two tabs. Add 1 when a change moves a cell
 * that a formula or the script reads. The next refresh then replaces the two
 * tabs of each spreadsheet.
 */
const LAYOUT_VERSION = 14;

/**
 * The cell of the tab Concentration.Exposure that holds the layout version.
 * A tab with no value in this cell holds a layout with no version, and a
 * refresh replaces it.
 */
const VERSION_CELL = "B3";

/**
 * The count of rows of each new tab.
 */
const TAB_ROWS = 1000;

/**
 * The count of columns of the report tab. The last column is Z.
 */
const REPORT_COLUMNS = 26;

/**
 * The two cells of the report tab that the person types in. Each one has a
 * name, the label in column A of its row, and the value of a new tab.
 * readInputs finds a cell by its label. When a label gets another text, make
 * readInputs find the old text too, or a replaced tab loses the typed value.
 */
const INPUTS = {
  threshold: { label: "Threshold", value: 0.01 },
  overlapMinimum: { label: "Overlap minimum", value: 0.1 },
};

/**
 * The count of rows at the top of a report tab that readInputs reads.
 */
const INPUT_ROWS = 40;

/**
 * The row of the status cell and the row of the total value cell of the
 * report tab. Both cells are in column B.
 */
const STATUS_ROW = 3;
const TOTAL_ROW = 6;

/**
 * The row of the header of the block of the security measures, and the row
 * of the header of the composition block.
 */
const STOCKS_ROW = 12;
const COMPOSITION_ROW = 18;

/**
 * The row of the threshold cell, and the row of the overlap minimum cell.
 * Both cells are in column B.
 */
const THRESHOLD_ROW = 23;
const OVERLAP_ROW = 24;

/**
 * The row of the report spill. The spill holds the section Holdings,
 * the header of the company table, the company table, and each section under
 * it. The row count of the section Holdings changes with the Holdings
 * tab, so the header of the company table is inside the spill.
 */
const REPORT_ROW = 26;

/**
 * The share of the portfolio at or above which a holding that the route
 * cannot look through gets a row of its own in the section of the funds not
 * looked through.
 */
const UNSEEN_FLOOR = 0.01;

/**
 * The title of the section of the funds not looked through, and the text of
 * each row of a holding with no mix.
 */
const UNSEEN_TITLE = "Funds not looked through";
const ADD_MIX_NOTE = "Add its fund mix: Concentration › Describe a fund";

/**
 * The count of days after which the report asks the person to check a mix.
 */
const MIX_AGE_DAYS = 182;

/**
 * The title of the section of the holdings that are not securities of a
 * company, below the company table, and the title of the section of the
 * groups of the Holdings tab, above it.
 */
const OTHER_TITLE = "Other holdings";
const HOLDINGS_TITLE = "Holdings";

/**
 * The disclaimer of the report tab, in B1.
 */
const DISCLAIMER =
  "This report is intended for informational purposes only. It does not constitute investment advice. The presented data might be inaccurate, incomplete, or out of date.";

/**
 * The first row of the spill of the fund table, in column D, and the last
 * row that the spill can use. The report spill starts in row REPORT_ROW.
 */
const FUND_ROW = 4;
const FUND_LAST_ROW = REPORT_ROW - 1;

/**
 * The header of the company table of the report spill, from column A. One
 * column for each id of the stock fund block follows it. The chart block
 * finds the Company column, the Direct column, and the first fund column by
 * this list.
 */
const COMPANY_COLUMNS = ["Rank", "Company", "Ticker", "Value", "% of portfolio", "", "Direct"];

/**
 * The header of the table of the section Holdings of the report spill,
 * from column A. The holdings chart block finds the Holding, the Ticker, the
 * Accounts, and the % of portfolio columns by this list.
 */
const HOLDING_COLUMNS = ["Holding", "Ticker", "Accounts", "Value", "% of portfolio"];

/**
 * The colors of the series of the company chart: one for Direct, one for each
 * of the CHART_FUNDS funds in rank order, and one gray for Other funds. The
 * one series of the holdings chart has the Direct color. The colors are
 * fixed, so the largest fund always gets the first fund color. Each pair of
 * fund colors differs in OKLCH lightness by 0.06 or more. Each two neighbors
 * of the stack, from Direct to Other funds, differ by an OKLab distance of 8
 * or more under simulated protan vision and simulated deutan vision, and by
 * 15 or more under normal vision.
 */
const CHART_DIRECT_COLOR = "#2a78d6";
const CHART_FUND_COLORS = ["#f26e3b", "#037952", "#eba007", "#c75d87", "#036503"];
const CHART_OTHER_COLOR = "#a9a7a0";

/**
 * The width in pixels of each column of the report tab from A to F.
 */
const REPORT_WIDTHS = { A: 250, B: 300, C: 150, D: 175, E: 110, F: 180 };

/**
 * The size of each bar chart in pixels, and the count of blank rows of the
 * band that the report spill holds for each chart. A chart spans the columns
 * A to D of the report tab. A band holds the height of a chart at the
 * default row height of 21 pixels, and one more row as a margin.
 */
const CHART_WIDTH = REPORT_WIDTHS.A + REPORT_WIDTHS.B + REPORT_WIDTHS.C + REPORT_WIDTHS.D;
const CHART_HEIGHT = 520;
const CHART_BAND_ROWS = Math.ceil(CHART_HEIGHT / 21) + 1;

/**
 * The largest length of the series name of a fund of the company chart. A
 * fund with no symbol has its description as its id, and the name is the
 * start of it. The holdings chart cuts the label of a holding group with no
 * ticker to the same length.
 */
const CHART_NAME_LENGTH = 24;

/**
 * The parts of the text of the header row of the company chart in the report
 * spill: the text before the threshold, the count of decimals of the
 * threshold as a percent, and the text after it. chartHeading builds the
 * expression of the spill formula from these parts, and headingText builds
 * the text that drawChart expects in that row, so the two texts agree.
 */
const CHART_HEADING = { before: "Companies at or over ", decimals: 2, after: " of the portfolio, by source" };

/**
 * The text of the header row of the holdings chart in the report spill. The
 * text is fixed. The holdings anchor cell finds it in column A of the
 * report tab, and drawHoldingsChart expects it in the row above the anchor
 * row.
 */
const HOLDINGS_HEADING = "Holdings by share of the portfolio";

/**
 * The time in milliseconds that drawBlockChart waits before it reads an
 * anchor cell again, and the largest total time of the wait.
 */
const CHART_WAIT_STEP = 500;
const CHART_WAIT_LIMIT = 20000;

/**
 * A reference to a range of the tab Concentration.Exposure.
 */
function exposure(range) {
  return `'${EXPOSURE_TAB}'!${range}`;
}

/**
 * The cell of the tab Concentration.Exposure that holds one measure.
 */
function measureCell(name) {
  return exposure(`B${FIRST_DATA_ROW + MEASURE_NAMES.indexOf(name)}`);
}

/**
 * The cell of the tab Concentration.Exposure that holds one field of the
 * equity block.
 */
function equityCell(name) {
  return exposure(`B${EQUITY_HEADER_ROW + 1 + EQUITY_NAMES.indexOf(name)}`);
}

/**
 * The column letters of the blocks of the tab Concentration.Exposure that the
 * report formulas read: the funds block, the overlaps block, the mix block,
 * the unseen block, the stock fund block, the own block, the group block,
 * the lines block, and the sources block. The sources block ends at the
 * column of position MAX_POSITIONS, because a request holds MAX_POSITIONS
 * positions at most.
 */
function exposureColumns() {
  const fund = (name) => columnLetter(FUND_COLUMN + FUND_FIELDS.indexOf(name));
  const mix = (name) => columnLetter(MIX_COLUMN + MIX_FIELDS.indexOf(name));
  const own = (name) => columnLetter(OWN_COLUMN + OWN_FIELDS.indexOf(name));
  const group = (name) => columnLetter(GROUP_COLUMN + GROUP_FIELDS.indexOf(name));
  const line = (name) => columnLetter(LINE_COLUMN + LINE_FIELDS.indexOf(name));
  return {
    fundId: fund("id"),
    fundWeight: fund("weight"),
    fundPart: fund("partWeight"),
    pairFirst: columnLetter(OVERLAP_COLUMN),
    pairSecond: columnLetter(OVERLAP_COLUMN + 1),
    pairOverlap: columnLetter(OVERLAP_COLUMN + 2),
    pairShared: columnLetter(OVERLAP_COLUMN + 3),
    mixId: mix("mixId"),
    mixWeight: mix("mixWeight"),
    mixEntered: mix("entered"),
    mixTicker: mix("partTicker"),
    mixPart: mix("partWeight"),
    mixSubstitute: mix("substitute"),
    unseenId: columnLetter(UNSEEN_COLUMN),
    unseenWeight: columnLetter(UNSEEN_COLUMN + 1),
    stockFund: columnLetter(STOCK_FUND_COLUMN),
    ownKey: own("ownKey"),
    ownName: own("ownName"),
    ownClass: own("ownClass"),
    ownWeight: own("ownWeight"),
    ownSources: own("ownSources"),
    groupClass: group("groupClass"),
    groupWeight: group("groupWeight"),
    groupCount: group("groupCount"),
    groupSources: group("groupSources"),
    key: line("key"),
    name: line("name"),
    ticker: line("ticker"),
    class: line("class"),
    weight: line("weight"),
    stock: line("stockWeight"),
    part: columnLetter(PART_COLUMN),
    direct: columnLetter(DIRECT_COLUMN),
    first: columnLetter(SOURCE_COLUMN),
    last: columnLetter(SOURCE_COLUMN + MAX_POSITIONS - 1),
  };
}

/**
 * A column of a block of the tab Concentration.Exposure, from the first data
 * row to the last row of the tab.
 */
function exposureColumn(letter) {
  return exposure(`$${letter}$${FIRST_DATA_ROW}:$${letter}`);
}

/**
 * The formula of B6: the sum of the Value column of the Holdings tab. The
 * formula finds the column by the header text in row 1, with the match rule
 * of findColumns. INDIRECT with the R1C1 text "C" and the column number reads
 * the whole column, so the formula holds no column letter. The header text in
 * row 1 adds nothing to the sum. When row 1 holds no Value column, B6 shows a
 * text in place of a number.
 */
const TOTAL_FORMULA = `=LET(c,XMATCH(TRUE,ARRAYFORMULA(EXACT(TRIM(Holdings!$1:$1),"Value"))),
IF(ISNA(c),"No Holdings column Value",SUM(INDIRECT("Holdings!C"&c,FALSE))))`;

/**
 * The formula of D4: the funds that the route looked through, with the
 * report date, the count of holdings, the weight, and the covered part. A
 * fund of a mix shows the holding and the ticker of the fund. When the route
 * looked through no fund, D4 shows a text.
 */
function fundsFormula() {
  const x = exposureColumns();
  const block = exposure(`$${x.fundId}$${FIRST_DATA_ROW}:$${x.fundPart}`);
  const part = FUND_FIELDS.indexOf("partWeight") + 1;
  return `=LET(r,IFNA(FILTER(${block},${exposureColumn(x.fundId)}<>""),""),
IF(INDEX(r,1,1)="","No fund looked through.",
HSTACK(ARRAYFORMULA(IF(CHOOSECOLS(r,${part})="",CHOOSECOLS(r,1),CHOOSECOLS(r,1)&" › "&CHOOSECOLS(r,2))),
CHOOSECOLS(r,3,5,6),ARRAYFORMULA(CHOOSECOLS(r,7)/CHOOSECOLS(r,6)))))`;
}

/**
 * The formula of B9: the seconds of the newest run in A16:B25 of the tab
 * Concentration.Exposure. B9 is empty when no run is recorded.
 */
function runLastFormula() {
  const cell = exposure(`B${RUN_HEADER_ROW + 1}`);
  return `=IF(ISNUMBER(${cell}),${cell},"")`;
}

/**
 * The formula of B10: the average seconds of the runs in A16:B25 of the tab
 * Concentration.Exposure. The block keeps RUN_LIMIT runs at most, so the
 * average uses each recorded run when fewer than RUN_LIMIT exist. B10 is
 * empty when no run is recorded.
 */
function runAverageFormula() {
  const block = exposure(`B${RUN_HEADER_ROW + 1}:B${RUN_HEADER_ROW + RUN_LIMIT}`);
  return `=IF(COUNT(${block})=0,"",AVERAGE(${block}))`;
}

/**
 * The formula of a cell that shows one field of the equity block. The cell is
 * empty when the field holds no number, so a portfolio with no stock shows no
 * error.
 */
function equityFormula(name) {
  return `=IF(ISNUMBER(${equityCell(name)}),${equityCell(name)},"")`;
}

/**
 * The formula of A17: one text when the last good answer holds lines and no
 * equity block. The cell is empty before the first good refresh.
 */
function noStockFormula() {
  return `=IF(AND(ISNUMBER(${measureCell("lineCount")}),NOT(ISNUMBER(${equityCell("weight")}))),"Portfolio holds no securities.","")`;
}

/**
 * The names at the start of the LET of a formula that reads the holdings
 * that the route cannot look through. uid holds each position id of the last
 * answer, as a column, and uu holds the weight of each position in the
 * unseen block. The script writes the unseen block: the weight of the lines
 * of the class unknown that came through each position with no mix, and 0
 * for a position with a mix.
 */
function unseenNames() {
  const x = exposureColumns();
  const id = exposureColumn(x.unseenId);
  return `uid,IFNA(FILTER(${id},${id}<>""),""),
uu,IFNA(FILTER(${exposureColumn(x.unseenWeight)},${id}<>""),0),`;
}

/**
 * The formula of B4, under the status cell: the count of the holdings that
 * the route cannot look through and that have no mix, with their share of
 * the portfolio. The formula reads the weights of the unseen block. The cell
 * is empty when no such holding exists.
 */
function unseenNoteFormula() {
  const x = exposureColumns();
  return `=LET(uu,${exposureColumn(x.unseenWeight)},
n,COUNTIF(uu,">1E-12"),s,SUM(uu),
IF(n=0,"",n&IF(n=1," holding ("," holdings (")&IF(s<0.01,"less than 1%",TEXT(s,"0%"))&" of your portfolio) "&IF(n=1,"is a fund","are funds")&" not looked through."))`;
}

/**
 * The formula of the coverage label in the header of the block of the
 * security measures: the share of the portfolio outside the lines of the class
 * unknown. The cell is empty when the answer holds no unknownWeight.
 */
function coverageFormula() {
  const cell = measureCell("unknownWeight");
  return `=IF(ISNUMBER(${cell}),"These measures cover "&TEXT(1-${cell},"0.0%")&" of the portfolio.","")`;
}

/**
 * The formula of one row of the Composition block: the sum of the stock
 * weight of the rows of the stock part that meet the comparison with the
 * threshold. The comparison is `>=` or `<`. The stock part of the cap line
 * holds many securities, so it adds to the row `<` whatever its weight.
 */
function stockSumFormula(comparison) {
  const x = exposureColumns();
  const stock = exposureColumn(x.stock);
  const part = exposureColumn(x.part);
  const key = exposureColumn(x.key);
  const sum = (sign, keyRule) =>
    `SUMIFS(${stock},${part},"${STOCK_PART}",${stock},"${sign}"&$B$${THRESHOLD_ROW},${key},"${keyRule}")`;
  if (comparison === ">=") return `=${sum(">=", `<>${CAP_KEY}`)}`;
  return `=${sum("<", `<>${CAP_KEY}`)}+SUMIFS(${stock},${part},"${STOCK_PART}",${key},"${CAP_KEY}")`;
}

/**
 * The formula of the last row of the Composition block: the sum of the part
 * of each line that is not stock. The part is the weight minus the stock
 * weight, on the rows of the other part.
 */
function otherSumFormula() {
  const x = exposureColumns();
  const part = exposureColumn(x.part);
  return `=SUMIFS(${exposureColumn(x.weight)},${part},"${OTHER_PART}")-SUMIFS(${exposureColumn(x.stock)},${part},"${OTHER_PART}")`;
}

/**
 * The note under the total of the report spill.
 */
const SUM_NOTE =
  "The positions in a fund report can add up to more than 100% of the net assets of the fund. This total can then be above 100%.";

/**
 * The note under the header of the fund overlap list.
 */
const OVERLAP_NOTE = "Overlap is the part of the two funds that sits in the same securities.";

/**
 * The note above the header of the company table.
 */
const TRUST_NOTE =
  "A trust or a closed-end fund, such as SPY, GLD, or IBIT, is not a company. Other holdings shows it as one line.";

/**
 * The text of the section Holdings when the Holdings tab holds no
 * Symbol, Description, or Value column, and the name of a group of rows with
 * an empty Symbol and an empty Description.
 */
const HOLDINGS_COLUMNS_NOTE = "The Holdings tab needs the columns Symbol, Description, and Value.";
const NO_NAME = "No symbol or description";

/**
 * The formula of the report spill, in column A of row REPORT_ROW. The spill
 * holds these blocks from the top: the section Holdings, the header row
 * of the holdings chart, the band of CHART_BAND_ROWS blank rows in which the
 * holdings chart sits, the note on trusts, the header of the company table,
 * the company table, the header row of the company chart, the band of
 * CHART_BAND_ROWS blank rows in which the company chart sits, the fund
 * overlap list, the section of the funds not looked through, the section of
 * the other holdings, and the total. The header row of the holdings chart
 * holds HOLDINGS_HEADING in column A, and the header row of the company
 * chart holds the text of chartHeading in column A, so the anchor cells of
 * the hidden tab find them.
 *
 * Each section starts in column A. The report tab formats each column of
 * the spill, not each section, so each section puts a value in column D and
 * a share of the portfolio in column E. A title, a note, and a text that
 * replaces an empty list are in column A. The fund overlap list holds the two
 * funds in A and B, and the overlap, the count of shared securities, and the
 * weight of each fund in E to H. The section of the funds not looked through
 * holds the holding, the fund in the mix, the value, the share, and the note
 * in A, B, D, E, and F. The section of the other holdings holds the line,
 * the kind, the value, the share, and the funds that the line came from in
 * A, B, D, E, and F. The total of all lines holds its label in A, the value
 * in D, and the share in E. The columns between the text and column D are
 * empty.
 *
 * The section Holdings reads the Holdings tab alone, so it changes with
 * no refresh. Each Holdings row with a number in Value goes into the group of
 * its key: its Symbol, or its Description when the Symbol is empty, cut to 64
 * characters, as buildPositions makes it. EXACT compares the keys, so two
 * keys that differ in case give two groups, as in the script. A group row
 * holds the Description of the first row of the group, the Symbol, the count
 * of rows, the sum of the values, and the sum divided by the total value
 * cell. The rows sort by value, largest first. A Total row with the total
 * value cell ends the section.
 *
 * LN, LT, LW, LX, LP, and LK are the columns of the lines block: the name,
 * the ticker, the weight, the stock weight, the part name, and the key. LS
 * is the sources block, and LH is the row of the position ids. LS and LH end
 * at the column of position MAX_POSITIONS. f holds each id of the stock fund
 * block, or one empty string when the block is empty. The header of the
 * company table holds one column for each id of f, from column H.
 *
 * Each row of the lines block is one part of a line. pw is the weight of the
 * part: the stock weight on a row of the stock part, and the weight minus the
 * stock weight on a row of the other part. dw is the direct weight column of
 * the lines block: the part of pw that came through no fund. A company row
 * is a row of the stock part with a stock weight at or above the threshold,
 * whatever the class of the line. The source cells of that row are the stock
 * part of each source, so the Direct column and the fund columns add up to
 * the row. The row of the securities under the threshold follows the company
 * rows with the next rank, the count of the company rows plus 1. The stock
 * part of the cap line, the line with the key CAP_KEY, holds many
 * securities. It goes into that row whatever its weight, and the count of the
 * row then ends with "and more".
 *
 * The section of the funds not looked through gives one row to each holding
 * with no mix whose lines of the class unknown hold UNSEEN_FLOOR or more of
 * the portfolio, and one row to the rest of those holdings. Then it gives one
 * row to each fund of each mix, with the substitute mark and the entry date
 * of the mix, and one Not described row to each mix with a total under 100%.
 * A mix older than MIX_AGE_DAYS asks the person to check the fact sheet. MI,
 * MW, ME, MT, MP, and MS are the columns of the mix block. hn gives the
 * Description of the Holdings tab for a Symbol, or the text that it gets.
 * Each piece of the section has a seventh column. The text "x" in it marks a
 * row that the section drops.
 *
 * The script writes the rows of the other part that get a row of their own
 * into the own block, and the groups of the other rows of the other part
 * into the group block. A row of the other part gets a row of its own when
 * the class of its line is not stock and the line is a residual line, is
 * the cap line, holds a direct position, or has the class unknown. Each
 * other row of the other part goes into the group of its class. The own row
 * of a residual line names the fund ticker, so oname shows the Description
 * of that fund. The route gives the class trust to a direct line alone, and
 * the class cash also to a direct line of a money market fund. Each such
 * line that holds a direct position gets a row of its own with the kind of
 * its class. The line of a trust in a fund mix that holds a fund has no
 * direct weight, so it goes into the group trust. keep selects the rows of
 * the own block, and grp selects the rows of the group block. Both hold 1
 * and 0, not TRUE and FALSE, because SUM adds no TRUE in an array, and nown
 * and ng count the rows with SUM. The groups of the classes stock and fund
 * always get a row. Another group with a value under 100 goes into one row
 * of small holdings.
 *
 * The formula finds the Symbol, the Description, and the Value columns of the
 * Holdings tab by the header text in row 1, as the total value cell finds the
 * Value column. When the total value cell holds no number, each value cell
 * stays empty. Each FILTER that can find no row has a fallback in IFNA or
 * IFERROR, or an IF before it that uses the FILTER only when a row matches.
 */
function reportFormula() {
  const x = exposureColumns();
  const column = (letter) => exposureColumn(letter);
  const floor = UNSEEN_FLOOR;
  return `=LET(LN,${column(x.name)},LT,${column(x.ticker)},LW,${column(x.weight)},LX,${column(x.stock)},LP,${column(x.part)},LK,${column(x.key)},
LS,${exposure(`$${x.first}$${FIRST_DATA_ROW}:$${x.last}`)},LH,${exposure(`$${x.first}$${HEADER_ROW}:$${x.last}$${HEADER_ROW}`)},
f,IFNA(FILTER(${column(x.stockFund)},${column(x.stockFund)}<>""),""),
hh,Holdings!$1:$1,
hs,XMATCH(TRUE,ARRAYFORMULA(EXACT(TRIM(hh),"Symbol"))),
hd,XMATCH(TRUE,ARRAYFORMULA(EXACT(TRIM(hh),"Description"))),
hv,XMATCH(TRUE,ARRAYFORMULA(EXACT(TRIM(hh),"Value"))),
tot,IF(ISNUMBER($B$${TOTAL_ROW}),$B$${TOTAL_ROW},NA()),thr,$B$${THRESHOLD_ROW},
yours,IF(ISNA(hs)+ISNA(hd)+ISNA(hv)>0,"${HOLDINGS_COLUMNS_NOTE}",
 LET(yv,INDIRECT("Holdings!C"&hv,FALSE),
  ys,ARRAYFORMULA(TRIM(INDIRECT("Holdings!C"&hs,FALSE))),yd,ARRAYFORMULA(TRIM(INDIRECT("Holdings!C"&hd,FALSE))),
  yok,ARRAYFORMULA(ISNUMBER(yv)*1),
  IF(SUM(yok)=0,"No holding on the Holdings tab has a value.",
   LET(fk,FILTER(ARRAYFORMULA(LEFT(IF(ys<>"",ys,yd),64)),yok),fs,FILTER(ys,yok),fd,FILTER(yd,yok),fv,FILTER(yv,yok),
    yu,UNIQUE(fk),
    yi,MAP(yu,LAMBDA(k,XMATCH(TRUE,ARRAYFORMULA(EXACT(fk,k))))),
    yn,MAP(yu,yi,LAMBDA(k,i,IF(INDEX(fd,i)<>"",INDEX(fd,i),IF(k<>"",k,"${NO_NAME}")))),
    yt,MAP(yi,LAMBDA(i,INDEX(fs,i))),
    yc,MAP(yu,LAMBDA(k,SUMPRODUCT(EXACT(fk,k)*1))),
    yw,MAP(yu,LAMBDA(k,SUMPRODUCT(EXACT(fk,k)*fv))),
    VSTACK(SORT(HSTACK(yn,yt,yc,yw,ARRAYFORMULA(IFERROR(yw/tot,""))),4,FALSE),
     HSTACK("Total","","",tot,IFERROR(tot/tot,""))))))),
MI,${column(x.mixId)},MW,${column(x.mixWeight)},ME,${column(x.mixEntered)},
MT,${column(x.mixTicker)},MP,${column(x.mixPart)},MS,${column(x.mixSubstitute)},
hn,LAMBDA(n,IF(ISNA(hs)+ISNA(hd),n,IFNA(XLOOKUP(n,INDIRECT("Holdings!C"&hs,FALSE),INDIRECT("Holdings!C"&hd,FALSE)),n))),
pw,ARRAYFORMULA(IF(LP="${STOCK_PART}",LX,IF(LP="${OTHER_PART}",LW-LX,0))),
dw,${column(x.direct)},
sel,ARRAYFORMULA((LP="${STOCK_PART}")*ISNUMBER(LX)*(LX>=thr)*(LK<>"${CAP_KEY}")),
top,IF(SUM(sel)=0,HSTACK("","No company at or above the threshold."),
 LET(blk,SORT(HSTACK(FILTER(LN,sel),FILTER(LT,sel),FILTER(LX,sel),FILTER(dw,sel),FILTER(LS,sel)),3,FALSE),
  sw,CHOOSECOLS(blk,3),mx,MAX(sw),
  fx,MAKEARRAY(ROWS(sw),ROWS(f),LAMBDA(i,j,IFERROR(N(INDEX(blk,i,4+XMATCH(INDEX(f,j),LH))),0))),
  HSTACK(SEQUENCE(ROWS(sw)),CHOOSECOLS(blk,1,2),ARRAYFORMULA(sw*tot),sw,
   MAP(sw,LAMBDA(x,SPARKLINE(x,{"charttype","bar";"max",mx;"color1","#2a78d6"}))),
   ARRAYFORMULA(ROUND(CHOOSECOLS(blk,4),12)),fx))),
rsel,ARRAYFORMULA((LP="${STOCK_PART}")*ISNUMBER(LX)*((LX<thr)+(LK="${CAP_KEY}")>0)),
rw,SUM(ARRAYFORMULA(IF(rsel,LX,0))),
rc,COUNTIFS(LP,"${STOCK_PART}",LX,"<"&thr,LX,"<>0",LK,"<>${CAP_KEY}"),
rk,COUNTIFS(LP,"${STOCK_PART}",LK,"${CAP_KEY}"),
rf,MAP(f,LAMBDA(x,IFERROR(SUM(FILTER(INDEX(LS,0,XMATCH(x,LH)),rsel)),0))),
rest,HSTACK(SUM(sel)+1,"Securities under "&TEXT(thr,"0.00%")&" ("&TEXT(rc,"#,##0")&IF(rk>0," and more","")&")","",rw*tot,rw,"",ROUND(SUM(IFNA(FILTER(dw,rsel),0)),12),TRANSPOSE(rf)),
${unseenNames()}
cx,{"","","","","","","x"},
cu,LET(q,IFNA(FILTER(uid,uu>=${floor}),""),
 IF(INDEX(q,1,1)="",cx,
  LET(qw,IFNA(FILTER(uu,uu>=${floor}),0),qe,MAKEARRAY(ROWS(q),1,LAMBDA(i,j,"")),
   SORT(HSTACK(MAP(q,LAMBDA(v,hn(v))),qe,ARRAYFORMULA(qw*tot),qw,MAKEARRAY(ROWS(q),1,LAMBDA(i,j,"${ADD_MIX_NOTE}")),qe,qe),4,FALSE)))),
csn,SUM(ARRAYFORMULA(IF((uu>1E-12)*(uu<${floor}),1,0))),
csw,SUM(ARRAYFORMULA(IF(uu<${floor},uu,0))),
cs,IF(csn=0,cx,HSTACK("Holdings under ${floor * 100}% ("&csn&")","",csw*tot,csw,"${ADD_MIX_NOTE}","","")),
cdn,SUM(ARRAYFORMULA(IF(MI<>"",1,0))),
cd,IF(cdn=0,cx,
 LET(ki,IFNA(FILTER(MI,MI<>""),""),kw,IFNA(FILTER(MW,MI<>""),0),ke,IFNA(FILTER(ME,MI<>""),0),
  kt,IFNA(FILTER(MT,MI<>""),""),kp,IFNA(FILTER(MP,MI<>""),0),ks,IFNA(FILTER(MS,MI<>""),FALSE),
  kn,MAP(ke,LAMBDA(d,"Mix entered "&TEXT(d,"mmmm yyyy")&IF(TODAY()-d>${MIX_AGE_DAYS}," — check the fact sheet",""))),
  kr,HSTACK(MAP(ki,LAMBDA(v,hn(v))),ARRAYFORMULA(kt&IF(ks," (substitute)","")),ARRAYFORMULA(kw*kp*tot),ARRAYFORMULA(kw*kp),kn,ki,
   MAKEARRAY(ROWS(ki),1,LAMBDA(i,j,0))),
  ku,UNIQUE(ki),
  kv,MAP(ku,LAMBDA(v,XLOOKUP(v,ki,kw)*(1-SUM(IFNA(FILTER(kp,ki=v),0))))),
  ko,HSTACK(MAP(ku,LAMBDA(v,hn(v))),MAKEARRAY(ROWS(ku),1,LAMBDA(i,j,"Not described")),ARRAYFORMULA(kv*tot),kv,
   MAP(ku,LAMBDA(v,XLOOKUP(v,ki,kn))),ku,MAKEARRAY(ROWS(ku),1,LAMBDA(i,j,1))),
  SORT(VSTACK(kr,IFNA(FILTER(ko,kv>1E-9),cx)),6,TRUE,7,TRUE))),
cv,VSTACK(cu,cs,cd),
can,IFNA(FILTER(CHOOSECOLS(cv,1,2,3,4,5),CHOOSECOLS(cv,7)<>"x"),{"Every fund is looked through.","","","",""}),
cb,MAKEARRAY(ROWS(can),1,LAMBDA(i,j,"")),
keep,ARRAYFORMULA(ISNUMBER(${column(x.ownWeight)})*1),
nown,SUM(keep),
on,FILTER(${column(x.ownName)},keep),ok,FILTER(${column(x.ownKey)},keep),oc,FILTER(${column(x.ownClass)},keep),
ow,FILTER(${column(x.ownWeight)},keep),came,FILTER(${column(x.ownSources)},keep),
oname,IF(ISNA(hs)+ISNA(hd),on,
 LET(sc,INDIRECT("Holdings!C"&hs,FALSE),sy,ARRAYFORMULA(IF(ROW(sc)=1,"",sc)),de,INDIRECT("Holdings!C"&hd,FALSE),
  MAP(on,LAMBDA(n,IFNA(XLOOKUP(n,sy,de),n))))),
kind,MAP(ok,oc,on,LAMBDA(k,c,n,IF(LEFT(k,9)="residual:","cash and other net assets",IF(k="${CAP_KEY}","many small lines, combined",
 IF(RIGHT(n,16)=" (not described)","not described",
 SWITCH(c,"fund","fund, no holdings data","trust","trust or closed-end fund, no holdings data","cash","money market fund",
  "unknown","not in the SEC data","preferred","preferred shares",c)))))),
ownRows,HSTACK(oname,kind,ARRAYFORMULA(ow*tot),ow,came),
grp,ARRAYFORMULA(ISNUMBER(${column(x.groupWeight)})*1),
cls,FILTER(${column(x.groupClass)},grp),
gw,FILTER(${column(x.groupWeight)},grp),
gn,FILTER(${column(x.groupCount)},grp),
gl,MAP(cls,gn,LAMBDA(x,m,SWITCH(x,"stock","Bonds and other securities of companies whose stock you hold","fund","Funds held by the funds, not looked through","cash","Cash and money market funds inside funds","derivative","Derivatives inside funds","treasury","Treasury securities inside funds","preferred","Preferred shares inside funds",
 "trust","Trusts and closed-end funds inside funds","other","Other holdings inside funds",x&" inside funds")&" ("&TEXT(m,"#,##0")&")")),
gk,MAP(cls,LAMBDA(x,SWITCH(x,"stock","bond or other security","fund","fund, not looked through","preferred","preferred shares",
 "trust","trust or closed-end fund",x))),
gf,FILTER(${column(x.groupSources)},grp),
grpAll,HSTACK(gl,gk,ARRAYFORMULA(gw*tot),gw,gf),
big,ARRAYFORMULA(IF((cls="stock")+(cls="fund")>0,1,IF(ISNUMBER(gw*tot),IF(ABS(gw*tot)>=100,1,0),1))),
sm,ARRAYFORMULA(1-big),
ng,IF(SUM(grp)=0,0,SUM(big)),
nsm,IF(SUM(grp)=0,0,SUM(sm)),
smallRow,HSTACK("Other small holdings inside funds ("&TEXT(SUM(FILTER(gn,sm)),"#,##0")&")",
 TEXTJOIN(", ",TRUE,FILTER(gk,sm)),SUM(FILTER(gw,sm))*tot,SUM(FILTER(gw,sm)),
 IFERROR(TEXTJOIN(", ",TRUE,UNIQUE(TRANSPOSE(ARRAYFORMULA(TRIM(SPLIT(TEXTJOIN(",",TRUE,FILTER(gf,sm)),",")))))),"")),
main,IF(nown=0,FILTER(grpAll,big),IF(ng=0,ownRows,VSTACK(ownRows,FILTER(grpAll,big)))),
nis,IF(nown+ng=0,IF(nsm=0,{"No holding other than securities.","","","",""},smallRow),
 IF(nsm=0,SORT(main,4,FALSE),VSTACK(SORT(main,4,FALSE),smallRow))),
blank,MAKEARRAY(ROWS(nis),1,LAMBDA(i,j,"")),
pa,${column(x.pairFirst)},pb,${column(x.pairSecond)},po,${column(x.pairOverlap)},pn,${column(x.pairShared)},
fw,${column(x.fundWeight)},
pwt,LAMBDA(v,IFNA(XLOOKUP(v,MI,MW),XLOOKUP(v,${column(x.fundId)},fw,""))),
psel,ARRAYFORMULA(ISNUMBER(po)*(po>=$B$${OVERLAP_ROW})),
plist,IF(SUM(psel)=0,"No pair of funds is at or above the overlap minimum.",
 LET(qa,FILTER(pa,psel),qb,FILTER(pb,psel),gap,MAKEARRAY(ROWS(qa),1,LAMBDA(i,j,"")),
  HSTACK(qa,qb,gap,gap,FILTER(po,psel),FILTER(pn,psel),
   MAP(qa,LAMBDA(v,pwt(v))),MAP(qb,LAMBDA(v,pwt(v)))))),
IFNA(VSTACK("${HOLDINGS_TITLE}",
 {${HOLDING_COLUMNS.map((text) => `"${text}"`).join(",")}},
 yours,
 "",
 "${HOLDINGS_HEADING}",
 MAKEARRAY(${CHART_BAND_ROWS},1,LAMBDA(i,j,"")),
 "${TRUST_NOTE}",
 HSTACK({${COMPANY_COLUMNS.map((text) => `"${text}"`).join(",")}},TRANSPOSE(f)),
 top,rest,"",
 ${chartHeading("thr")},
 MAKEARRAY(${CHART_BAND_ROWS},1,LAMBDA(i,j,"")),
 "Fund overlap",
 "${OVERLAP_NOTE}",
 {"Fund 1","Fund 2","","","Overlap","Shared securities","Fund 1 weight","Fund 2 weight"},
 plist,
 "",
 "${UNSEEN_TITLE}",
 {"Holding","Fund in the mix","","Value","% of portfolio","Note"},
 HSTACK(CHOOSECOLS(can,1,2),cb,CHOOSECOLS(can,3,4,5)),
 "",
 "${OTHER_TITLE}",
 {"Line","Kind","","Value","% of portfolio","Came from"},
 HSTACK(CHOOSECOLS(nis,1,2),blank,CHOOSECOLS(nis,3,4,5)),
 "",
 HSTACK("Total of all lines","","",SUM(pw)*tot,SUM(pw)),
 "${SUM_NOTE}"),""))`;
}

/**
 * The text of the chart header row of the report spill, as an expression of
 * a formula. `threshold` is the expression of the threshold, such as a name
 * of a LET or a cell reference. The text holds the threshold as a percent
 * with CHART_HEADING.decimals decimals, so it changes with each edit of the
 * threshold cell.
 */
function chartHeading(threshold) {
  const format = `0.${"0".repeat(CHART_HEADING.decimals)}%`;
  return `"${CHART_HEADING.before}"&TEXT(${threshold},"${format}")&"${CHART_HEADING.after}"`;
}

/**
 * The text that the chart header row of the report spill shows for a
 * threshold, a share from 0 to 1. The percent has CHART_HEADING.decimals
 * decimals, and a half rounds up. The function first rounds the percent to
 * 15 significant digits, so a binary error such as 100.49999999999999 for
 * 100.5 does not change the last decimal.
 */
function headingText(threshold) {
  const scale = 10 ** CHART_HEADING.decimals;
  const percent = Math.round(Number((threshold * 100 * scale).toPrecision(15))) / scale;
  return `${CHART_HEADING.before}${percent.toFixed(CHART_HEADING.decimals)}%${CHART_HEADING.after}`;
}

/**
 * True when the text of the chart header row in the sheet and the text of
 * headingText are the same after each comma becomes a dot. TEXT in the spill
 * formula uses the decimal mark of the spreadsheet locale, so a locale with a
 * comma decimal mark shows 10,00% where headingText gives 10.00%.
 */
function sameHeading(shown, heading) {
  return shown.replaceAll(",", ".") === heading.replaceAll(",", ".");
}

/**
 * The formula of the anchor cell of the company chart, in the tab
 * Concentration.Exposure: the row of the report tab under the header row of
 * the company chart in the report spill, which is the first row of the band.
 * The formula finds the text of the chart header row in column A of the
 * report tab, with the threshold cell, and adds 1. The cell is empty when no
 * row holds the text.
 */
function anchorFormula() {
  const heading = chartHeading(`'${REPORT_TAB}'!$B$${THRESHOLD_ROW}`);
  return `=IFERROR(XMATCH(${heading},'${REPORT_TAB}'!$A:$A)+1,"")`;
}

/**
 * The formula of the holdings anchor cell of the tab Concentration.Exposure:
 * the row of the report tab under the header row of the holdings chart in
 * the report spill, which is the first row of its band. The formula finds
 * HOLDINGS_HEADING in column A of the report tab and adds 1. The cell is
 * empty when no row holds the text.
 */
function holdingsAnchorFormula() {
  return `=IFERROR(XMATCH("${HOLDINGS_HEADING}",'${REPORT_TAB}'!$A:$A)+1,"")`;
}

/**
 * The formulas of the company chart block of the tab Concentration.Exposure:
 * the header row, and the row that each of the CHART_ROWS data rows holds.
 * The columns are the company name, Direct, CHART_FUNDS fund columns, and
 * Other funds. The header cells hold the series names. A header cell that
 * holds an empty string marks a column with no series. writeChartBlock writes
 * the formulas.
 *
 * Each formula reads the company table of the report spill alone, so a
 * change of the threshold or of the Holdings tab changes the block with no
 * refresh. h is the row of the header of the company table inside the
 * spill: the row with Rank in column A and Company in column B. The company
 * rows under it hold the rank 1, 2, 3, and so on in column A, largest first.
 * The row of the securities under the threshold holds the next rank and is
 * not a company row. m counts the company rows, up to CHART_ROWS. A data
 * row shows the company of its rank r: its row minus HEADER_ROW. A data row
 * with no company of its rank is blank.
 *
 * g holds the fund columns of the company rows, and s holds the sum of each
 * fund column. The order o puts the funds with a sum other than 0 first,
 * largest sum first, and the order of the stock fund block decides a tie.
 * The fund column k of the block shows the fund at place k of o while k is
 * not above t, the count of the funds with a sum other than 0, up to
 * CHART_FUNDS. Its header is the id of the fund, cut to CHART_NAME_LENGTH
 * characters. The Other funds column adds each fund after place
 * CHART_FUNDS, and it is blank when nz, the count of the funds with a sum
 * other than 0, is not above CHART_FUNDS.
 *
 * The header of the Direct column is empty when the direct weights of the
 * company rows add up to 0. The header of the Other funds column is empty
 * when nz is not above CHART_FUNDS. Each header is empty when the company
 * table holds no company.
 */
function chartFormulas() {
  const spill = (letter) => `'${REPORT_TAB}'!$${letter}$${REPORT_ROW}:$${letter}`;
  const company = (name) => spill(columnLetter(COMPANY_COLUMNS.indexOf(name) + 1));
  const funds = `'${REPORT_TAB}'!$${columnLetter(COMPANY_COLUMNS.length + 1)}$${REPORT_ROW}:$${columnLetter(REPORT_COLUMNS)}`;
  const row = `r,ROW()-${HEADER_ROW},`;
  const find = `sa,${spill("A")},
h,XMATCH(1,ARRAYFORMULA((sa="${COMPANY_COLUMNS[0]}")*(${company("Company")}="Company"))),
w,CHOOSEROWS(sa,SEQUENCE(MIN(${CHART_ROWS},ROWS(sa)-h),1,h+1)),
wb,CHOOSEROWS(${company("Company")},SEQUENCE(ROWS(w),1,h+1)),
m,SUMPRODUCT((w=SEQUENCE(ROWS(w)))*(LEFT(wb,17)<>"Securities under ")*1),`;
  const rank = `g,CHOOSEROWS(${funds},SEQUENCE(m,1,h+1)),
s,BYCOL(g,LAMBDA(c,SUM(c))),
z,ARRAYFORMULA((s<>0)*1),
nz,SUM(z),
t,MIN(${CHART_FUNDS},nz),
n,COLUMNS(g),
o,SORT(SEQUENCE(n),TRANSPOSE(z),FALSE,TRANSPOSE(s),FALSE,SEQUENCE(n),TRUE),`;
  const ranks = Array.from({ length: CHART_FUNDS }, (_, i) => i + 1);
  const cell = (body) => `=IFERROR(LET(${body}),"")`;
  return {
    header: [
      `="Company"`,
      cell(`${find}
IF(m=0,"",IF(SUM(CHOOSEROWS(${company("Direct")},SEQUENCE(m,1,h+1)))=0,"","Direct"))`),
      ...ranks.map((k) =>
        cell(
          `${find}
IF(m=0,"",LET(${rank}
IF(${k}>t,"",LEFT(INDEX(${funds},h,INDEX(o,${k})),${CHART_NAME_LENGTH}))))`,
        ),
      ),
      cell(`${find}
IF(m=0,"",LET(${rank}
IF(nz<=${CHART_FUNDS},"","Other funds")))`),
    ],
    row: [
      cell(`${row}${find}
IF(r>m,"",INDEX(${company("Company")},h+r))`),
      cell(`${row}${find}
IF(r>m,"",INDEX(${company("Direct")},h+r))`),
      ...ranks.map((k) =>
        cell(`${row}${find}
IF(r>m,"",LET(${rank}
IF(${k}>t,"",INDEX(g,r,INDEX(o,${k})))))`),
      ),
      cell(`${row}${find}
IF(r>m,"",LET(${rank}
IF(nz<=${CHART_FUNDS},"",SUM(MAP(SEQUENCE(n),LAMBDA(j,IF(XMATCH(j,o)>${CHART_FUNDS},N(INDEX(g,r,j)),0)))))))`),
    ],
  };
}

/**
 * The formulas of the holdings chart block of the tab Concentration.Exposure:
 * the header row, and the row that each of the HOLDINGS_CHART_ROWS data rows
 * holds. The columns are the label of each holding group and its share of the
 * portfolio. The header cell of the share column holds the series name. When
 * it holds an empty string, the block has no series. writeChartBlock writes
 * the formulas.
 *
 * Each formula reads the section Holdings of the report spill alone,
 * so a change of the Holdings tab changes the block with no refresh. h is
 * the row of the header of the table of the section inside the spill: the
 * row with Holding in column A and Ticker in column B. The group rows under
 * it hold the count of accounts, a number, in column C, in the order of the
 * table, largest value first. The Total row and a note row hold no number
 * there. m counts the rows with a number in column C from row h + 1 up to
 * the first row with no number, HOLDINGS_CHART_ROWS at most. A data row
 * shows the group at place r of the table: its row minus HEADER_ROW. A data
 * row with no group at its place is blank.
 *
 * The label is the Ticker of the group, or the Holding cut to
 * CHART_NAME_LENGTH characters when the Ticker is empty. The share is the %
 * of portfolio of the group. The header of the share column is empty when
 * the table holds no group, or when no group has a share, as when the total
 * value cell holds no number.
 */
function holdingsChartFormulas() {
  const spill = (name) => {
    const letter = columnLetter(HOLDING_COLUMNS.indexOf(name) + 1);
    return `'${REPORT_TAB}'!$${letter}$${REPORT_ROW}:$${letter}`;
  };
  const share = spill("% of portfolio");
  const row = `r,ROW()-${HEADER_ROW},`;
  const find = `sa,${spill("Holding")},
h,XMATCH(1,ARRAYFORMULA((sa="${HOLDING_COLUMNS[0]}")*(${spill("Ticker")}="${HOLDING_COLUMNS[1]}"))),
w,CHOOSEROWS(${spill("Accounts")},SEQUENCE(MIN(${HOLDINGS_CHART_ROWS},ROWS(sa)-h),1,h+1)),
m,IFNA(XMATCH(0,ARRAYFORMULA(ISNUMBER(w)*1)),ROWS(w)+1)-1,`;
  const cell = (body) => `=IFERROR(LET(${body}),"")`;
  return {
    header: [
      `="${HOLDING_COLUMNS[0]}"`,
      cell(`${find}
IF(m=0,"",IF(COUNT(CHOOSEROWS(${share},SEQUENCE(m,1,h+1)))=0,"","${HOLDING_COLUMNS[4]}"))`),
    ],
    row: [
      cell(`${row}${find}
IF(r>m,"",IF(INDEX(${spill("Ticker")},h+r)<>"",INDEX(${spill("Ticker")},h+r),LEFT(INDEX(sa,h+r),${CHART_NAME_LENGTH})))`),
      cell(`${row}${find}
IF(r>m,"",INDEX(${share},h+r))`),
    ],
  };
}

/**
 * The color of the series of the column at an offset of the company chart
 * block.
 */
function seriesColor(offset) {
  if (offset === 1) return CHART_DIRECT_COLOR;
  if (offset <= 1 + CHART_FUNDS) return CHART_FUND_COLORS[offset - 2];
  return CHART_OTHER_COLOR;
}

/**
 * The two bar charts of the report tab. Each one names its block of the tab
 * Concentration.Exposure: the first column, the count of columns, and the
 * count of data rows under the header row HEADER_ROW. It also names the
 * column of its anchor cell, a name for the log, the color of the series of
 * each column offset of the block, the stacking, the position of the
 * legend, and the other options of setOption.
 *
 * The company chart stacks the series and puts the legend at the top. The
 * total data label of a stacked chart shows the share of each company at
 * the end of its bar. The value of a segment shows on hover alone.
 *
 * The holdings chart has one series and no legend. The data label of the
 * series shows the share of each holding group at the end of its bar, with
 * the number format of the block.
 */
function companyChart() {
  return {
    name: "company",
    column: CHART_COLUMN,
    width: CHART_BLOCK_WIDTH,
    rows: CHART_ROWS,
    anchor: ANCHOR_COLUMN,
    color: seriesColor,
    stacked: true,
    legend: "top",
    options: { "annotations.total.enabled": true },
  };
}

/**
 * The holdings chart. companyChart states the fields.
 */
function holdingsChart() {
  return {
    name: "holdings",
    column: HOLDINGS_CHART_COLUMN,
    width: HOLDINGS_BLOCK_WIDTH,
    rows: HOLDINGS_CHART_ROWS,
    anchor: HOLDINGS_ANCHOR_COLUMN,
    color: () => CHART_DIRECT_COLOR,
    stacked: false,
    legend: "none",
    options: { series: { 0: { dataLabel: "value", dataLabelPlacement: "outsideEnd", hasAnnotations: true } } },
  };
}

/**
 * Draw the company chart of the report tab again from the company chart
 * block. A good refresh and onEdit call this function. `threshold` is the
 * threshold that the report spill must show, a share from 0 to 1: the
 * refresh gives the value of readInputs, and onEdit gives the value of the
 * edit. The header row of the company chart holds the threshold, so the
 * function waits for the text of headingText of that threshold.
 */
function drawChart(book, threshold) {
  drawBlockChart(book, companyChart(), headingText(threshold));
}

/**
 * Draw the holdings chart of the report tab again from the holdings chart
 * block. A good refresh calls this function. The header row of the holdings
 * chart holds HOLDINGS_HEADING, so the function waits for that text.
 */
function drawHoldingsChart(book) {
  drawBlockChart(book, holdingsChart(), HOLDINGS_HEADING);
}

/**
 * Draw one bar chart of the report tab again from its block of the tab
 * Concentration.Exposure. `chart` is companyChart or holdingsChart, and
 * `heading` is the text that the header row of the chart in the report
 * spill must show. The function uses the spreadsheet alone, and it sends no
 * request.
 *
 * The function changes nothing when a tab is absent, or when the hidden tab
 * holds another layout version. Else it calls SpreadsheetApp.flush, so the
 * block and the anchor cell show the values of the current answer and the
 * current cells. Then freshAnchorRow reads the anchor cell of the chart and
 * the header row above it. The anchor cell is a formula on the report
 * spill, and a read right after an edit can run ahead of the calculation of
 * the spill and give an old row. So freshAnchorRow compares the header row
 * with `heading`, and waits and reads again while the two differ, for up to
 * CHART_WAIT_LIMIT milliseconds. When the text never matches, the function
 * changes no chart and logs one warning.
 *
 * Then the function reads the header row of the block. Each column of the
 * block after the first one whose header cell holds a text is a series. The
 * first column gives the bar labels. Each range of the chart runs from the
 * header row to the last data row of the block, so it never moves.
 *
 * The function builds the new chart first. Then it removes each chart of the
 * report tab that reads a column of this block, and it inserts the new
 * chart. So one chart of each block stays, the chart of the other block
 * stays, and an error of the build keeps the old chart. When no header cell
 * after the first one holds a text, the function removes the old chart of
 * the block and inserts none.
 *
 * The chart anchors at column A of the row in the anchor cell. It spans the
 * columns A to D, it is CHART_HEIGHT pixels high, and it has no title. The
 * axis shows the share of the portfolio with the number format 0%.
 */
function drawBlockChart(book, chart, heading) {
  const report = book.getSheetByName(REPORT_TAB);
  const hidden = book.getSheetByName(EXPOSURE_TAB);
  if (report === null || hidden === null || hidden.getRange(VERSION_CELL).getValue() !== LAYOUT_VERSION) return;
  SpreadsheetApp.flush();
  const row = freshAnchorRow(report, hidden, chart.anchor, heading);
  if (row === null) {
    console.warn(
      `The ${chart.name} chart did not change: no row of the report showed "${heading}" in ${CHART_WAIT_LIMIT / 1000} s.`,
    );
    return;
  }
  const header = hidden.getRange(HEADER_ROW, chart.column, 1, chart.width).getValues()[0];
  const series = header.map((_, offset) => offset).filter((offset) => offset > 0 && cellText(header[offset]) !== "");
  let built = null;
  if (series.length > 0) {
    const builder = report.newChart().asBarChart();
    for (const offset of [0, ...series]) {
      builder.addRange(hidden.getRange(HEADER_ROW, chart.column + offset, chart.rows + 1, 1));
    }
    builder.setNumHeaders(1);
    if (chart.stacked) builder.setStacked();
    builder.setColors(series.map(chart.color));
    builder.setPosition(row, 1, 0, 0);
    builder.setOption("useFirstColumnAsDomain", true);
    builder.setOption("legend", { position: chart.legend });
    builder.setOption("hAxis", { format: "0%" });
    for (const [name, value] of Object.entries(chart.options)) builder.setOption(name, value);
    builder.setOption("width", CHART_WIDTH);
    builder.setOption("height", CHART_HEIGHT);
    built = builder.build();
  }
  const ownBlock = (range) =>
    range.getSheet().getName() === EXPOSURE_TAB &&
    range.getColumn() >= chart.column &&
    range.getColumn() < chart.column + chart.width;
  for (const old of report.getCharts()) {
    if (old.getRanges().some(ownBlock)) report.removeChart(old);
  }
  if (built !== null) report.insertChart(built);
}

/**
 * The row in the anchor cell of the column `anchor` of the tab
 * Concentration.Exposure when column A of the report tab holds the text
 * `heading` in the row above it, as sameHeading compares them, else null.
 * The function reads the anchor cell and that cell of the report tab. While
 * the text differs, it waits CHART_WAIT_STEP milliseconds with
 * Utilities.sleep and reads the two cells again, until the total wait
 * reaches CHART_WAIT_LIMIT milliseconds.
 */
function freshAnchorRow(report, hidden, anchor, heading) {
  for (let waited = 0; ; waited += CHART_WAIT_STEP) {
    const row = hidden.getRange(FIRST_DATA_ROW, anchor).getValue();
    const fresh =
      Number.isInteger(row) &&
      row > 1 &&
      row <= report.getMaxRows() &&
      sameHeading(cellText(report.getRange(row - 1, 1).getValue()), heading);
    if (fresh) return row;
    if (waited >= CHART_WAIT_LIMIT) return null;
    Utilities.sleep(CHART_WAIT_STEP);
  }
}

/**
 * The layout of the tab Concentration.Exposure. The cells hold the layout
 * version and the labels. The refresh writes the status, the time, the
 * answer, the mix block, the unseen block, the stock fund block, the own
 * block, the group block, and the direct weights. A good refresh also writes
 * its start time and its seconds into the run-time block A16:B25, newest
 * first. The grid holds the source column of each position up to
 * MAX_POSITIONS, then the columns of the company chart block, the anchor
 * column of the company chart, the columns of the holdings chart block, and
 * the anchor column of the holdings chart. One narrow empty column
 * separates each block from the next. The layout holds the column widths
 * and the formats of the two chart blocks and the labels of the two anchor
 * cells. writeChartBlock writes the formulas of the blocks and of the anchor
 * cells.
 */
function exposureLayout() {
  const x = exposureColumns();
  const letter = columnLetter;
  const span = (first, count) => `${letter(first)}:${letter(first + count - 1)}`;
  const header = (first, count) => `${letter(first)}${HEADER_ROW}:${letter(first + count - 1)}${HEADER_ROW}`;
  const data = (name) => `${name}${FIRST_DATA_ROW}:${name}`;
  const fund = (name) => letter(FUND_COLUMN + FUND_FIELDS.indexOf(name));
  const lastMeasure = FIRST_DATA_ROW + MEASURE_NAMES.length - 1;
  const runFirst = RUN_HEADER_ROW + 1;
  const runLast = RUN_HEADER_ROW + RUN_LIMIT;
  const equityFirst = EQUITY_HEADER_ROW + 1;
  const equityLast = EQUITY_HEADER_ROW + EQUITY_NAMES.length;
  const chartLast = CHART_COLUMN + CHART_BLOCK_WIDTH - 1;
  const holdingsShare = letter(HOLDINGS_CHART_COLUMN + 1);
  return {
    name: EXPOSURE_TAB,
    rows: TAB_ROWS,
    columns: HOLDINGS_ANCHOR_COLUMN,
    hidden: true,
    frozenRows: 4,
    columnWidths: {
      A: 170,
      B: 150,
      C: 24,
      [span(FUND_COLUMN, FUND_FIELDS.length)]: 110,
      [letter(OVERLAP_COLUMN - 1)]: 24,
      [span(OVERLAP_COLUMN, OVERLAP_WIDTH)]: 110,
      [letter(MIX_COLUMN - 1)]: 24,
      [span(MIX_COLUMN, MIX_FIELDS.length)]: 110,
      [letter(UNSEEN_COLUMN - 1)]: 24,
      [span(UNSEEN_COLUMN, UNSEEN_FIELDS.length)]: 110,
      [letter(STOCK_FUND_COLUMN - 1)]: 24,
      [letter(STOCK_FUND_COLUMN)]: 110,
      [letter(OWN_COLUMN - 1)]: 24,
      [span(OWN_COLUMN, OWN_FIELDS.length)]: 110,
      [letter(GROUP_COLUMN - 1)]: 24,
      [span(GROUP_COLUMN, GROUP_FIELDS.length)]: 110,
      [letter(LINE_COLUMN - 1)]: 24,
      [letter(LINE_COLUMN)]: 220,
      [letter(LINE_COLUMN + 1)]: 260,
      [`${letter(LINE_COLUMN + 2)}:${x.direct}`]: 90,
      [`${x.first}:${x.last}`]: 110,
      [letter(CHART_COLUMN - 1)]: 24,
      [letter(CHART_COLUMN)]: 220,
      [span(CHART_COLUMN + 1, CHART_BLOCK_WIDTH - 1)]: 90,
      [letter(ANCHOR_COLUMN - 1)]: 24,
      [letter(ANCHOR_COLUMN)]: 110,
      [letter(HOLDINGS_CHART_COLUMN - 1)]: 24,
      [letter(HOLDINGS_CHART_COLUMN)]: 220,
      [span(HOLDINGS_CHART_COLUMN + 1, HOLDINGS_BLOCK_WIDTH - 1)]: 90,
      [letter(HOLDINGS_ANCHOR_COLUMN - 1)]: 24,
      [letter(HOLDINGS_ANCHOR_COLUMN)]: 110,
    },
    cells: [
      { range: "A1:A2", values: [["Status"], ["Last run"]] },
      { range: `A3:${VERSION_CELL}`, values: [["layoutVersion", LAYOUT_VERSION]] },
      { range: `A4:A${lastMeasure}`, values: [["measure"], ...MEASURE_NAMES.map((name) => [name])] },
      { range: "B4", values: [["value"]] },
      { range: header(FUND_COLUMN, FUND_FIELDS.length), values: [FUND_FIELDS] },
      {
        range: header(OVERLAP_COLUMN, OVERLAP_WIDTH),
        values: [["firstId", "secondId", "overlap", "sharedLineCount"]],
      },
      { range: header(MIX_COLUMN, MIX_FIELDS.length), values: [MIX_FIELDS] },
      { range: header(UNSEEN_COLUMN, UNSEEN_FIELDS.length), values: [UNSEEN_FIELDS] },
      { range: header(STOCK_FUND_COLUMN, STOCK_FUND_FIELDS.length), values: [STOCK_FUND_FIELDS] },
      { range: header(OWN_COLUMN, OWN_FIELDS.length), values: [OWN_FIELDS] },
      { range: header(GROUP_COLUMN, GROUP_FIELDS.length), values: [GROUP_FIELDS] },
      { range: header(LINE_COLUMN, LINE_FIELDS.length + 2), values: [[...LINE_FIELDS, "part", "directWeight"]] },
      { range: `A${RUN_HEADER_ROW}:B${RUN_HEADER_ROW}`, values: [["runStart", "seconds"]] },
      { range: `A${EQUITY_HEADER_ROW}:B${EQUITY_HEADER_ROW}`, values: [["equity", "value"]] },
      { range: `A${equityFirst}:A${equityLast}`, values: EQUITY_NAMES.map((name) => [name]) },
      { range: header(ANCHOR_COLUMN, 1), values: [["chartRow"]] },
      { range: header(HOLDINGS_ANCHOR_COLUMN, 1), values: [["holdingsChartRow"]] },
    ],
    styles: [
      { range: `A4:${x.last}4`, bold: true, background: "#f7f6f1" },
      { range: header(CHART_COLUMN, CHART_BLOCK_WIDTH), bold: true, background: "#f7f6f1" },
      {
        range: `${letter(CHART_COLUMN + 1)}${FIRST_DATA_ROW}:${letter(chartLast)}${HEADER_ROW + CHART_ROWS}`,
        numberFormat: "0.00%",
      },
      { range: header(ANCHOR_COLUMN, 1), bold: true, background: "#f7f6f1" },
      { range: header(HOLDINGS_CHART_COLUMN, HOLDINGS_BLOCK_WIDTH), bold: true, background: "#f7f6f1" },
      {
        range: `${holdingsShare}${FIRST_DATA_ROW}:${holdingsShare}${HEADER_ROW + HOLDINGS_CHART_ROWS}`,
        numberFormat: "0.00%",
      },
      { range: header(HOLDINGS_ANCHOR_COLUMN, 1), bold: true, background: "#f7f6f1" },
      { range: `A${RUN_HEADER_ROW}:B${RUN_HEADER_ROW}`, bold: true, background: "#f7f6f1" },
      { range: `A${EQUITY_HEADER_ROW}:B${EQUITY_HEADER_ROW}`, bold: true, background: "#f7f6f1" },
      { range: `A${runFirst}:A${runLast}`, numberFormat: "yyyy-mm-dd hh:mm:ss" },
      { range: `B${runFirst}:B${runLast}`, numberFormat: "0.000" },
      { range: "B2", numberFormat: "yyyy-mm-dd hh:mm:ss" },
      { range: "B5", numberFormat: "#,##0" },
      { range: "B6", numberFormat: "0.000000" },
      { range: "B7", numberFormat: "0.0" },
      { range: "B8", numberFormat: "0.00" },
      { range: `B9:B${lastMeasure}`, numberFormat: "0.000000" },
      { range: `B${equityFirst}`, numberFormat: "0.000000" },
      { range: `B${equityFirst + 1}`, numberFormat: "#,##0" },
      { range: `B${equityFirst + 2}`, numberFormat: "0.000000" },
      { range: `B${equityFirst + 3}`, numberFormat: "0.0" },
      { range: `B${equityLast}`, numberFormat: "0.00" },
      { range: data(fund("reportDate")), numberFormat: "yyyy-mm-dd" },
      { range: data(fund("holdingCount")), numberFormat: "#,##0" },
      { range: `${fund("weight")}${FIRST_DATA_ROW}:${fund("coveredWeight")}`, numberFormat: "0.00000" },
      { range: `${fund("mergedByTicker")}${FIRST_DATA_ROW}:${fund("mergedByName")}`, numberFormat: "#,##0" },
      { range: data(x.fundPart), numberFormat: "0.00000" },
      { range: data(x.pairOverlap), numberFormat: "0.00000" },
      { range: data(x.pairShared), numberFormat: "#,##0" },
      { range: data(x.mixWeight), numberFormat: "0.00000" },
      { range: data(x.mixEntered), numberFormat: "yyyy-mm-dd" },
      { range: data(x.mixPart), numberFormat: "0.00000" },
      { range: data(x.unseenWeight), numberFormat: "0.00000" },
      { range: data(x.ownWeight), numberFormat: "0.00000" },
      { range: data(x.groupWeight), numberFormat: "0.00000" },
      { range: data(x.groupCount), numberFormat: "#,##0" },
      { range: `${x.weight}5:${x.stock}`, numberFormat: "0.00000" },
      { range: data(x.direct), numberFormat: "0.00000" },
      { range: `${x.first}5:${x.last}`, numberFormat: "0.00000" },
    ],
    conditional: [],
    validation: [],
  };
}

/**
 * The layout of the tab Concentration. `inputs` holds the value of the
 * threshold cell and of the overlap minimum cell, by the names of INPUTS. An
 * absent name gets the value of INPUTS. A person can change both cells. B1
 * holds the disclaimer. B4, under the status cell, counts the holdings that
 * are funds not looked through. B9 and B10 show the seconds of the last good
 * refresh and the average of the recorded refreshes. The header of the
 * security measures carries the coverage label.
 *
 * The report spill starts in row REPORT_ROW, and its row count changes with
 * the Holdings tab. So the conditional formats find the titles, the chart
 * header rows, the headers, the totals, and the notes of the spill by their
 * text. The number formats and the alignment of the columns apply from row
 * REPORT_ROW to the last row.
 */
function reportLayout(inputs = {}) {
  const threshold = inputs.threshold === undefined ? INPUTS.threshold.value : inputs.threshold;
  const overlapMinimum = inputs.overlapMinimum === undefined ? INPUTS.overlapMinimum.value : inputs.overlapMinimum;
  const last = columnLetter(REPORT_COLUMNS);
  const first = REPORT_ROW;
  const spill = `A${first}:${last}`;
  const status = STATUS_ROW;
  const total = `$B$${TOTAL_ROW}`;
  const stocks = STOCKS_ROW;
  const comp = COMPOSITION_ROW;
  const fund = FUND_ROW;
  return {
    name: REPORT_TAB,
    rows: TAB_ROWS,
    columns: REPORT_COLUMNS,
    hidden: false,
    frozenRows: 2,
    columnWidths: { ...REPORT_WIDTHS, [`G:${last}`]: 96 },
    cells: [
      { range: "A1:B1", values: [["Concentration", DISCLAIMER]] },
      { range: "B2", values: [["Securities by company, with a look inside each fund."]] },
      {
        range: "D2",
        values: [
          [
            "A fund that holds other funds uses the newest report of each held fund. The date of a held report can differ from the date in this table.",
          ],
        ],
      },
      {
        range: `A${status}:B${status + 7}`,
        values: [
          ["Status", `=${exposure("B1")}`],
          ["", unseenNoteFormula()],
          ["Last run", `=${exposure("B2")}`],
          ["Total value", TOTAL_FORMULA],
          ["Looked through", `=${measureCell("lookedThroughWeight")}`],
          ["Not looked through", `=${measureCell("notLookedThroughWeight")}`],
          ["Last run time", runLastFormula()],
          ["Average (last 10)", runAverageFormula()],
        ],
      },
      {
        range: `D${fund - 1}:H${fund - 1}`,
        values: [["Fund looked through", "Report date", "Holdings", "Weight", "Covered"]],
      },
      { range: `D${fund}`, values: [[fundsFormula()]] },
      {
        range: `A${stocks}:B${stocks + 5}`,
        values: [
          ["Securities alone", coverageFormula()],
          ["Securities, share of the portfolio", equityFormula("weight")],
          ["Top 10 securities, share of the securities", equityFormula("top10Weight")],
          ["HHI of the securities, 0 to 10,000", equityFormula("hhi")],
          ["Effective number of securities", equityFormula("effectiveCount")],
          [noStockFormula(), ""],
        ],
      },
      {
        range: `A${comp}:C${comp + 3}`,
        values: [
          ["Composition", "Value", "% of portfolio"],
          [
            `="Securities at "&TEXT($B$${THRESHOLD_ROW},"0.00%")&" or more"`,
            `=IF(ISNUMBER(${total}),C${comp + 1}*${total},"")`,
            stockSumFormula(">="),
          ],
          [
            `="Securities under "&TEXT($B$${THRESHOLD_ROW},"0.00%")`,
            `=IF(ISNUMBER(${total}),C${comp + 2}*${total},"")`,
            stockSumFormula("<"),
          ],
          [OTHER_TITLE, `=IF(ISNUMBER(${total}),C${comp + 3}*${total},"")`, otherSumFormula()],
        ],
      },
      {
        range: `A${THRESHOLD_ROW}:C${OVERLAP_ROW}`,
        values: [
          [
            INPUTS.threshold.label,
            threshold,
            "Type a percent. Each company whose securities are at or above it gets a row.",
          ],
          [
            INPUTS.overlapMinimum.label,
            overlapMinimum,
            "Type a percent. The fund overlap list under the company table shows each pair of funds at or above it.",
          ],
        ],
      },
      { range: `A${first}`, values: [[reportFormula()]] },
    ],
    styles: [
      { range: "A1", bold: true, fontSize: 16 },
      { range: "B1:B2", color: "#6b6962", italic: true },
      { range: "D2", color: "#6b6962", italic: true },
      { range: `A${status}:A${comp + 3}`, color: "#57554f" },
      { range: `B${status + 2}`, numberFormat: "yyyy-mm-dd hh:mm" },
      { range: `B${TOTAL_ROW}`, numberFormat: "$#,##0", bold: true },
      { range: `B${status + 4}:B${status + 5}`, numberFormat: "0.00%" },
      { range: `B${status + 6}:B${status + 7}`, numberFormat: '0.0" s"' },
      { range: `B${status}:B${status + 7}`, align: "right" },
      { range: `B${status + 1}`, align: "left", color: "#6b6962", italic: true, wrap: true },
      { range: `A${stocks}:B${stocks}`, bold: true, background: "#f7f6f1", color: "#1d1c1a" },
      { range: `B${stocks}`, bold: false, italic: true, color: "#57554f" },
      { range: `B${stocks + 1}:B${stocks + 2}`, numberFormat: "0.00%" },
      { range: `B${stocks + 3}`, numberFormat: "#,##0" },
      { range: `B${stocks + 4}`, numberFormat: "0.0" },
      { range: `A${stocks + 5}`, color: "#6b6962", italic: true },
      { range: `A${comp}:C${comp}`, bold: true, background: "#f7f6f1", color: "#1d1c1a" },
      { range: `B${comp}:C${comp}`, align: "right" },
      { range: `B${comp + 1}:B${comp + 3}`, numberFormat: "$#,##0" },
      { range: `C${comp + 1}:C${comp + 3}`, numberFormat: "0.00%" },
      { range: `A${THRESHOLD_ROW}:A${OVERLAP_ROW}`, bold: true },
      {
        range: `B${THRESHOLD_ROW}:B${OVERLAP_ROW}`,
        numberFormat: "0.00%",
        bold: true,
        background: "#fff4c7",
        align: "right",
      },
      { range: `C${THRESHOLD_ROW}:C${OVERLAP_ROW}`, color: "#6b6962", italic: true },
      { range: `A${first}:A`, numberFormat: "0" },
      { range: `D${first}:D`, numberFormat: "$#,##0", align: "right" },
      { range: `E${first}:E`, numberFormat: "0.00%", align: "right" },
      { range: `F${first}:F`, numberFormat: "#,##0", wrap: true },
      { range: `G${first}:${last}`, numberFormat: '0.00%;-0.00%;""', align: "right" },
      { range: `D${fund - 1}:H${fund - 1}`, bold: true, background: "#f7f6f1" },
      { range: `D${fund}:D${FUND_LAST_ROW}`, bold: true },
      { range: `E${fund}:E${FUND_LAST_ROW}`, numberFormat: "yyyy-mm-dd" },
      { range: `F${fund}:F${FUND_LAST_ROW}`, numberFormat: "#,##0" },
      { range: `G${fund}:G${FUND_LAST_ROW}`, numberFormat: "0.00%" },
      { range: `H${fund}:H${FUND_LAST_ROW}`, numberFormat: "0.0%" },
    ],
    conditional: [
      { range: `B${status}`, formula: `=$B$${status}="OK"`, background: "#dcefe2", color: "#1b5e34", bold: true },
      { range: `B${status}`, formula: `=$B$${status}<>"OK"`, background: "#f7d4d4", color: "#8a1c1c", bold: true },
      {
        range: spill,
        formula: `=LEFT($B${first},17)="Securities under "`,
        color: "#3c3b37",
        italic: true,
      },
      {
        range: spill,
        formula: `=OR($A${first}="${HOLDINGS_TITLE}",$A${first}="${HOLDINGS_HEADING}",$A${first}=${chartHeading(`$B$${THRESHOLD_ROW}`)},$A${first}="Fund overlap",$A${first}="${UNSEEN_TITLE}",$A${first}="${OTHER_TITLE}")`,
        bold: true,
      },
      {
        range: spill,
        formula:
          `=OR(AND($A${first}="Holding",$B${first}="Ticker"),AND($A${first}="Rank",$B${first}="Company"),` +
          `AND($A${first}="Fund 1",$B${first}="Fund 2"),AND($A${first}="Holding",$B${first}="Fund in the mix"),AND($A${first}="Line",$B${first}="Kind"))`,
        bold: true,
        background: "#f7f6f1",
      },
      {
        range: spill,
        formula: `=OR(AND($A${first}="Total",$B${first}="",$C${first}="",$F${first}=""),$A${first}="Total of all lines")`,
        bold: true,
      },
      {
        range: spill,
        formula: `=OR($A${first}="${TRUST_NOTE}",$A${first}="${OVERLAP_NOTE}",$A${first}="${SUM_NOTE}")`,
        color: "#6b6962",
        italic: true,
      },
      {
        range: `F${first}:F`,
        formula: `=RIGHT($F${first},20)="check the fact sheet"`,
        color: "#8a1c1c",
      },
      {
        range: `B${first}:B`,
        formula: `=AND(ISNUMBER($A${first}),LEFT($B${first},17)<>"Securities under ",N($G${first})>0,SUM($H${first}:$${last}${first})>0)`,
        bold: true,
      },
    ],
    validation: [
      { range: `B${THRESHOLD_ROW}`, min: 0, max: 1, message: "Type a percent from 0% to 100%, such as 1%." },
      { range: `B${OVERLAP_ROW}`, min: 0, max: 1, message: "Type a percent from 0% to 100%, such as 10%." },
    ],
  };
}

/**
 * The value of each cell of a report tab that the person types in, by the
 * names of INPUTS. The function finds a cell by its label in column A, so it
 * also reads a tab of another layout version. A name gets the value of INPUTS
 * when the tab is absent, when no row holds the label, or when the cell holds
 * no number from 0 to 1.
 */
function readInputs(report) {
  const rows =
    report === null || report.getMaxColumns() < 2
      ? []
      : report.getRange(1, 1, Math.min(INPUT_ROWS, report.getMaxRows()), 2).getValues();
  const inputs = {};
  for (const [name, input] of Object.entries(INPUTS)) {
    const at = labelIndex(rows, input.label);
    const typed = at < 0 ? null : rows[at][1];
    inputs[name] = isNumber(typed) && typed >= 0 && typed <= 1 ? typed : input.value;
  }
  return inputs;
}

/**
 * The index of the first row whose cell in column A holds the label, or -1.
 * `rows` holds the top rows of a report tab from column A. readInputs and
 * onEdit find the cell of an input by its label with this function.
 */
function labelIndex(rows, label) {
  return rows.findIndex((cells) => cellText(cells[0]) === label);
}

/**
 * Make sure that the two tabs of the report hold the layout of this file.
 *
 * The function reads the layout version of Concentration.Exposure. When the
 * version equals LAYOUT_VERSION and Concentration exists, both tabs stay as
 * they are. When the version equals LAYOUT_VERSION and Concentration is
 * absent, the function creates Concentration. In each other condition, such
 * as an older version, no version, or no hidden tab, the function deletes
 * each of the two tabs that exists and creates both from the layout. A new
 * Concentration tab gets the threshold and the overlap minimum of the old
 * one.
 *
 * The function creates Concentration.Exposure first, because the formulas of
 * Concentration read it. It hides Concentration.Exposure after it creates
 * Concentration, because a spreadsheet must keep one visible tab through
 * each step. It changes no other tab.
 *
 * The last step writes the chart blocks and the anchor cells of
 * Concentration.Exposure, each time the function replaced a tab. Their
 * formulas read Concentration. A formula that reads a tab which the same
 * run deletes and creates again stays stale, so the function writes them
 * after both tabs exist. When Concentration alone is absent, the formulas
 * of the current hidden tab read a deleted tab, so the function writes them
 * again.
 */
function ensureTabs(book) {
  const hidden = book.getSheetByName(EXPOSURE_TAB);
  const report = book.getSheetByName(REPORT_TAB);
  const current = hidden !== null && hidden.getRange(VERSION_CELL).getValue() === LAYOUT_VERSION;
  if (current && report !== null) return;
  const inputs = readInputs(report);
  const exposureTab = current ? hidden : replaceTab(book, hidden, exposureLayout());
  replaceTab(book, report, reportLayout(inputs));
  if (!current) exposureTab.hideSheet();
  writeChartBlock(exposureTab);
}

/**
 * Write the formulas of the two chart blocks and of the two anchor cells
 * into the tab Concentration.Exposure, with four calls. The first call
 * writes the company chart block: the header row HEADER_ROW and the
 * CHART_ROWS data rows under it, from the column CHART_COLUMN. The second
 * call writes the anchor cell of the company chart, in the row
 * FIRST_DATA_ROW of the column ANCHOR_COLUMN. The third call writes the
 * holdings chart block: the header row and the HOLDINGS_CHART_ROWS data rows
 * under it, from the column HOLDINGS_CHART_COLUMN. The fourth call writes
 * the holdings anchor cell, in the row FIRST_DATA_ROW of the column
 * HOLDINGS_ANCHOR_COLUMN. The formulas read the report tab, so ensureTabs
 * calls this function after the report tab exists. The layout of the hidden
 * tab holds the column widths, the formats, and the labels of the anchor
 * cells.
 */
function writeChartBlock(hidden) {
  const chart = chartFormulas();
  const rows = [chart.header, ...Array.from({ length: CHART_ROWS }, () => chart.row)];
  hidden.getRange(HEADER_ROW, CHART_COLUMN, CHART_ROWS + 1, CHART_BLOCK_WIDTH).setValues(rows);
  hidden.getRange(FIRST_DATA_ROW, ANCHOR_COLUMN).setValues([[anchorFormula()]]);
  const holdings = holdingsChartFormulas();
  const holdingRows = [holdings.header, ...Array.from({ length: HOLDINGS_CHART_ROWS }, () => holdings.row)];
  hidden
    .getRange(HEADER_ROW, HOLDINGS_CHART_COLUMN, HOLDINGS_CHART_ROWS + 1, HOLDINGS_BLOCK_WIDTH)
    .setValues(holdingRows);
  hidden.getRange(FIRST_DATA_ROW, HOLDINGS_ANCHOR_COLUMN).setValues([[holdingsAnchorFormula()]]);
}

/**
 * Create a tab from its layout in the place of an old tab, and return the
 * new tab. The function deletes the old tab first, and the new tab takes its
 * position. When the old tab is null, the new tab goes to the end of the
 * spreadsheet.
 */
function replaceTab(book, old, layout) {
  let index = book.getNumSheets();
  if (old !== null) {
    index = old.getIndex() - 1;
    book.deleteSheet(old);
  }
  return createTab(book, layout, index);
}

/**
 * Add a tab at the given position of the spreadsheet, apply its layout, and
 * return the tab. The layout holds the grid size, the cells, the styles, the
 * column widths, the frozen rows, the conditional formats, and the data
 * validation. The caller hides a tab with the hidden flag.
 */
function createTab(book, layout, index) {
  const sheet = book.insertSheet(layout.name, index);
  sizeGrid(sheet, layout.rows, layout.columns);
  for (const cell of layout.cells) sheet.getRange(cell.range).setValues(cell.values);
  for (const style of layout.styles) applyStyle(sheet.getRange(style.range), style);
  for (const [columns, width] of Object.entries(layout.columnWidths)) {
    const [first, last] = columns.split(":");
    const start = columnNumber(first);
    sheet.setColumnWidths(start, columnNumber(last || first) - start + 1, width);
  }
  sheet.setFrozenRows(layout.frozenRows);
  sheet.setConditionalFormatRules(layout.conditional.map((rule) => conditionalRule(sheet, rule)));
  for (const check of layout.validation) {
    const rule = SpreadsheetApp.newDataValidation()
      .requireNumberBetween(check.min, check.max)
      .setAllowInvalid(false)
      .setHelpText(check.message)
      .build();
    sheet.getRange(check.range).setDataValidation(rule);
  }
  return sheet;
}

/**
 * Add or delete rows and columns at the end of the grid until the grid holds
 * the given count of rows and columns.
 */
function sizeGrid(sheet, rows, columns) {
  const haveRows = sheet.getMaxRows();
  if (rows > haveRows) sheet.insertRowsAfter(haveRows, rows - haveRows);
  if (rows < haveRows) sheet.deleteRows(rows + 1, haveRows - rows);
  const haveColumns = sheet.getMaxColumns();
  if (columns > haveColumns) sheet.insertColumnsAfter(haveColumns, columns - haveColumns);
  if (columns < haveColumns) sheet.deleteColumns(columns + 1, haveColumns - columns);
}

/**
 * The column number of a column letter, such as 1 for A and 27 for AA.
 */
function columnNumber(letters) {
  return [...letters].reduce((n, c) => n * 26 + c.charCodeAt(0) - 64, 0);
}

/**
 * The column letter of a column number, such as A for 1 and AA for 27.
 */
function columnLetter(number) {
  let letters = "";
  for (let n = number; n > 0; n = Math.floor((n - 1) / 26)) {
    letters = String.fromCharCode(65 + ((n - 1) % 26)) + letters;
  }
  return letters;
}

/**
 * Apply one style entry of a layout to a range. The keys are bold, italic,
 * fontSize, color, background, numberFormat, align, and wrap. The function
 * changes only the properties that the entry names.
 */
function applyStyle(range, style) {
  if (style.bold !== undefined) range.setFontWeight(style.bold ? "bold" : "normal");
  if (style.italic !== undefined) range.setFontStyle(style.italic ? "italic" : "normal");
  if (style.fontSize !== undefined) range.setFontSize(style.fontSize);
  if (style.color !== undefined) range.setFontColor(style.color);
  if (style.background !== undefined) range.setBackground(style.background);
  if (style.numberFormat !== undefined) range.setNumberFormat(style.numberFormat);
  if (style.align !== undefined) range.setHorizontalAlignment(style.align);
  if (style.wrap !== undefined) range.setWrap(style.wrap);
}

/**
 * The conditional format rule of one entry of a layout: a custom formula
 * and the style of a cell that meets it.
 */
function conditionalRule(sheet, rule) {
  const builder = SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied(rule.formula)
    .setRanges([sheet.getRange(rule.range)]);
  if (rule.background !== undefined) builder.setBackground(rule.background);
  if (rule.color !== undefined) builder.setFontColor(rule.color);
  if (rule.bold !== undefined) builder.setBold(rule.bold);
  if (rule.italic !== undefined) builder.setItalic(rule.italic);
  return builder.build();
}
