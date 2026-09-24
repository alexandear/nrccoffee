/** Fetches per run, so one run stays well inside the 6 minute limit. */
const MAX_FETCHES = 60;

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Кава')
    .addItem('Заповнити з посилань', 'fillActiveSheet')
    .addItem('Оновити статистику MH', 'fillStats')
    .addItem('Графік по місяцях', 'buildMonthlyChart')
    .addItem('Впорядкувати аркуші', 'sortSheets')
    .addToUi();
}

/**
 * Fills "Назва", "Обсмажка", "Помел" and "Ціна, грн" on the active sheet
 * from every MadHeads cart link in the "Посилання..." column.
 */
function fillActiveSheet() {
  const sheet = SpreadsheetApp.getActiveSheet();
  const values = sheet.getDataRange().getValues();

  const headerRow = findHeaderRow(values);
  const columns = headerRow < 0 ? null : findColumns(values[headerRow]);
  if (!columns) {
    SpreadsheetApp.getUi().alert('На аркуші "' + sheet.getName() + '" немає потрібних колонок.');
    return;
  }

  const cache = {};
  const errors = [];
  let fetches = 0;
  let filled = 0;
  let foreign = 0;

  for (let row = headerRow + 1; row < values.length; row++) {
    const link = String(values[row][columns.link]).trim();
    if (!madHeadsSlug(link)) {
      if (link) foreign++;
      continue;
    }

    try {
      if (!cache[link]) {
        if (fetches >= MAX_FETCHES) break;
        fetches++;
        cache[link] = madHeadsItem(link);
      }
      const item = cache[link];
      // One cell at a time, so that columns this script does not own keep their formulas.
      sheet.getRange(row + 1, columns.name + 1).setValue(item.name);
      sheet.getRange(row + 1, columns.roast + 1).setValue(item.roast);
      sheet.getRange(row + 1, columns.grind + 1).setValue(item.grind);
      sheet.getRange(row + 1, columns.price + 1).setValue(item.price);
      filled++;
    } catch (error) {
      errors.push('рядок ' + (row + 1) + ': ' + error.message);
    }
  }

  if (!filled && foreign) {
    SpreadsheetApp.getUi().alert('Посилань MadHeads не знайдено. Інші магазини поки не підтримуються.');
    return;
  }

  const report = 'Заповнено рядків: ' + filled +
    (foreign ? '\nПропущено посилань інших магазинів: ' + foreign : '');
  SpreadsheetApp.getUi().alert(errors.length ? report + '\n\nПомилки:\n' + errors.join('\n') : report);
}

function findHeaderRow(values) {
  for (let row = 0; row < Math.min(values.length, 10); row++) {
    if (values[row].filter(isLinkHeader).length) return row;
  }
  return -1;
}

/** Column indexes by header text, or null when the sheet uses a different layout. */
function findColumns(header) {
  const columns = {
    link: indexOfHeader(header, isLinkHeader),
    name: indexOfHeader(header, function (cell) { return cell === 'назва'; }),
    roast: indexOfHeader(header, function (cell) { return cell.indexOf('обсмаж') === 0; }),
    grind: indexOfHeader(header, function (cell) { return cell === 'помел'; }),
    price: indexOfHeader(header, function (cell) { return cell.indexOf('ціна, грн') === 0; }),
  };
  const found = Object.keys(columns).filter(function (key) { return columns[key] >= 0; });
  return found.length === Object.keys(columns).length ? columns : null;
}

function isLinkHeader(cell) {
  return normalize(cell).indexOf('посилання') === 0;
}

function indexOfHeader(header, matches) {
  for (let i = 0; i < header.length; i++) {
    if (matches(normalize(header[i]))) return i;
  }
  return -1;
}

function normalize(cell) {
  return String(cell).trim().toLowerCase();
}

const STATS_SHEET = 'Статистика MH';

/** Sheet name of an order, e.g. "21.12.2025 ✅" or " 09.04.2026 MH ✅". */
const SHEET_DATE = /(\d{1,2})\.(\d{1,2})\.(\d{4})/;

/** A sheet of the other shop, e.g. "27.11.2025 3champs ✅" or "08.09.2026 (3chmp) ✅". */
const OTHER_SHOP_SHEET = /3\s*ch/i;

/** Appends a row per order sheet that "Статистика MH" does not list yet, oldest first. */
function fillStats() {
  const stats = SpreadsheetApp.getActive().getSheetByName(STATS_SHEET);
  if (!stats) {
    SpreadsheetApp.getUi().alert('Аркуша "' + STATS_SHEET + '" немає.');
    return;
  }

  const listed = {};
  const dates = stats.getRange(1, 1, stats.getLastRow(), 1).getValues();
  for (let row = 0; row < dates.length; row++) {
    listed[normalizeDate(dates[row][0])] = true;
  }

  // Two sheets can carry the same date, and the statistics keeps one row per date.
  const missing = {};
  const sheets = SpreadsheetApp.getActive().getSheets();
  for (let i = 0; i < sheets.length; i++) {
    const name = sheets[i].getName();
    const date = sheetDate(name);
    if (!date || listed[date] || OTHER_SHOP_SHEET.test(name)) continue;
    const order = summarizeOrder(sheets[i]);
    if (!order) continue;
    missing[date] = missing[date] ? addOrders(missing[date], order) : order;
  }

  const ordered = Object.keys(missing).sort(function (a, b) {
    return sortableDate(a) < sortableDate(b) ? -1 : 1;
  });

  let row = stats.getLastRow();
  for (let i = 0; i < ordered.length; i++) {
    const order = missing[ordered[i]];
    row++;
    // One cell at a time, so that columns this script does not own keep their formulas.
    // "Дата" is day.month.year text, and a written date is parsed in the spreadsheet's locale,
    // which reads 05.12.2022 as the fifth of December or as the twelfth of May depending on it.
    stats.getRange(row, 1).setNumberFormat('@').setValue(ordered[i]);
    stats.getRange(row, 2).setValue(order.coffee);
    stats.getRange(row, 3).setValue(order.drip);
    stats.getRange(row, 4).setValue(order.tonic);
    stats.getRange(row, 5).setValue(order.black);
    stats.getRange(row, 6).setValue(order.filter);
    stats.getRange(row, 7).setValue(order.sum);
  }

  SpreadsheetApp.getUi().alert('Додано рядків: ' + ordered.length);
}

/** Counts per category and the order total, or null when the sheet is not a MadHeads order. */
function summarizeOrder(sheet) {
  const values = sheet.getDataRange().getValues();
  const headerRow = findHeaderRow(values);
  if (headerRow < 0) return null;

  const header = values[headerRow];
  const link = indexOfHeader(header, isLinkHeader);
  const name = indexOfHeader(header, function (cell) { return cell === 'назва'; });
  const quantity = indexOfHeader(header, function (cell) { return cell === 'кількість'; });
  const sum = indexOfHeader(header, function (cell) { return cell === 'сума'; });
  if (quantity < 0 || sum < 0) return null;

  const order = { coffee: 0, drip: 0, tonic: 0, black: 0, filter: 0, sum: 0 };
  let counted = 0;
  let foreign = 0;
  for (let row = headerRow + 1; row < values.length; row++) {
    const url = link < 0 ? '' : String(values[row][link]).trim();
    const title = name < 0 ? '' : String(values[row][name]).trim();
    if (typeof values[row][quantity] !== 'number' || (!title && !url)) continue;
    // Rows are classified by their link: one sheet mixes both shops under a MadHeads header.
    if (isOtherShop(url)) {
      foreign++;
      continue;
    }
    order[madHeadsCategory(title, url)] += values[row][quantity];
    if (typeof values[row][sum] === 'number') order.sum += values[row][sum];
    counted++;
  }
  if (!counted) return null;

  // "Всього грн" is a label with the hand-kept total under it, and it covers the whole sheet.
  const total = foreign ? null : labelledValue(values, 'всього грн');
  if (total !== null) order.sum = total;
  return order;
}

/** True for a link to a shop that is not MadHeads; a bare product name is not one. */
function isOtherShop(url) {
  return /^https?:\/\//.test(url) && url.indexOf('madheadscoffee.com') < 0;
}

/** Value in the cell under the first cell holding this label, or null. */
function labelledValue(values, label) {
  for (let row = 0; row + 1 < values.length; row++) {
    for (let column = 0; column < values[row].length; column++) {
      if (normalize(values[row][column]).replace(',', '') !== label) continue;
      const below = values[row + 1][column];
      if (typeof below === 'number') return below;
    }
  }
  return null;
}

function sheetDate(sheetName) {
  const matched = SHEET_DATE.exec(sheetName);
  return matched ? pad(matched[1]) + '.' + pad(matched[2]) + '.' + matched[3] : '';
}

/** "Дата" is kept as text, but a sheet whose column is dated returns a Date. */
function normalizeDate(cell) {
  if (cell instanceof Date) {
    return pad(cell.getDate()) + '.' + pad(cell.getMonth() + 1) + '.' + cell.getFullYear();
  }
  return sheetDate(String(cell));
}

function sortableDate(date) {
  return date.slice(6) + date.slice(3, 5) + date.slice(0, 2);
}

function pad(number) {
  return String(number).length < 2 ? '0' + number : String(number);
}

function addOrders(order, other) {
  Object.keys(order).forEach(function (key) { order[key] += other[key]; });
  return order;
}

const MONTHLY_CHART_TITLE = 'Замовлення Madheads за місяць';

/** First column of the monthly table, right of the "Всього..." totals. */
const MONTHLY_COLUMN = 15;

/** Rebuilds the monthly table and the chart over it, below the charts already on the sheet. */
function buildMonthlyChart() {
  const stats = SpreadsheetApp.getActive().getSheetByName(STATS_SHEET);
  if (!stats || stats.getLastRow() < 2) {
    SpreadsheetApp.getUi().alert('Аркуш "' + STATS_SHEET + '" порожній.');
    return;
  }

  const months = monthlyTotals(stats.getRange(2, 1, stats.getLastRow() - 1, 7).getValues());
  writeMonthlyTable(stats, months);
  replaceMonthlyChart(stats, months.length);
  SpreadsheetApp.getUi().alert('Місяців у графіку: ' + months.length);
}

/** Totals per calendar month, oldest first, with a month without orders kept as a zero. */
function monthlyTotals(values) {
  const sums = {};
  for (let row = 0; row < values.length; row++) {
    const date = normalizeDate(values[row][0]);
    if (!date) continue;
    const month = date.slice(3);
    if (!sums[month]) sums[month] = { month: month, coffee: 0, drip: 0, sum: 0 };
    sums[month].coffee += number(values[row][1]);
    sums[month].drip += number(values[row][2]);
    sums[month].sum += number(values[row][6]);
  }

  const months = Object.keys(sums).sort(function (a, b) { return sortableMonth(a) < sortableMonth(b) ? -1 : 1; });
  if (!months.length) return [];

  const totals = [];
  for (let month = months[0]; ; month = nextMonth(month)) {
    totals.push(sums[month] || { month: month, coffee: 0, drip: 0, sum: 0 });
    if (month === months[months.length - 1]) return totals;
  }
}

/** The table is the script's own block, so it is written in one go. */
function writeMonthlyTable(sheet, months) {
  const rows = [['Місяць', 'Пачок кави', 'Пачок дріпів', 'Сума, грн']];
  for (let i = 0; i < months.length; i++) {
    rows.push([months[i].month, months[i].coffee, months[i].drip, months[i].sum]);
  }
  sheet.getRange(1, MONTHLY_COLUMN, sheet.getMaxRows(), rows[0].length).clearContent();
  // "12.2025" is a number to a spreadsheet, and a date to some locales, so the column stays text.
  sheet.getRange(1, MONTHLY_COLUMN, sheet.getMaxRows(), 1).setNumberFormat('@');
  sheet.getRange(1, MONTHLY_COLUMN, rows.length, rows[0].length).setValues(rows);
}

function replaceMonthlyChart(sheet, months) {
  const charts = sheet.getCharts();
  let anchorRow = 2;
  let anchorColumn = 9;
  for (let i = 0; i < charts.length; i++) {
    if (String(charts[i].getOptions().get('title')) === MONTHLY_CHART_TITLE) {
      sheet.removeChart(charts[i]);
      continue;
    }
    // Roughly a chart's height in rows: setPosition takes rows, and the chart's own height is in pixels.
    const position = charts[i].getContainerInfo();
    anchorRow = Math.max(anchorRow, position.getAnchorRow() + 22);
    anchorColumn = position.getAnchorColumn();
  }

  const chart = sheet.newChart()
    .setChartType(Charts.ChartType.COMBO)
    .addRange(sheet.getRange(1, MONTHLY_COLUMN, months + 1, 4))
    .setNumHeaders(1)
    .setPosition(anchorRow, anchorColumn, 0, 0)
    .setOption('title', MONTHLY_CHART_TITLE)
    .setOption('width', 900)
    .setOption('height', 400)
    .setOption('isStacked', true)
    .setOption('series', {
      0: { type: 'bars', targetAxisIndex: 0 },
      1: { type: 'bars', targetAxisIndex: 0 },
      2: { type: 'line', targetAxisIndex: 1 },
    })
    .setOption('vAxes', { 0: { title: 'Пачок' }, 1: { title: 'Сума, грн' } })
    .setOption('hAxis', { title: 'Місяць', slantedText: true })
    .setOption('trendlines', { 2: { type: 'linear', lineWidth: 1, opacity: 0.3 } })
    .build();
  sheet.insertChart(chart);
}

function nextMonth(month) {
  const value = Number(month.slice(0, 2));
  return value < 12 ? pad(value + 1) + month.slice(2) : '01.' + (Number(month.slice(3)) + 1);
}

function sortableMonth(month) {
  return month.slice(3) + month.slice(0, 2);
}

function number(cell) {
  return typeof cell === 'number' ? cell : 0;
}

/** Orders the tabs newest first, with the sheets that are not orders kept in front. */
function sortSheets() {
  const spreadsheet = SpreadsheetApp.getActive();
  const sheets = spreadsheet.getSheets().sort(byDateDescending);
  const active = spreadsheet.getActiveSheet();

  let moved = 0;
  for (let i = 0; i < sheets.length; i++) {
    if (spreadsheet.getSheets()[i].getSheetId() === sheets[i].getSheetId()) continue;
    // A hidden sheet cannot be activated, and moving one is only possible through activating it.
    const hidden = sheets[i].isSheetHidden();
    if (hidden) sheets[i].showSheet();
    spreadsheet.setActiveSheet(sheets[i]);
    spreadsheet.moveActiveSheet(i + 1);
    if (hidden) sheets[i].hideSheet();
    moved++;
  }

  spreadsheet.setActiveSheet(active);
  SpreadsheetApp.getUi().alert('Переміщено аркушів: ' + moved);
}

function byDateDescending(first, second) {
  const left = sheetDate(first.getName());
  const right = sheetDate(second.getName());
  if (!left || !right) return left === right ? 0 : (left ? 1 : -1);
  if (sortableDate(left) === sortableDate(right)) return 0;
  return sortableDate(left) > sortableDate(right) ? -1 : 1;
}
