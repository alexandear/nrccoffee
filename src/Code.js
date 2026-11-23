/** Fetches per run, so one run stays well inside the 6 minute limit. */
const MAX_FETCHES = 60;

/** Sheet row of the first order line: rows 2 and 3 sit under the header but are not orders. */
const FIRST_ORDER_ROW = 4;

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Кава')
    .addItem('Заповнити з посилань', 'fillActiveSheet')
    .addSubMenu(SpreadsheetApp.getUi()
      .createMenu('Оновити статистику MH')
      .addItem('2026', 'fillStats2026')
      .addItem('2025', 'fillStats2025')
      .addItem('2024', 'fillStats2024')
      .addItem('2023', 'fillStats2023')
      .addItem('2022', 'fillStats2022'))
    .addItem('Впорядкувати аркуші', 'sortSheets')
    .addToUi();
}

/**
 * Fills the empty "Назва", "Обсмажка", "Помел" and "Ціна, грн" cells on the active sheet
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
  const warnings = [];
  const prices = [];
  let fetches = 0;
  let filled = 0;
  let madHeads = 0;
  const foreign = [];

  for (let row = Math.max(headerRow + 1, FIRST_ORDER_ROW - 1); row < values.length; row++) {
    const link = String(values[row][columns.link]).trim();
    if (!madHeadsSlug(link)) {
      if (link) foreign.push(row + 1);
      continue;
    }
    madHeads++;

    try {
      if (!cache[link]) {
        if (fetches >= MAX_FETCHES) break;
        fetches++;
        cache[link] = madHeadsItem(link);
      }
      const item = cache[link];
      const differs = mismatches(values[row], columns, item);
      if (differs.length) {
        warnings.push('рядок ' + (row + 1) + ': ' + differs.join(', '));
        continue;
      }
      const line = row + 1;
      // The price is reported rather than refreshed, because the cell may be the price that was paid.
      const price = String(values[row][columns.price]).trim();
      if (price && Number(price) !== item.price) {
        prices.push('рядок ' + line + ': ' + price + ' ≠ ' + item.price);
      }

      // A cell someone filled in is kept, so only the blanks are written, and one at a time,
      // so that columns this script does not own keep their formulas.
      const blanks = [[columns.name, item.name], [columns.roast, item.roast],
        [columns.grind, item.grind], [columns.price, item.price]];
      let wrote = 0;
      for (let i = 0; i < blanks.length; i++) {
        if (String(values[row][blanks[i][0]]).trim()) continue;
        sheet.getRange(line, blanks[i][0] + 1).setValue(blanks[i][1]);
        wrote++;
      }
      if (wrote) filled++;
    } catch (error) {
      errors.push('рядок ' + (row + 1) + ': ' + error.message);
    }
  }

  if (!madHeads && foreign.length) {
    SpreadsheetApp.getUi().alert('Посилань MadHeads не знайдено. Інші магазини поки не підтримуються.');
    return;
  }

  const report = 'Заповнено рядків: ' + filled +
    (foreign.length ? '\nПропущено посилань інших магазинів, рядки: ' + foreign.join(', ') : '') +
    (prices.length ? '\n\nЦіна змінилася, комірку не змінено:\n' + prices.join('\n') : '') +
    (warnings.length ? '\n\nНе збігається з посиланням, рядок не змінено:\n' + warnings.join('\n') : '') +
    (errors.length ? '\n\nПомилки:\n' + errors.join('\n') : '');
  SpreadsheetApp.getUi().alert(report);
}

/**
 * Filled-in cells that disagree with the variant the link points at, e.g. 'Помел "V-60" ≠ "у зерні"'.
 * The link is what gets ordered, so a row like that is a typo in one or the other, and only a person can tell which.
 */
function mismatches(row, columns, item) {
  const checked = [['Назва', columns.name, item.name], ['Обсмажка', columns.roast, item.roast],
    ['Помел', columns.grind, item.grind]];
  const differs = [];
  for (let i = 0; i < checked.length; i++) {
    const current = String(row[checked[i][1]]).trim();
    if (isDripMark(current, item)) continue;
    if (current && !sameWords(current, checked[i][2])) {
      differs.push(checked[i][0] + ' "' + current + '" ≠ "' + checked[i][2] + '"');
    }
  }
  return differs;
}

/** Older sheets mark a drip's roast and grind as "дріпи", where the shop has a roast and no grind. */
function isDripMark(cell, item) {
  return normalize(cell) === 'дріпи' && item.grind === '-';
}

/** True for the same words in any order: older sheets put "Easy Drip" first, the shop puts it last. */
function sameWords(first, second) {
  return words(first) === words(second);
}

function words(text) {
  // "Easy Drip  KENYA" -> "drip easy kenya"
  return normalize(text).split(/\s+/).sort().join(' ');
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

/** Category of each "грн" column of "Статистика MH", from column H on, and its header. */
const STATS_MONEY = [
  ['coffee', 'Кава, грн'],
  ['drip', 'Дріпи, грн'],
  ['tonic', 'Coffee Tonic, грн'],
  ['black', 'Black Coffee, грн'],
  ['filter', 'Фільтри V60, грн'],
];
const STATS_MONEY_COLUMN = 8;

// A menu item calls a function by name, without arguments, so each year needs its own.
// One year at a time, because every order sheet together does not fit in the 6 minute limit.
function fillStats2026() { fillStats(2026); }
function fillStats2025() { fillStats(2025); }
function fillStats2024() { fillStats(2024); }
function fillStats2023() { fillStats(2023); }
function fillStats2022() { fillStats(2022); }

/**
 * Adds a row per order sheet of the year that "Статистика MH" does not list yet, in its place among the dates,
 * and recounts the rows it already lists. Without a year it reads every order sheet.
 */
function fillStats(year) {
  const stats = SpreadsheetApp.getActive().getSheetByName(STATS_SHEET);
  if (!stats) {
    SpreadsheetApp.getUi().alert('Аркуша "' + STATS_SHEET + '" немає.');
    return;
  }

  // Two sheets can carry the same date, and the statistics keeps one row per date.
  const orders = {};
  const sheets = SpreadsheetApp.getActive().getSheets();
  for (let i = 0; i < sheets.length; i++) {
    const name = sheets[i].getName();
    const date = sheetDate(name);
    if (!date || OTHER_SHOP_SHEET.test(name)) continue;
    if (year && date.slice(6) !== String(year)) continue;
    const order = summarizeOrder(sheets[i]);
    if (!order) continue;
    orders[date] = orders[date] ? addOrders(orders[date], order) : order;
  }

  const lastColumn = STATS_MONEY_COLUMN + STATS_MONEY.length - 1;
  const values = stats.getRange(1, 1, stats.getLastRow(), lastColumn).getValues();
  for (let i = 0; i < STATS_MONEY.length; i++) {
    if (values[0][STATS_MONEY_COLUMN - 1 + i] === '') {
      stats.getRange(1, STATS_MONEY_COLUMN + i).setValue(STATS_MONEY[i][1]);
    }
  }

  const listed = {};
  let updated = 0;
  for (let row = 0; row < values.length; row++) {
    const date = normalizeDate(values[row][0]);
    listed[date] = true;
    if (!date || (year && date.slice(6) !== String(year))) continue;
    // Only the cells that differ, so that a formula which already gives the right value stays.
    const counts = orders[date] ? statsRow(orders[date]) : zeroFilled(values[row]);
    let written = 0;
    for (let i = 0; i < counts.length; i++) {
      if (counts[i] === null || values[row][i + 1] === counts[i]) continue;
      stats.getRange(row + 1, i + 2).setValue(counts[i]);
      written++;
    }
    if (written) updated++;
  }

  const missing = Object.keys(orders).filter(function (date) { return !listed[date]; });
  missing.sort(function (a, b) { return sortableDate(a) < sortableDate(b) ? -1 : 1; });

  for (let i = 0; i < missing.length; i++) {
    const row = statsRowFor(stats, values, missing[i]);
    const counts = statsRow(orders[missing[i]]);
    values[row - 1] = [missing[i]].concat(counts);
    // One cell at a time, so that columns this script does not own keep their formulas.
    // "Дата" is day.month.year text, and a written date is parsed in the spreadsheet's locale,
    // which reads 05.12.2022 as the fifth of December or as the twelfth of May depending on it.
    stats.getRange(row, 1).setNumberFormat('@').setValue(missing[i]);
    for (let j = 0; j < counts.length; j++) {
      if (counts[j] !== null) stats.getRange(row, j + 2).setValue(counts[j]);
    }
  }

  // An old sheet whose rows do not add up leaves cells for a person to fill.
  const unknown = Object.keys(orders).filter(function (date) { return statsRow(orders[date]).indexOf(null) >= 0; });
  unknown.sort(function (a, b) { return sortableDate(a) < sortableDate(b) ? -1 : 1; });
  SpreadsheetApp.getUi().alert('Додано рядків: ' + missing.length + '\nОновлено рядків: ' + updated +
    (unknown.length ? '\nНе вдалося порахувати, заповніть вручну: ' + unknown.join(', ') : ''));
}

/**
 * Sheet row for a date the statistics does not list: the row after the latest earlier date when that row is empty,
 * as in the empty rows kept for a year, or a row inserted there, which moves the later dates down.
 * values mirrors columns A:L and is kept in step with an inserted row.
 */
function statsRowFor(stats, values, date) {
  // Without an earlier date, the new row goes above the first date, under the header and the totals.
  let previous = -1;
  let first = -1;
  for (let row = 0; row < values.length; row++) {
    const listed = normalizeDate(values[row][0]);
    if (!listed) continue;
    if (first < 0) first = row;
    if (sortableDate(listed) < sortableDate(date)) previous = row;
  }
  if (previous < 0) previous = first < 0 ? values.length - 1 : first - 1;

  const next = previous + 1;
  if (next >= values.length || isEmptyRow(values[next])) return next + 1;
  stats.insertRowAfter(previous + 1);
  values.splice(next, 0, []);
  return next + 1;
}

function isEmptyRow(row) {
  return row.every(function (cell) { return cell === ''; });
}

/**
 * Columns B:L of a row with no order sheet to recount it from, such as a sheet without "Сума":
 * an empty count is 0, and so is the "грн" of a category with none, but an unknown "грн" stays empty.
 */
function zeroFilled(row) {
  const cells = row.slice(1, STATS_MONEY_COLUMN + STATS_MONEY.length - 1);
  for (let i = 0; i < STATS_MONEY.length; i++) {
    const money = STATS_MONEY_COLUMN - 2 + i;
    if (cells[i] === '') cells[i] = 0;
    if (cells[i] === 0 && cells[money] === '') cells[money] = 0;
  }
  return cells;
}

/** Columns B:L of "Статистика MH" for one order: the counts, the total, then the "грн" of each category. */
function statsRow(order) {
  const row = [order.coffee, order.drip, order.tonic, order.black, order.filter, order.sum];
  for (let i = 0; i < STATS_MONEY.length; i++) {
    row.push(order[STATS_MONEY[i][0] + 'Sum']);
  }
  return row;
}

/**
 * Counts and "Сума" per category and the order total, or null when the sheet is not a MadHeads order.
 * The order total can be less than its categories together: "Всього грн" holds the price after a discount.
 */
function summarizeOrder(sheet) {
  const values = sheet.getDataRange().getValues();
  const headerRow = findHeaderRow(values);
  if (headerRow < 0) return null;

  const header = values[headerRow];
  const link = indexOfHeader(header, isLinkHeader);
  const name = indexOfHeader(header, function (cell) { return cell === 'назва'; });
  const quantity = indexOfHeader(header, function (cell) { return cell === 'кількість'; });
  const sum = indexOfHeader(header, function (cell) { return cell === 'сума'; });
  if (quantity < 0 || sum < 0) return summarizeOldOrder(values, headerRow);

  const order = {
    coffee: 0, drip: 0, tonic: 0, black: 0, filter: 0, sum: 0,
    coffeeSum: 0, dripSum: 0, tonicSum: 0, blackSum: 0, filterSum: 0,
  };
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
    const category = madHeadsCategory(title, url);
    order[category] += values[row][quantity];
    if (typeof values[row][sum] === 'number') {
      order.sum += values[row][sum];
      order[category + 'Sum'] += values[row][sum];
    }
    counted++;
  }
  if (!counted) return null;

  // "Всього грн" is a label with the hand-kept total under it, and it covers the whole sheet.
  const total = foreign ? null : labelledValue(values, 'всього грн');
  if (total !== null) order.sum = total;
  return order;
}

/** A drip in a sheet older than April 2024, e.g. "MIX BOX", "Cyato Box", "Дріпи мікс" or ".../izidrip/mix-box/". */
const OLD_DRIP = /izidrip|\bbox\b|дріп/i;

/** Labels of the counts in a sheet older than April 2024, each with its number in the cell under it. */
const OLD_COUNT_LABELS = {
  coffee: ['пачок кави'],
  drip: ['пачок дріпів'],
  tonic: ['coffee tonic'],
  black: ['blackcoffee', 'black coffee'],
  filter: [],
};

/**
 * Counts and "Сума" per category of a sheet older than April 2024, which has no "Сума" column, or null.
 * Its counts and total are labelled cells, and each row holds its own price, so a category's "Сума" is the sum of
 * those prices. A value that cannot be told is null: a count with no label when rows of its category exist,
 * and every "Сума" when the prices do not add up to the total or the drips found do not match "Пачок дріпів".
 */
function summarizeOldOrder(values, headerRow) {
  const total = labelledValue(values, 'всього грн');
  const sum = total !== null ? total : labelledValue(values, 'всього');
  if (sum === null) return null;

  const header = values[headerRow];
  const link = indexOfHeader(header, isLinkHeader);
  const name = indexOfHeader(header, function (cell) { return cell === 'назва'; });
  const note = indexOfHeader(header, function (cell) { return cell === 'примітки'; });
  const price = indexOfHeader(header, function (cell) { return cell === 'ціна, грн' || cell === 'ціна'; });
  const dripMark = indexOfHeader(header, function (cell) { return cell === 'пачок дріпів'; });

  const lines = [];
  // A bare link is the same product as that link on a row that names it.
  const dripLinks = {};
  for (let row = headerRow + 1; price >= 0 && row < values.length; row++) {
    const url = link < 0 ? '' : String(values[row][link]).trim();
    const text = [name, note].map(function (column) {
      return column < 0 ? '' : String(values[row][column]).trim();
    }).join(' ').trim();
    if (typeof values[row][price] !== 'number' || (!text && !url)) continue;
    let category = madHeadsCategory(text, url);
    if (category === 'coffee' && (OLD_DRIP.test(text + ' ' + url) || values[row][dripMark] === 1)) category = 'drip';
    if (category === 'drip' && url) dripLinks[url] = true;
    lines.push({ category: category, url: url, price: values[row][price] });
  }

  const sums = { coffee: 0, drip: 0, tonic: 0, black: 0, filter: 0 };
  const rows = { coffee: 0, drip: 0, tonic: 0, black: 0, filter: 0 };
  let priced = 0;
  lines.forEach(function (line) {
    const category = line.category === 'coffee' && dripLinks[line.url] ? 'drip' : line.category;
    sums[category] += line.price;
    rows[category]++;
    priced += line.price;
  });

  const order = { sum: sum };
  Object.keys(OLD_COUNT_LABELS).forEach(function (category) {
    let count = null;
    OLD_COUNT_LABELS[category].forEach(function (label) {
      if (count === null) count = labelledValue(values, label);
    });
    order[category] = count !== null ? count : (rows[category] ? null : 0);
  });
  const known = price >= 0 && priced === sum && rows.drip === order.drip;
  Object.keys(sums).forEach(function (category) {
    order[category + 'Sum'] = known ? sums[category] : null;
  });
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
  Object.keys(order).forEach(function (key) {
    order[key] = order[key] === null || other[key] === null ? null : order[key] + other[key];
  });
  return order;
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
