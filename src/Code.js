/** Fetches per run, so one run stays well inside the 6 minute limit. */
const MAX_FETCHES = 60;

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Кава')
    .addItem('Заповнити з посилань', 'fillActiveSheet')
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
