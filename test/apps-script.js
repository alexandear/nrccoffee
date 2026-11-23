// Loads src/ the way Apps Script does, into one global scope, and provides in-memory stand-ins for the services it
// calls: SpreadsheetApp, UrlFetchApp and Charts.
const fs = require('node:fs');
const path = require('node:path');
const vm = require('node:vm');

for (const file of ['MadHeads.js', 'Code.js']) {
  const filename = path.join(__dirname, '..', 'src', file);
  vm.runInThisContext(fs.readFileSync(filename, 'utf8'), { filename });
}

class FakeRange {
  constructor(sheet, row, column, numRows = 1, numColumns = 1) {
    this.sheet = sheet;
    this.row = row;
    this.column = column;
    this.numRows = numRows;
    this.numColumns = numColumns;
  }

  each(callback) {
    for (let r = 0; r < this.numRows; r++) {
      for (let c = 0; c < this.numColumns; c++) callback(this.row + r, this.column + c, r, c);
    }
  }

  getValues() {
    const rows = [];
    this.each((row, column, r) => {
      if (!rows[r]) rows[r] = [];
      rows[r].push(this.sheet.cell(row, column));
    });
    return rows;
  }

  setValue(value) {
    this.sheet.write(this.row, this.column, value);
    return this;
  }

  setValues(values) {
    if (values.length !== this.numRows || values[0].length !== this.numColumns) {
      throw new Error('setValues: ' + values.length + 'x' + values[0].length + ' into ' +
        this.numRows + 'x' + this.numColumns);
    }
    this.each((row, column, r, c) => this.sheet.write(row, column, values[r][c]));
    return this;
  }

  setNumberFormat(format) {
    this.each((row, column) => { this.sheet.formats[row + ',' + column] = format; });
    return this;
  }

  clear() {
    this.each((row, column) => {
      delete this.sheet.formats[row + ',' + column];
      if (this.sheet.values[row - 1] && column <= this.sheet.values[row - 1].length) {
        this.sheet.values[row - 1][column - 1] = '';
      }
    });
    return this;
  }
}

class FakeChart {
  constructor(options, anchorRow, anchorColumn, ranges = [], type = null) {
    this.options = options;
    this.anchorRow = anchorRow;
    this.anchorColumn = anchorColumn;
    this.ranges = ranges;
    this.type = type;
  }

  getOptions() {
    return { get: (key) => (key in this.options ? this.options[key] : null) };
  }

  getContainerInfo() {
    return { getAnchorRow: () => this.anchorRow, getAnchorColumn: () => this.anchorColumn };
  }
}

class FakeChartBuilder {
  constructor() {
    this.options = {};
    this.ranges = [];
  }

  setChartType(type) { this.type = type; return this; }
  addRange(range) { this.ranges.push(range); return this; }
  setNumHeaders(headers) { this.headers = headers; return this; }
  setPosition(row, column) { this.row = row; this.column = column; return this; }
  setOption(key, value) { this.options[key] = value; return this; }
  build() { return new FakeChart(this.options, this.row, this.column, this.ranges, this.type); }
}

let nextSheetId = 1;

class FakeSheet {
  /** values is the sheet's content from A1, one array per row. */
  constructor(name, values = [], { hidden = false, charts = [], maxRows = 1000 } = {}) {
    this.name = name;
    this.values = values.map((row) => row.slice());
    this.hidden = hidden;
    this.charts = charts;
    this.maxRows = maxRows;
    this.formats = {};
    this.writes = [];
    this.id = nextSheetId++;
  }

  cell(row, column) {
    const line = this.values[row - 1];
    return line && column <= line.length && line[column - 1] !== undefined ? line[column - 1] : '';
  }

  write(row, column, value) {
    this.writes.push([row, column, value]);
    while (this.values.length < row) this.values.push([]);
    const line = this.values[row - 1];
    while (line.length < column) line.push('');
    line[column - 1] = value;
  }

  getName() { return this.name; }
  getSheetId() { return this.id; }
  isSheetHidden() { return this.hidden; }
  showSheet() { this.hidden = false; return this; }
  hideSheet() { this.hidden = true; return this; }
  getMaxRows() { return this.maxRows; }

  getLastRow() {
    for (let row = this.values.length; row > 0; row--) {
      if (this.values[row - 1].some((cell) => cell !== '' && cell !== undefined)) return row;
    }
    return 0;
  }

  getDataRange() {
    const width = Math.max(1, ...this.values.map((row) => row.length));
    return new FakeRange(this, 1, 1, Math.max(1, this.getLastRow()), width);
  }

  getRange(row, column, numRows, numColumns) {
    return new FakeRange(this, row, column, numRows, numColumns);
  }

  getCharts() { return this.charts.slice(); }
  removeChart(chart) { this.charts = this.charts.filter((c) => c !== chart); }
  newChart() { return new FakeChartBuilder(); }
  insertChart(chart) { this.charts.push(chart); }
}

class FakeSpreadsheet {
  constructor(sheets) {
    this.sheets = sheets.slice();
    this.active = sheets[0];
  }

  getSheets() { return this.sheets.slice(); }
  getSheetByName(name) { return this.sheets.find((sheet) => sheet.getName() === name) || null; }
  getActiveSheet() { return this.active; }

  setActiveSheet(sheet) {
    if (sheet.isSheetHidden()) throw new Error('Cannot activate hidden sheet ' + sheet.getName());
    this.active = sheet;
    return sheet;
  }

  moveActiveSheet(position) {
    this.sheets.splice(this.sheets.indexOf(this.active), 1);
    this.sheets.splice(position - 1, 0, this.active);
  }
}

class FakeUi {
  constructor() {
    this.alerts = [];
    this.menus = [];
  }

  alert(message) { this.alerts.push(message); }

  createMenu(title) {
    const menus = this.menus;
    const menu = { title, items: [] };
    const builder = {
      menu,
      addItem(caption, functionName) { menu.items.push([caption, functionName]); return builder; },
      addSubMenu(subMenu) { menu.items.push(subMenu.menu); return builder; },
      addToUi() { menus.push(menu); },
    };
    return builder;
  }
}

/** Points SpreadsheetApp at these sheets, the first one active, and returns its UI to read alerts from. */
function useSpreadsheet(sheets) {
  const spreadsheet = new FakeSpreadsheet(sheets);
  const ui = new FakeUi();
  globalThis.SpreadsheetApp = {
    getActive: () => spreadsheet,
    getActiveSheet: () => spreadsheet.getActiveSheet(),
    getUi: () => ui,
  };
  globalThis.Charts = { ChartType: { COMBO: 'COMBO' } };
  return { spreadsheet, ui };
}

/** The SvelteKit flattening that unflatten undoes: every value is stored once and referenced by its index. */
function flatten(value) {
  const flat = [];
  function add(item) {
    const index = flat.length;
    flat.push(null);
    if (Array.isArray(item)) {
      flat[index] = item.map((element) => add(element));
    } else if (item !== null && typeof item === 'object') {
      const object = {};
      for (const key of Object.keys(item)) object[key] = add(item[key]);
      flat[index] = object;
    } else {
      flat[index] = item;
    }
    return index;
  }
  add(value);
  return flat;
}

/** A shop product in the shape the endpoint returns, with the variant under test next to a sibling. */
function shopProduct({ slug, label, country = '', roast, grind = '', wholesale = 450, retail = 600 }) {
  const options = [{ option: { property_id: 27, label: roast } }];
  if (grind) options.push({ option: { property_id: 31, label: grind } });
  return {
    label: label + ' (internal)',
    i18n: [
      { lang: 'en', label: 'English ' + label, description_short: 'English country' },
      { lang: 'uk', label, description_short: country },
    ],
    variants: [
      { slug: slug + '-sibling', price_wholesale: 1, price_retail: 2, options: [] },
      { slug, price_wholesale: wholesale, price_retail: retail, options },
    ],
  };
}

/**
 * Serves products from __data.json the way the shop does, and 404 for any other slug.
 * Returns the list of fetched URLs.
 */
function useShop(products) {
  const fetched = [];
  globalThis.UrlFetchApp = {
    fetch(url) {
      fetched.push(url);
      // https://madheadscoffee.com/p/5tnVGD8qJTy-3cff1938/__data.json
      const matched = /^https:\/\/madheadscoffee\.com\/p\/([^/]+)\/__data\.json$/.exec(url);
      const product = matched && products.find((p) => p.variants.some((v) => v.slug === matched[1]));
      const body = product ? JSON.stringify({
        type: 'data',
        nodes: [null, { type: 'skip' }, { type: 'data', data: flatten({ user: null }) },
          { type: 'data', data: flatten({ product }) }],
      }) : 'Not found';
      return { getResponseCode: () => (product ? 200 : 404), getContentText: () => body };
    },
  };
  return fetched;
}

module.exports = { FakeChart, FakeSheet, flatten, shopProduct, useShop, useSpreadsheet };
