const test = require('node:test');
const assert = require('node:assert/strict');
const { FakeSheet, shopProduct, useShop, useSpreadsheet } = require('./apps-script');

const BEANS = shopProduct({
  slug: 'beans-1', label: 'Biloya Oro Gr.1', country: 'ETHIOPIA', roast: 'Фільтр', grind: 'У зерні', wholesale: 450,
});
const DRIP = shopProduct({
  slug: 'drip-1', label: 'Biloya Oro Gr.1 Easy Drip', country: 'ETHIOPIA', roast: 'Еспресо', wholesale: 300,
});
const BEANS_LINK = 'https://madheadscoffee.com/p/beans-1';
const DRIP_LINK = 'https://madheadscoffee.com/p/drip-1';
const OTHER_LINK = 'https://3champsroastery.com/p/kenya';

const ORDER_HEADER = ['Посилання на товар', 'Назва', 'Обсмажка', 'Помел', 'Кількість', 'Ціна, грн', 'Сума'];
const BEANS_ITEM = { name: 'ETHIOPIA Biloya Oro Gr.1', roast: 'фільтр', grind: 'у зерні', price: 450 };
const DRIP_ITEM = { name: 'ETHIOPIA Biloya Oro Gr.1 Easy Drip', roast: 'еспресо', grind: '-', price: 300 };
const COLUMNS = { link: 0, name: 1, roast: 2, grind: 3, price: 5 };

test('onOpen adds the Кава menu', () => {
  const { ui } = useSpreadsheet([new FakeSheet('Шаблон')]);
  onOpen();
  assert.deepEqual(ui.menus, [{
    title: 'Кава',
    items: [
      ['Заповнити з посилань', 'fillActiveSheet'],
      {
        title: 'Оновити статистику MH',
        items: [
          ['2026', 'fillStats2026'],
          ['2025', 'fillStats2025'],
          ['2024', 'fillStats2024'],
        ],
      },
      ['Впорядкувати аркуші', 'sortSheets'],
    ],
  }]);
});

test('fillActiveSheet', async (t) => {
  await t.test('fills the blanks and reports what it left alone', () => {
    const rows = [
      ORDER_HEADER,
      [BEANS_LINK, '', '', '', 1, '', 'not an order'],
      ['', '', '', '', '', '', ''],
      [BEANS_LINK, '', '', '', 2, 400, '=E4*F4'],
      [DRIP_LINK, 'Easy Drip ETHIOPIA Biloya Oro Gr.1', 'дріпи', 'дріпи', 1, 250, '=E5*F5'],
      [BEANS_LINK, 'ETHIOPIA Biloya Oro Gr.1', 'Фільтр', 'V-60', 1, 400, '=E6*F6'],
      [BEANS_LINK, 'ETHIOPIA Biloya Oro Gr.1', 'Фільтр', 'у зерні', 1, 450, '=E7*F7'],
      [OTHER_LINK, 'Kenya', '', '', 1, '', ''],
      ['https://madheadscoffee.com/p/gone', '', '', '', 1, '', ''],
    ];
    const sheet = new FakeSheet('21.12.2025', rows.map((row) => [...row]));
    const { ui } = useSpreadsheet([sheet]);
    const fetched = useShop([BEANS, DRIP]);

    fillActiveSheet();

    const filled = rows.map((row) => [...row]);
    filled[3] = [BEANS_LINK, 'ETHIOPIA Biloya Oro Gr.1', 'фільтр', 'у зерні', 2, 400, '=E4*F4'];
    assert.deepEqual(sheet.values, filled, 'only the empty cells are written');
    assert.deepEqual(sheet.writes.map(([row, column]) => [row, column]), [[4, 2], [4, 3], [4, 4]]);
    assert.equal(fetched.length, 3, 'a repeated link is fetched once');
    assert.deepEqual(ui.alerts, ['Заповнено рядків: 1' +
      '\nПропущено посилань інших магазинів, рядки: 8' +
      '\n\nЦіна змінилася, комірку не змінено:\nрядок 4: 400 ≠ 450\nрядок 5: 250 ≠ 300' +
      '\n\nНе збігається з посиланням, рядок не змінено:\nрядок 6: Помел "V-60" ≠ "у зерні"' +
      '\n\nПомилки:\nрядок 9: 404 from https://madheadscoffee.com/p/gone/__data.json']);
  });

  await t.test('a lower header row is followed by orders right away', () => {
    const header = ['Посилання', 'Назва', 'Обсмаження', 'Помел', 'Кількість', 'Ціна, грн/шт', 'Сума'];
    const sheet = new FakeSheet('05.12.2022', [
      ['Замовлення'], [], [], [],
      header,
      [BEANS_LINK, '', '', '', 1, '', ''],
    ]);
    const { ui } = useSpreadsheet([sheet]);
    useShop([BEANS]);

    fillActiveSheet();

    assert.deepEqual(sheet.values[5], [BEANS_LINK, 'ETHIOPIA Biloya Oro Gr.1', 'фільтр', 'у зерні', 1, 450, '']);
    assert.deepEqual(ui.alerts, ['Заповнено рядків: 1']);
  });

  await t.test('a sheet without the columns is reported', () => {
    const sheet = new FakeSheet('Шаблон', [['Посилання', 'Назва', 'Кількість']]);
    const { ui } = useSpreadsheet([sheet]);
    fillActiveSheet();
    assert.deepEqual(ui.alerts, ['На аркуші "Шаблон" немає потрібних колонок.']);
  });

  await t.test('a header below the tenth row is not found', () => {
    const rows = Array.from({ length: 10 }, () => ['']);
    const sheet = new FakeSheet('Шаблон', [...rows, ORDER_HEADER]);
    const { ui } = useSpreadsheet([sheet]);
    fillActiveSheet();
    assert.deepEqual(ui.alerts, ['На аркуші "Шаблон" немає потрібних колонок.']);
  });

  await t.test('a sheet of another shop says it is not supported', () => {
    const sheet = new FakeSheet('27.11.2025 3champs', [ORDER_HEADER, [], [], [OTHER_LINK, 'Kenya']]);
    const { ui } = useSpreadsheet([sheet]);
    useShop([]);
    fillActiveSheet();
    assert.deepEqual(ui.alerts, ['Посилань MadHeads не знайдено. Інші магазини поки не підтримуються.']);
    assert.deepEqual(sheet.writes, []);
  });

  await t.test('one run stops at MAX_FETCHES products', () => {
    const products = [];
    const rows = [ORDER_HEADER, [], []];
    for (let i = 0; i < 61; i++) {
      products.push(shopProduct({ slug: 'p' + i, label: 'Coffee ' + i, roast: 'Omni', grind: 'V-60' }));
      rows.push(['https://madheadscoffee.com/p/p' + i, '', '', '', 1, '', '']);
    }
    const sheet = new FakeSheet('01.01.2026', rows);
    const { ui } = useSpreadsheet([sheet]);
    const fetched = useShop(products);

    fillActiveSheet();

    assert.equal(fetched.length, 60);
    assert.deepEqual(ui.alerts, ['Заповнено рядків: 60']);
    assert.equal(sheet.values[62][1], 'Coffee 59');
    assert.equal(sheet.values[63][1], '', 'the 61st link waits for the next run');
  });
});

test('mismatches', async (t) => {
  await t.test('empty cells match anything', () => {
    assert.deepEqual(mismatches([BEANS_LINK, '', '', '', 1, ''], COLUMNS, BEANS_ITEM), []);
  });

  await t.test('same words in any order and case', () => {
    const row = [DRIP_LINK, ' easy drip  ETHIOPIA Biloya Oro Gr.1 ', 'Еспресо', '-', 1, 1];
    assert.deepEqual(mismatches(row, COLUMNS, DRIP_ITEM), []);
  });

  await t.test('every differing cell is listed', () => {
    const row = [BEANS_LINK, 'KENYA', 'Еспресо', 'V-60', 1, 1];
    assert.deepEqual(mismatches(row, COLUMNS, BEANS_ITEM), [
      'Назва "KENYA" ≠ "ETHIOPIA Biloya Oro Gr.1"',
      'Обсмажка "Еспресо" ≠ "фільтр"',
      'Помел "V-60" ≠ "у зерні"',
    ]);
  });

  await t.test('"дріпи" matches a drip only', () => {
    assert.deepEqual(mismatches([DRIP_LINK, '', 'дріпи', 'Дріпи', 1, 1], COLUMNS, DRIP_ITEM), []);
    assert.deepEqual(mismatches([BEANS_LINK, '', 'дріпи', '', 1, 1], COLUMNS, BEANS_ITEM),
      ['Обсмажка "дріпи" ≠ "фільтр"']);
  });

  await t.test('the price is not one of the compared cells', () => {
    assert.deepEqual(mismatches([BEANS_LINK, '', '', '', 1, 999], COLUMNS, BEANS_ITEM), []);
  });
});

test('isDripMark', () => {
  assert.equal(isDripMark(' Дріпи ', DRIP_ITEM), true);
  assert.equal(isDripMark('дріпи', BEANS_ITEM), false);
  assert.equal(isDripMark('еспресо', DRIP_ITEM), false);
});

test('sameWords and words', () => {
  assert.equal(words('Easy Drip  KENYA'), 'drip easy kenya');
  assert.equal(sameWords('Easy Drip KENYA', 'KENYA Easy Drip'), true);
  assert.equal(sameWords('KENYA', 'KENYA AA'), false);
});

test('findHeaderRow', () => {
  assert.equal(findHeaderRow([ORDER_HEADER]), 0);
  assert.equal(findHeaderRow([['Замовлення'], ['', ' ПОСИЛАННЯ ']]), 1);
  assert.equal(findHeaderRow([['Назва']]), -1);
  assert.equal(findHeaderRow([]), -1);
});

test('findColumns', () => {
  assert.deepEqual(findColumns(ORDER_HEADER), COLUMNS);
  assert.deepEqual(findColumns(['Назва', 'Обсмаження', 'Помел', 'Ціна, грн/шт', 'Посилання (MH)']),
    { link: 4, name: 0, roast: 1, grind: 2, price: 3 });
  assert.equal(findColumns(['Посилання', 'Назва', 'Обсмажка', 'Помел']), null);
});

const OLD_HEADER = [
  "Ім'я", 'Посилання на каву', 'Примітки', 'Ціна, грн', 'Ціна з людини',
  'Пачок кави', 'Пачок дріпів', 'Coffee Tonic', 'Всього пачок', 'Всього, грн', 'Назва',
];

function oldOrderRows({ dripMark = 1 } = {}) {
  return [
    OLD_HEADER,
    ['Хто', 'Що', 'Примітка', 'Ціна оптова', '', 2, 2, 22, 26, 2301],
    ['', 'Посилання без помолу і обсмажування вважається "Фільтр у зерні"'],
    ['A', 'https://madheadscoffee.com/p/OcYG8qMawKc', '', 220, 523],
    ['', 'https://madheadscoffee.com/p/q46wOSi6umo', '', 303],
    ['B', 'https://madheadscoffee.com/p/WueUkw88Y4O', '', 240, 480, '', '', '', '', '', 'RWANDA Cyato Box'],
    ['', 'https://madheadscoffee.com/p/Ae8oEACSG6', '', 240, '', '', dripMark],
    ['C', 'Coffee Tonic', 22, 1298, 1298],
    ['', '', '', '', 0],
  ];
}

test('summarizeOldOrder', async (t) => {
  await t.test('takes the labelled counts and sums each row\'s price by category', () => {
    assert.deepEqual(summarizeOrder(new FakeSheet('09.01.2024', oldOrderRows())), {
      coffee: 2, drip: 2, tonic: 22, black: 0, filter: 0, sum: 2301,
      coffeeSum: 523, dripSum: 480, tonicSum: 1298, blackSum: 0, filterSum: 0,
    });
  });

  await t.test('leaves every Сума unknown when a drip cannot be told', () => {
    assert.deepEqual(summarizeOrder(new FakeSheet('02.03.2024', oldOrderRows({ dripMark: '' }))), {
      coffee: 2, drip: 2, tonic: 22, black: 0, filter: 0, sum: 2301,
      coffeeSum: null, dripSum: null, tonicSum: null, blackSum: null, filterSum: null,
    });
  });

  await t.test('a drip named in a note makes its bare link a drip too', () => {
    const sheet = new FakeSheet('17.02.2024', [
      OLD_HEADER,
      ['Хто', 'Що', 'Примітка', 'Ціна оптова', '', 1, 2, 0, 3, 915],
      ['A', 'https://madheadscoffee.com/p/ekEsWs4mIaq', '', 395, 915],
      ['', 'https://madheadscoffee.com/p/Ae8oEACSG6', 'Дріпи мікс', 260],
      ['B', 'https://madheadscoffee.com/p/Ae8oEACSG6', '', 260, 260],
    ]);
    assert.deepEqual(summarizeOrder(sheet), {
      coffee: 1, drip: 2, tonic: 0, black: 0, filter: 0, sum: 915,
      coffeeSum: 395, dripSum: 520, tonicSum: 0, blackSum: 0, filterSum: 0,
    });
  });

  await t.test('reads the 2022 layout with its labels at the bottom', () => {
    const sheet = new FakeSheet('22.08.2022', [
      ["Ім'я", 'Посилання на каву', 'Кількість', 'Кількість Coffee Tonic', 'Кількість BlackCoffee', 'Примітки', 'Ціна'],
      ['A', 'https://madheadscoffee.com/shop/mh-drinks/coffee-tonic/', '', 5, '', '', 245],
      ['B', 'https://madheadscoffee.com/shop/mh-drinks/black-coffee/', '', '', 22, '', 1078],
      ['C', 'https://madheadscoffee.com/shop/izidrip/mix-box/', 1, '', '', '', 300],
      [],
      ['', '', 'Пачок дріпів', 'Coffee Tonic', 'BlackCoffee', '', 'Всього'],
      ['', '', 1, 5, 22, '', 1623],
    ]);
    assert.deepEqual(summarizeOrder(sheet), {
      coffee: 0, drip: 1, tonic: 5, black: 22, filter: 0, sum: 1623,
      coffeeSum: 0, dripSum: 300, tonicSum: 245, blackSum: 1078, filterSum: 0,
    });
  });

  await t.test('a count without a label is unknown when rows of its category exist', () => {
    const rows = oldOrderRows();
    rows[0] = rows[0].map((cell) => (cell === 'Coffee Tonic' ? 'Тонік' : cell));
    const order = summarizeOrder(new FakeSheet('x', rows));
    assert.equal(order.tonic, null);
    assert.equal(order.black, 0);
  });

  await t.test('fillStats keeps the cells it cannot count and names their dates', () => {
    const stats = new FakeSheet('Статистика MH', [
      [...STATS_HEADER, ...MONEY_HEADER],
      ['02.03.2024', 2, 2, '', 0, 0, 2301],
    ]);
    const { ui } = useSpreadsheet([stats, new FakeSheet('02.03.2024', oldOrderRows({ dripMark: '' }))]);

    fillStats(2024);

    assert.deepEqual(stats.values[1], ['02.03.2024', 2, 2, 22, 0, 0, 2301]);
    assert.deepEqual(ui.alerts, [
      'Додано рядків: 0\nОновлено рядків: 1\nНе вдалося порахувати, заповніть вручну: 02.03.2024',
    ]);
  });
});

test('isLinkHeader, indexOfHeader and normalize', () => {
  assert.equal(isLinkHeader(' Посилання на товар'), true);
  assert.equal(isLinkHeader('Моє посилання'), false);
  assert.equal(indexOfHeader(['A', ' Б '], (cell) => cell === 'б'), 1);
  assert.equal(indexOfHeader(['A'], (cell) => cell === 'б'), -1);
  assert.equal(normalize('  Сума '), 'сума');
  assert.equal(normalize(42), '42');
});

const STATS_HEADER = ['Дата', 'Кава', 'Дріпи', 'Тонік', 'Блек', 'Фільтри', 'Сума'];
const MONEY_HEADER = ['Кава, грн', 'Дріпи, грн', 'Coffee Tonic, грн', 'Black Coffee, грн', 'Фільтри V60, грн'];

function orderSheet(name, rows) {
  return new FakeSheet(name, [['Посилання', 'Назва', 'Кількість', 'Сума'], ...rows]);
}

test('fillStats', async (t) => {
  await t.test('appends the missing dates oldest first, recounts the listed ones and merges sheets of one date', () => {
    const stats = new FakeSheet('Статистика MH', [
      STATS_HEADER,
      ['', '=SUM(B3:B300)'],
      ['21.12.2025', 1, 0, 0, 0, 0, 450],
      [new Date(2025, 0, 5), 2, 0, 0, 0, 0, 900, 777],
    ]);
    const { ui } = useSpreadsheet([
      stats,
      new FakeSheet('Шаблон', [['Посилання', 'Назва', 'Кількість', 'Сума'], [BEANS_LINK, 'Kenya', 1, 1]]),
      orderSheet('21.12.2025 ✅', [[BEANS_LINK, 'Kenya', 5, 5]]),
      orderSheet('5.1.2025', [[BEANS_LINK, 'Kenya', 5, 5]]),
      orderSheet(' 9.4.2026 MH ✅', [[BEANS_LINK, 'Kenya', 2, 900], [DRIP_LINK, 'Kenya Easy Drip', 10, 3000]]),
      orderSheet('09.04.2026 (2)', [[BEANS_LINK, 'Cold Brew Tonic', 3, 600]]),
      orderSheet('27.11.2025 3champs ✅', [[BEANS_LINK, 'Kenya', 1, 450]]),
      orderSheet('01.02.2026', [['', 'HARIO V60', 1, 200], ['', 'Black Coffee', 4, 800]]),
      new FakeSheet('02.03.2026', [['Нотатки']]),
    ]);

    fillStats();

    assert.deepEqual(stats.values, [
      [...STATS_HEADER, ...MONEY_HEADER],
      ['', '=SUM(B3:B300)'],
      ['21.12.2025', 5, 0, 0, 0, 0, 5, 5, 0, 0, 0, 0],
      [new Date(2025, 0, 5), 5, 0, 0, 0, 0, 5, 5, 0, 0, 0, 0],
      ['01.02.2026', 0, 0, 0, 4, 1, 1000, 0, 0, 0, 800, 200],
      ['09.04.2026', 2, 10, 3, 0, 0, 4500, 900, 3000, 600, 0, 0],
    ]);
    assert.equal(stats.formats['5,1'], '@');
    assert.equal(stats.formats['6,1'], '@');
    assert.deepEqual(ui.alerts, ['Додано рядків: 2\nОновлено рядків: 2']);
  });

  await t.test('reads only the sheets of the given year', () => {
    const stats = new FakeSheet('Статистика MH', [
      [...STATS_HEADER, ...MONEY_HEADER],
      ['21.12.2025', 1, 0, 0, 0, 0, 450],
    ]);
    const { ui } = useSpreadsheet([
      stats,
      orderSheet('21.12.2025', [[BEANS_LINK, 'Kenya', 1, 450]]),
      orderSheet('09.04.2026', [[BEANS_LINK, 'Kenya', 2, 900]]),
    ]);

    fillStats2026();

    assert.deepEqual(stats.values, [
      [...STATS_HEADER, ...MONEY_HEADER],
      ['21.12.2025', 1, 0, 0, 0, 0, 450],
      ['09.04.2026', 2, 0, 0, 0, 0, 900, 900, 0, 0, 0, 0],
    ]);
    assert.deepEqual(ui.alerts, ['Додано рядків: 1\nОновлено рядків: 0']);
  });

  await t.test('fills the empty rows kept for a year', () => {
    const stats = new FakeSheet('Статистика MH', [
      [...STATS_HEADER, ...MONEY_HEADER],
      ['Дата/Всього', '=SUM(B3:B300)'],
      ['26.12.2024', 1],
      [],
      [],
      [],
      ['07.01.2026', 1],
    ]);
    useSpreadsheet([
      stats,
      orderSheet('05.06.2025', [[BEANS_LINK, 'Kenya', 2, 900]]),
      orderSheet('10.01.2025', [[BEANS_LINK, 'Kenya', 1, 450]]),
    ]);

    fillStats(2025);

    assert.deepEqual(stats.values, [
      [...STATS_HEADER, ...MONEY_HEADER],
      ['Дата/Всього', '=SUM(B3:B300)'],
      ['26.12.2024', 1],
      ['10.01.2025', 1, 0, 0, 0, 0, 450, 450, 0, 0, 0, 0],
      ['05.06.2025', 2, 0, 0, 0, 0, 900, 900, 0, 0, 0, 0],
      [],
      ['07.01.2026', 1],
    ]);
    assert.equal(stats.inserted, undefined);
  });

  await t.test('inserts a row among the dates when there is no empty one', () => {
    const stats = new FakeSheet('Статистика MH', [
      [...STATS_HEADER, ...MONEY_HEADER],
      ['Дата/Всього', '=SUM(B3:B300)'],
      ['07.01.2026', 1],
      ['20.01.2026', 1],
    ]);
    const { ui } = useSpreadsheet([
      stats,
      orderSheet('10.01.2026', [[BEANS_LINK, 'Kenya', 1, 450]]),
      orderSheet('20.12.2025', [[BEANS_LINK, 'Kenya', 2, 900]]),
      orderSheet('01.02.2026', [[BEANS_LINK, 'Kenya', 3, 1350]]),
    ]);

    fillStats();

    assert.deepEqual(stats.values.map((row) => row[0]), [
      'Дата', 'Дата/Всього', '20.12.2025', '07.01.2026', '10.01.2026', '20.01.2026', '01.02.2026',
    ]);
    assert.deepEqual(stats.values[4], ['10.01.2026', 1, 0, 0, 0, 0, 450, 450, 0, 0, 0, 0]);
    assert.deepEqual(stats.inserted, [2, 4]);
    assert.deepEqual(stats.values[3], ['07.01.2026', 1, 0, 0, 0, 0, '', '', 0, 0, 0, 0]);
    assert.deepEqual(ui.alerts, ['Додано рядків: 3\nОновлено рядків: 2']);
  });

  await t.test('fills 0 into the empty cells of a row without an order sheet, and leaves 2023 alone', () => {
    const stats = new FakeSheet('Статистика MH', [
      [...STATS_HEADER, ...MONEY_HEADER],
      ['05.06.2024', '', 1, 24, '', '', 1656],
      ['05.06.2023', 8, '', '', '', '', 1411],
    ]);
    const { ui } = useSpreadsheet([stats, orderSheet('10.01.2023', [[BEANS_LINK, 'Kenya', 1, 450]])]);

    fillStats();

    assert.deepEqual(stats.values, [
      [...STATS_HEADER, ...MONEY_HEADER],
      ['05.06.2024', 0, 1, 24, 0, 0, 1656, 0, '', '', 0, 0],
      ['05.06.2023', 8, '', '', '', '', 1411],
    ]);
    assert.deepEqual(ui.alerts, ['Додано рядків: 0\nОновлено рядків: 1']);
  });

  await t.test('nothing to change', () => {
    const stats = new FakeSheet('Статистика MH', [
      [...STATS_HEADER, ...MONEY_HEADER],
      ['21.12.2025', 1, 0, 0, 0, 0, 450, 450, 0, 0, 0, 0],
    ]);
    const { ui } = useSpreadsheet([stats, orderSheet('21.12.2025', [[BEANS_LINK, 'Kenya', 1, 450]])]);
    fillStats();
    assert.deepEqual(stats.writes, []);
    assert.deepEqual(ui.alerts, ['Додано рядків: 0\nОновлено рядків: 0']);
  });

  await t.test('no statistics sheet', () => {
    const { ui } = useSpreadsheet([new FakeSheet('Шаблон')]);
    fillStats();
    assert.deepEqual(ui.alerts, ['Аркуша "Статистика MH" немає.']);
  });
});

test('summarizeOrder', async (t) => {
  await t.test('counts per category and sums the Сума column', () => {
    const sheet = orderSheet('01.02.2026', [
      [BEANS_LINK, 'Kenya', 2, 900],
      [DRIP_LINK, 'Kenya Easy Drip', 10, 3000],
      ['', 'Еспресо-тонік', 3, 600],
      ['', 'Black Coffee', 1, 200],
      ['', 'Фільтри паперові', 1, 100],
      ['', 'Фільтр', 1, ''],
      ['', 'Без кількості', 'два', 100],
      ['', '', 5, 100],
      ['', 'Всього', '', 9999],
    ]);
    assert.deepEqual(summarizeOrder(sheet), {
      coffee: 3, drip: 10, tonic: 3, black: 1, filter: 1, sum: 4800,
      coffeeSum: 900, dripSum: 3000, tonicSum: 600, blackSum: 200, filterSum: 100,
    });
  });

  await t.test('"Всього грн" overrides the Сума column', () => {
    const sheet = new FakeSheet('01.02.2026', [
      ['Посилання', 'Назва', 'Кількість', 'Сума', '', 'Всього, грн'],
      [BEANS_LINK, 'Kenya', 2, 900, '', 850],
    ]);
    assert.equal(summarizeOrder(sheet).sum, 850);
    assert.equal(summarizeOrder(sheet).coffeeSum, 900);
  });

  await t.test('rows of another shop are dropped along with the hand-kept total', () => {
    const sheet = new FakeSheet('01.02.2026', [
      ['Посилання', 'Назва', 'Кількість', 'Сума', 'Всього грн'],
      [BEANS_LINK, 'Kenya', 2, 900, 2000],
      [OTHER_LINK, 'Kenya', 3, 1100],
    ]);
    assert.deepEqual(summarizeOrder(sheet), {
      coffee: 2, drip: 0, tonic: 0, black: 0, filter: 0, sum: 900,
      coffeeSum: 900, dripSum: 0, tonicSum: 0, blackSum: 0, filterSum: 0,
    });
  });

  await t.test('not an order', () => {
    assert.equal(summarizeOrder(new FakeSheet('x', [['Назва', 'Кількість', 'Сума']])), null);
    assert.equal(summarizeOrder(new FakeSheet('x', [['Посилання', 'Назва', 'Сума'], [BEANS_LINK, 'K', 1]])), null);
    assert.equal(summarizeOrder(new FakeSheet('x', [['Посилання', 'Назва', 'Кількість']])), null);
    assert.equal(summarizeOrder(orderSheet('x', [])), null);
    assert.equal(summarizeOrder(orderSheet('x', [[OTHER_LINK, 'Kenya', 1, 1]])), null);
  });
});

test('isOtherShop', () => {
  assert.equal(isOtherShop(OTHER_LINK), true);
  assert.equal(isOtherShop('http://example.com'), true);
  assert.equal(isOtherShop(BEANS_LINK), false);
  assert.equal(isOtherShop('Kenya'), false);
  assert.equal(isOtherShop(''), false);
});

test('labelledValue', () => {
  assert.equal(labelledValue([['', 'Всього грн'], ['', 100]], 'всього грн'), 100);
  assert.equal(labelledValue([[' ВСЬОГО, ГРН '], [100]], 'всього грн'), 100);
  assert.equal(labelledValue([['Всього грн', 'Всього грн'], ['', 70]], 'всього грн'), 70);
  assert.equal(labelledValue([['Всього грн'], ['сто']], 'всього грн'), null);
  assert.equal(labelledValue([['Всього грн']], 'всього грн'), null);
});

test('sheetDate', () => {
  assert.equal(sheetDate('21.12.2025 ✅'), '21.12.2025');
  assert.equal(sheetDate(' 9.4.2026 MH ✅'), '09.04.2026');
  assert.equal(sheetDate('Шаблон'), '');
});

test('normalizeDate', () => {
  assert.equal(normalizeDate(new Date(2025, 0, 5)), '05.01.2025');
  assert.equal(normalizeDate('5.1.2025'), '05.01.2025');
  assert.equal(normalizeDate('Дата'), '');
  assert.equal(normalizeDate(''), '');
});

test('sortableDate and pad', () => {
  assert.equal(sortableDate('21.12.2025'), '20251221');
  assert.equal(pad(5), '05');
  assert.equal(pad('12'), '12');
});

test('addOrders', () => {
  const first = { coffee: 1, drip: 2, tonic: 0, black: 0, filter: 0, sum: 100 };
  const second = { coffee: 3, drip: 0, tonic: 1, black: 0, filter: 1, sum: 50 };
  assert.equal(addOrders(first, second), first);
  assert.deepEqual(first, { coffee: 4, drip: 2, tonic: 1, black: 0, filter: 1, sum: 150 });
});

test('sortSheets', async (t) => {
  await t.test('newest first, sheets without a date in front, hidden ones kept hidden', () => {
    const sheets = [
      new FakeSheet('01.02.2026'),
      new FakeSheet('Шаблон'),
      new FakeSheet('15.03.2026', [], { hidden: true }),
      new FakeSheet('Статистика MH'),
      new FakeSheet('27.11.2025 3champs ✅'),
      new FakeSheet('15.03.2026 (2)'),
    ];
    const { spreadsheet, ui } = useSpreadsheet(sheets);

    sortSheets();

    assert.deepEqual(spreadsheet.getSheets().map((sheet) => sheet.getName()), [
      'Шаблон', 'Статистика MH', '15.03.2026', '15.03.2026 (2)', '01.02.2026', '27.11.2025 3champs ✅',
    ]);
    assert.equal(sheets[2].isSheetHidden(), true);
    assert.equal(spreadsheet.getActiveSheet(), sheets[0]);
    assert.deepEqual(ui.alerts, ['Переміщено аркушів: 4']);
  });

  await t.test('already sorted', () => {
    const { ui } = useSpreadsheet([new FakeSheet('Шаблон'), new FakeSheet('02.01.2026'), new FakeSheet('01.01.2026')]);
    sortSheets();
    assert.deepEqual(ui.alerts, ['Переміщено аркушів: 0']);
  });
});

test('byDateDescending', () => {
  const sheet = (name) => new FakeSheet(name);
  assert.equal(byDateDescending(sheet('02.01.2026'), sheet('01.01.2026')), -1);
  assert.equal(byDateDescending(sheet('01.01.2026'), sheet('31.12.2025')), -1);
  assert.equal(byDateDescending(sheet('01.01.2026'), sheet('02.01.2026')), 1);
  assert.equal(byDateDescending(sheet('01.01.2026'), sheet('1.1.2026 (2)')), 0);
  assert.equal(byDateDescending(sheet('Шаблон'), sheet('01.01.2026')), -1);
  assert.equal(byDateDescending(sheet('01.01.2026'), sheet('Шаблон')), 1);
  assert.equal(byDateDescending(sheet('Шаблон'), sheet('Статистика MH')), 0);
});
