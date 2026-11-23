const test = require('node:test');
const assert = require('node:assert/strict');
const { flatten, shopProduct, useShop } = require('./apps-script');

const BEANS = shopProduct({
  slug: 'beans-1', label: 'Biloya Oro Gr.1', country: 'ETHIOPIA', roast: 'Фільтр', grind: 'У зерні', wholesale: 450,
});
const DRIP = shopProduct({
  slug: 'drip-1', label: 'Biloya Oro Gr.1 Easy Drip', country: 'ETHIOPIA', roast: 'Еспресо', wholesale: 300,
});

test('madHeadsSlug', async (t) => {
  const cases = [
    ['https://madheadscoffee.com/p/5tnVGD8qJTy-3cff1938', '5tnVGD8qJTy-3cff1938'],
    ['http://www.madheadscoffee.com/p/abc_D-1?utm=x', 'abc_D-1'],
    ['  https://madheadscoffee.com/p/abc/  ', 'abc'],
    ['https://madheadscoffee.com/catalog/coffee', ''],
    ['https://3champsroastery.com/p/abc', ''],
    ['madheadscoffee.com/p/abc', ''],
    ['', ''],
    [42, ''],
  ];
  for (const [url, slug] of cases) {
    await t.test(String(url), () => assert.equal(madHeadsSlug(url), slug));
  }
});

test('madHeadsItem', async (t) => {
  await t.test('beans take the wholesale price and lowercase Cyrillic options', () => {
    useShop([BEANS, DRIP]);
    assert.deepEqual(madHeadsItem('https://madheadscoffee.com/p/beans-1'), {
      name: 'ETHIOPIA Biloya Oro Gr.1', roast: 'фільтр', grind: 'у зерні', price: 450,
    });
  });

  await t.test('a drip has no grind', () => {
    useShop([BEANS, DRIP]);
    assert.deepEqual(madHeadsItem('https://madheadscoffee.com/p/drip-1'), {
      name: 'ETHIOPIA Biloya Oro Gr.1 Easy Drip', roast: 'еспресо', grind: '-', price: 300,
    });
  });

  await t.test('Latin options keep their case', () => {
    useShop([shopProduct({ slug: 'latin', label: 'Kenya', roast: 'Omni', grind: 'V-60' })]);
    const item = madHeadsItem('https://madheadscoffee.com/p/latin');
    assert.equal(item.roast, 'Omni');
    assert.equal(item.grind, 'V-60');
  });

  await t.test('fetches the variant slug from __data.json', () => {
    const fetched = useShop([BEANS]);
    madHeadsItem('https://madheadscoffee.com/p/beans-1?x=1');
    assert.deepEqual(fetched, ['https://madheadscoffee.com/p/beans-1/__data.json']);
  });

  await t.test('another shop is null and fetches nothing', () => {
    const fetched = useShop([BEANS]);
    assert.equal(madHeadsItem('https://3champsroastery.com/p/beans-1'), null);
    assert.deepEqual(fetched, []);
  });

  await t.test('an unknown product throws with the status', () => {
    useShop([]);
    assert.throws(() => madHeadsItem('https://madheadscoffee.com/p/gone'),
      { message: '404 from https://madheadscoffee.com/p/gone/__data.json' });
  });

  await t.test('a product without the linked variant throws', () => {
    const product = shopProduct({ slug: 'other', label: 'Kenya', roast: 'Omni' });
    globalThis.UrlFetchApp = {
      fetch: () => ({
        getResponseCode: () => 200,
        getContentText: () => JSON.stringify({ nodes: [{ type: 'data', data: flatten({ product }) }] }),
      }),
    };
    assert.throws(() => madHeadsItem('https://madheadscoffee.com/p/missing'),
      { message: 'No variant missing in product Kenya (internal)' });
  });

  await t.test('a payload without a product throws', () => {
    globalThis.UrlFetchApp = {
      fetch: () => ({
        getResponseCode: () => 200,
        getContentText: () => JSON.stringify({ nodes: [{ type: 'data', data: flatten({ user: null }) }] }),
      }),
    };
    assert.throws(() => madHeadsItem('https://madheadscoffee.com/p/x'),
      { message: 'No product data at https://madheadscoffee.com/p/x/__data.json' });
  });
});

test('productName', async (t) => {
  await t.test('country and Ukrainian label', () => {
    const i18n = [{ lang: 'uk', label: 'Kenya AA', description_short: 'KENYA' }];
    assert.equal(productName({ label: 'internal', i18n }), 'KENYA Kenya AA');
  });
  await t.test('no country', () => {
    assert.equal(productName({ label: 'internal', i18n: [{ lang: 'uk', label: 'Blend' }] }), 'Blend');
  });
  await t.test('no Ukrainian translation', () => {
    assert.equal(productName({ label: 'internal', i18n: [{ lang: 'en', label: 'Blend' }] }), 'internal');
    assert.equal(productName({ label: 'internal' }), 'internal');
  });
});

test('variantOption', () => {
  const variant = { options: [{ option: null }, { option: { property_id: 27, label: 'Омні' } }] };
  assert.equal(variantOption(variant, 27), 'Омні');
  assert.equal(variantOption(variant, 31), '');
  assert.equal(variantOption({}, 27), '');
});

test('lowerIfCyrillic', () => {
  assert.equal(lowerIfCyrillic('У зерні'), 'у зерні');
  assert.equal(lowerIfCyrillic('Під еспресо V-60'), 'під еспресо v-60');
  assert.equal(lowerIfCyrillic('Espresso'), 'Espresso');
  assert.equal(lowerIfCyrillic(''), '');
});

test('firstMatch', () => {
  assert.equal(firstMatch([1, 2, 3, 4], (n) => n % 2 === 0), 2);
  assert.equal(firstMatch([1, 3], (n) => n % 2 === 0), null);
  assert.equal(firstMatch(undefined, () => true), null);
});

test('unflatten', async (t) => {
  await t.test('round trips a flattened object', () => {
    const value = { product: { label: 'Kenya', price: 450, inStock: true, tags: ['a', 'b'], note: null } };
    assert.deepEqual(unflatten(flatten(value)), value);
  });

  await t.test('numbers are indexes, not values', () => {
    assert.deepEqual(unflatten([{ price: 1 }, 450]), { price: 450 });
  });

  await t.test('a shared index is one object', () => {
    const built = unflatten([{ first: 1, second: 1 }, { label: 2 }, 'x']);
    assert.equal(built.first, built.second);
  });

  await t.test('a cycle points back at the object', () => {
    const built = unflatten([{ self: 0, label: 1 }, 'x']);
    assert.equal(built.self, built);
  });

  await t.test('a negative index is null', () => {
    assert.deepEqual(unflatten([{ missing: -1 }]), { missing: null });
  });

  await t.test('a primitive at the root', () => {
    assert.equal(unflatten(['x']), 'x');
  });
});

test('madHeadsCategory', async (t) => {
  const cases = [
    ['ETHIOPIA Biloya Oro Gr.1', '', 'coffee'],
    ['Фільтр', '', 'coffee'],
    ['ETHIOPIA Biloya Oro Gr.1 Easy Drip', '', 'drip'],
    ['Дріп-пакети', '', 'coffee'],
    ['Drip Bag', '', 'coffee'],
    ['MIX BOX', '', 'drip'],
    ['Cold Brew Tonic', '', 'tonic'],
    ['Еспресо-тонік', '', 'tonic'],
    ['Black Coffee 0.25', '', 'black'],
    ['', 'https://madheadscoffee.com/p/blackcoffee-1', 'black'],
    ['Паперові фільтри', '', 'filter'],
    ['HARIO V60 02', '', 'filter'],
    ['Tonic Easy Drip', '', 'tonic'],
  ];
  for (const [name, link, category] of cases) {
    await t.test(name || link, () => assert.equal(madHeadsCategory(name, link), category));
  }
});
