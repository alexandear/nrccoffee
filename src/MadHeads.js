const MADHEADS_HOST = 'https://madheadscoffee.com';
const ROAST_PROPERTY_ID = 27;
const GRIND_PROPERTY_ID = 31;

// https://madheadscoffee.com/p/5tnVGD8qJTy-3cff1938
const PRODUCT_URL = /^https?:\/\/(?:www\.)?madheadscoffee\.com\/p\/([A-Za-z0-9_-]+)/;

/** Variant slug of a cart link, or '' when the link is not a MadHeads product. */
function madHeadsSlug(url) {
  const matched = PRODUCT_URL.exec(String(url).trim());
  return matched ? matched[1] : '';
}

/**
 * Name, roast, grind and wholesale price of the variant the link points at.
 * Returns null for anything that is not a MadHeads product link.
 */
function madHeadsItem(url) {
  const slug = madHeadsSlug(url);
  if (!slug) return null;

  const product = fetchMadHeadsProduct(slug);
  const variant = firstMatch(product.variants, function (v) { return v.slug === slug; });
  if (!variant) throw new Error('No variant ' + slug + ' in product ' + product.label);

  const grind = variantOption(variant, GRIND_PROPERTY_ID);
  return {
    name: productName(product),
    roast: lowerIfCyrillic(variantOption(variant, ROAST_PROPERTY_ID)),
    grind: grind ? lowerIfCyrillic(grind) : '-',
    price: variant.price_wholesale,
  };
}

function fetchMadHeadsProduct(slug) {
  const url = MADHEADS_HOST + '/p/' + slug + '/__data.json';
  const response = UrlFetchApp.fetch(url, { muteHttpExceptions: true });
  if (response.getResponseCode() !== 200) {
    throw new Error(response.getResponseCode() + ' from ' + url);
  }

  const nodes = JSON.parse(response.getContentText()).nodes || [];
  for (let i = 0; i < nodes.length; i++) {
    if (!nodes[i] || nodes[i].type !== 'data') continue;
    const data = unflatten(nodes[i].data);
    if (data && data.product) return data.product;
  }
  throw new Error('No product data at ' + url);
}

/** Country plus label, the two headings the shop shows: "ETHIOPIA Biloya Oro Gr.1 Easy Drip". */
function productName(product) {
  const uk = firstMatch(product.i18n, function (t) { return t.lang === 'uk'; }) || {};
  const country = uk.description_short || '';
  const label = uk.label || product.label;
  return country ? country + ' ' + label : label;
}

function variantOption(variant, propertyId) {
  const option = firstMatch(variant.options, function (o) {
    return o.option && o.option.property_id === propertyId;
  });
  return option ? option.option.label : '';
}

/** The shop capitalizes its Ukrainian labels ("У зерні"); the sheet keeps them lowercase. */
function lowerIfCyrillic(label) {
  return /[Ѐ-ӿ]/.test(label) ? label.toLowerCase() : label;
}

function firstMatch(items, predicate) {
  const found = (items || []).filter(predicate);
  return found.length ? found[0] : null;
}

/**
 * SvelteKit flattens its payload into an array where every number is an index
 * into that same array. Rebuild the object graph from index 0.
 */
function unflatten(flat) {
  const built = new Array(flat.length);
  const done = new Array(flat.length);

  function at(index) {
    if (index < 0) return null;
    if (done[index]) return built[index];
    done[index] = true;

    const value = flat[index];
    if (value === null || typeof value !== 'object') {
      built[index] = value;
      return value;
    }
    if (Object.prototype.toString.call(value) === '[object Array]') {
      const list = [];
      built[index] = list;
      for (let i = 0; i < value.length; i++) list.push(at(value[i]));
      return list;
    }
    const object = {};
    built[index] = object;
    const keys = Object.keys(value);
    for (let i = 0; i < keys.length; i++) object[keys[i]] = at(value[keys[i]]);
    return object;
  }

  return at(0);
}

/**
 * Statistics category of an order row: 'coffee', 'drip', 'tonic', 'black' or 'filter'.
 * The shop names variants, not categories, so the name is all there is to go on.
 * "Фільтр" alone is a roast, only "Фільтри" and Hario are the paper filters.
 */
function madHeadsCategory(name, link) {
  const text = (String(name) + ' ' + String(link)).toLowerCase();
  if (text.indexOf('tonic') >= 0 || text.indexOf('тонік') >= 0) return 'tonic';
  if (text.indexOf('black coffee') >= 0 || text.indexOf('blackcoffee') >= 0) return 'black';
  if (text.indexOf('drip') >= 0 || text.indexOf('дріп') >= 0) return 'drip';
  if (text.indexOf('фільтри') >= 0 || text.indexOf('hario') >= 0) return 'filter';
  return 'coffee';
}
