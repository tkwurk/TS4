// ========== PRODUCT OPTIONS / VARIANTS (base + option → sell SKU) ==========
// Products sheet: menu bases (Lil Swirl, Big Swirl).
// ProductOptions: choices (Vanilla, Choco…) grouped by OptionGroupId.
// ProdOptions: links base → option group (one group can attach to many bases).
// ProductVariants: resolved sell SKU per base+option → Orders + BOM Unit Product Code.

function sortCompositeOptionParts_(parts, sortOrderByCode) {
  sortOrderByCode = sortOrderByCode || {};
  return parts.slice().sort(function(a, b) {
    var sa = sortOrderByCode[a];
    var sb = sortOrderByCode[b];
    var hasA = sa !== undefined && sa !== null && sa !== '';
    var hasB = sb !== undefined && sb !== null && sb !== '';
    if (hasA && hasB && Number(sa) !== Number(sb)) return Number(sa) - Number(sb);
    if (hasA && !hasB) return -1;
    if (!hasA && hasB) return 1;
    return a.localeCompare(b);
  });
}

/**
 * Composite sort rank: option group step (ProdOptions SortOrder) then option SortOrder.
 * Ensures e.g. VANILLA_FLAVOR+SMALL not SMALL+VANILLA_FLAVOR when both use SortOrder 1.
 */
function buildOptionCompositeSortOrders_(optionLines, prodLinkLines) {
  var groupStep = {};
  (prodLinkLines || []).forEach(function(link) {
    var gid = (link.optionGroupId || '').toString().trim();
    if (!gid) return;
    var step = Number(link.sortOrder) || 0;
    if (groupStep[gid] === undefined || step < groupStep[gid]) {
      groupStep[gid] = step;
    }
  });

  var sortOrders = {};
  (optionLines || []).forEach(function(opt) {
    var code = normalizeConfiguratorCode_(opt.optionCode);
    if (!code) return;
    var gid = (opt.optionGroupId || '').toString().trim();
    var groupRank = groupStep[gid];
    if (groupRank === undefined || groupRank === null || groupRank === '') {
      groupRank = 999;
    }
    sortOrders[code] = (Number(groupRank) || 0) * 1000 + (Number(opt.sortOrder) || 0);
  });
  return sortOrders;
}

var COMPOSITE_OPTION_SEP = '_';

function getKnownOptionCodeList_(optionsCatalog) {
  var list = [];
  (optionsCatalog.lines || []).forEach(function(opt) {
    if (opt.optionType === 'mix_trigger') return;
    var code = normalizeConfiguratorCode_(opt.optionCode);
    if (code) list.push(code);
  });
  return list;
}

function splitCompositeOptionCode_(composite, knownCodes) {
  var raw = normalizeConfiguratorCode_(composite);
  if (!raw) return [];

  if (raw.indexOf('+') >= 0) {
    return raw.split('+').map(function(part) {
      return normalizeConfiguratorCode_(part);
    }).filter(function(part) { return !!part; });
  }

  var codeList = (knownCodes || []).slice().sort(function(a, b) {
    return b.length - a.length;
  });
  if (!codeList.length) {
    return [raw];
  }
  if (codeList.indexOf(raw) >= 0) {
    return [raw];
  }

  var parts = [];
  var pos = 0;
  while (pos < raw.length) {
    var matched = false;
    for (var i = 0; i < codeList.length; i++) {
      var code = codeList[i];
      if (raw.substr(pos, code.length) !== code) continue;
      var next = pos + code.length;
      if (next !== raw.length && raw.charAt(next) !== COMPOSITE_OPTION_SEP) continue;
      parts.push(code);
      pos = next < raw.length ? next + 1 : next;
      matched = true;
      break;
    }
    if (!matched) {
      return [raw];
    }
  }
  return parts.length ? parts : [raw];
}

function joinCompositeOptionParts_(parts, sortOrderByCode) {
  var normalized = (parts || []).map(function(part) {
    return normalizeConfiguratorCode_(part);
  }).filter(function(part) { return !!part; });
  if (normalized.length <= 1) {
    return normalized[0] || '';
  }
  return sortCompositeOptionParts_(normalized, sortOrderByCode || {}).join(COMPOSITE_OPTION_SEP);
}

function isCompositeOptionMix_(parts, optionsByCode) {
  var groupCounts = {};
  (parts || []).forEach(function(code) {
    var opt = optionsByCode[normalizeConfiguratorCode_(code)];
    if (!opt) return;
    var gid = opt.optionGroupId || '';
    groupCounts[gid] = (groupCounts[gid] || 0) + 1;
  });
  return Object.keys(groupCounts).some(function(gid) {
    return groupCounts[gid] >= 2;
  });
}

function normalizeCompositeOptionCode_(optionCode, sortOrderByCode, knownCodes) {
  var raw = (optionCode || '').toString().trim();
  if (!raw) return '';
  var parts = splitCompositeOptionCode_(raw, knownCodes);
  if (parts.length <= 1) {
    return parts[0] || normalizeConfiguratorCode_(raw);
  }
  return joinCompositeOptionParts_(parts, sortOrderByCode);
}

function isMixTriggerOptionCode_(optionCode) {
  return normalizeConfiguratorCode_(optionCode) === 'MIX_BLAST_FLAVOR';
}

function normalizeConfiguratorCode_(code) {
  if (typeof normalizeProductCodeKey_ === 'function') {
    return normalizeProductCodeKey_(code);
  }
  return (code || '').toString().trim().toUpperCase().replace(/\s+/g, '_');
}

function isConfiguratorRowActive_(status) {
  if (typeof isBomRowActive_ === 'function') {
    return isBomRowActive_(status);
  }
  var s = (status || '').toString().trim().toLowerCase();
  if (!s) return true;
  return s === 'active' || s === 'aktif' || s === 'yes' || s === 'y' || s === '1' || s === 'true';
}

function ensureConfiguratorSheet_(sheetKey, headers) {
  var ss = getGlobalSpreadsheet();
  var name = SHEETS[sheetKey];
  var sheet = ss.getSheetByName(name);
  if (!sheet) {
    sheet = ss.insertSheet(name);
    sheet.appendRow(headers);
    Logger.log('Created ' + name + ' sheet');
    return sheet;
  }
  if (sheet.getLastRow() === 0) {
    sheet.appendRow(headers);
  }
  return sheet;
}

function ensureProductOptionsSheet_() {
  return ensureConfiguratorSheet_('PRODUCT_OPTIONS', SHEET_HEADER_ROWS.PRODUCT_OPTIONS);
}

function ensureProdOptionsSheet_() {
  return ensureConfiguratorSheet_('PROD_OPTIONS', SHEET_HEADER_ROWS.PROD_OPTIONS);
}

function ensureProductVariantsSheet_() {
  return ensureConfiguratorSheet_('PRODUCT_VARIANTS', SHEET_HEADER_ROWS.PRODUCT_VARIANTS);
}

function ensureProductConfiguratorSheets_() {
  ensureProductOptionsSheet_();
  ensureProdOptionsSheet_();
  ensureProductVariantsSheet_();
}

function findHeaderIndex_(headers, pattern) {
  for (var i = 0; i < headers.length; i++) {
    var h = headers[i];
    if (!h) continue;
    if (pattern.test(h.toString())) return i;
  }
  return -1;
}

/**
 * @returns {{ byCode: Object, byGroup: Object, optionCodes: Object, lines: Array }}
 */
function loadProductOptionsCatalog_() {
  ensureProductOptionsSheet_();
  var sheet = getGlobalSpreadsheet().getSheetByName(SHEETS.PRODUCT_OPTIONS);
  var byCode = {};
  var byGroup = {};
  var optionCodes = {};
  var lines = [];
  if (!sheet || sheet.getLastRow() < 2) {
    return { byCode: byCode, byGroup: byGroup, optionCodes: optionCodes, lines: lines };
  }

  var data = sheet.getDataRange().getValues();
  var headers = data[0];
  var codeIdx = findHeaderIndex_(headers, /^option\s*code$/i);
  var groupIdx = findHeaderIndex_(headers, /option\s*group\s*id/i);
  var nameIdx = findHeaderIndex_(headers, /display\s*name|^name$/i);
  var sortIdx = findHeaderIndex_(headers, /sort\s*order|^sort$/i);
  var priceAdjIdx = findHeaderIndex_(headers, /price\s*adjust|priceadjust/i);
  var typeIdx = findHeaderIndex_(headers, /option\s*type/i);
  var statusIdx = findHeaderIndex_(headers, /^status$/i);

  if (codeIdx === -1 || groupIdx === -1) {
    return { byCode: byCode, byGroup: byGroup, optionCodes: optionCodes, lines: lines };
  }

  for (var r = 1; r < data.length; r++) {
    var row = data[r];
    var optionCode = normalizeConfiguratorCode_(row[codeIdx]);
    var groupId = (row[groupIdx] || '').toString().trim();
    if (!optionCode && !groupId) continue;
    if (statusIdx >= 0 && !isConfiguratorRowActive_(row[statusIdx])) continue;
    if (!optionCode || !groupId) continue;

    var entry = {
      optionCode: optionCode,
      optionGroupId: groupId,
      displayName: nameIdx >= 0 ? (row[nameIdx] || optionCode).toString().trim() : optionCode,
      sortOrder: sortIdx >= 0 ? (Number(row[sortIdx]) || 0) : 0,
      priceAdjust: priceAdjIdx >= 0 ? (Number(row[priceAdjIdx]) || 0) : 0,
      optionType: typeIdx >= 0 ? (row[typeIdx] || '').toString().trim().toLowerCase() : ''
    };
    if (!entry.optionType && isMixTriggerOptionCode_(optionCode)) {
      entry.optionType = 'mix_trigger';
    } else if (!entry.optionType) {
      entry.optionType = 'single';
    }
    byCode[optionCode] = entry;
    optionCodes[optionCode] = true;
    if (!byGroup[groupId]) byGroup[groupId] = [];
    byGroup[groupId].push(entry);
    lines.push(entry);
  }

  Object.keys(byGroup).forEach(function(gid) {
    byGroup[gid].sort(function(a, b) {
      return (a.sortOrder || 0) - (b.sortOrder || 0) ||
        (a.displayName || '').localeCompare(b.displayName || '');
    });
  });

  return { byCode: byCode, byGroup: byGroup, optionCodes: optionCodes, lines: lines };
}

/**
 * @returns {{ byBase: Object, lines: Array }}
 */
function loadProdOptionsLinks_() {
  ensureProdOptionsSheet_();
  var sheet = getGlobalSpreadsheet().getSheetByName(SHEETS.PROD_OPTIONS);
  var byBase = {};
  var lines = [];
  if (!sheet || sheet.getLastRow() < 2) {
    return { byBase: byBase, lines: lines };
  }

  var data = sheet.getDataRange().getValues();
  var headers = data[0];
  var baseIdx = findHeaderIndex_(headers, /base\s*product\s*code/i);
  var groupIdx = findHeaderIndex_(headers, /option\s*group\s*id/i);
  var reqIdx = findHeaderIndex_(headers, /^required$/i);
  var sortIdx = findHeaderIndex_(headers, /sort\s*order|^sort$/i);
  var statusIdx = findHeaderIndex_(headers, /^status$/i);

  if (baseIdx === -1 || groupIdx === -1) {
    return { byBase: byBase, lines: lines };
  }

  for (var r = 1; r < data.length; r++) {
    var row = data[r];
    var baseCode = normalizeConfiguratorCode_(row[baseIdx]);
    var groupId = (row[groupIdx] || '').toString().trim();
    if (!baseCode && !groupId) continue;
    if (statusIdx >= 0 && !isConfiguratorRowActive_(row[statusIdx])) continue;
    if (!baseCode || !groupId) continue;

    var requiredRaw = reqIdx >= 0 ? (row[reqIdx] || '').toString().trim().toLowerCase() : 'y';
    var link = {
      baseProductCode: baseCode,
      optionGroupId: groupId,
      required: !requiredRaw || requiredRaw === 'y' || requiredRaw === 'yes' || requiredRaw === '1' || requiredRaw === 'true',
      sortOrder: sortIdx >= 0 ? (Number(row[sortIdx]) || 0) : 0
    };
    if (!byBase[baseCode]) byBase[baseCode] = [];
    byBase[baseCode].push(link);
    lines.push(link);
  }

  Object.keys(byBase).forEach(function(base) {
    byBase[base].sort(function(a, b) {
      return (a.sortOrder || 0) - (b.sortOrder || 0);
    });
  });

  return { byBase: byBase, lines: lines };
}

/**
 * @returns {{ byKey: Object, byVariantCode: Object, lines: Array }}
 */
function loadProductVariants_() {
  ensureProductVariantsSheet_();
  var optionsCatalog = loadProductOptionsCatalog_();
  var prodLinks = loadProdOptionsLinks_();
  var optionSortOrders = buildOptionCompositeSortOrders_(
    optionsCatalog.lines || [],
    prodLinks.lines || []
  );
  var knownCodes = getKnownOptionCodeList_(optionsCatalog);

  var sheet = getGlobalSpreadsheet().getSheetByName(SHEETS.PRODUCT_VARIANTS);
  var byKey = {};
  var byVariantCode = {};
  var lines = [];
  if (!sheet || sheet.getLastRow() < 2) {
    return { byKey: byKey, byVariantCode: byVariantCode, lines: lines };
  }

  var data = sheet.getDataRange().getValues();
  var headers = data[0];
  var variantIdx = findHeaderIndex_(headers, /variant\s*product\s*code/i);
  var baseIdx = findHeaderIndex_(headers, /base\s*product\s*code/i);
  var optionIdx = findHeaderIndex_(headers, /^option\s*code$/i);
  var nameIdx = findHeaderIndex_(headers, /display\s*name|^name$/i);
  var priceIdx = findHeaderIndex_(headers, /^price$/i);
  var statusIdx = findHeaderIndex_(headers, /^status$/i);

  if (variantIdx === -1 || baseIdx === -1 || optionIdx === -1) {
    return { byKey: byKey, byVariantCode: byVariantCode, lines: lines };
  }

  for (var r = 1; r < data.length; r++) {
    var row = data[r];
    var variantCode = normalizeConfiguratorCode_(row[variantIdx]);
    var baseCode = normalizeConfiguratorCode_(row[baseIdx]);
    var optionCodeRaw = (row[optionIdx] || '').toString().trim();
    var optionCode = normalizeCompositeOptionCode_(optionCodeRaw, optionSortOrders, knownCodes);
    if (!variantCode && !baseCode && !optionCode) continue;
    if (statusIdx >= 0 && !isConfiguratorRowActive_(row[statusIdx])) continue;
    if (!variantCode || !baseCode || !optionCode) continue;

    var priceRaw = priceIdx >= 0 ? row[priceIdx] : '';
    var price = (priceRaw === '' || priceRaw === null || priceRaw === undefined)
      ? null
      : (Number(priceRaw) || 0);

    var variant = {
      variantProductCode: variantCode,
      baseProductCode: baseCode,
      optionCode: optionCode,
      displayName: nameIdx >= 0 ? (row[nameIdx] || '').toString().trim() : '',
      price: price
    };
    var key = baseCode + '|' + optionCode;
    byKey[key] = variant;
    byVariantCode[variantCode] = variant;
    lines.push(variant);
  }

  return { byKey: byKey, byVariantCode: byVariantCode, lines: lines };
}

function loadBaseProductMeta_(baseCodes, outlet, options) {
  options = options || {};
  var skipOutletFilter = !!options.skipOutletFilter;

  var meta = {};
  if (!baseCodes || !baseCodes.length) return meta;

  var sheet = getGlobalSpreadsheet().getSheetByName(SHEETS.PRODUCTS);
  if (!sheet || sheet.getLastRow() < 2) return meta;

  var data = sheet.getDataRange().getValues();
  var headers = data[0].map(function(h) { return (h || '').toString().trim(); });
  var codeIdx = headers.findIndex(function(h) { return /^productcode$|^code$/i.test(h); });
  var nameIdx = headers.findIndex(function(h) { return /^name$/i.test(h); });
  var priceIdx = headers.findIndex(function(h) { return /^price$/i.test(h); });
  var catIdx = headers.findIndex(function(h) { return /^category$/i.test(h); });
  var navIdx = headers.findIndex(function(h) { return /^navcategory$/i.test(h); });
  var imgIdx = headers.findIndex(function(h) { return /^imagefileid$/i.test(h); });
  var outletIdx = headers.findIndex(function(h) { return /^outlet$/i.test(h); });
  var activeIdx = headers.findIndex(function(h) { return /^active$/i.test(h); });
  var typeIdx = headers.findIndex(function(h) { return /^type$/i.test(h); });

  var wanted = {};
  baseCodes.forEach(function(c) { wanted[normalizeConfiguratorCode_(c)] = true; });

  for (var i = 1; i < data.length; i++) {
    var row = data[i];
    if (activeIdx >= 0 && typeof isProductActive_ === 'function' && !isProductActive_(row[activeIdx])) continue;
    if (typeIdx >= 0) {
      var typeVal = (row[typeIdx] || 'Sell').toString().trim().toLowerCase();
      if (typeVal && typeVal !== 'sell') continue;
    }
    var code = codeIdx >= 0 ? normalizeConfiguratorCode_(row[codeIdx]) : '';
    if (!code || !wanted[code]) continue;

    if (!skipOutletFilter && outlet && outletIdx >= 0) {
      var rowOutlet = (row[outletIdx] || '').toString().trim();
      if (rowOutlet && rowOutlet.toLowerCase().indexOf('all') < 0) {
        var list = rowOutlet.split(',').map(function(o) { return o.trim().toLowerCase(); });
        if (list.indexOf(outlet.toString().trim().toLowerCase()) < 0) continue;
      }
    }

    meta[code] = {
      productCode: code,
      name: nameIdx >= 0 ? (row[nameIdx] || code).toString().trim() : code,
      price: priceIdx >= 0 ? (Number(row[priceIdx]) || 0) : 0,
      category: catIdx >= 0 ? (row[catIdx] || '').toString().trim() : '',
      navCategory: navIdx >= 0 ? (row[navIdx] || '').toString().trim() : '',
      imageFileId: imgIdx >= 0 ? (row[imgIdx] || '').toString().trim() : ''
    };
  }

  return meta;
}

/**
 * POS payload: bases with option groups, variant lookup, codes to hide from menu.
 * @param {string} [outlet]
 */
function getProductConfiguratorData_(outlet) {
  ensureProductConfiguratorSheets_();

  var options = loadProductOptionsCatalog_();
  var links = loadProdOptionsLinks_();
  var variants = loadProductVariants_();

  var baseCodes = Object.keys(links.byBase || {});
  var baseMeta = loadBaseProductMeta_(baseCodes, outlet, { skipOutletFilter: true });

  var bases = {};
  baseCodes.forEach(function(baseCode) {
    var meta = baseMeta[baseCode] || {
      productCode: baseCode,
      name: baseCode,
      price: 0,
      category: '',
      navCategory: '',
      imageFileId: ''
    };

    var groups = (links.byBase[baseCode] || []).map(function(link) {
      var allOpts = (options.byGroup[link.optionGroupId] || []).map(function(opt) {
        return {
          optionCode: opt.optionCode,
          displayName: opt.displayName,
          sortOrder: opt.sortOrder,
          priceAdjust: opt.priceAdjust,
          optionType: opt.optionType || 'single'
        };
      });
      var singles = allOpts.filter(function(opt) { return opt.optionType !== 'mix_trigger'; });
      var mixTriggers = allOpts.filter(function(opt) { return opt.optionType === 'mix_trigger'; });
      return {
        optionGroupId: link.optionGroupId,
        required: link.required !== false,
        sortOrder: link.sortOrder,
        options: allOpts,
        singles: singles,
        mixTriggers: mixTriggers,
        mixPickCount: 2
      };
    }).filter(function(g) { return g.options && g.options.length; });

    if (!groups.length) return;

    bases[baseCode] = {
      productCode: baseCode,
      name: meta.name,
      price: meta.price,
      category: meta.category,
      navCategory: meta.navCategory || meta.category,
      imageFileId: meta.imageFileId,
      groups: groups
    };
  });

  var knownCodes = getKnownOptionCodeList_(options);

  var variantMap = {};
  Object.keys(variants.byKey || {}).forEach(function(key) {
    var v = variants.byKey[key];
    var base = bases[v.baseProductCode];
    var optionParts = splitCompositeOptionCode_(v.optionCode, knownCodes);
    var price = v.price;
    if (price === null || price === undefined) {
      var basePrice = base ? (Number(base.price) || 0) : 0;
      var adj = 0;
      if (optionParts.length > 1) {
        optionParts.forEach(function(part) {
          var optPart = options.byCode[normalizeConfiguratorCode_(part)];
          if (optPart) adj += Number(optPart.priceAdjust) || 0;
        });
      } else {
        var opt = options.byCode[v.optionCode];
        adj = opt ? (Number(opt.priceAdjust) || 0) : 0;
      }
      price = basePrice + adj;
    }
    var displayName = v.displayName;
    if (!displayName) {
      if (optionParts.length > 1 && isCompositeOptionMix_(optionParts, options.byCode)) {
        var mixNames = optionParts.map(function(part) {
          var optPart = options.byCode[normalizeConfiguratorCode_(part)];
          return optPart ? optPart.displayName : part;
        });
        displayName = (base ? base.name : v.baseProductCode) + ' \u2013 Mix (' + mixNames.join(' & ') + ')';
      } else if (optionParts.length > 1) {
        var partNames = optionParts.map(function(part) {
          var optPart = options.byCode[normalizeConfiguratorCode_(part)];
          return optPart ? optPart.displayName : part;
        });
        displayName = (base ? base.name : v.baseProductCode) + ' \u2013 ' + partNames.join(', ');
      } else {
        var optSingle = options.byCode[v.optionCode];
        displayName = (base ? base.name : v.baseProductCode) + ' \u2013 ' +
          (optSingle ? optSingle.displayName : v.optionCode);
      }
    }

    variantMap[key] = {
      variantProductCode: v.variantProductCode,
      baseProductCode: v.baseProductCode,
      optionCode: v.optionCode,
      displayName: displayName,
      price: price,
      isMix: isCompositeOptionMix_(optionParts, options.byCode)
    };
  });

  var optionSortOrders = buildOptionCompositeSortOrders_(
    options.lines || [],
    links.lines || []
  );

  return {
    bases: bases,
    variants: variantMap,
    hideOptionCodes: options.optionCodes || {},
    optionSortOrders: optionSortOrders,
    hasConfigurator: Object.keys(bases).length > 0
  };
}

function getProductConfiguratorData(outlet) {
  return JSON.stringify(getProductConfiguratorData_(outlet || ''));
}

function getProductConfiguratorSecure(params) {
  var authResult = validateAuthenticationContext(params._authContext);
  var outlet = authResult.outlet;
  if (typeof resolveUserOutlet_ === 'function') {
    outlet = resolveUserOutlet_(authResult.username, outlet);
  }
  try {
    return getProductConfiguratorData(outlet || '');
  } catch (e) {
    Logger.log('getProductConfiguratorSecure error: ' + e.toString());
    throw e;
  }
}

// ========== GENERATE ProductVariants FROM ProductOptions + ProdOptions ==========

function cartesianChoiceProduct_(choiceSets) {
  if (!choiceSets.length) return [{ choices: [] }];
  return choiceSets.reduce(function(acc, set) {
    var result = [];
    acc.forEach(function(prev) {
      set.forEach(function(choice) {
        result.push({ choices: prev.choices.concat([choice]) });
      });
    });
    return result;
  }, [{ choices: [] }]);
}

/**
 * Choices for one option group: each single, plus mix pairs when mix_trigger exists.
 */
function getOptionGroupChoices_(groupId, optionsCatalog, sortByCode) {
  var opts = optionsCatalog.byGroup[groupId] || [];
  var singles = opts.filter(function(o) { return o.optionType !== 'mix_trigger'; });
  var hasMix = opts.some(function(o) { return o.optionType === 'mix_trigger'; });
  var choices = [];

  singles.forEach(function(s) {
    choices.push({
      codes: [normalizeConfiguratorCode_(s.optionCode)],
      label: s.displayName || s.optionCode,
      isMix: false
    });
  });

  if (hasMix && singles.length >= 2) {
    for (var i = 0; i < singles.length; i++) {
      for (var j = i + 1; j < singles.length; j++) {
        var a = singles[i];
        var b = singles[j];
        var codes = sortCompositeOptionParts_([
          normalizeConfiguratorCode_(a.optionCode),
          normalizeConfiguratorCode_(b.optionCode)
        ], sortByCode);
        var nameByCode = {};
        nameByCode[normalizeConfiguratorCode_(a.optionCode)] = a.displayName || a.optionCode;
        nameByCode[normalizeConfiguratorCode_(b.optionCode)] = b.displayName || b.optionCode;
        var mixNames = codes.map(function(c) { return nameByCode[c]; }).join(' & ');
        choices.push({
          codes: codes,
          label: 'Mix (' + mixNames + ')',
          isMix: true
        });
      }
    }
  }

  return choices;
}

function buildVariantProductCode_(baseCode, optionCode, sortByCode, knownCodes) {
  var normalizedOption = normalizeCompositeOptionCode_(optionCode, sortByCode, knownCodes);
  return normalizeConfiguratorCode_(baseCode) + '_' + normalizedOption;
}

/**
 * Read ProductOptions + ProdOptions and build ProductVariants rows (all group combinations).
 */
function buildProductVariantRowsFromConfigurator_() {
  ensureProductConfiguratorSheets_();

  var optionsCatalog = loadProductOptionsCatalog_();
  var prodLinks = loadProdOptionsLinks_();
  var sortByCode = buildOptionCompositeSortOrders_(
    optionsCatalog.lines || [],
    prodLinks.lines || []
  );
  var knownCodes = getKnownOptionCodeList_(optionsCatalog);

  var baseCodes = Object.keys(prodLinks.byBase || {});
  if (!baseCodes.length) return [];

  var baseMeta = loadBaseProductMeta_(baseCodes, '', { skipOutletFilter: true });
  var rows = [];
  var seenKeys = {};

  baseCodes.forEach(function(baseCode) {
    var links = (prodLinks.byBase[baseCode] || []).slice().sort(function(a, b) {
      return (a.sortOrder || 0) - (b.sortOrder || 0);
    });

    var seenGroups = {};
    var groupIds = [];
    links.forEach(function(link) {
      var gid = (link.optionGroupId || '').toString().trim();
      if (!gid || seenGroups[gid]) return;
      seenGroups[gid] = true;
      groupIds.push(gid);
    });
    if (!groupIds.length) return;

    var choiceSets = groupIds.map(function(gid) {
      return getOptionGroupChoices_(gid, optionsCatalog, sortByCode);
    }).filter(function(choices) { return choices.length > 0; });

    if (!choiceSets.length || choiceSets.some(function(choices) { return !choices.length; })) {
      Logger.log('buildProductVariantRowsFromConfigurator_: skip base ' + baseCode + ' (missing option choices)');
      return;
    }

    var baseLabel = (baseMeta[baseCode] && baseMeta[baseCode].name) || baseCode;
    var combinations = cartesianChoiceProduct_(choiceSets);

    combinations.forEach(function(combo) {
      var allCodes = [];
      combo.choices.forEach(function(choice) {
        allCodes = allCodes.concat(choice.codes);
      });
      var optionKey = joinCompositeOptionParts_(allCodes, sortByCode);
      var dedupeKey = baseCode + '|' + optionKey;
      if (seenKeys[dedupeKey]) return;
      seenKeys[dedupeKey] = true;

      var labelParts = combo.choices.map(function(c) { return c.label; });
      var displayName = baseLabel + ' \u2013 ' + labelParts.join(', ');

      rows.push([
        buildVariantProductCode_(baseCode, optionKey, sortByCode, knownCodes),
        baseCode,
        optionKey,
        displayName,
        '',
        'Active'
      ]);
    });
  });

  return rows;
}

function writeConfiguratorSheetRows_(sheetKey, headers, rows, force) {
  var sheet = ensureConfiguratorSheet_(sheetKey, headers);
  if (!sheet) return 0;
  if (sheet.getLastRow() > 1 && !force) {
    Logger.log(SHEETS[sheetKey] + ': skipped (already has data). Pass { force: true } to replace.');
    return 0;
  }
  if (force && sheet.getLastRow() > 1) {
    sheet.getRange(2, 1, sheet.getLastRow() - 1, sheet.getLastColumn()).clearContent();
  }
  if (!rows.length) return 0;
  if (typeof appendRowsAtomic_ === 'function') {
    appendRowsAtomic_(sheet, rows);
  } else {
    rows.forEach(function(r) { sheet.appendRow(r); });
  }
  return rows.length;
}

/**
 * Regenerate ProductVariants from existing ProductOptions + ProdOptions sheets.
 * Pass { force: true } to replace existing variant rows.
 */
function seedProductVariantsFromConfigurator_(options) {
  options = options || {};
  var force = !!options.force;

  ensureProductConfiguratorSheets_();
  var variantRows = buildProductVariantRowsFromConfigurator_();
  if (!variantRows.length) {
    throw new Error('No ProductVariants generated. Add data to ProductOptions and ProdOptions first.');
  }

  var nVar = writeConfiguratorSheetRows_(
    'PRODUCT_VARIANTS',
    SHEET_HEADER_ROWS.PRODUCT_VARIANTS,
    variantRows,
    force
  );

  var msg = 'ProductVariants generated: ' + nVar + ' rows from ProductOptions + ProdOptions';
  Logger.log(msg);
  return {
    productOptions: 0,
    prodOptions: 0,
    productVariants: nVar,
    bases: Object.keys(loadProdOptionsLinks_().byBase || {}).length,
    message: msg
  };
}

/** @deprecated Use seedProductVariantsFromConfigurator_ */
function seedEsKrimConfigurator_(options) {
  return seedProductVariantsFromConfigurator_(options);
}

/**
 * Admin: regenerate ProductVariants from ProductOptions + ProdOptions.
 */
function seedProductVariantsFromConfiguratorSecure(params) {
  params = params || {};
  if (typeof requireAdminOrdersSession_ === 'function') {
    requireAdminOrdersSession_(params);
  }
  var result = seedProductVariantsFromConfigurator_({
    force: !!params.force
  });
  return { success: true, data: result };
}

/** @deprecated Use seedProductVariantsFromConfiguratorSecure */
function seedEsKrimConfiguratorSecure(params) {
  return seedProductVariantsFromConfiguratorSecure(params);
}
