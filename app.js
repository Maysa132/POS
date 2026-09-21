/* ====================================================================
   ร้านมัทฉะในฝัน — ระบบ POS
   ลอจิกเดียวกับไฟล์ ร้านมัทฉะในฝัน.xlsx
     ต้นทุน/แก้ว = ผลรวม(ปริมาณ × ต้นทุนต่อหน่วยวัตถุดิบ)
     ค่า GP      = ยอดขาย × %GP ของช่องทาง
     ยอดสุทธิ    = ยอดขาย − ค่า GP
     กำไรสุทธิ   = ยอดสุทธิ − ต้นทุน
     ตัดสต็อก    = ปริมาณในสูตร × จำนวนแก้วที่ขาย

   สต็อกคิดเป็นล็อต — เลือกลำดับการหยิบใช้ได้ที่หน้าตั้งค่า
     FEFO (ค่าเริ่มต้น) = ใช้ล็อตที่จะหมดอายุก่อน เหมาะกับผงมัทฉะและนม
     FIFO                = ใช้ล็อตที่รับเข้ามาก่อน
     ทุกครั้งที่รับของเข้าจะสร้าง “ล็อต” ที่จำราคาซื้อและวันหมดอายุของรอบนั้นไว้
     ต้นทุนของบิลจึงเป็นราคาจริงของล็อตที่ถูกใช้
     ถ้าของในล็อตหัวคิวไม่พอ จะไล่ไปล็อตถัดไปโดยอัตโนมัติ (บิลเดียวอาจมีหลายราคา)

   กำไรจริง = กำไรขั้นต้น − ค่าใช้จ่ายคงที่เฉลี่ยต่อวัน (ค่าเช่า ค่าแรง ค่าไฟ ฯลฯ)
   ==================================================================== */
(function () {
'use strict';

/* ------------------------------------------------ helpers */
var $  = function (s, r) { return (r || document).querySelector(s); };
var $$ = function (s, r) { return Array.prototype.slice.call((r || document).querySelectorAll(s)); };

var seq = 0;
function uid(p) { seq += 1; return (p || 'id') + '_' + Date.now().toString(36) + seq.toString(36); }
function r2(n) { return Math.round((Number(n) + Number.EPSILON) * 100) / 100; }
function rq(n) { return Math.round((Number(n) + Number.EPSILON) * 1e6) / 1e6; }   /* ปริมาณ — ปัดละเอียดกว่าเงิน */
function baht(n) { return r2(n).toLocaleString('th-TH', { minimumFractionDigits: 2, maximumFractionDigits: 2 }); }
function qtyStr(n) { return r2(n).toLocaleString('th-TH', { maximumFractionDigits: 2 }); }
/** ติดลบให้ขึ้นเครื่องหมายหน้าสัญลักษณ์เงิน เช่น −฿1,200.00 */
function signedBaht(n) { return (r2(n) < 0 ? '−฿' : '฿') + baht(Math.abs(n)); }
function esc(s) {
  return String(s == null ? '' : s).replace(/[&<>"']/g, function (c) {
    return { '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c];
  });
}
function num(v, dflt) { var n = parseFloat(v); return isFinite(n) ? n : (dflt || 0); }

function dayKey(ts) {
  var d = new Date(ts);
  return d.getFullYear() + '-' + String(d.getMonth() + 1).padStart(2, '0') + '-' + String(d.getDate()).padStart(2, '0');
}
function fmtDateTime(ts) {
  return new Date(ts).toLocaleString('th-TH', { day: '2-digit', month: 'short', year: '2-digit', hour: '2-digit', minute: '2-digit' });
}
function fmtDayShort(key) {
  var p = key.split('-');
  return Number(p[2]) + '/' + Number(p[1]);
}
function fmtDayFull(key) {
  var p = key.split('-');
  return new Date(Number(p[0]), Number(p[1]) - 1, Number(p[2]))
    .toLocaleDateString('th-TH', { weekday: 'long', day: 'numeric', month: 'long', year: 'numeric' });
}

function toast(msg, danger) {
  var el = document.createElement('div');
  el.className = 'toast' + (danger ? ' is-danger' : '');
  el.textContent = msg;
  $('#toastWrap').appendChild(el);
  setTimeout(function () { el.remove(); }, 2600);
}

function download(filename, content, mime) {
  var blob = new Blob([content], { type: mime || 'text/plain;charset=utf-8' });
  var url = URL.createObjectURL(blob);
  var a = document.createElement('a');
  a.href = url; a.download = filename;
  document.body.appendChild(a); a.click();
  setTimeout(function () { URL.revokeObjectURL(url); a.remove(); }, 800);
}

/* ------------------------------------------------ storage */
var KEY = 'matcha_pos_v1';
var db = null;

function save() {
  try { localStorage.setItem(KEY, JSON.stringify(db)); }
  catch (e) { toast('บันทึกข้อมูลลงเครื่องไม่สำเร็จ', true); }
}
function loadStored() {
  try {
    var raw = localStorage.getItem(KEY);
    return raw ? JSON.parse(raw) : null;
  } catch (e) { return null; }
}

function buildFresh() {
  var s = JSON.parse(JSON.stringify(window.SEED));
  var d = {
    version: 4,
    shop: s.shop,
    platforms: s.platforms,
    payMethods: s.payMethods,
    fixedCosts: s.fixedCosts,
    ingredients: s.ingredients,
    menus: s.menus,
    sales: [],
    moves: [],
    lots: [],
    closings: [],
    billSeq: 0,
    lotSeq: 0,
    closeSeq: 0
  };
  /* คลี่ตัวเลือกมาตรฐานเข้าไปในแต่ละเมนู เพื่อให้แก้ไขแยกเมนูได้ */
  d.menus.forEach(function (m) {
    m.options = (m.optionKeys || []).map(function (k) {
      return s.optionPresets[k] ? JSON.parse(JSON.stringify(s.optionPresets[k])) : null;
    }).filter(Boolean);
    delete m.optionKeys;
  });
  var base = Date.now() - 86400000;
  d.ingredients.forEach(function (i) {
    if (i.openingQty) {
      receiveStock(d, i.id, i.openingQty, costPerUnit(i), 'สต็อกตั้งต้น (จากไฟล์ Excel)', base,
        num(i.shelfLifeDays) > 0 ? addDays(num(i.shelfLifeDays) - 1) : null);
    }
    delete i.openingQty;
  });
  (s.seedSales || []).forEach(function (row) {
    commitSale(d, {
      platformId: row.platformId,
      payMethodId: row.payMethodId,
      items: row.items,
      discount: 0,
      ts: Date.now() - row.minutesAgo * 60000,
      note: 'ข้อมูลจากชีต Sale'
    });
  });
  return d;
}

function migrate(d) {
  if (!d || typeof d !== 'object') return null;
  if (!Array.isArray(d.sales)) d.sales = [];
  if (!Array.isArray(d.moves)) d.moves = [];
  if (!Array.isArray(d.menus) || !Array.isArray(d.ingredients) || !Array.isArray(d.platforms)) return null;
  if (typeof d.billSeq !== 'number') d.billSeq = d.sales.length;
  if (!d.shop) d.shop = { name: 'ร้านมัทฉะในฝัน', tagline: '' };
  if (!Array.isArray(d.lots) || d.version < 2) rebuildLots(d);
  if (d.version < 3) upgradeToV3(d);
  if (d.version < 4) upgradeToV4(d);
  return d;
}

/** เวอร์ชัน 4 — วันหมดอายุต่อล็อต (FEFO) และค่าใช้จ่ายคงที่ */
function upgradeToV4(d) {
  if (!d.shop.stockMethod) d.shop.stockMethod = 'fefo';
  if (d.shop.expiryWarnDays == null) d.shop.expiryWarnDays = 14;
  if (d.shop.openDaysPerMonth == null) d.shop.openDaysPerMonth = 26;
  if (!Array.isArray(d.fixedCosts)) d.fixedCosts = JSON.parse(JSON.stringify(window.SEED.fixedCosts));

  /* เติมอายุสินค้าให้วัตถุดิบเดิมจากค่าตั้งต้น ถ้าชื่อตรงกัน */
  var seedIng = {};
  window.SEED.ingredients.forEach(function (i) { seedIng[i.id] = i; });
  d.ingredients.forEach(function (i) {
    if (i.shelfLifeDays == null) i.shelfLifeDays = seedIng[i.id] ? num(seedIng[i.id].shelfLifeDays) : 0;
  });

  /* ล็อตเดิมไม่มีวันหมดอายุ — ปล่อยว่างไว้ (ไปต่อท้ายคิว FEFO) ให้ผู้ใช้มาเติมเอง */
  d.lots.forEach(function (l) { if (l.expiry === undefined) l.expiry = null; });

  d.version = 4;
}

/** เวอร์ชัน 3 — เพิ่ม VAT, วิธีชำระเงิน, ตัวเลือกเมนู และการปิดยอดรายวัน */
function upgradeToV3(d) {
  if (!d.shop.vat) d.shop.vat = { enabled: true, rate: 7, mode: 'inclusive' };
  if (d.shop.openingFloat == null) d.shop.openingFloat = 1000;
  if (!Array.isArray(d.payMethods)) d.payMethods = JSON.parse(JSON.stringify(window.SEED.payMethods));
  if (!Array.isArray(d.closings)) d.closings = [];
  if (typeof d.closeSeq !== 'number') d.closeSeq = 0;

  d.platforms.forEach(function (p) {
    if (!p.settlement) p.settlement = num(p.gp) > 0 ? 'platform' : 'direct';
  });
  /* เมนูเดิมยังไม่มีตัวเลือก — ใส่ชุดมาตรฐานให้เมนูที่เข้าเงื่อนไข (มีน้ำเชื่อม/ผงมัทฉะในสูตร) */
  d.menus.forEach(function (m) {
    if (Array.isArray(m.options) && m.options.length) return;
    m.options = ['sweet', 'powder', 'strength']
      .map(function (k) { return buildPresetGroup(d, m, k); })
      .filter(Boolean);
  });

  /* บิลเดิมไม่มี VAT — ถือว่าราคาที่บันทึกไว้คือยอดที่ลูกค้าจ่ายและไม่มีภาษี
     (บิลที่ขายไปแล้วไม่ถูกคิด VAT ย้อนหลัง ตัวเลขเดิมจึงไม่เปลี่ยน) */
  d.sales.forEach(function (s) {
    if (s.payable == null) s.payable = s.afterDiscount;
    if (s.vat == null) s.vat = 0;
    if (s.exVat == null) s.exVat = s.afterDiscount;
    if (!s.settlement) {
      var pl = findPlatform(d, s.platformId);
      s.settlement = pl ? pl.settlement : 'direct';
    }
    if (!s.payMethodId) s.payMethodId = s.settlement === 'direct' ? 'cash' : 'platform';
    if (!s.payMethodName) {
      var pm = payMethod(d, s.payMethodId);
      s.payMethodName = pm ? pm.name : 'แพลตฟอร์ม';
    }
    (s.items || []).forEach(function (it) { if (!it.selections) it.selections = {}; });
  });

  d.version = 3;
}

/**
 * แปลงข้อมูลเวอร์ชันเดิม (ไม่มีล็อต) มาเป็น FIFO
 * ไล่รายการเคลื่อนไหวตามเวลา: รับเข้า → สร้างล็อต, จ่ายออก → ตัดแบบ FIFO
 * แล้วคำนวณต้นทุน/กำไรของบิลเก่าใหม่ตามล็อตที่ถูกใช้จริง
 */
function rebuildLots(d) {
  d.lots = [];
  d.lotSeq = 0;
  var ordered = d.moves.slice().sort(function (a, b) { return a.ts - b.ts; });
  var rebuilt = [];
  var costBySale = {};

  ordered.forEach(function (m) {
    if (m.qty > 0) {
      d.lotSeq += 1;
      var unitCost = m.unitCost != null ? m.unitCost : costPerUnit(findIng(d, m.ingId));
      var lot = {
        id: uid('lot'), ingId: m.ingId, ts: m.ts, seq: d.lotSeq,
        qtyIn: rq(m.qty), qtyLeft: rq(m.qty), unitCost: unitCost, note: m.note || '', moveId: m.id
      };
      d.lots.push(lot);
      m.lotId = lot.id;
      m.unitCost = unitCost;
      m.cost = r2(m.qty * unitCost);
    } else if (m.qty < 0) {
      var take = makeLedger(d);
      var res = take(m.ingId, -m.qty, true);
      m.allocs = res.allocs;
      m.cost = res.cost;
      if (m.ref) costBySale[m.ref] = r2((costBySale[m.ref] || 0) + res.cost);
    }
    rebuilt.push(m);
  });
  d.moves = rebuilt;

  d.sales.forEach(function (s) {
    if (costBySale[s.id] == null) return;
    var old = s.cost || 0;
    s.cost = costBySale[s.id];
    (s.items || []).forEach(function (it) {       /* เกลี่ยต้นทุนใหม่ตามสัดส่วนเดิมของแต่ละรายการ */
      it.lineCost = old > 0 ? r2(it.lineCost / old * s.cost) : r2(s.cost / Math.max(s.items.length, 1));
      it.unitCost = it.qty ? r2(it.lineCost / it.qty) : 0;
    });
    s.profit = r2(s.net - s.cost);
  });

  d.version = 2;
}

/* ------------------------------------------------ calculations */
function findIng(d, id) { for (var i = 0; i < d.ingredients.length; i++) if (d.ingredients[i].id === id) return d.ingredients[i]; return null; }
function findMenu(d, id) { for (var i = 0; i < d.menus.length; i++) if (d.menus[i].id === id) return d.menus[i]; return null; }
function findPlatform(d, id) { for (var i = 0; i < d.platforms.length; i++) if (d.platforms[i].id === id) return d.platforms[i]; return d.platforms[0]; }
function payMethod(d, id) {
  var list = d.payMethods || [];
  for (var i = 0; i < list.length; i++) if (list[i].id === id) return list[i];
  return null;
}
function isDirect(plat) { return !plat || plat.settlement !== 'platform'; }

function costPerUnit(ing) { return ing && ing.qtyPerPurchase > 0 ? ing.purchasePrice / ing.qtyPerPurchase : 0; }

/** ต้นทุนตามราคาซื้อล่าสุดที่ตั้งไว้ (ใช้อ้างอิง/เปรียบเทียบ ไม่ใช่ต้นทุนจริงที่ตัดขาย) */
function menuCostStd(d, menu) {
  if (!menu || !menu.recipe) return 0;
  return menu.recipe.reduce(function (sum, line) {
    return sum + line.qty * costPerUnit(findIng(d, line.ingId));
  }, 0);
}

/* ================================================================
   FIFO — ล็อตสินค้า
   lot = { id, ingId, ts, seq, qtyIn, qtyLeft, unitCost, note, moveId }
   ================================================================ */

/* ลำดับการหยิบใช้
     FIFO = เรียงตามวันที่รับเข้า
     FEFO = เรียงตามวันหมดอายุ (ล็อตที่ไม่มีวันหมดอายุไปอยู่ท้ายแถว) แล้วค่อยดูวันรับเข้า */
function stockMethod(d) { return (d && d.shop && d.shop.stockMethod) === 'fifo' ? 'fifo' : 'fefo'; }
function methodLabel(d) { return stockMethod(d) === 'fifo' ? 'FIFO' : 'FEFO'; }

function lotOrder(a, b) { return (a.ts - b.ts) || ((a.seq || 0) - (b.seq || 0)); }

function lotOrderFor(d) {
  if (stockMethod(d) === 'fifo') return lotOrder;
  return function (a, b) {
    var ea = a.expiry || '9999-99-99';
    var eb = b.expiry || '9999-99-99';
    if (ea !== eb) return ea < eb ? -1 : 1;
    return lotOrder(a, b);
  };
}

function lotsOf(d, ingId) {
  return d.lots.filter(function (l) { return l.ingId === ingId; }).sort(lotOrderFor(d));
}

/* ---- วันหมดอายุ ---- */
function todayKey() { return dayKey(Date.now()); }
function daysUntil(dateKey) {
  if (!dateKey) return null;
  var p = dateKey.split('-');
  var target = new Date(Number(p[0]), Number(p[1]) - 1, Number(p[2])).getTime();
  var now = new Date();
  var today = new Date(now.getFullYear(), now.getMonth(), now.getDate()).getTime();
  return Math.round((target - today) / 86400000);
}
function addDays(days) {
  var d = new Date();
  d.setDate(d.getDate() + Number(days || 0));
  return dayKey(d.getTime());
}
function lotExpiryState(d, lot) {
  if (!lot.expiry) return 'none';
  var left = daysUntil(lot.expiry);
  if (left < 0) return 'expired';
  if (left <= num(d.shop.expiryWarnDays, 14)) return 'soon';
  return 'ok';
}
/** ล็อตที่หมดอายุแล้วหรือใกล้หมดอายุ และยังมีของเหลือ */
function expiringLots(d) {
  var out = [];
  d.lots.forEach(function (l) {
    if (l.qtyLeft <= 1e-9) return;
    var st = lotExpiryState(d, l);
    if (st === 'expired' || st === 'soon') out.push(l);
  });
  return out.sort(function (a, b) { return (a.expiry || '') < (b.expiry || '') ? -1 : 1; });
}
function openLots(d, ingId) {
  return lotsOf(d, ingId).filter(function (l) { return l.qtyLeft > 1e-9; });
}

/** ราคาที่ใช้เมื่อของหมดล็อตแล้วยังต้องตัดต่อ — ยึดราคาล็อตล่าสุด ถ้าไม่มีเลยใช้ราคาซื้อที่ตั้งไว้ */
function fallbackCost(d, ingId) {
  var all = lotsOf(d, ingId);
  if (all.length) return all[all.length - 1].unitCost;
  return costPerUnit(findIng(d, ingId));
}

/** ต้นทุนของหน่วยถัดไปที่จะถูกหยิบใช้ (ล็อตหัวคิว) */
function nextUnitCost(d, ingId) {
  var open = openLots(d, ingId);
  return open.length ? open[0].unitCost : fallbackCost(d, ingId);
}

/** มูลค่าสต็อกคงเหลือตามราคาจริงของแต่ละล็อต */
function stockValue(d, ingId) {
  return r2(openLots(d, ingId).reduce(function (s, l) { return s + l.qtyLeft * l.unitCost; }, 0));
}
function lotQtyLeft(d, ingId) {
  return rq(openLots(d, ingId).reduce(function (s, l) { return s + l.qtyLeft; }, 0));
}
/** ต้นทุนถัวเฉลี่ยของของที่ยังอยู่ในสต็อก */
function avgCost(d, ingId) {
  var q = lotQtyLeft(d, ingId);
  return q > 0 ? stockValue(d, ingId) / q : fallbackCost(d, ingId);
}

/**
 * สร้าง “กระดานทด” ของล็อตคงเหลือ แล้วคืนฟังก์ชัน take()
 *   take(ingId, need, commit) → { allocs, cost, shortage }
 * commit = true จะตัดยอดล็อตจริง, false = แค่คำนวณล่วงหน้า (ใช้ตอนพรีวิวตะกร้า)
 * กระดานเดียวกันหยิบต่อเนื่องได้ ทำให้ยอดพรีวิวตรงกับตอนบันทึกจริงเสมอ
 */
function makeLedger(d) {
  var pools = {};
  var order = lotOrderFor(d);
  d.lots.forEach(function (l) {
    if (l.qtyLeft <= 1e-9) return;
    (pools[l.ingId] = pools[l.ingId] || []).push({ lot: l, left: l.qtyLeft });
  });
  Object.keys(pools).forEach(function (k) {
    pools[k].sort(function (a, b) { return order(a.lot, b.lot); });
  });

  return function take(ingId, need, commit) {
    var arr = pools[ingId] || [];
    var allocs = [], cost = 0, remain = rq(need), shortage = 0;
    for (var i = 0; i < arr.length && remain > 1e-9; i++) {
      var t = Math.min(arr[i].left, remain);
      if (t <= 1e-9) continue;
      arr[i].left = rq(arr[i].left - t);
      if (commit) arr[i].lot.qtyLeft = arr[i].left;
      allocs.push({ lotId: arr[i].lot.id, qty: rq(t), unitCost: arr[i].lot.unitCost });
      cost += t * arr[i].lot.unitCost;
      remain = rq(remain - t);
    }
    if (remain > 1e-9) {                       /* ของไม่พอ — คิดต้นทุนด้วยราคาล็อตล่าสุด */
      var fb = fallbackCost(d, ingId);
      allocs.push({ lotId: null, qty: rq(remain), unitCost: fb });
      cost += remain * fb;
      shortage = rq(remain);
    }
    return { allocs: allocs, cost: r2(cost), shortage: shortage };
  };
}

/** คืนของกลับเข้าล็อตเดิม (ใช้ตอนลบบิล) */
function releaseAllocs(d, allocs) {
  (allocs || []).forEach(function (a) {
    if (!a.lotId) return;
    for (var i = 0; i < d.lots.length; i++) {
      if (d.lots[i].id === a.lotId) {
        d.lots[i].qtyLeft = rq(Math.min(d.lots[i].qtyIn, d.lots[i].qtyLeft + a.qty));
        return;
      }
    }
  });
}

/** รับของเข้า = สร้างล็อตใหม่ 1 ล็อต */
function receiveStock(d, ingId, qty, unitCost, note, ts, expiry) {
  ts = ts || Date.now();
  d.lotSeq = (d.lotSeq || 0) + 1;
  var move = {
    id: uid('mv'), ts: ts, ingId: ingId, qty: rq(qty),
    type: 'in', note: note || '', unitCost: unitCost, cost: r2(qty * unitCost)
  };
  var lot = {
    id: uid('lot'), ingId: ingId, ts: ts, seq: d.lotSeq,
    qtyIn: rq(qty), qtyLeft: rq(qty), unitCost: unitCost,
    expiry: expiry || null, note: note || '', moveId: move.id
  };
  move.lotId = lot.id;
  d.lots.push(lot);
  d.moves.push(move);
  return lot;
}

/* ================================================================
   ค่าใช้จ่ายคงที่ & จุดคุ้มทุน
   ================================================================ */

function openDays(d) { return Math.max(1, num(d.shop.openDaysPerMonth, 26)); }

/** ค่าใช้จ่ายคงที่เฉลี่ยต่อวันเปิดร้าน */
function dailyFixedCost(d) {
  return r2((d.fixedCosts || []).reduce(function (s, c) {
    if (c.active === false) return s;
    return s + (c.period === 'day' ? num(c.amount) : num(c.amount) / openDays(d));
  }, 0));
}
function hasFixedCost(d) { return dailyFixedCost(d) > 0; }

/**
 * จุดคุ้มทุน — ใช้กำไรเฉลี่ยต่อแก้วจากยอดขายจริงที่ผ่านมา
 * ถ้ายังไม่มียอดขาย ประเมินจากเมนูที่เปิดขายอยู่ (ขายหน้าร้าน ตัวเลือกค่าเริ่มต้น)
 */
function breakEven(d, rows) {
  var fixed = dailyFixedCost(d);
  var cups = 0, profit = 0, payable = 0;
  (rows || []).forEach(function (s) { cups += s.cups; profit += s.profit; payable += num(s.payable, s.afterDiscount); });

  var perCup, pricePerCup, basis;
  if (cups > 0 && profit > 0) {
    perCup = profit / cups;
    pricePerCup = payable / cups;
    basis = 'จากยอดขายจริง ' + cups + ' แก้ว';
  } else {
    var active = d.menus.filter(function (m) { return m.active !== false; });
    if (!active.length || fixed <= 0) return { fixed: fixed, perCup: 0, cups: 0, sales: 0, basis: '', ready: false };
    var tp = 0, tc = 0;
    active.forEach(function (m) {
      var price = variantPrice(m, defaultSelections(m));
      var vatCfg = d.shop.vat || {};
      var rate = vatCfg.enabled ? num(vatCfg.rate) / 100 : 0;
      var ex = vatCfg.mode === 'exclusive' ? price : r2(price - price * rate / (1 + rate));
      tp += ex - menuCost(d, m);
      tc += price;
    });
    perCup = tp / active.length;
    pricePerCup = tc / active.length;
    basis = 'ประเมินจากเมนูที่เปิดขาย (ยังไม่มียอดขายในช่วงนี้)';
  }

  if (perCup <= 0) return { fixed: fixed, perCup: r2(perCup), cups: 0, sales: 0, basis: basis, ready: false };
  var needCups = Math.ceil(fixed / perCup);
  return {
    fixed: fixed,
    perCup: r2(perCup),
    pricePerCup: r2(pricePerCup),
    cups: needCups,
    sales: r2(needCups * pricePerCup),
    basis: basis,
    ready: fixed > 0
  };
}

/** จำนวนวันที่ใช้เฉลี่ยค่าใช้จ่ายคงที่ในช่วงที่เลือก */
function daysInRange(rows, period) {
  if (period === 'today') return 1;
  if (period === '7d') return 7;
  if (period === 'month') return new Date().getDate();
  if (!rows.length) return 1;
  var min = Infinity, max = -Infinity;
  rows.forEach(function (s) { min = Math.min(min, s.ts); max = Math.max(max, s.ts); });
  return Math.max(1, Math.round((new Date(dayKey(max)).getTime() - new Date(dayKey(min)).getTime()) / 86400000) + 1);
}

/** ตัดสต็อกออกนอกการขาย (ของเสีย / ปรับยอด) */
function issueStock(d, ingId, qty, type, note, ts) {
  var take = makeLedger(d);
  var res = take(ingId, qty, true);
  d.moves.push({
    id: uid('mv'), ts: ts || Date.now(), ingId: ingId, qty: -rq(qty),
    type: type || 'adjust', note: note || '', allocs: res.allocs, cost: res.cost
  });
  return res;
}

/* ================================================================
   ตัวเลือกเมนู (ความหวาน / ผงมัทฉะ / ความเข้ม ฯลฯ)
   group = { id, name, type:'scale'|'swap', target:<ingId เดิมในสูตร>, defaultId, alwaysShow, choices }
   choice = { id, label, factor, ingId (เฉพาะ swap), priceDelta }
   ================================================================ */

function optionGroups(menu) { return (menu && menu.options) || []; }

function choiceOf(group, choiceId) {
  var list = group.choices || [];
  for (var i = 0; i < list.length; i++) if (list[i].id === choiceId) return list[i];
  for (var j = 0; j < list.length; j++) if (list[j].id === group.defaultId) return list[j];
  return list[0] || null;
}

function defaultSelections(menu) {
  var sel = {};
  optionGroups(menu).forEach(function (g) {
    var c = choiceOf(g, g.defaultId);
    if (c) sel[g.id] = c.id;
  });
  return sel;
}

/** สูตรจริงหลังใส่ตัวเลือกแล้ว — จับคู่กับวัตถุดิบ "เดิม" ในสูตร ลำดับตัวเลือกจึงไม่มีผล */
function resolveRecipe(menu, sel) {
  var lines = ((menu && menu.recipe) || []).map(function (l) {
    return { src: l.ingId, ingId: l.ingId, qty: l.qty };
  });
  optionGroups(menu).forEach(function (g) {
    var c = choiceOf(g, sel && sel[g.id]);
    if (!c) return;
    lines.forEach(function (l) {
      if (l.src !== g.target) return;
      if (g.type === 'swap' && c.ingId) l.ingId = c.ingId;
      if (c.factor != null) l.qty = rq(l.qty * c.factor);
    });
  });
  return lines.filter(function (l) { return l.qty > 1e-9; });
}

/** ราคาขายหลังบวกส่วนเพิ่มของตัวเลือก */
function variantPrice(menu, sel) {
  var p = num(menu.price);
  optionGroups(menu).forEach(function (g) {
    var c = choiceOf(g, sel && sel[g.id]);
    if (c) p += num(c.priceDelta);
  });
  return r2(p);
}

/** ป้ายกำกับตัวเลือก — ซ่อนตัวที่เป็นค่าเริ่มต้น ยกเว้นกลุ่มที่ตั้ง alwaysShow */
function variantLabel(menu, sel) {
  return optionGroups(menu).map(function (g) {
    var c = choiceOf(g, sel && sel[g.id]);
    if (!c) return null;
    return (g.alwaysShow || c.id !== g.defaultId) ? c.label : null;
  }).filter(Boolean).join(' · ');
}

/** คีย์รวมเมนู+ตัวเลือก ใช้แยกบรรทัดในตะกร้า */
function variantKey(menuId, sel) {
  var keys = Object.keys(sel || {}).sort();
  return menuId + '|' + keys.map(function (k) { return k + '=' + sel[k]; }).join(',');
}

/**
 * สร้างกลุ่มตัวเลือกสำเร็จรูป แล้วเดาว่าควรผูกกับวัตถุดิบตัวไหนในสูตรของเมนูนี้
 * key: 'sweet' = น้ำเชื่อม, 'powder' = สลับผงมัทฉะ, 'strength' = ปริมาณผงมัทฉะ
 */
function buildPresetGroup(d, menu, key, fallbackTarget) {
  var p = window.SEED.optionPresets && window.SEED.optionPresets[key];
  if (!p) return null;
  var g = JSON.parse(JSON.stringify(p));
  g.id = uid('g');
  var defIdx = 0;
  p.choices.forEach(function (c, i) { if (c.id === p.defaultId) defIdx = i; });
  g.choices.forEach(function (c) { c.id = uid('c'); });
  g.defaultId = g.choices[defIdx].id;

  var target = null;
  (menu.recipe || []).forEach(function (l) {
    var ing = findIng(d, l.ingId);
    if (!ing || target) return;
    if (key === 'sweet' && (l.ingId === p.target || ing.name.indexOf('เชื่อม') >= 0)) target = l.ingId;
    if ((key === 'powder' || key === 'strength') && ing.group === 'matcha') target = l.ingId;
  });
  target = target || fallbackTarget;
  if (!target) return null;
  g.target = target;
  if (key === 'powder') {
    g.choices.forEach(function (c) { if (!findIng(d, c.ingId)) c.ingId = target; });
  }
  return g;
}

/** ต้นทุนจริงของเมนูถ้าทำเพิ่มอีก 1 แก้วตอนนี้ (ไล่ตามล็อต) */
function menuCost(d, menu, ledger, sel) {
  if (!menu || !menu.recipe) return 0;
  var take = ledger || makeLedger(d);
  return r2(resolveRecipe(menu, sel || defaultSelections(menu)).reduce(function (sum, line) {
    return sum + take(line.ingId, line.qty, false).cost;
  }, 0));
}

function balance(d, ingId) {
  var b = 0;
  for (var i = 0; i < d.moves.length; i++) if (d.moves[i].ingId === ingId) b += d.moves[i].qty;
  return r2(b);
}
function movedIn(d, ingId) {
  var b = 0;
  for (var i = 0; i < d.moves.length; i++) if (d.moves[i].ingId === ingId && d.moves[i].qty > 0) b += d.moves[i].qty;
  return r2(b);
}
function movedOut(d, ingId) {
  var b = 0;
  for (var i = 0; i < d.moves.length; i++) if (d.moves[i].ingId === ingId && d.moves[i].qty < 0) b -= d.moves[i].qty;
  return r2(b);
}

function maxCups(d, menu, sel) {
  var lines = resolveRecipe(menu, sel || defaultSelections(menu));
  if (!lines.length) return 0;
  var min = Infinity;
  lines.forEach(function (line) {
    if (line.qty > 0) min = Math.min(min, Math.floor(balance(d, line.ingId) / line.qty));
  });
  return min === Infinity ? 0 : Math.max(0, min);
}

/**
 * คำนวณยอดของบิล ใช้ทั้งตอนพรีวิวในตะกร้าและตอนบันทึกจริง
 * ต้นทุนตัดจากล็อตแบบ FIFO ผ่านกระดานเดียวกันทั้งบิล
 * opt.commit = true → ตัดยอดล็อตจริง, opt.draws = อาร์เรย์รับรายละเอียดการตัดแต่ละวัตถุดิบ
 */
function calcSale(d, itemsIn, platformId, discount, opt) {
  opt = opt || {};
  var commit = !!opt.commit;
  var take = opt.ledger || makeLedger(d);
  var draws = opt.draws || null;
  var plat = findPlatform(d, platformId);
  var items = [];
  var shortage = 0;
  itemsIn.forEach(function (it) {
    var m = findMenu(d, it.menuId);
    if (!m || !(it.qty > 0)) return;
    var sel = it.selections || defaultSelections(m);
    var price = variantPrice(m, sel);
    var lineCost = 0;
    resolveRecipe(m, sel).forEach(function (line) {
      var need = rq(line.qty * it.qty);
      if (!(need > 0)) return;
      var res = take(line.ingId, need, commit);
      lineCost += res.cost;
      if (res.shortage > 0) shortage += 1;
      if (draws) draws.push({
        menuId: m.id, menuName: m.name, menuQty: it.qty,
        ingId: line.ingId, qty: need, allocs: res.allocs, cost: res.cost, shortage: res.shortage
      });
    });
    items.push({
      key: it.key || variantKey(m.id, sel),
      menuId: m.id, name: m.name, variant: variantLabel(m, sel), selections: sel,
      qty: it.qty, price: price,
      unitCost: r2(lineCost / it.qty),
      lineTotal: r2(price * it.qty),
      lineCost: r2(lineCost)
    });
  });

  var gross = r2(items.reduce(function (s, i) { return s + i.lineTotal; }, 0));
  var disc = Math.min(r2(discount || 0), gross);
  var afterDiscount = r2(gross - disc);

  /* ---- VAT ---- */
  var vatCfg = (d.shop && d.shop.vat) || { enabled: false, rate: 7, mode: 'inclusive' };
  var rate = vatCfg.enabled ? num(vatCfg.rate) / 100 : 0;
  var payable, vat, exVat;
  if (rate <= 0) {
    payable = afterDiscount; vat = 0; exVat = afterDiscount;
  } else if (vatCfg.mode === 'exclusive') {          /* ราคายังไม่รวม VAT → บวกเพิ่ม */
    exVat = afterDiscount;
    vat = r2(afterDiscount * rate);
    payable = r2(afterDiscount + vat);
  } else {                                           /* ราคารวม VAT แล้ว → ถอดออกมา */
    payable = afterDiscount;
    vat = r2(afterDiscount * rate / (1 + rate));
    exVat = r2(afterDiscount - vat);
  }

  var gpRate = num(plat ? plat.gp : 0);
  var gpAmount = r2(payable * gpRate / 100);         /* แพลตฟอร์มหักจากยอดที่ลูกค้าจ่าย */
  var net = r2(payable - gpAmount);                  /* เงินที่ร้านได้รับจริง (ยังมี VAT ที่ต้องนำส่ง) */
  var cost = r2(items.reduce(function (s, i) { return s + i.lineCost; }, 0));
  var profit = r2(exVat - gpAmount - cost);          /* กำไรคิดจากรายได้ที่ไม่รวม VAT */

  return {
    platformId: plat ? plat.id : '', platformName: plat ? plat.name : '-', gpRate: gpRate,
    items: items, cups: items.reduce(function (s, i) { return s + i.qty; }, 0),
    gross: gross, discount: disc, afterDiscount: afterDiscount,
    vatRate: rate * 100, vatMode: vatCfg.mode, vat: vat, exVat: exVat, payable: payable,
    gpAmount: gpAmount, net: net, cost: cost, profit: profit,
    shortage: shortage
  };
}

function commitSale(d, opt) {
  var ts = opt.ts || Date.now();
  var draws = [];
  var t = calcSale(d, opt.items, opt.platformId, opt.discount, { commit: true, draws: draws });
  d.billSeq = (d.billSeq || 0) + 1;
  var plat = findPlatform(d, t.platformId);
  var payId = opt.payMethodId || (plat && plat.settlement === 'direct' ? 'cash' : 'platform');
  var pm = payMethod(d, payId);
  var sale = {
    id: uid('s'),
    no: 'B' + String(d.billSeq).padStart(4, '0'),
    ts: ts,
    platformId: t.platformId, platformName: t.platformName, gpRate: t.gpRate,
    settlement: plat ? plat.settlement : 'direct',
    payMethodId: payId, payMethodName: pm ? pm.name : 'แพลตฟอร์ม',
    items: t.items, cups: t.cups,
    gross: t.gross, discount: t.discount, afterDiscount: t.afterDiscount,
    vatRate: t.vatRate, vatMode: t.vatMode, vat: t.vat, exVat: t.exVat, payable: t.payable,
    gpAmount: t.gpAmount, net: t.net, cost: t.cost, profit: t.profit,
    received: opt.received != null ? r2(opt.received) : null,
    change: opt.received != null ? r2(opt.received - t.payable) : null,
    note: opt.note || ''
  };
  d.sales.push(sale);

  draws.forEach(function (dr) {
    d.moves.push({
      id: uid('mv'), ts: ts, ingId: dr.ingId, qty: -dr.qty,
      type: 'sale', ref: sale.id,
      note: dr.menuName + ' ×' + dr.menuQty + ' (' + sale.no + ')',
      allocs: dr.allocs, cost: dr.cost
    });
  });
  return sale;
}

/** ลบบิล + คืนของเข้าล็อตเดิมทุกล็อตที่เคยถูกตัดไป */
function deleteSale(d, saleId) {
  d.moves.forEach(function (m) {
    if (m.ref === saleId) releaseAllocs(d, m.allocs);
  });
  d.sales = d.sales.filter(function (s) { return s.id !== saleId; });
  d.moves = d.moves.filter(function (m) { return m.ref !== saleId; });
}

/* ------------------------------------------------ modal */
var modalFooterHandlers = [];
function openModal(title, bodyHtml, buttons) {
  $('#modalTitle').textContent = title;
  $('#modalBody').innerHTML = bodyHtml;
  var foot = $('#modalFoot');
  foot.innerHTML = '';
  modalFooterHandlers = [];
  (buttons || []).forEach(function (b, i) {
    var btn = document.createElement('button');
    btn.className = 'btn ' + (b.cls || '');
    btn.textContent = b.label;
    btn.addEventListener('click', function () { b.onClick && b.onClick(); });
    modalFooterHandlers.push(btn);
    foot.appendChild(btn);
  });
  $('#modalRoot').classList.remove('is-hidden');
  var first = $('#modalBody input, #modalBody select');
  if (first) setTimeout(function () { first.focus(); }, 30);
}
function closeModal() { $('#modalRoot').classList.add('is-hidden'); }

function confirmDialog(title, message, onYes, yesLabel) {
  openModal(title, '<p>' + message + '</p>', [
    { label: 'ยกเลิก', onClick: closeModal },
    { label: yesLabel || 'ยืนยัน', cls: 'btn-danger-ghost', onClick: function () { closeModal(); onYes(); } }
  ]);
}

/* ------------------------------------------------ app state */
var state = {
  view: 'pos',
  platformId: null,
  payMethodId: 'cash',
  received: null,
  closeDay: '',
  cart: [],
  search: '',
  discount: 0,
  period: '7d',
  histFrom: '',
  histTo: '',
  histPlatform: 'all'
};

/* ============================================================
   VIEW: ขาย
   ============================================================ */
/** แถบความคืบหน้าสู่จุดคุ้มทุนของวันนี้ */
function renderBreakEvenBar() {
  var bar = $('#beBar');
  if (!hasFixedCost(db)) { bar.classList.add('is-hidden'); return; }
  bar.classList.remove('is-hidden');

  var today = salesOfDay(db, todayKey());
  var sum = summarize(today);
  var fixed = dailyFixedCost(db);
  var be = breakEven(db, db.sales.slice(-200));
  var netAfterFixed = r2(sum.profit - fixed);
  var pct = fixed > 0 ? Math.min(100, Math.max(0, sum.profit / fixed * 100)) : 100;
  var done = netAfterFixed >= 0;

  bar.innerHTML =
    '<div class="be-head">' +
      '<span>' + (done
        ? '🎉 <strong>วันนี้คุ้มทุนแล้ว</strong> — กำไรจริงหลังหักค่าใช้จ่ายคงที่ ' + signedBaht(netAfterFixed)
        : 'วันนี้ทำได้ <strong>฿' + baht(sum.profit) + '</strong> จาก ฿' + baht(fixed) +
          ' · ขาดอีก <strong>฿' + baht(fixed - sum.profit) + '</strong>' +
          (be.perCup > 0 ? ' (≈ ' + Math.ceil((fixed - sum.profit) / be.perCup) + ' แก้ว)' : '')) + '</span>' +
      '<span class="be-sub">ขายแล้ว ' + sum.cups + ' แก้ว · เป้า ' + (be.cups || '–') + ' แก้ว/วัน</span>' +
    '</div>' +
    '<div class="be-track"><div class="be-fill' + (done ? ' is-done' : '') + '" style="width:' + pct + '%"></div></div>';
}

function renderPOS() {
  renderBreakEvenBar();

  /* ช่องทางขาย */
  $('#platformChips').innerHTML = db.platforms.map(function (p) {
    return '<button class="chip' + (p.id === state.platformId ? ' is-active' : '') + '" data-plat="' + esc(p.id) + '">' +
      esc(p.name) + (p.gp ? ' <small>' + qtyStr(p.gp) + '%</small>' : '') + '</button>';
  }).join('');
  $$('#platformChips .chip').forEach(function (el) {
    el.addEventListener('click', function () { state.platformId = el.getAttribute('data-plat'); renderPOS(); });
  });

  /* วิธีชำระเงิน — เฉพาะช่องทางที่ร้านรับเงินเอง */
  var plat0 = findPlatform(db, state.platformId);
  if (isDirect(plat0)) {
    $('#payRow').classList.remove('is-hidden');
    $('#payChips').innerHTML = (db.payMethods || []).map(function (p) {
      return '<button class="chip' + (p.id === state.payMethodId ? ' is-active' : '') + '" data-pay="' + esc(p.id) + '">' +
        esc(p.name) + '</button>';
    }).join('');
    $$('#payChips .chip').forEach(function (el) {
      el.addEventListener('click', function () { state.payMethodId = el.getAttribute('data-pay'); renderPOS(); });
    });
  } else {
    $('#payRow').classList.add('is-hidden');
  }

  /* เมนู */
  var term = state.search.trim().toLowerCase();
  var list = db.menus.filter(function (m) {
    return m.active !== false && (!term || m.name.toLowerCase().indexOf(term) >= 0);
  });
  var plat = findPlatform(db, state.platformId);

  if (!list.length) {
    $('#menuGrid').innerHTML = '<div class="empty">ยังไม่มีเมนู — เพิ่มได้ที่แท็บ “เมนู &amp; สูตร”</div>';
  } else {
    $('#menuGrid').innerHTML = list.map(function (m) {
      var cost = menuCost(db, m);
      var profit = m.price * (1 - num(plat && plat.gp) / 100) - cost;
      var can = maxCups(db, m);
      var opts = optionGroups(m).length;
      var badge = can <= 0
        ? '<span class="m-badge is-out">วัตถุดิบหมด</span>'
        : (can <= 5 ? '<span class="m-badge">เหลือ ' + can + ' แก้ว</span>' : '');
      return '<button class="menu-card' + (can <= 0 ? ' is-out' : '') + '" data-menu="' + esc(m.id) + '">' +
        badge +
        '<div class="m-swatch" style="background:' + esc(m.color || '#4f7a34') + '"></div>' +
        '<div class="m-name">' + esc(m.name) + '</div>' +
        '<div class="m-price">฿' + baht(m.price) + '</div>' +
        '<div class="m-meta">ต้นทุน ' + methodLabel(db) + ' ฿' + baht(cost) + ' · กำไร ฿' + baht(profit) +
          (opts ? '<br><span class="m-opt">เลือกตัวเลือกได้ ' + opts + ' อย่าง</span>' : '') + '</div>' +
        '</button>';
    }).join('');
    $$('#menuGrid .menu-card').forEach(function (el) {
      el.addEventListener('click', function () {
        var m = findMenu(db, el.getAttribute('data-menu'));
        if (!m) return;
        if (optionGroups(m).length) optionDialog(m);
        else addToCart(m.id, {}, 1);
      });
    });
  }

  renderCart();
}

/** หน้าต่างเลือกความหวาน/ผงมัทฉะ/ความเข้ม พร้อมพรีวิวราคา ต้นทุน และปริมาณที่จะตัดจริง */
function optionDialog(menu) {
  var sel = defaultSelections(menu);
  var qty = 1;

  openModal('เลือกตัวเลือก — ' + menu.name,
    '<div id="optGroups"></div>' +
    '<div class="opt-qty">' +
      '<span>จำนวน</span>' +
      '<button class="qbtn" id="optDec">−</button>' +
      '<span class="n" id="optQty">1</span>' +
      '<button class="qbtn" id="optInc">+</button>' +
    '</div>' +
    '<div class="opt-preview" id="optPreview"></div>',
    [
      { label: 'ยกเลิก', onClick: closeModal },
      { label: 'เพิ่มลงตะกร้า', cls: 'btn-primary', onClick: function () {
        addToCart(menu.id, sel, qty);
        closeModal();
      } }
    ]);

  function draw() {
    $('#optGroups').innerHTML = optionGroups(menu).map(function (g) {
      return '<div class="opt-group"><div class="opt-label">' + esc(g.name) + '</div><div class="chips">' +
        (g.choices || []).map(function (c) {
          return '<button class="chip' + (sel[g.id] === c.id ? ' is-active' : '') + '" ' +
            'data-g="' + esc(g.id) + '" data-c="' + esc(c.id) + '">' + esc(c.label) +
            (num(c.priceDelta) ? ' <small>+' + qtyStr(c.priceDelta) + '฿</small>' : '') + '</button>';
        }).join('') + '</div></div>';
    }).join('');
    $$('#optGroups .chip').forEach(function (el) {
      el.addEventListener('click', function () {
        sel[el.getAttribute('data-g')] = el.getAttribute('data-c');
        draw();
      });
    });
    $('#optQty').textContent = qty;

    var price = variantPrice(menu, sel);
    var cost = menuCost(db, menu, null, sel);
    var lines = resolveRecipe(menu, sel).filter(function (l) {
      return optionGroups(menu).some(function (g) { return g.target === l.src; });
    });
    $('#optPreview').innerHTML =
      '<div class="opt-sum"><span>ราคา ฿' + baht(price) + ' × ' + qty + '</span>' +
      '<strong>฿' + baht(price * qty) + '</strong></div>' +
      '<div class="hint">ต้นทุน ' + methodLabel(db) + ' ฿' + baht(cost) + ' /แก้ว · กำไรขั้นต้น ฿' + baht(price - cost) + '</div>' +
      (lines.length ? '<div class="hint">จะตัดสต็อก: ' + lines.map(function (l) {
        var ing = findIng(db, l.ingId);
        return esc(ing ? ing.name : l.ingId) + ' ' + qtyStr(l.qty * qty) + ' ' + esc(ing ? ing.unit : '');
      }).join(' · ') + '</div>' : '');
  }

  $('#optDec').addEventListener('click', function () { qty = Math.max(1, qty - 1); draw(); });
  $('#optInc').addEventListener('click', function () { qty += 1; draw(); });
  draw();
}

function addToCart(menuId, selections, qty) {
  var m = findMenu(db, menuId);
  if (!m) return;
  var sel = selections && Object.keys(selections).length ? selections : defaultSelections(m);
  var key = variantKey(menuId, sel);
  var line = null;
  state.cart.forEach(function (c) { if (c.key === key) line = c; });
  if (line) line.qty += (qty || 1);
  else state.cart.push({ key: key, menuId: menuId, selections: sel, qty: qty || 1 });
  renderCart();
}

function setQty(key, delta) {
  state.cart.forEach(function (c) { if (c.key === key) c.qty += delta; });
  state.cart = state.cart.filter(function (c) { return c.qty > 0; });
  renderCart();
}

function renderCart() {
  var box = $('#cartLines');
  var t = calcSale(db, state.cart, state.platformId, state.discount);

  if (!t.items.length) {
    box.innerHTML = '<div class="cart-empty">ยังไม่มีรายการ<br>แตะที่เมนูเพื่อเพิ่มลงตะกร้า</div>';
  } else {
    box.innerHTML = t.items.map(function (it) {
      var m = findMenu(db, it.menuId);
      var warn = it.qty > maxCups(db, m, it.selections) ? ' <span class="pill pill-danger">วัตถุดิบไม่พอ</span>' : '';
      return '<div class="cart-line">' +
        '<div class="cl-name">' + esc(it.name) + warn +
          (it.variant ? '<br><span class="cl-variant">' + esc(it.variant) + '</span>' : '') + '</div>' +
        '<div class="cl-total">฿' + baht(it.lineTotal) + '</div>' +
        '<div class="cl-sub">฿' + baht(it.price) + ' /แก้ว · ต้นทุน ' + methodLabel(db) + ' ฿' + baht(it.lineCost) + '</div>' +
        '<div class="cl-qty">' +
          '<button class="qbtn" data-dec="' + esc(it.key) + '" aria-label="ลด">−</button>' +
          '<span class="n">' + it.qty + '</span>' +
          '<button class="qbtn" data-inc="' + esc(it.key) + '" aria-label="เพิ่ม">+</button>' +
        '</div></div>';
    }).join('');
    $$('#cartLines [data-dec]').forEach(function (b) { b.addEventListener('click', function () { setQty(b.getAttribute('data-dec'), -1); }); });
    $$('#cartLines [data-inc]').forEach(function (b) { b.addEventListener('click', function () { setQty(b.getAttribute('data-inc'), 1); }); });
  }

  var vatOn = t.vatRate > 0;
  $('#cartTotals').innerHTML =
    row('ยอดขาย (' + t.cups + ' แก้ว)', '฿' + baht(t.gross)) +
    (t.discount ? row('ส่วนลด', '−฿' + baht(t.discount), 'row-minus') : '') +
    (vatOn ? row('ยอดก่อน VAT', '฿' + baht(t.exVat)) : '') +
    (vatOn ? row('VAT ' + qtyStr(t.vatRate) + '%' + (t.vatMode === 'inclusive' ? ' (รวมในราคา)' : ''), '฿' + baht(t.vat)) : '') +
    row('ยอดที่ลูกค้าจ่าย', '฿' + baht(t.payable), 'row-strong') +
    row('ค่า GP ' + esc(t.platformName) + ' ' + qtyStr(t.gpRate) + '%',
        t.gpAmount ? '−฿' + baht(t.gpAmount) : '฿0.00', t.gpAmount ? 'row-minus' : '') +
    row('ยอดสุทธิที่ได้รับ', '฿' + baht(t.net)) +
    row('ต้นทุนวัตถุดิบ (' + methodLabel(db) + ')', t.cost ? '−฿' + baht(t.cost) : '฿0.00', t.cost ? 'row-minus' : '') +
    row('กำไรสุทธิ', '฿' + baht(t.profit), 'row-strong row-profit');

  /* ช่องรับเงิน/ทอน — เฉพาะเงินสดหน้าร้าน */
  var plat = findPlatform(db, state.platformId);
  var pm = payMethod(db, state.payMethodId);
  var cashMode = isDirect(plat) && pm && pm.inDrawer;
  $('#cashRow').classList.toggle('is-hidden', !cashMode || !t.items.length);
  if (cashMode && t.items.length) {
    var got = state.received;
    $('#changeOut').textContent = (got != null && got >= t.payable)
      ? 'เงินทอน ฿' + baht(got - t.payable)
      : (got != null && got > 0 ? 'ยังขาดอีก ฿' + baht(t.payable - got) : '');
    $('#quickCash').innerHTML = quickCashAmounts(t.payable).map(function (v) {
      return '<button class="chip" data-cash="' + v + '">' + (v === t.payable ? 'พอดี' : qtyStr(v)) + '</button>';
    }).join('');
    $$('#quickCash .chip').forEach(function (b) {
      b.addEventListener('click', function () {
        state.received = num(b.getAttribute('data-cash'));
        $('#received').value = state.received;
        renderCart();
      });
    });
  }

  $('#checkoutBtn').disabled = !state.cart.length;
  $('#checkoutBtn').textContent = state.cart.length ? 'ชำระเงิน · ฿' + baht(t.payable) : 'ชำระเงิน';

  function row(label, value, cls) {
    var c = cls || '';
    return '<dt class="' + c + '">' + label + '</dt><dd class="' + c + '">' + value + '</dd>';
  }
}

/** ปุ่มลัดรับเงิน — ยอดพอดี + แบงก์ที่ใช้จ่ายจริงบ่อย ๆ */
function quickCashAmounts(payable) {
  var out = [r2(payable)];
  [20, 50, 100, 500, 1000].forEach(function (note) {
    var v = Math.ceil(payable / note) * note;
    if (v > payable && out.indexOf(v) < 0) out.push(v);
  });
  return out.sort(function (a, b) { return a - b; }).slice(0, 5);
}

function checkout() {
  if (!state.cart.length) return;
  var shortages = [];
  state.cart.forEach(function (c) {
    var m = findMenu(db, c.menuId);
    if (m) resolveRecipe(m, c.selections).forEach(function (l) {
      var need = l.qty * c.qty;
      var have = balance(db, l.ingId);
      if (need > have) {
        var ing = findIng(db, l.ingId);
        shortages.push(esc(ing ? ing.name : l.ingId) + ' (ต้องใช้ ' + qtyStr(need) + ' เหลือ ' + qtyStr(have) + ' ' + esc(ing ? ing.unit : '') + ')');
      }
    });
  });

  if (shortages.length) {
    confirmDialog('วัตถุดิบไม่พอ',
      'รายการต่อไปนี้จะทำให้สต็อกติดลบ:<br><br>• ' + shortages.join('<br>• ') + '<br><br>ต้องการบันทึกการขายต่อไปหรือไม่?',
      doCheckout, 'บันทึกต่อไป');
    return;
  }
  doCheckout();
}

function doCheckout() {
  var plat = findPlatform(db, state.platformId);
  var pm = payMethod(db, state.payMethodId);
  var sale = commitSale(db, {
    platformId: state.platformId,
    payMethodId: isDirect(plat) ? state.payMethodId : 'platform',
    received: (isDirect(plat) && pm && pm.inDrawer) ? state.received : null,
    items: state.cart.map(function (c) {
      return { menuId: c.menuId, qty: c.qty, selections: c.selections, key: c.key };
    }),
    discount: state.discount
  });
  state.cart = [];
  state.discount = 0;
  state.received = null;
  $('#discount').value = 0;
  $('#received').value = '';
  save();
  renderPOS();
  showReceipt(sale);
  toast('บันทึกบิล ' + sale.no + ' แล้ว · กำไร ฿' + baht(sale.profit));
}

function receiptHtml(sale) {
  var vatOn = num(sale.vatRate) > 0;
  return '<div class="receipt">' +
    '<div class="r-shop">' + esc(db.shop.name) + '</div>' +
    '<div class="r-meta">' + esc(db.shop.tagline || '') +
      (db.shop.taxId ? '<br>เลขประจำตัวผู้เสียภาษี ' + esc(db.shop.taxId) : '') +
      (vatOn ? '<br>ใบเสร็จรับเงิน / ใบกำกับภาษีอย่างย่อ' : '') + '<br>' +
      'บิล ' + esc(sale.no) + ' · ' + fmtDateTime(sale.ts) + '<br>' +
      'ช่องทาง: ' + esc(sale.platformName) +
      (sale.payMethodName ? ' · ' + esc(sale.payMethodName) : '') + '</div>' +
    '<table><tbody>' +
      sale.items.map(function (i) {
        return '<tr><td>' + esc(i.name) +
          (i.variant ? '<br><small>' + esc(i.variant) + '</small>' : '') +
          '<br><small>' + i.qty + ' × ฿' + baht(i.price) + '</small></td>' +
          '<td class="num">฿' + baht(i.lineTotal) + '</td></tr>';
      }).join('') +
    '</tbody></table>' +
    '<div class="r-total">' +
      '<span>ยอดรวม</span><span class="num">฿' + baht(sale.gross) + '</span>' +
      (sale.discount ? '<span>ส่วนลด</span><span class="num">−฿' + baht(sale.discount) + '</span>' : '') +
      (vatOn ? '<span>มูลค่าก่อน VAT</span><span class="num">฿' + baht(sale.exVat) + '</span>' +
               '<span>VAT ' + qtyStr(sale.vatRate) + '%</span><span class="num">฿' + baht(sale.vat) + '</span>' : '') +
      '<span class="strong">ยอดชำระ</span><span class="num strong">฿' + baht(sale.payable) + '</span>' +
      (sale.received != null ? '<span>รับเงิน</span><span class="num">฿' + baht(sale.received) + '</span>' +
                               '<span>เงินทอน</span><span class="num">฿' + baht(sale.change) + '</span>' : '') +
      (sale.gpRate ? '<span>ค่า GP ' + qtyStr(sale.gpRate) + '%</span><span class="num">−฿' + baht(sale.gpAmount) + '</span>' +
                     '<span>ร้านได้รับสุทธิ</span><span class="num">฿' + baht(sale.net) + '</span>' : '') +
      '<span>ต้นทุน (' + methodLabel(db) + ')</span>' + '<span class="num">฿' + baht(sale.cost) + '</span>' +
      '<span>กำไร</span><span class="num">฿' + baht(sale.profit) + '</span>' +
    '</div>' +
    '<div class="r-thanks">ขอบคุณที่อุดหนุน 🍵</div></div>';
}

function showReceipt(sale) {
  openModal('ใบเสร็จ ' + sale.no, receiptHtml(sale), [
    { label: 'ปิด', onClick: closeModal },
    { label: 'พิมพ์ใบเสร็จ', cls: 'btn-primary', onClick: function () { printReceipt(sale); } }
  ]);
}
function printReceipt(sale) {
  $('#printArea').innerHTML = receiptHtml(sale);
  window.print();
}

/* ============================================================
   VIEW: สต็อก
   ============================================================ */
function renderStock() {
  var low = db.ingredients.filter(function (i) { return balance(db, i.id) <= num(i.minQty); });
  var alerts = low.length
    ? '<div class="alert ' + (low.some(function (i) { return balance(db, i.id) <= 0; }) ? 'alert-danger' : 'alert-warn') + '">' +
      '<span>⚠️</span><div><strong>ต้องสั่งของเพิ่ม:</strong> ' +
      low.map(function (i) { return esc(i.name) + ' (เหลือ ' + qtyStr(balance(db, i.id)) + ' ' + esc(i.unit) + ')'; }).join(', ') +
      '</div></div>'
    : '';

  /* แจ้งเตือนวันหมดอายุ */
  var exp = expiringLots(db);
  var expired = exp.filter(function (l) { return lotExpiryState(db, l) === 'expired'; });
  var soon = exp.filter(function (l) { return lotExpiryState(db, l) === 'soon'; });

  if (expired.length) {
    alerts += '<div class="alert alert-danger"><span>🚫</span><div>' +
      '<strong>หมดอายุแล้ว ' + expired.length + ' ล็อต — ห้ามนำไปใช้:</strong> ' +
      expired.map(function (l) {
        var ing = findIng(db, l.ingId);
        return esc(ing ? ing.name : '') + ' L' + l.seq + ' (' + qtyStr(l.qtyLeft) + ' ' + esc(ing ? ing.unit : '') +
          ' · หมดอายุ ' + fmtDayShort(l.expiry) + ')';
      }).join(', ') +
      '<div class="btn-row"><button class="btn btn-sm btn-danger-ghost" id="dumpExpiredBtn">ตัดของหมดอายุทิ้งทั้งหมด (฿' +
      baht(expired.reduce(function (s, l) { return s + l.qtyLeft * l.unitCost; }, 0)) + ')</button></div>' +
      '</div></div>';
  }
  if (soon.length) {
    alerts += '<div class="alert alert-warn"><span>⏳</span><div><strong>ใกล้หมดอายุใน ' +
      qtyStr(num(db.shop.expiryWarnDays, 14)) + ' วัน — เร่งใช้ก่อน:</strong> ' +
      soon.map(function (l) {
        var ing = findIng(db, l.ingId);
        return esc(ing ? ing.name : '') + ' L' + l.seq + ' (เหลือ ' + qtyStr(l.qtyLeft) + ' ' + esc(ing ? ing.unit : '') +
          ' · อีก ' + daysUntil(l.expiry) + ' วัน)';
      }).join(', ') + '</div></div>';
  }

  $('#stockAlerts').innerHTML = alerts;

  if (expired.length) {
    $('#dumpExpiredBtn').addEventListener('click', function () {
      confirmDialog('ตัดของหมดอายุทิ้ง',
        'ตัดของที่หมดอายุแล้วออกจากสต็อกทั้งหมด ' + expired.length + ' ล็อต? จะบันทึกเป็น “ของเสีย” ให้',
        function () {
          var total = 0;
          expired.forEach(function (l) {
            var ing = findIng(db, l.ingId);
            var q = l.qtyLeft;
            total += q * l.unitCost;
            l.qtyLeft = 0;
            db.moves.push({
              id: uid('mv'), ts: Date.now(), ingId: l.ingId, qty: -rq(q), type: 'waste',
              note: 'หมดอายุ ' + l.expiry + ' (L' + l.seq + ')',
              allocs: [{ lotId: l.id, qty: rq(q), unitCost: l.unitCost }],
              cost: r2(q * l.unitCost)
            });
            if (!ing) return;
          });
          save(); render(); toast('ตัดของหมดอายุทิ้งแล้ว มูลค่า ฿' + baht(total));
        }, 'ตัดทิ้ง');
    });
  }

  var fefo = stockMethod(db) === 'fefo';
  $('#lotTitle').textContent = 'คิวล็อต — ' + (fefo ? 'FEFO: ของที่จะหมดอายุก่อน ถูกใช้ก่อน' : 'FIFO: ของที่รับเข้าก่อน ถูกใช้ก่อน');
  $('#lotNote').innerHTML = fefo
    ? 'ล็อตที่ยังไม่ระบุวันหมดอายุจะไปต่อท้ายคิว — คลิกช่องวันหมดอายุในตารางเพื่อเติมได้เลย'
    : 'เรียงตามวันที่รับเข้าอย่างเดียว ไม่สนวันหมดอายุ (เปลี่ยนเป็น FEFO ได้ที่แท็บตั้งค่า)';

  $('#matchaTable').innerHTML = stockTable('matcha');
  $('#otherTable').innerHTML = stockTable('other');
  $('#lotTable').innerHTML = lotTable();
  bindStockButtons();

  var moves = db.moves.slice().sort(function (a, b) { return b.ts - a.ts; }).slice(0, 60);
  $('#movesTable').innerHTML = moves.length ? '<table><thead><tr>' +
    '<th>เวลา</th><th>วัตถุดิบ</th><th>ประเภท</th><th class="num">จำนวน</th><th class="num">มูลค่า</th><th>ล็อตที่ใช้</th><th>หมายเหตุ</th>' +
    '</tr></thead><tbody>' + moves.map(function (m) {
      var ing = findIng(db, m.ingId);
      var label = { in: 'รับเข้า', sale: 'ขาย', adjust: 'ปรับยอด', waste: 'ของเสีย' }[m.type] || m.type;
      return '<tr><td>' + fmtDateTime(m.ts) + '</td>' +
        '<td>' + esc(ing ? ing.name : m.ingId) + '</td>' +
        '<td><span class="pill ' + (m.qty >= 0 ? 'pill-ok' : 'pill-plain') + '">' + label + '</span></td>' +
        '<td class="num ' + (m.qty >= 0 ? 'pos' : 'neg') + '">' + (m.qty >= 0 ? '+' : '') + qtyStr(m.qty) + ' ' + esc(ing ? ing.unit : '') + '</td>' +
        '<td class="num">' + (m.cost != null ? '฿' + baht(m.cost) : '–') + '</td>' +
        '<td>' + allocLabel(m) + '</td>' +
        '<td>' + esc(m.note || '') + '</td></tr>';
    }).join('') + '</tbody></table>'
    : '<div class="empty">ยังไม่มีความเคลื่อนไหว</div>';
}

/** ชื่อย่อของล็อต เช่น L3 · 20 ก.ย. */
function lotLabel(d, lotId) {
  for (var i = 0; i < d.lots.length; i++) {
    if (d.lots[i].id === lotId) {
      return 'L' + d.lots[i].seq + ' (฿' + baht(d.lots[i].unitCost) + ')';
    }
  }
  return 'ล็อตที่ถูกลบ';
}

/** สรุปว่าการเคลื่อนไหวนี้หยิบจากล็อตไหนบ้าง */
function allocLabel(m) {
  if (m.qty > 0) return m.lotId ? lotLabel(db, m.lotId) : '–';
  if (!m.allocs || !m.allocs.length) return '–';
  return m.allocs.map(function (a) {
    var ing = findIng(db, m.ingId);
    var q = qtyStr(a.qty) + (ing ? ' ' + ing.unit : '');
    return a.lotId
      ? '<span class="pill pill-plain">' + esc(lotLabel(db, a.lotId)) + ' · ' + q + '</span>'
      : '<span class="pill pill-danger">เกินสต็อก · ' + q + '</span>';
  }).join(' ');
}

/** ตารางคิวล็อต FIFO — ล็อตบนสุดของแต่ละวัตถุดิบคือล็อตที่กำลังถูกใช้ */
function lotTable() {
  var groups = db.ingredients.map(function (ing) {
    return { ing: ing, lots: openLots(db, ing.id) };
  }).filter(function (g) { return g.lots.length; });

  if (!groups.length) return '<div class="empty">ยังไม่มีล็อตคงเหลือ</div>';

  var body = '';
  groups.forEach(function (g) {
    g.lots.forEach(function (l, idx) {
      var st = lotExpiryState(db, l);
      var left = l.expiry ? daysUntil(l.expiry) : null;
      var expCell = '<input type="date" class="input input-sm lot-exp' +
        (st === 'expired' ? ' is-expired' : (st === 'soon' ? ' is-soon' : '')) +
        '" data-lotexp="' + esc(l.id) + '" value="' + esc(l.expiry || '') + '">' +
        (st === 'expired' ? '<span class="pill pill-danger">หมดอายุแล้ว</span>'
          : st === 'soon' ? '<span class="pill pill-warn">อีก ' + left + ' วัน</span>'
          : st === 'ok' ? '<span class="hint">อีก ' + left + ' วัน</span>'
          : '<span class="hint">ไม่ระบุ</span>');
      body += '<tr>' +
        '<td>' + (idx === 0 ? '<strong>' + esc(g.ing.name) + '</strong>' : '<span class="text-faint">↳</span>') + '</td>' +
        '<td><span class="pill ' + (idx === 0 ? 'pill-ok' : 'pill-plain') + '">' +
          (idx === 0 ? 'ใช้ก่อน' : 'คิวที่ ' + (idx + 1)) + '</span></td>' +
        '<td>L' + l.seq + '</td>' +
        '<td>' + fmtDateTime(l.ts) + '</td>' +
        '<td class="lot-exp-cell">' + expCell + '</td>' +
        '<td class="num">฿' + baht(l.unitCost) + '</td>' +
        '<td class="num">' + qtyStr(l.qtyIn) + '</td>' +
        '<td class="num">' + qtyStr(rq(l.qtyIn - l.qtyLeft)) + '</td>' +
        '<td class="num"><strong>' + qtyStr(l.qtyLeft) + '</strong> ' + esc(g.ing.unit) + '</td>' +
        '<td class="num">฿' + baht(l.qtyLeft * l.unitCost) + '</td>' +
        '<td>' + esc(l.note || '') + '</td></tr>';
    });
  });

  var total = db.ingredients.reduce(function (s, i) { return s + stockValue(db, i.id); }, 0);
  return '<table><thead><tr>' +
    '<th>วัตถุดิบ</th><th>ลำดับใช้</th><th>ล็อต</th><th>วันที่รับเข้า</th><th>วันหมดอายุ</th><th class="num">ต้นทุน/หน่วย</th>' +
    '<th class="num">รับเข้า</th><th class="num">ใช้ไปแล้ว</th><th class="num">คงเหลือ</th><th class="num">มูลค่า</th><th>หมายเหตุ</th>' +
    '</tr></thead><tbody>' + body + '</tbody>' +
    '<tfoot><tr><td colspan="9">มูลค่าสต็อกคงเหลือรวม</td><td class="num">฿' + baht(total) + '</td><td></td></tr></tfoot></table>';
}

function stockTable(group) {
  var rows = db.ingredients.filter(function (i) { return i.group === group; });
  if (!rows.length) return '<div class="empty">ยังไม่มีรายการ</div>';
  var totalValue = 0;
  var body = rows.map(function (i) {
    var bal = balance(db, i.id);
    var open = openLots(db, i.id);
    var value = stockValue(db, i.id);
    totalValue += value;
    var next = nextUnitCost(db, i.id);
    var avg = avgCost(db, i.id);
    var spread = open.length > 1
      ? ' <span class="pill pill-warn" title="มีหลายราคาในสต็อก">' + open.length + ' ล็อต</span>' : '';
    var first = open[0];
    if (first) {
      var st = lotExpiryState(db, first);
      if (st === 'expired') spread += ' <span class="pill pill-danger">หมดอายุ</span>';
      else if (st === 'soon') spread += ' <span class="pill pill-warn">อีก ' + daysUntil(first.expiry) + ' วัน</span>';
    }
    var status = bal <= 0 ? '<span class="pill pill-danger">หมด</span>'
      : (bal <= num(i.minQty) ? '<span class="pill pill-warn">ใกล้หมด</span>' : '<span class="pill pill-ok">ปกติ</span>');
    return '<tr>' +
      '<td><strong>' + esc(i.name) + '</strong>' + spread + '</td>' +
      '<td class="num">' + qtyStr(i.qtyPerPurchase) + ' ' + esc(i.unit) + ' / ' + esc(i.purchaseUnit) + '</td>' +
      '<td class="num"><strong>฿' + baht(next) + '</strong></td>' +
      '<td class="num">฿' + baht(avg) + '</td>' +
      '<td class="num">' + qtyStr(movedIn(db, i.id)) + '</td>' +
      '<td class="num">' + qtyStr(movedOut(db, i.id)) + '</td>' +
      '<td class="num"><strong>' + qtyStr(bal) + '</strong> ' + esc(i.unit) + '</td>' +
      '<td class="num">฿' + baht(value) + '</td>' +
      '<td>' + status + '</td>' +
      '<td><button class="btn btn-sm" data-receive="' + esc(i.id) + '">รับเข้า</button> ' +
          '<button class="btn btn-sm" data-adjust="' + esc(i.id) + '">ตัดออก</button> ' +
          '<button class="btn btn-sm" data-editing="' + esc(i.id) + '">แก้ไข</button></td>' +
      '</tr>';
  }).join('');
  return '<table><thead><tr>' +
    '<th>วัตถุดิบ</th><th class="num">ขนาดที่ซื้อ</th>' +
    '<th class="num" title="ต้นทุนของล็อตที่กำลังถูกหยิบใช้">ต้นทุนล็อตปัจจุบัน</th>' +
    '<th class="num" title="ต้นทุนถัวเฉลี่ยของของที่ยังเหลือในสต็อก">ถัวเฉลี่ยคงเหลือ</th>' +
    '<th class="num">รับเข้า</th><th class="num">จ่ายออก</th><th class="num">คงเหลือ</th>' +
    '<th class="num">มูลค่า (' + methodLabel(db) + ')</th>' + '<th>สถานะ</th><th></th>' +
    '</tr></thead><tbody>' + body + '</tbody>' +
    '<tfoot><tr><td colspan="7">มูลค่าสต็อกคงเหลือรวม</td><td class="num">฿' + baht(totalValue) + '</td><td colspan="2"></td></tr></tfoot></table>';
}

function bindStockButtons() {
  $$('[data-lotexp]').forEach(function (el) {
    el.addEventListener('change', function () {
      var id = el.getAttribute('data-lotexp');
      db.lots.forEach(function (l) { if (l.id === id) l.expiry = el.value || null; });
      save(); render();
      toast(el.value ? 'อัปเดตวันหมดอายุแล้ว — คิวการใช้ถูกจัดใหม่' : 'ลบวันหมดอายุแล้ว');
    });
  });
  $$('[data-receive]').forEach(function (b) { b.addEventListener('click', function () { receiveDialog(b.getAttribute('data-receive')); }); });
  $$('[data-adjust]').forEach(function (b) { b.addEventListener('click', function () { adjustDialog(b.getAttribute('data-adjust')); }); });
  $$('[data-editing]').forEach(function (b) { b.addEventListener('click', function () { ingredientDialog(b.getAttribute('data-editing')); }); });
}

/** ตัดของเสีย / ปรับยอดลง — ตัดจากล็อตเก่าสุดเหมือนการขาย */
function adjustDialog(ingId) {
  var ing = findIng(db, ingId);
  if (!ing) return;
  var open = openLots(db, ingId);
  openModal('ตัดสต็อกออก — ' + ing.name,
    '<p class="hint">ตัดจากล็อตหัวคิวก่อนตามหลัก ' + methodLabel(db) + ' เช่นเดียวกับการขาย</p>' +
    '<div class="form-row">' +
      '<label class="field"><span>จำนวนที่ตัดออก (' + esc(ing.unit) + ')</span><input class="input" id="adjQty" type="number" min="0" step="0.01" value="0"></label>' +
      '<label class="field"><span>เหตุผล</span><select class="input" id="adjType">' +
        '<option value="waste">ของเสีย / หมดอายุ / ทำหก</option>' +
        '<option value="adjust">ปรับยอดให้ตรงของจริง</option>' +
      '</select></label>' +
    '</div>' +
    '<label class="field"><span>หมายเหตุ</span><input class="input" id="adjNote" placeholder="ไม่บังคับ"></label>' +
    '<p class="hint">คงเหลือตอนนี้ ' + qtyStr(balance(db, ingId)) + ' ' + esc(ing.unit) +
      (open.length ? ' · ล็อตหัวคิว L' + open[0].seq + ' ราคา ฿' + baht(open[0].unitCost) + '/' + esc(ing.unit) : '') + '</p>' +
    '<p class="hint" id="adjPreview"></p>',
    [
      { label: 'ยกเลิก', onClick: closeModal },
      { label: 'ตัดออก', cls: 'btn-danger-ghost', onClick: function () {
        var q = num($('#adjQty').value);
        if (q <= 0) { toast('ใส่จำนวนที่จะตัดออก', true); return; }
        var type = $('#adjType').value;
        var res = issueStock(db, ingId, q, type,
          $('#adjNote').value || (type === 'waste' ? 'ตัดของเสีย' : 'ปรับยอดสต็อก'));
        save(); closeModal(); render();
        toast('ตัด ' + ing.name + ' ออก ' + qtyStr(q) + ' ' + ing.unit + ' · มูลค่า ฿' + baht(res.cost));
      } }
    ]);

  function preview() {
    var q = num($('#adjQty').value);
    if (q <= 0) { $('#adjPreview').textContent = ''; return; }
    var res = makeLedger(db)(ingId, q, false);
    $('#adjPreview').innerHTML = 'มูลค่าที่จะตัดออก ฿' + baht(res.cost) + ' — ' +
      res.allocs.map(function (a) {
        return a.lotId ? esc(lotLabel(db, a.lotId)) + ' × ' + qtyStr(a.qty) : 'เกินสต็อก ' + qtyStr(a.qty);
      }).join(' + ');
  }
  $('#adjQty').addEventListener('input', preview);
}

function receiveDialog(ingId) {
  var opts = db.ingredients.map(function (i) {
    return '<option value="' + esc(i.id) + '"' + (i.id === ingId ? ' selected' : '') + '>' + esc(i.name) + '</option>';
  }).join('');
  openModal('รับสินค้าเข้าสต็อก',
    '<p class="hint">ของที่รับเข้าจะถูกเก็บเป็น “ล็อตใหม่” พร้อมราคาซื้อของรอบนี้ ' +
    'ของเก่าที่ยังเหลือจะถูกใช้ก่อนด้วยราคาเดิมของมัน</p>' +
    '<div class="form-row">' +
      '<label class="field"><span>วัตถุดิบ</span><select class="input" id="rcvIng">' + opts + '</select></label>' +
      '<label class="field"><span>จำนวนที่ซื้อ</span><input class="input" id="rcvPacks" type="number" min="0" step="0.01" value="1"></label>' +
    '</div>' +
    '<div class="form-row">' +
      '<label class="field"><span>ราคาซื้อรวมรอบนี้ (บาท) — เว้นว่างถ้าราคาเดิม</span><input class="input" id="rcvPrice" type="number" min="0" step="0.01" placeholder="ราคาเดิม"></label>' +
      '<label class="field"><span>วันหมดอายุของล็อตนี้</span><input class="input" id="rcvExpiry" type="date"></label>' +
    '</div>' +
    '<label class="field"><span>หมายเหตุ</span><input class="input" id="rcvNote" placeholder="เช่น ซื้อจากร้าน…"></label>' +
    '<p class="hint" id="rcvPreview"></p>',
    [
      { label: 'ยกเลิก', onClick: closeModal },
      { label: 'บันทึกรับเข้า', cls: 'btn-primary', onClick: function () {
        var ing = findIng(db, $('#rcvIng').value);
        var packs = num($('#rcvPacks').value);
        if (!ing || packs <= 0) { toast('กรอกจำนวนให้ถูกต้อง', true); return; }
        var priceTotal = $('#rcvPrice').value;
        if (priceTotal !== '' && num(priceTotal) > 0) ing.purchasePrice = r2(num(priceTotal) / packs);
        var qtyIn = rq(packs * ing.qtyPerPurchase);
        var unitCost = ing.purchasePrice / ing.qtyPerPurchase;
        var lot = receiveStock(db, ing.id, qtyIn, unitCost,
          $('#rcvNote').value || ('รับเข้า ' + qtyStr(packs) + ' ' + ing.purchaseUnit),
          null, $('#rcvExpiry').value || null);
        save(); closeModal(); render();
        toast('รับ ' + ing.name + ' เข้าล็อต L' + lot.seq + ' · ฿' + baht(unitCost) + '/' + ing.unit);
      } }
    ]);

  function fillExpiry() {
    var ing = findIng(db, $('#rcvIng').value);
    $('#rcvExpiry').value = (ing && num(ing.shelfLifeDays) > 0) ? addDays(num(ing.shelfLifeDays)) : '';
  }

  function preview() {
    var ing = findIng(db, $('#rcvIng').value);
    var packs = num($('#rcvPacks').value);
    if (!ing) return;
    var priceTotal = $('#rcvPrice').value;
    var unitCost = (priceTotal !== '' && num(priceTotal) > 0)
      ? num(priceTotal) / packs / ing.qtyPerPurchase
      : costPerUnit(ing);
    var open = openLots(db, ing.id);
    var exp = $('#rcvExpiry').value;
    var older = open.filter(function (l) { return !exp || !l.expiry || l.expiry <= exp; }).length;
    $('#rcvPreview').innerHTML =
      'จะสร้างล็อตใหม่ ' + qtyStr(packs * ing.qtyPerPurchase) + ' ' + esc(ing.unit) +
      ' ที่ต้นทุน ฿' + baht(unitCost) + '/' + esc(ing.unit) +
      ' (1 ' + esc(ing.purchaseUnit) + ' = ' + qtyStr(ing.qtyPerPurchase) + ' ' + esc(ing.unit) + ')' +
      (num(ing.shelfLifeDays) > 0 ? '<br>อายุสินค้าที่ตั้งไว้ ' + qtyStr(ing.shelfLifeDays) + ' วัน (แก้วันหมดอายุเองได้)' : '') +
      (open.length
        ? '<br>ยังมีของเก่าค้างอยู่ ' + qtyStr(lotQtyLeft(db, ing.id)) + ' ' + esc(ing.unit) +
          ' ใน ' + open.length + ' ล็อต' +
          (older ? ' — ' + older + ' ล็อตจะถูกใช้ก่อนล็อตนี้' : ' — ล็อตนี้จะถูกใช้ก่อน เพราะหมดอายุเร็วกว่า')
        : '');
  }
  $('#rcvIng').addEventListener('change', function () { fillExpiry(); preview(); });
  $('#rcvExpiry').addEventListener('change', preview);
  fillExpiry();
  $('#rcvPacks').addEventListener('input', preview);
  $('#rcvPrice').addEventListener('input', preview);
  preview();
}

function ingredientDialog(ingId) {
  var ing = ingId ? findIng(db, ingId) : null;
  var isNew = !ing;
  if (isNew) ing = { id: uid('ing'), name: '', group: 'other', purchaseUnit: 'ขวด', purchasePrice: 0, qtyPerPurchase: 1, unit: 'มล', minQty: 0 };
  openModal(isNew ? 'เพิ่มวัตถุดิบ' : 'แก้ไข ' + ing.name,
    '<div class="form-row">' +
      '<label class="field"><span>ชื่อ</span><input class="input" id="ingName" value="' + esc(ing.name) + '"></label>' +
      '<label class="field"><span>ประเภท</span><select class="input" id="ingGroup">' +
        '<option value="matcha"' + (ing.group === 'matcha' ? ' selected' : '') + '>ผงมัทฉะ</option>' +
        '<option value="other"' + (ing.group === 'other' ? ' selected' : '') + '>วัตถุดิบ/บรรจุภัณฑ์</option>' +
      '</select></label>' +
    '</div>' +
    '<div class="form-row">' +
      '<label class="field"><span>ซื้อเป็นหน่วย</span><input class="input" id="ingPUnit" value="' + esc(ing.purchaseUnit) + '"></label>' +
      '<label class="field"><span>ราคาต่อหน่วยซื้อ (บาท)</span><input class="input" id="ingPrice" type="number" min="0" step="0.01" value="' + ing.purchasePrice + '"></label>' +
    '</div>' +
    '<div class="form-row">' +
      '<label class="field"><span>ปริมาณต่อหน่วยซื้อ</span><input class="input" id="ingQty" type="number" min="0" step="0.01" value="' + ing.qtyPerPurchase + '"></label>' +
      '<label class="field"><span>หน่วยที่ใช้ในสูตร</span><input class="input" id="ingUnit" value="' + esc(ing.unit) + '"></label>' +
      '<label class="field"><span>แจ้งเตือนเมื่อเหลือต่ำกว่า</span><input class="input" id="ingMin" type="number" min="0" step="0.01" value="' + num(ing.minQty) + '"></label>' +
      '<label class="field"><span>อายุสินค้าหลังรับเข้า (วัน, 0 = ไม่มีวันหมดอายุ)</span><input class="input" id="ingShelf" type="number" min="0" step="1" value="' + num(ing.shelfLifeDays) + '"></label>' +
    '</div>' +
    (isNew ? '<div class="form-row"><label class="field"><span>สต็อกเริ่มต้น (' + esc(ing.unit) + ')</span><input class="input" id="ingOpen" type="number" min="0" step="0.01" value="0"></label></div>' : '') +
    '<p class="hint" id="ingPreview"></p>',
    [
      { label: 'ยกเลิก', onClick: closeModal }
    ].concat(isNew ? [] : [{ label: 'ลบวัตถุดิบ', cls: 'btn-danger-ghost', onClick: function () { removeIngredient(ing.id); } }])
     .concat([{ label: 'บันทึก', cls: 'btn-primary', onClick: function () {
        var name = $('#ingName').value.trim();
        if (!name) { toast('ใส่ชื่อวัตถุดิบด้วย', true); return; }
        ing.name = name;
        ing.group = $('#ingGroup').value;
        ing.purchaseUnit = $('#ingPUnit').value.trim() || 'หน่วย';
        ing.purchasePrice = num($('#ingPrice').value);
        ing.qtyPerPurchase = Math.max(num($('#ingQty').value), 0.0001);
        ing.unit = $('#ingUnit').value.trim() || 'หน่วย';
        ing.minQty = num($('#ingMin').value);
        ing.shelfLifeDays = num($('#ingShelf').value);
        if (isNew) {
          db.ingredients.push(ing);
          var open = num($('#ingOpen').value);
          if (open > 0) receiveStock(db, ing.id, open, costPerUnit(ing), 'สต็อกเริ่มต้น', null,
            ing.shelfLifeDays > 0 ? addDays(ing.shelfLifeDays) : null);
        }
        save(); closeModal(); render();
        toast('บันทึก ' + ing.name + ' แล้ว');
      } }]));

  function preview() {
    var p = num($('#ingPrice').value), q = Math.max(num($('#ingQty').value), 0.0001);
    $('#ingPreview').textContent = 'ต้นทุนต่อ ' + ($('#ingUnit').value || 'หน่วย') + ' = ฿' + baht(p / q);
  }
  ['#ingPrice', '#ingQty', '#ingUnit'].forEach(function (s) { $(s).addEventListener('input', preview); });
  preview();
}

function removeIngredient(ingId) {
  var used = db.menus.filter(function (m) { return m.recipe.some(function (l) { return l.ingId === ingId; }); });
  if (used.length) {
    toast('ลบไม่ได้ — ยังถูกใช้ในเมนู: ' + used.map(function (m) { return m.name; }).join(', '), true);
    return;
  }
  confirmDialog('ลบวัตถุดิบ', 'ลบวัตถุดิบนี้พร้อมล็อตและประวัติสต็อกทั้งหมด?', function () {
    db.ingredients = db.ingredients.filter(function (i) { return i.id !== ingId; });
    db.moves = db.moves.filter(function (m) { return m.ingId !== ingId; });
    db.lots = db.lots.filter(function (l) { return l.ingId !== ingId; });
    save(); closeModal(); render(); toast('ลบแล้ว');
  }, 'ลบ');
}

/* ============================================================
   VIEW: เมนู & สูตร
   ============================================================ */
function renderMenus() {
  if (!db.menus.length) {
    $('#menuCards').innerHTML = '<div class="empty">ยังไม่มีเมนู</div>';
    return;
  }
  $('#menuCards').innerHTML = db.menus.map(function (m) {
    var take = makeLedger(db);            /* กระดานเดียว — ยอดรายบรรทัดจึงรวมได้เท่ากับต้นทุนจริง */
    var cost = 0;
    var spans = 0;
    var recipeRows = m.recipe.map(function (l) {
      var ing = findIng(db, l.ingId);
      var res = take(l.ingId, l.qty, false);
      cost += res.cost;
      if (res.allocs.length > 1) spans += 1;
      var lots = res.allocs.map(function (a) {
        return a.lotId ? esc(lotLabel(db, a.lotId)) : 'เกินสต็อก';
      }).join(' + ');
      return '<tr><td>' + esc(ing ? ing.name : '—') + '</td>' +
        '<td class="num">' + qtyStr(l.qty) + ' ' + esc(ing ? ing.unit : '') + '</td>' +
        '<td class="num">฿' + baht(l.qty > 0 ? res.cost / l.qty : 0) + '</td>' +
        '<td class="num">฿' + baht(res.cost) + '</td>' +
        '<td>' + lots + '</td></tr>';
    }).join('');
    cost = r2(cost);
    var stdCost = r2(menuCostStd(db, m));
    var margin = m.price > 0 ? (m.price - cost) / m.price * 100 : 0;
    var platRows = db.platforms.map(function (p) {
      var net = m.price * (1 - num(p.gp) / 100);
      var profit = net - cost;
      return '<tr><td>' + esc(p.name) + '</td>' +
        '<td class="num">' + qtyStr(p.gp) + '%</td>' +
        '<td class="num">฿' + baht(net) + '</td>' +
        '<td class="num ' + (profit >= 0 ? 'pos' : 'neg') + '">฿' + baht(profit) + '</td></tr>';
    }).join('');

    return '<div class="recipe-card">' +
      '<div class="recipe-head">' +
        '<div><h3>' + esc(m.name) + (m.active === false ? ' <span class="pill pill-plain">ปิดขาย</span>' : '') +
        (spans ? ' <span class="pill pill-warn">แก้วนี้ใช้ของ ' + (spans + 1) + ' ราคา</span>' : '') + '</h3>' +
        '<div class="r-sub">ราคาขาย ฿' + baht(m.price) + ' · ต้นทุนแก้วถัดไป (' + methodLabel(db) + ') ฿' + baht(cost) +
        ' · กำไรขั้นต้น ' + qtyStr(margin) + '% · ทำได้อีก ' + maxCups(db, m) + ' แก้ว' +
        (r2(Math.abs(stdCost - cost)) >= 0.01
          ? '<br>ถ้าคิดด้วยราคาซื้อล่าสุดจะเป็น ฿' + baht(stdCost) +
            ' (ต่างกัน ' + (cost > stdCost ? '+' : '−') + '฿' + baht(Math.abs(cost - stdCost)) + ')'
          : '') +
        '</div></div>' +
        '<div class="view-actions">' +
          '<button class="btn btn-sm" data-menuedit="' + esc(m.id) + '">แก้ไขสูตร</button>' +
          '<button class="btn btn-sm" data-menudup="' + esc(m.id) + '">คัดลอก</button>' +
          '<button class="btn btn-sm btn-danger-ghost" data-menudel="' + esc(m.id) + '">ลบ</button>' +
        '</div>' +
      '</div>' +
      (optionGroups(m).length
        ? '<div class="opt-chiplist">' + optionGroups(m).map(function (g) {
            var ing = findIng(db, g.target);
            return '<span class="opt-chiplist-g"><strong>' + esc(g.name) + '</strong>' +
              '<span class="hint"> (' + (g.type === 'swap' ? 'สลับ' : 'ปรับปริมาณ') + ' ' + esc(ing ? ing.name : '?') + ')</span>: ' +
              (g.choices || []).map(function (c) {
                return esc(c.label) + (num(c.priceDelta) ? ' +' + qtyStr(c.priceDelta) + '฿' : '') +
                  (g.type === 'scale' && num(c.factor, 1) !== 1 ? ' ×' + qtyStr(c.factor) : '');
              }).join(' / ') + '</span>';
          }).join('') + '</div>'
        : '') +
      '<div class="recipe-body">' +
        '<div class="table-wrap"><table class="mini-table"><thead><tr><th>วัตถุดิบ</th><th class="num">ปริมาณ/แก้ว</th><th class="num">ต้นทุน/หน่วย</th><th class="num">รวม</th><th>ตัดจากล็อต</th></tr></thead>' +
        '<tbody>' + recipeRows + '</tbody>' +
        '<tfoot><tr><td colspan="3">ต้นทุนรวมต่อแก้ว</td><td class="num">฿' + baht(cost) + '</td><td></td></tr></tfoot></table></div>' +
        '<div class="table-wrap"><table class="mini-table"><thead><tr><th>ช่องทาง</th><th class="num">GP</th><th class="num">รับสุทธิ</th><th class="num">กำไร/แก้ว</th></tr></thead>' +
        '<tbody>' + platRows + '</tbody></table></div>' +
      '</div></div>';
  }).join('');

  $$('[data-menuedit]').forEach(function (b) { b.addEventListener('click', function () { menuDialog(b.getAttribute('data-menuedit')); }); });
  $$('[data-menudup]').forEach(function (b) { b.addEventListener('click', function () {
    var m = findMenu(db, b.getAttribute('data-menudup'));
    if (!m) return;
    var copy = JSON.parse(JSON.stringify(m));
    copy.id = uid('menu'); copy.name = m.name + ' (สำเนา)';
    db.menus.push(copy); save(); render(); toast('คัดลอกเมนูแล้ว');
  }); });
  $$('[data-menudel]').forEach(function (b) { b.addEventListener('click', function () {
    var id = b.getAttribute('data-menudel');
    confirmDialog('ลบเมนู', 'ลบเมนูนี้? ประวัติการขายเดิมจะยังอยู่', function () {
      db.menus = db.menus.filter(function (m) { return m.id !== id; });
      state.cart = state.cart.filter(function (c) { return c.menuId !== id; });
      save(); render(); toast('ลบเมนูแล้ว');
    }, 'ลบ');
  }); });
}

function menuDialog(menuId) {
  var m = menuId ? findMenu(db, menuId) : null;
  var isNew = !m;
  var draft = isNew
    ? { id: uid('menu'), name: '', price: 0, color: '#4f7a34', active: true, recipe: [], options: [] }
    : JSON.parse(JSON.stringify(m));
  if (!Array.isArray(draft.options)) draft.options = [];
  var optionsDrawn = false;

  openModal(isNew ? 'เพิ่มเมนูใหม่' : 'แก้ไข ' + m.name,
    '<div class="form-row">' +
      '<label class="field"><span>ชื่อเมนู</span><input class="input" id="mName" value="' + esc(draft.name) + '"></label>' +
      '<label class="field"><span>ราคาขาย (บาท)</span><input class="input" id="mPrice" type="number" min="0" step="1" value="' + draft.price + '"></label>' +
      '<label class="field"><span>สีประจำเมนู</span><input class="input" id="mColor" type="color" value="' + esc(draft.color || '#4f7a34') + '"></label>' +
    '</div>' +
    '<label class="field-inline" style="margin-bottom:12px"><input type="checkbox" id="mActive"' + (draft.active !== false ? ' checked' : '') + '><span>เปิดขายในหน้า POS</span></label>' +
    '<h4 class="section-title" style="margin-top:0">สูตร (ต่อ 1 แก้ว)</h4>' +
    '<div id="recipeRows"></div>' +
    '<button class="btn btn-sm" id="addLineBtn">+ เพิ่มวัตถุดิบ</button>' +
    '<p class="hint" id="mCostPreview"></p>' +
    '<h4 class="section-title">ตัวเลือกของเมนูนี้</h4>' +
    '<p class="hint">ตัวเลือกจะปรับปริมาณวัตถุดิบและราคาขายตอนสั่งจริง เช่น “หวานน้อย” = น้ำเชื่อมครึ่งเดียว</p>' +
    '<div id="optionRows"></div>' +
    '<div class="btn-row">' +
      '<button class="btn btn-sm" data-preset="sweet">+ ความหวาน</button>' +
      '<button class="btn btn-sm" data-preset="powder">+ เลือกผงมัทฉะ</button>' +
      '<button class="btn btn-sm" data-preset="strength">+ ความเข้ม</button>' +
    '</div>',
    [
      { label: 'ยกเลิก', onClick: closeModal },
      { label: 'บันทึก', cls: 'btn-primary', onClick: function () {
        collect();
        if (!draft.name.trim()) { toast('ใส่ชื่อเมนูด้วย', true); return; }
        draft.recipe = draft.recipe.filter(function (l) { return l.ingId && l.qty > 0; });
        if (isNew) db.menus.push(draft);
        else {
          var idx = db.menus.findIndex(function (x) { return x.id === menuId; });
          db.menus[idx] = draft;
        }
        save(); closeModal(); render(); toast('บันทึกเมนูแล้ว');
      } }
    ]);

  function collect() {
    draft.name = $('#mName').value;
    draft.price = num($('#mPrice').value);
    draft.color = $('#mColor').value;
    draft.active = $('#mActive').checked;
    draft.recipe = $$('#recipeRows .recipe-row').map(function (row) {
      return { ingId: $('select', row).value, qty: num($('input', row).value) };
    });
    collectOptions();
  }

  function ingOptions(selId) {
    return db.ingredients.map(function (i) {
      return '<option value="' + esc(i.id) + '"' + (i.id === selId ? ' selected' : '') + '>' + esc(i.name) + '</option>';
    }).join('');
  }

  function collectOptions() {
    var wrap = $('#optionRows');
    if (!wrap || !optionsDrawn) return;   /* ยังไม่ได้วาด — อย่าเพิ่งอ่านค่าจาก DOM ว่างเปล่า */
    draft.options = $$('.opt-edit', wrap).map(function (el, gi) {
      var old = draft.options[gi] || {};
      var choices = $$('.opt-choice', el).map(function (ce, ci) {
        var oc = (old.choices && old.choices[ci]) || {};
        var ingSel = $('[data-f="ingId"]', ce);
        return {
          id: oc.id || uid('c'),
          label: $('[data-f="label"]', ce).value.trim() || 'ตัวเลือก',
          factor: num($('[data-f="factor"]', ce).value, 1),
          ingId: ingSel ? ingSel.value : undefined,
          priceDelta: num($('[data-f="priceDelta"]', ce).value)
        };
      });
      var defIdx = 0;
      $$('.opt-choice [data-f="def"]', el).forEach(function (r, i) { if (r.checked) defIdx = i; });
      return {
        id: old.id || uid('g'),
        name: $('[data-f="name"]', el).value.trim() || 'ตัวเลือก',
        type: $('[data-f="type"]', el).value,
        target: $('[data-f="target"]', el).value,
        alwaysShow: !!old.alwaysShow,
        defaultId: choices[defIdx] ? choices[defIdx].id : null,
        choices: choices
      };
    }).filter(function (g) { return g.choices.length; });
  }

  function drawOptions() {
    $('#optionRows').innerHTML = draft.options.length ? draft.options.map(function (g, gi) {
      var isSwap = g.type === 'swap';
      return '<div class="opt-edit" data-gi="' + gi + '">' +
        '<div class="opt-edit-head">' +
          '<input class="input input-sm" data-f="name" value="' + esc(g.name) + '" placeholder="ชื่อกลุ่ม">' +
          '<select class="input input-sm" data-f="type">' +
            '<option value="scale"' + (!isSwap ? ' selected' : '') + '>ปรับปริมาณ</option>' +
            '<option value="swap"' + (isSwap ? ' selected' : '') + '>สลับวัตถุดิบ</option>' +
          '</select>' +
          '<select class="input input-sm" data-f="target" title="ปรับวัตถุดิบตัวไหนในสูตร">' + ingOptions(g.target) + '</select>' +
          '<button class="icon-btn" data-rmgroup="' + gi + '" title="ลบกลุ่ม">✕</button>' +
        '</div>' +
        '<div class="opt-choice-head' + (isSwap ? ' is-swap' : '') + '"><span></span><span>ชื่อที่ลูกค้าเห็น</span>' +
          (isSwap ? '<span>ใช้วัตถุดิบ</span>' : '') + '<span>×ปริมาณ</span><span>+ราคา</span><span></span></div>' +
        (g.choices || []).map(function (c, ci) {
          return '<div class="opt-choice' + (isSwap ? ' is-swap' : '') + '" data-ci="' + ci + '">' +
            '<input type="radio" name="optdef' + gi + '" data-f="def"' + (c.id === g.defaultId ? ' checked' : '') + ' title="ตั้งเป็นค่าเริ่มต้น">' +
            '<input class="input input-sm" data-f="label" value="' + esc(c.label) + '">' +
            (isSwap ? '<select class="input input-sm" data-f="ingId">' + ingOptions(c.ingId) + '</select>' : '') +
            '<input class="input input-sm" type="number" step="0.05" min="0" data-f="factor" value="' + num(c.factor, 1) + '">' +
            '<input class="input input-sm" type="number" step="1" data-f="priceDelta" value="' + num(c.priceDelta) + '">' +
            '<button class="icon-btn" data-rmchoice="' + gi + ':' + ci + '" title="ลบ">✕</button>' +
          '</div>';
        }).join('') +
        '<button class="btn btn-sm" data-addchoice="' + gi + '">+ ตัวเลือกย่อย</button>' +
        '</div>';
    }).join('') : '<p class="hint">ยังไม่มีตัวเลือก — กดปุ่มด้านล่างเพื่อเพิ่มแบบสำเร็จรูป</p>';

    $$('#optionRows [data-rmgroup]').forEach(function (b) {
      b.addEventListener('click', function () {
        collectOptions();
        draft.options.splice(Number(b.getAttribute('data-rmgroup')), 1);
        drawOptions();
      });
    });
    $$('#optionRows [data-rmchoice]').forEach(function (b) {
      b.addEventListener('click', function () {
        collectOptions();
        var p = b.getAttribute('data-rmchoice').split(':');
        draft.options[Number(p[0])].choices.splice(Number(p[1]), 1);
        drawOptions();
      });
    });
    $$('#optionRows [data-addchoice]').forEach(function (b) {
      b.addEventListener('click', function () {
        collectOptions();
        var g = draft.options[Number(b.getAttribute('data-addchoice'))];
        g.choices.push({ id: uid('c'), label: 'ตัวเลือกใหม่', factor: 1, ingId: g.target, priceDelta: 0 });
        drawOptions();
      });
    });
    $$('#optionRows [data-f="type"], #optionRows [data-f="target"]').forEach(function (el) {
      el.addEventListener('change', function () { collectOptions(); drawOptions(); });
    });
    optionsDrawn = true;
  }

  function addPreset(key) {
    collect();
    var fallback = (draft.recipe[0] || {}).ingId || (db.ingredients[0] || {}).id;
    var g = buildPresetGroup(db, draft, key, fallback);
    if (!g) { toast('เพิ่มตัวเลือกไม่ได้ — ยังไม่มีวัตถุดิบในสูตร', true); return; }
    draft.options.push(g);
    drawOptions();
    toast('เพิ่ม “' + g.name + '” แล้ว — ปรับรายละเอียดได้ในตาราง');
  }

  function drawRows() {
    $('#recipeRows').innerHTML = draft.recipe.map(function (l, idx) {
      var ing = findIng(db, l.ingId);
      return '<div class="recipe-row" data-idx="' + idx + '">' +
        '<select class="input input-sm">' + db.ingredients.map(function (i) {
          return '<option value="' + esc(i.id) + '"' + (i.id === l.ingId ? ' selected' : '') + '>' + esc(i.name) + '</option>';
        }).join('') + '</select>' +
        '<input class="input input-sm" type="number" min="0" step="0.01" value="' + l.qty + '">' +
        '<span class="unit">' + esc(ing ? ing.unit : '') + '</span>' +
        '<button class="icon-btn" data-rmline="' + idx + '" title="ลบ">✕</button>' +
        '</div>';
    }).join('') || '<p class="hint">ยังไม่มีวัตถุดิบในสูตร</p>';

    $$('#recipeRows [data-rmline]').forEach(function (b) {
      b.addEventListener('click', function () {
        collect();
        draft.recipe.splice(Number(b.getAttribute('data-rmline')), 1);
        drawRows();
      });
    });
    $$('#recipeRows select, #recipeRows input').forEach(function (el) {
      el.addEventListener('change', function () { collect(); drawRows(); });
    });
    costPreview();
  }

  function costPreview() {
    var take = makeLedger(db);
    var c = r2(draft.recipe.reduce(function (s, l) { return s + take(l.ingId, l.qty, false).cost; }, 0));
    var price = num($('#mPrice').value);
    $('#mCostPreview').textContent = 'ต้นทุนต่อแก้ว (' + methodLabel(db) + ') ฿' + baht(c) +
      (price > 0 ? ' · กำไรขั้นต้น ฿' + baht(price - c) + ' (' + qtyStr((price - c) / price * 100) + '%)' : '');
  }

  $('#addLineBtn').addEventListener('click', function () {
    collect();
    draft.recipe.push({ ingId: db.ingredients[0] ? db.ingredients[0].id : '', qty: 1 });
    drawRows();
  });
  $$('#modalBody [data-preset]').forEach(function (b) {
    b.addEventListener('click', function () { addPreset(b.getAttribute('data-preset')); });
  });
  $('#mPrice').addEventListener('input', costPreview);
  drawRows();
  drawOptions();
}

/* ============================================================
   VIEW: ประวัติการขาย
   ============================================================ */
function filteredSales() {
  var from = state.histFrom ? new Date(state.histFrom + 'T00:00:00').getTime() : -Infinity;
  var to = state.histTo ? new Date(state.histTo + 'T23:59:59').getTime() : Infinity;
  return db.sales.filter(function (s) {
    if (s.ts < from || s.ts > to) return false;
    if (state.histPlatform !== 'all' && s.platformId !== state.histPlatform) return false;
    return true;
  }).sort(function (a, b) { return b.ts - a.ts; });
}

function renderHistory() {
  var sel = $('#historyPlatform');
  sel.innerHTML = '<option value="all">ทั้งหมด</option>' + db.platforms.map(function (p) {
    return '<option value="' + esc(p.id) + '"' + (p.id === state.histPlatform ? ' selected' : '') + '>' + esc(p.name) + '</option>';
  }).join('');
  sel.value = state.histPlatform;

  var rows = filteredSales();
  var sum = summarize(rows);
  $('#historySummary').innerHTML =
    kpi('จำนวนบิล', String(sum.bills), sum.cups + ' แก้ว') +
    kpi('ยอดขายรวม', '฿' + baht(sum.gross)) +
    (sum.vat ? kpi('VAT', '฿' + baht(sum.vat), 'ต้องนำส่งสรรพากร', 'is-cost') : '') +
    kpi('ค่า GP', '฿' + baht(sum.gp), 'หักโดยแพลตฟอร์ม', 'is-cost') +
    kpi('ต้นทุน', '฿' + baht(sum.cost), '', 'is-cost') +
    kpi('กำไรสุทธิ', '฿' + baht(sum.profit), sum.gross ? qtyStr(sum.profit / sum.gross * 100) + '% ของยอดขาย' : '', 'is-profit');

  $('#historyTable').innerHTML = rows.length ? '<table><thead><tr>' +
    '<th>บิล</th><th>วันเวลา</th><th>รายการ</th><th class="num">แก้ว</th><th>ช่องทาง</th>' +
    '<th class="num">ยอดขาย</th><th class="num">ส่วนลด</th><th class="num">VAT</th><th class="num">GP</th><th class="num">สุทธิ</th>' +
    '<th class="num">ต้นทุน (' + methodLabel(db) + ')</th>' + '<th class="num">กำไร</th><th></th>' +
    '</tr></thead><tbody>' + rows.map(function (s) {
      return '<tr' + (s.closingId ? ' class="is-closed"' : '') + '>' +
        '<td><strong>' + esc(s.no) + '</strong>' +
          (s.closingId ? ' <span class="pill pill-plain" title="ปิดยอดแล้ว">🔒</span>' : '') + '</td>' +
        '<td>' + fmtDateTime(s.ts) + '</td>' +
        '<td>' + esc(s.items.map(function (i) {
          return i.name + (i.variant ? ' (' + i.variant + ')' : '') + ' ×' + i.qty;
        }).join(', ')) + '</td>' +
        '<td class="num">' + s.cups + '</td>' +
        '<td><span class="pill pill-plain">' + esc(s.platformName) + '</span>' +
          (s.payMethodName && s.settlement !== 'platform' ? ' <span class="pill pill-plain">' + esc(s.payMethodName) + '</span>' : '') + '</td>' +
        '<td class="num">฿' + baht(s.gross) + '</td>' +
        '<td class="num">' + (s.discount ? '−฿' + baht(s.discount) : '–') + '</td>' +
        '<td class="num">' + (num(s.vat) ? '฿' + baht(s.vat) : '–') + '</td>' +
        '<td class="num">' + (s.gpAmount ? '−฿' + baht(s.gpAmount) : '–') + '</td>' +
        '<td class="num">฿' + baht(s.net) + '</td>' +
        '<td class="num">฿' + baht(s.cost) + '</td>' +
        '<td class="num ' + (s.profit >= 0 ? 'pos' : 'neg') + '"><strong>฿' + baht(s.profit) + '</strong></td>' +
        '<td><button class="btn btn-sm" data-bill="' + esc(s.id) + '">ใบเสร็จ</button> ' +
            '<button class="btn btn-sm" data-costof="' + esc(s.id) + '">ต้นทุน</button> ' +
            '<button class="btn btn-sm btn-danger-ghost" data-delbill="' + esc(s.id) + '">ลบ</button></td>' +
        '</tr>';
    }).join('') + '</tbody>' +
    '<tfoot><tr><td colspan="5">รวม</td>' +
      '<td class="num">฿' + baht(sum.gross) + '</td>' +
      '<td class="num">' + (sum.discount ? '−฿' + baht(sum.discount) : '–') + '</td>' +
      '<td class="num">' + (sum.vat ? '฿' + baht(sum.vat) : '–') + '</td>' +
      '<td class="num">' + (sum.gp ? '−฿' + baht(sum.gp) : '–') + '</td>' +
      '<td class="num">฿' + baht(sum.net) + '</td>' +
      '<td class="num">฿' + baht(sum.cost) + '</td>' +
      '<td class="num">฿' + baht(sum.profit) + '</td><td></td></tr></tfoot></table>'
    : '<div class="empty">ไม่มีบิลในช่วงที่เลือก</div>';

  $$('[data-bill]').forEach(function (b) {
    b.addEventListener('click', function () {
      var s = db.sales.filter(function (x) { return x.id === b.getAttribute('data-bill'); })[0];
      if (s) showReceipt(s);
    });
  });
  $$('[data-costof]').forEach(function (b) {
    b.addEventListener('click', function () { costBreakdownDialog(b.getAttribute('data-costof')); });
  });
  $$('[data-delbill]').forEach(function (b) {
    b.addEventListener('click', function () {
      var id = b.getAttribute('data-delbill');
      var sale = db.sales.filter(function (x) { return x.id === id; })[0];
      if (sale && sale.closingId) {
        openModal('บิลนี้ปิดยอดแล้ว',
          '<p>บิล <strong>' + esc(sale.no) + '</strong> อยู่ในวันที่ปิดยอดไปแล้ว จึงแก้ไขไม่ได้</p>' +
          '<p class="hint">ถ้าจำเป็นต้องแก้จริง ๆ ให้ไปที่แท็บ <strong>ปิดยอด</strong> เลือกวันนั้น แล้วกด “ยกเลิกการปิดยอด” ก่อน</p>',
          [{ label: 'เข้าใจแล้ว', cls: 'btn-primary', onClick: closeModal }]);
        return;
      }
      confirmDialog('ลบบิล', 'ลบบิลนี้และคืนวัตถุดิบกลับเข้าล็อตเดิม?', function () {
        deleteSale(db, id); save(); render(); toast('ลบบิลแล้ว · คืนสต็อกเรียบร้อย');
      }, 'ลบบิล');
    });
  });
}

/** แจกแจงว่าบิลนี้ตัดวัตถุดิบจากล็อตไหน ราคาเท่าไหร่ */
function costBreakdownDialog(saleId) {
  var sale = db.sales.filter(function (s) { return s.id === saleId; })[0];
  if (!sale) return;
  var moves = db.moves.filter(function (m) { return m.ref === saleId; });

  var rows = moves.map(function (m) {
    var ing = findIng(db, m.ingId);
    var detail = (m.allocs || []).map(function (a) {
      return a.lotId
        ? esc(lotLabel(db, a.lotId)) + ' × ' + qtyStr(a.qty)
        : '<span class="neg">เกินสต็อก × ' + qtyStr(a.qty) + '</span>';
    }).join('<br>') || '–';
    return '<tr><td>' + esc(ing ? ing.name : m.ingId) + '</td>' +
      '<td class="num">' + qtyStr(-m.qty) + ' ' + esc(ing ? ing.unit : '') + '</td>' +
      '<td>' + detail + '</td>' +
      '<td class="num">฿' + baht(m.cost != null ? m.cost : 0) + '</td></tr>';
  }).join('');

  openModal('ต้นทุนตามล็อต — บิล ' + sale.no,
    '<p class="hint">ทุกบรรทัดคือวัตถุดิบที่ถูกตัดออกจริง เรียงตามลำดับที่ระบบหยิบใช้ (' + methodLabel(db) + ')</p>' +
    (rows
      ? '<div class="table-wrap"><table class="mini-table"><thead><tr>' +
        '<th>วัตถุดิบ</th><th class="num">ใช้ไป</th><th>ตัดจากล็อต</th><th class="num">ต้นทุน</th>' +
        '</tr></thead><tbody>' + rows + '</tbody>' +
        '<tfoot><tr><td colspan="3">ต้นทุนรวมของบิล</td><td class="num">฿' + baht(sale.cost) + '</td></tr>' +
        '<tr><td colspan="3">กำไรสุทธิ</td><td class="num">฿' + baht(sale.profit) + '</td></tr></tfoot></table></div>'
      : '<div class="empty">บิลนี้ไม่มีรายละเอียดการตัดสต็อก</div>'),
    [{ label: 'ปิด', cls: 'btn-primary', onClick: closeModal }]);
}

function summarize(rows) {
  var s = { bills: rows.length, cups: 0, gross: 0, discount: 0, vat: 0, exVat: 0, payable: 0, gp: 0, net: 0, cost: 0, profit: 0 };
  rows.forEach(function (x) {
    s.cups += x.cups; s.gross += x.gross; s.discount += x.discount;
    s.vat += num(x.vat); s.exVat += num(x.exVat, x.afterDiscount); s.payable += num(x.payable, x.afterDiscount);
    s.gp += x.gpAmount; s.net += x.net; s.cost += x.cost; s.profit += x.profit;
  });
  ['gross', 'discount', 'vat', 'exVat', 'payable', 'gp', 'net', 'cost', 'profit'].forEach(function (k) { s[k] = r2(s[k]); });
  return s;
}

function kpi(label, value, sub, cls) {
  return '<div class="kpi ' + (cls || '') + '"><div class="k-label">' + label + '</div>' +
    '<div class="k-value">' + value + '</div>' +
    (sub ? '<div class="k-sub">' + sub + '</div>' : '') + '</div>';
}

function exportSalesCsv() {
  var rows = filteredSales().slice().reverse();
  var head = ['เลขบิล', 'วันที่', 'เวลา', 'เมนู', 'จำนวน', 'ราคา/หน่วย', 'ช่องทาง', 'ยอดขาย', 'ส่วนลด', 'GP%', 'ค่า GP', 'ยอดสุทธิ', 'ต้นทุน/แก้ว (' + methodLabel(db) + ')', 'ต้นทุนรวม (' + methodLabel(db) + ')', 'กำไรสุทธิ'];
  var lines = [head];
  rows.forEach(function (s) {
    var d = new Date(s.ts);
    s.items.forEach(function (it, idx) {
      lines.push([
        s.no,
        dayKey(s.ts),
        d.toLocaleTimeString('th-TH', { hour: '2-digit', minute: '2-digit' }),
        it.name, it.qty, r2(it.price), s.platformName,
        r2(it.lineTotal),
        idx === 0 ? r2(s.discount) : 0,
        r2(s.gpRate),
        idx === 0 ? r2(s.gpAmount) : 0,
        idx === 0 ? r2(s.net) : 0,
        r2(it.unitCost != null ? it.unitCost : (it.qty ? it.lineCost / it.qty : 0)),
        r2(it.lineCost),
        idx === 0 ? r2(s.profit) : 0
      ]);
    });
  });
  var csv = lines.map(function (r) {
    return r.map(function (c) {
      var v = String(c);
      return /[",\n]/.test(v) ? '"' + v.replace(/"/g, '""') + '"' : v;
    }).join(',');
  }).join('\r\n');
  download('sales_' + dayKey(Date.now()) + '.csv', '﻿' + csv, 'text/csv;charset=utf-8');
  toast('ดาวน์โหลด CSV แล้ว (เปิดด้วย Excel ได้)');
}

/* ============================================================
   VIEW: ปิดยอดรายวัน (Z-report)
   ============================================================ */

function salesOfDay(d, key) {
  return d.sales.filter(function (s) { return dayKey(s.ts) === key; })
    .sort(function (a, b) { return a.ts - b.ts; });
}
function closingOf(d, key) {
  var list = d.closings || [];
  for (var i = 0; i < list.length; i++) if (list[i].dayKey === key) return list[i];
  return null;
}

/** รวมตัวเลขทุกอย่างของวันหนึ่ง ใช้ทั้งตอนแสดงฟอร์มและตอนบันทึกเป็นสแนปช็อต */
function buildDaySummary(d, key) {
  var rows = salesOfDay(d, key);
  var s = summarize(rows);

  var byPlatform = {};
  var byPay = {};
  var cashSales = 0, directNonCash = 0, platformReceivable = 0;

  rows.forEach(function (x) {
    var pk = x.platformName;
    if (!byPlatform[pk]) byPlatform[pk] = { name: pk, bills: 0, cups: 0, payable: 0, gp: 0, net: 0, settlement: x.settlement };
    byPlatform[pk].bills++; byPlatform[pk].cups += x.cups;
    byPlatform[pk].payable += num(x.payable, x.afterDiscount);
    byPlatform[pk].gp += x.gpAmount; byPlatform[pk].net += x.net;

    var direct = x.settlement !== 'platform';
    var pmName = direct ? (x.payMethodName || 'เงินสด') : x.platformName;
    if (!byPay[pmName]) byPay[pmName] = { name: pmName, bills: 0, amount: 0, inDrawer: false };
    byPay[pmName].bills++;
    byPay[pmName].amount += direct ? num(x.payable, x.afterDiscount) : x.net;

    if (direct) {
      var pm = payMethod(d, x.payMethodId);
      if (pm && pm.inDrawer) { cashSales += num(x.payable, x.afterDiscount); byPay[pmName].inDrawer = true; }
      else directNonCash += num(x.payable, x.afterDiscount);
    } else {
      platformReceivable += x.net;
    }
  });

  var waste = d.moves.filter(function (m) {
    return dayKey(m.ts) === key && (m.type === 'waste' || m.type === 'adjust');
  });
  var received = d.moves.filter(function (m) { return dayKey(m.ts) === key && m.type === 'in'; });

  return {
    dayKey: key,
    bills: s.bills, cups: s.cups,
    gross: s.gross, discount: s.discount, vat: s.vat, exVat: s.exVat, payable: s.payable,
    gp: s.gp, net: s.net, cost: s.cost, profit: s.profit,
    byPlatform: Object.keys(byPlatform).map(function (k) {
      var v = byPlatform[k];
      ['payable', 'gp', 'net'].forEach(function (f) { v[f] = r2(v[f]); });
      return v;
    }),
    byPay: Object.keys(byPay).map(function (k) {
      byPay[k].amount = r2(byPay[k].amount);
      return byPay[k];
    }),
    cashSales: r2(cashSales),
    directNonCash: r2(directNonCash),
    platformReceivable: r2(platformReceivable),
    fixedCost: dailyFixedCost(d),
    netProfit: r2(s.profit - dailyFixedCost(d)),
    breakEvenCups: breakEven(d, d.sales.slice(-200)).cups,
    wasteCount: waste.length,
    wasteValue: r2(waste.reduce(function (a, m) { return a + num(m.cost); }, 0)),
    receivedCount: received.length,
    receivedValue: r2(received.reduce(function (a, m) { return a + num(m.cost); }, 0)),
    stockWorth: r2(d.ingredients.reduce(function (a, i) { return a + stockValue(d, i.id); }, 0))
  };
}

function renderClose() {
  if (!state.closeDay) state.closeDay = dayKey(Date.now());
  $('#closeDate').value = state.closeDay;

  var key = state.closeDay;
  var closed = closingOf(db, key);
  var sum = closed ? closed.summary : buildDaySummary(db, key);
  var isFuture = key > dayKey(Date.now());

  var kpis =
    kpi('ยอดขายรวม', '฿' + baht(sum.gross), sum.bills + ' บิล · ' + sum.cups + ' แก้ว') +
    (sum.vat ? kpi('VAT ที่ต้องนำส่ง', '฿' + baht(sum.vat), 'ฐานภาษี ฿' + baht(sum.exVat), 'is-cost') : '') +
    kpi('ค่า GP', '฿' + baht(sum.gp), '', 'is-cost') +
    kpi('ต้นทุนวัตถุดิบ', '฿' + baht(sum.cost), '', 'is-cost') +
    kpi('กำไรขั้นต้น', '฿' + baht(sum.profit), sum.gross ? qtyStr(sum.profit / sum.gross * 100) + '% ของยอดขาย' : '', 'is-profit') +
    (num(sum.fixedCost) > 0
      ? kpi('ค่าใช้จ่ายคงที่/วัน', '฿' + baht(sum.fixedCost), '', 'is-cost') +
        kpi('กำไรจริงของวันนี้', signedBaht(sum.netProfit),
            num(sum.netProfit) >= 0 ? 'คุ้มทุนแล้ว' : 'ต้องขายอีก ' + Math.max(0, num(sum.breakEvenCups) - sum.cups) + ' แก้ว',
            num(sum.netProfit) >= 0 ? 'is-profit' : 'is-cost')
      : '');

  var platTable = sum.byPlatform.length
    ? '<table><thead><tr><th>ช่องทาง</th><th class="num">บิล</th><th class="num">แก้ว</th><th class="num">ลูกค้าจ่าย</th><th class="num">ค่า GP</th><th class="num">ร้านได้รับ</th></tr></thead><tbody>' +
      sum.byPlatform.map(function (v) {
        return '<tr><td>' + esc(v.name) + '</td><td class="num">' + v.bills + '</td><td class="num">' + v.cups + '</td>' +
          '<td class="num">฿' + baht(v.payable) + '</td><td class="num neg">' + (v.gp ? '−฿' + baht(v.gp) : '–') + '</td>' +
          '<td class="num">฿' + baht(v.net) + '</td></tr>';
      }).join('') + '</tbody></table>'
    : '<div class="empty">ไม่มีบิลในวันนี้</div>';

  var payTable = sum.byPay.length
    ? '<table><thead><tr><th>รับเงินทาง</th><th class="num">บิล</th><th class="num">จำนวนเงิน</th><th></th></tr></thead><tbody>' +
      sum.byPay.map(function (v) {
        return '<tr><td>' + esc(v.name) + '</td><td class="num">' + v.bills + '</td>' +
          '<td class="num">฿' + baht(v.amount) + '</td>' +
          '<td>' + (v.inDrawer ? '<span class="pill pill-ok">เข้าลิ้นชัก</span>' : '<span class="pill pill-plain">ไม่เข้าลิ้นชัก</span>') + '</td></tr>';
      }).join('') + '</tbody></table>'
    : '<div class="empty">ไม่มีบิลในวันนี้</div>';

  var html =
    '<div class="kpi-row">' + kpis + '</div>' +
    '<div class="two-col">' +
      '<div class="card"><h3 class="section-title">แยกตามช่องทาง</h3><div class="table-wrap">' + platTable + '</div></div>' +
      '<div class="card"><h3 class="section-title">แยกตามวิธีรับเงิน</h3><div class="table-wrap">' + payTable + '</div>' +
        '<p class="hint">ยอดรอรับจากแพลตฟอร์ม ฿' + baht(sum.platformReceivable) +
        ' · ของเสีย/ปรับยอดวันนี้ ' + sum.wasteCount + ' รายการ มูลค่า ฿' + baht(sum.wasteValue) +
        ' · รับของเข้า ' + sum.receivedCount + ' รายการ ฿' + baht(sum.receivedValue) + '</p>' +
      '</div>' +
    '</div>';

  if (closed) {
    html +=
      '<div class="card">' +
        '<h3 class="section-title">ปิดยอดแล้ว — ' + esc(closed.no) + '</h3>' +
        '<div class="kpi-row">' +
          kpi('เงินทอนตั้งต้น', '฿' + baht(closed.openingFloat)) +
          kpi('ยอดขายเงินสด', '฿' + baht(closed.summary.cashSales)) +
          kpi('เงินสดที่ควรมี', '฿' + baht(closed.expectedCash)) +
          kpi('นับได้จริง', '฿' + baht(closed.countedCash)) +
          kpi('ส่วนต่าง', (closed.diff >= 0 ? '+' : '−') + '฿' + baht(Math.abs(closed.diff)),
              closed.diff === 0 ? 'ตรงพอดี' : (closed.diff > 0 ? 'เงินเกิน' : 'เงินขาด'),
              closed.diff === 0 ? '' : 'is-cost') +
        '</div>' +
        '<p class="hint">ปิดเมื่อ ' + fmtDateTime(closed.ts) + (closed.by ? ' โดย ' + esc(closed.by) : '') +
        (closed.note ? '<br>หมายเหตุ: ' + esc(closed.note) : '') + '</p>' +
        '<div class="btn-row">' +
          '<button class="btn btn-primary" id="printCloseBtn">พิมพ์ใบปิดยอด</button>' +
          '<button class="btn btn-danger-ghost" id="undoCloseBtn">ยกเลิกการปิดยอด</button>' +
        '</div>' +
      '</div>';
  } else {
    var float0 = state.closeFloat != null ? state.closeFloat : num(db.shop.openingFloat);
    var expected = r2(float0 + sum.cashSales);
    var counted = state.countedCash;
    var diff = counted == null ? null : r2(counted - expected);
    html +=
      '<div class="card">' +
        '<h3 class="section-title">กระทบยอดเงินสดในลิ้นชัก</h3>' +
        '<div class="form-row">' +
          '<label class="field"><span>เงินทอนตั้งต้นตอนเปิดร้าน</span><input class="input" id="closeFloat" type="number" min="0" step="1" value="' + float0 + '"></label>' +
          '<label class="field"><span>นับเงินในลิ้นชักได้จริง</span><input class="input" id="countedCash" type="number" min="0" step="1" value="' + (counted == null ? '' : counted) + '" placeholder="กรอกยอดที่นับได้"></label>' +
        '</div>' +
        '<dl class="totals">' +
          '<dt>เงินทอนตั้งต้น</dt><dd>฿' + baht(float0) + '</dd>' +
          '<dt>+ ยอดขายที่รับเป็นเงินสด</dt><dd>฿' + baht(sum.cashSales) + '</dd>' +
          '<dt class="row-strong">เงินสดที่ควรมีในลิ้นชัก</dt><dd class="row-strong">฿' + baht(expected) + '</dd>' +
          (diff == null ? ''
            : '<dt class="row-strong">ส่วนต่าง (นับได้ − ที่ควรมี)</dt>' +
              '<dd class="row-strong ' + (diff === 0 ? 'row-profit' : 'row-minus') + '">' +
              (diff >= 0 ? '+' : '−') + '฿' + baht(Math.abs(diff)) +
              (diff === 0 ? ' ตรงพอดี' : (diff > 0 ? ' เงินเกิน' : ' เงินขาด')) + '</dd>') +
        '</dl>' +
        '<p class="hint">ยอดโอน/บัตรหน้าร้าน ฿' + baht(sum.directNonCash) + ' ไม่นับรวมในลิ้นชัก</p>' +
        '<div class="form-row" style="margin-top:12px">' +
          '<label class="field"><span>ผู้ปิดยอด</span><input class="input" id="closeBy" value="' + esc(state.closeBy || '') + '" placeholder="ชื่อคนปิดร้าน"></label>' +
          '<label class="field"><span>หมายเหตุ</span><input class="input" id="closeNote" placeholder="ไม่บังคับ"></label>' +
        '</div>' +
        '<button class="btn btn-primary btn-lg" id="doCloseBtn"' + (isFuture ? ' disabled' : '') + '>' +
          (isFuture ? 'ยังปิดยอดของวันในอนาคตไม่ได้' : 'ปิดยอดวันที่ ' + fmtDayFull(key)) + '</button>' +
        '<p class="hint">เมื่อปิดยอดแล้ว บิลของวันนี้จะถูกล็อกไม่ให้ลบหรือแก้ไข</p>' +
      '</div>';
  }

  $('#closeBody').innerHTML = html;

  if (closed) {
    $('#printCloseBtn').addEventListener('click', function () { printClosing(closed); });
    $('#undoCloseBtn').addEventListener('click', function () {
      confirmDialog('ยกเลิกการปิดยอด',
        'ปลดล็อกบิลของวันที่ ' + fmtDayFull(key) + ' เพื่อแก้ไข? ใบปิดยอด ' + esc(closed.no) + ' จะถูกลบ',
        function () {
          db.sales.forEach(function (s) { if (s.closingId === closed.id) delete s.closingId; });
          db.closings = db.closings.filter(function (c) { return c.id !== closed.id; });
          save(); render(); toast('ยกเลิกการปิดยอดแล้ว');
        }, 'ยกเลิกการปิดยอด');
    });
  } else {
    $('#closeFloat').addEventListener('input', function () { state.closeFloat = num(this.value); renderClose(); });
    $('#countedCash').addEventListener('input', function () {
      state.countedCash = this.value === '' ? null : num(this.value);
      renderClose();
      var el = $('#countedCash');     /* ช่องตัวเลขไม่รองรับการกำหนดตำแหน่งเคอร์เซอร์ — โฟกัสอย่างเดียวพอ */
      if (el) el.focus();
    });
    $('#closeBy').addEventListener('input', function () { state.closeBy = this.value; });
    $('#doCloseBtn').addEventListener('click', function () { doClose(key); });
  }

  renderCloseHistory();
}

function doClose(key) {
  var sum = buildDaySummary(db, key);
  var float0 = state.closeFloat != null ? state.closeFloat : num(db.shop.openingFloat);
  var expected = r2(float0 + sum.cashSales);
  if (state.countedCash == null) {
    toast('กรอกยอดเงินสดที่นับได้ก่อนปิดยอด', true);
    var el = $('#countedCash'); if (el) el.focus();
    return;
  }
  var counted = r2(state.countedCash);
  var diff = r2(counted - expected);

  var finish = function () {
    db.closeSeq = (db.closeSeq || 0) + 1;
    var rec = {
      id: uid('z'), no: 'Z' + String(db.closeSeq).padStart(4, '0'),
      dayKey: key, ts: Date.now(),
      by: (state.closeBy || '').trim(),
      note: ($('#closeNote') ? $('#closeNote').value : '').trim(),
      openingFloat: r2(float0), expectedCash: expected, countedCash: counted, diff: diff,
      summary: sum
    };
    db.closings.push(rec);
    salesOfDay(db, key).forEach(function (s) { s.closingId = rec.id; });
    db.shop.openingFloat = r2(float0);
    state.countedCash = null;
    state.closeFloat = null;
    save(); render();
    toast('ปิดยอด ' + rec.no + ' เรียบร้อย · กำไรวันนี้ ฿' + baht(sum.profit));
    openModal('ปิดยอดเรียบร้อย — ' + rec.no, closingHtml(rec), [
      { label: 'ปิด', onClick: closeModal },
      { label: 'พิมพ์ใบปิดยอด', cls: 'btn-primary', onClick: function () { printClosing(rec); } }
    ]);
  };

  if (diff !== 0) {
    confirmDialog('เงินสดไม่ตรง',
      'เงินสดที่ควรมี ฿' + baht(expected) + ' แต่นับได้ ฿' + baht(counted) +
      '<br>ส่วนต่าง <strong>' + (diff > 0 ? 'เกิน' : 'ขาด') + ' ฿' + baht(Math.abs(diff)) + '</strong>' +
      '<br><br>ยืนยันปิดยอดโดยบันทึกส่วนต่างนี้ไว้?',
      finish, 'ยืนยันปิดยอด');
  } else {
    finish();
  }
}

function renderCloseHistory() {
  var list = (db.closings || []).slice().sort(function (a, b) { return b.dayKey < a.dayKey ? -1 : 1; });
  $('#closeHistory').innerHTML = list.length
    ? '<table><thead><tr><th>เลขที่</th><th>วันที่</th><th class="num">บิล</th><th class="num">ยอดขาย</th>' +
      '<th class="num">VAT</th><th class="num">ต้นทุน</th><th class="num">กำไรจริง</th>' +
      '<th class="num">เงินสดควรมี</th><th class="num">นับได้</th><th class="num">ส่วนต่าง</th><th>ผู้ปิด</th><th></th></tr></thead><tbody>' +
      list.map(function (c) {
        return '<tr><td><strong>' + esc(c.no) + '</strong></td>' +
          '<td>' + fmtDayFull(c.dayKey) + '</td>' +
          '<td class="num">' + c.summary.bills + '</td>' +
          '<td class="num">฿' + baht(c.summary.gross) + '</td>' +
          '<td class="num">' + (c.summary.vat ? '฿' + baht(c.summary.vat) : '–') + '</td>' +
          '<td class="num">฿' + baht(c.summary.cost) + '</td>' +
          '<td class="num ' + (num(c.summary.netProfit, c.summary.profit) >= 0 ? 'pos' : 'neg') + '">' +
            signedBaht(num(c.summary.netProfit, c.summary.profit)) + '</td>' +
          '<td class="num">฿' + baht(c.expectedCash) + '</td>' +
          '<td class="num">฿' + baht(c.countedCash) + '</td>' +
          '<td class="num ' + (c.diff === 0 ? '' : 'neg') + '">' + (c.diff === 0 ? 'ตรง' : (c.diff > 0 ? '+' : '−') + '฿' + baht(Math.abs(c.diff))) + '</td>' +
          '<td>' + esc(c.by || '–') + '</td>' +
          '<td><button class="btn btn-sm" data-openz="' + esc(c.dayKey) + '">ดู</button></td></tr>';
      }).join('') + '</tbody></table>'
    : '<div class="empty">ยังไม่เคยปิดยอด</div>';

  $$('[data-openz]').forEach(function (b) {
    b.addEventListener('click', function () { state.closeDay = b.getAttribute('data-openz'); renderClose(); window.scrollTo(0, 0); });
  });
}

function closingHtml(c) {
  var s = c.summary;
  return '<div class="receipt">' +
    '<div class="r-shop">' + esc(db.shop.name) + '</div>' +
    '<div class="r-meta">ใบปิดยอดรายวัน ' + esc(c.no) + '<br>' +
      'ประจำวันที่ ' + fmtDayFull(c.dayKey) + '<br>' +
      'ปิดเมื่อ ' + fmtDateTime(c.ts) + (c.by ? ' โดย ' + esc(c.by) : '') + '</div>' +
    '<div class="r-total">' +
      '<span>จำนวนบิล</span><span class="num">' + s.bills + '</span>' +
      '<span>จำนวนแก้ว</span><span class="num">' + s.cups + '</span>' +
      '<span>ยอดขายรวม</span><span class="num">฿' + baht(s.gross) + '</span>' +
      (s.discount ? '<span>ส่วนลด</span><span class="num">−฿' + baht(s.discount) + '</span>' : '') +
      (s.vat ? '<span>มูลค่าก่อน VAT</span><span class="num">฿' + baht(s.exVat) + '</span>' +
               '<span>VAT</span><span class="num">฿' + baht(s.vat) + '</span>' : '') +
      '<span>ค่า GP</span><span class="num">−฿' + baht(s.gp) + '</span>' +
      '<span>ต้นทุนวัตถุดิบ</span><span class="num">−฿' + baht(s.cost) + '</span>' +
      '<span class="strong">กำไรขั้นต้น</span><span class="num strong">฿' + baht(s.profit) + '</span>' +
      (num(s.fixedCost) > 0
        ? '<span>ค่าใช้จ่ายคงที่/วัน</span><span class="num">−฿' + baht(s.fixedCost) + '</span>' +
          '<span class="strong">กำไรจริงของวัน</span><span class="num strong">' + signedBaht(s.netProfit) + '</span>'
        : '') +
    '</div>' +
    '<table style="margin-top:10px"><tbody>' +
      s.byPay.map(function (v) {
        return '<tr><td>' + esc(v.name) + ' (' + v.bills + ' บิล)</td><td class="num">฿' + baht(v.amount) + '</td></tr>';
      }).join('') +
    '</tbody></table>' +
    '<div class="r-total">' +
      '<span>เงินทอนตั้งต้น</span><span class="num">฿' + baht(c.openingFloat) + '</span>' +
      '<span>ยอดขายเงินสด</span><span class="num">฿' + baht(s.cashSales) + '</span>' +
      '<span class="strong">เงินสดที่ควรมี</span><span class="num strong">฿' + baht(c.expectedCash) + '</span>' +
      '<span>นับได้จริง</span><span class="num">฿' + baht(c.countedCash) + '</span>' +
      '<span class="strong">ส่วนต่าง</span><span class="num strong">' +
        (c.diff === 0 ? 'ตรงพอดี' : (c.diff > 0 ? 'เกิน ฿' : 'ขาด ฿') + baht(Math.abs(c.diff))) + '</span>' +
      '<span>ยอดรอรับจากแพลตฟอร์ม</span><span class="num">฿' + baht(s.platformReceivable) + '</span>' +
      '<span>มูลค่าสต็อกคงเหลือ</span><span class="num">฿' + baht(s.stockWorth) + '</span>' +
    '</div>' +
    (c.note ? '<div class="r-thanks">หมายเหตุ: ' + esc(c.note) + '</div>' : '') +
    '<div class="r-thanks">ลงชื่อ ............................................</div></div>';
}

function printClosing(c) {
  $('#printArea').innerHTML = closingHtml(c);
  window.print();
}

/* ============================================================
   VIEW: สรุป
   ============================================================ */
var PERIODS = [
  { id: 'today', label: 'วันนี้' },
  { id: '7d', label: '7 วัน' },
  { id: 'month', label: 'เดือนนี้' },
  { id: 'all', label: 'ทั้งหมด' }
];

function periodSales() {
  var now = new Date();
  var start;
  if (state.period === 'today') start = new Date(now.getFullYear(), now.getMonth(), now.getDate()).getTime();
  else if (state.period === '7d') start = new Date(now.getFullYear(), now.getMonth(), now.getDate() - 6).getTime();
  else if (state.period === 'month') start = new Date(now.getFullYear(), now.getMonth(), 1).getTime();
  else start = -Infinity;
  return db.sales.filter(function (s) { return s.ts >= start; });
}

function renderReport() {
  $('#periodChips').innerHTML = PERIODS.map(function (p) {
    return '<button class="chip' + (p.id === state.period ? ' is-active' : '') + '" data-period="' + p.id + '">' + p.label + '</button>';
  }).join('');
  $$('#periodChips .chip').forEach(function (el) {
    el.addEventListener('click', function () { state.period = el.getAttribute('data-period'); renderReport(); });
  });

  var rows = periodSales();
  var s = summarize(rows);
  var stockWorth = db.ingredients.reduce(function (sum, i) { return sum + stockValue(db, i.id); }, 0);

  var days = daysInRange(rows, state.period);
  var fixedTotal = r2(dailyFixedCost(db) * days);
  var trueProfit = r2(s.profit - fixedTotal);

  $('#reportKpi').innerHTML =
    kpi('ยอดขายรวม', '฿' + baht(s.gross), s.bills + ' บิล · ' + s.cups + ' แก้ว') +
    (s.vat ? kpi('VAT ที่ต้องนำส่ง', '฿' + baht(s.vat), '', 'is-cost') : '') +
    kpi('ค่า GP ที่ถูกหัก', '฿' + baht(s.gp), '', 'is-cost') +
    kpi('ต้นทุนวัตถุดิบ', '฿' + baht(s.cost), s.cups ? '฿' + baht(s.cost / s.cups) + ' /แก้ว' : '', 'is-cost') +
    kpi('กำไรขั้นต้น', '฿' + baht(s.profit), s.gross ? qtyStr(s.profit / s.gross * 100) + '% ของยอดขาย' : '', 'is-profit') +
    (hasFixedCost(db)
      ? kpi('ค่าใช้จ่ายคงที่', '฿' + baht(fixedTotal), days + ' วัน × ฿' + baht(dailyFixedCost(db)), 'is-cost') +
        kpi('กำไรจริงหลังหักทุกอย่าง', signedBaht(trueProfit),
            trueProfit >= 0 ? 'คุ้มทุนแล้ว' : 'ยังขาดทุนอยู่', trueProfit >= 0 ? 'is-profit' : 'is-cost')
      : '') +
    kpi('มูลค่าสต็อกคงเหลือ', '฿' + baht(stockWorth), 'ตีราคาตามล็อตจริง');

  /* จุดคุ้มทุน */
  var be = breakEven(db, rows.length ? rows : db.sales);
  if (hasFixedCost(db)) {
    var avgCups = r2(s.cups / days);
    $('#breakEvenCard').classList.remove('is-hidden');
    $('#breakEvenBody').innerHTML =
      '<div class="kpi-row">' +
        kpi('ต้องขายวันละ', be.cups + ' แก้ว', '≈ ฿' + baht(be.sales) + ' /วัน') +
        kpi('ตอนนี้เฉลี่ยวันละ', qtyStr(avgCups) + ' แก้ว',
            avgCups >= be.cups ? 'เกินจุดคุ้มทุน' : 'ขาดอีก ' + qtyStr(r2(be.cups - avgCups)) + ' แก้ว/วัน',
            avgCups >= be.cups ? 'is-profit' : 'is-cost') +
        kpi('กำไรเฉลี่ยต่อแก้ว', '฿' + baht(be.perCup), esc(be.basis)) +
        kpi('ค่าใช้จ่ายคงที่/วัน', '฿' + baht(dailyFixedCost(db)), 'เปิดร้านเดือนละ ' + openDays(db) + ' วัน', 'is-cost') +
      '</div>';
  } else {
    $('#breakEvenCard').classList.add('is-hidden');
  }

  /* กราฟกำไรรายวัน */
  var byDay = {};
  rows.forEach(function (x) { byDay[dayKey(x.ts)] = (byDay[dayKey(x.ts)] || 0) + x.profit; });
  var days = Object.keys(byDay).sort();
  if (!days.length) {
    $('#profitChart').innerHTML = '<div class="empty" style="margin:auto">ยังไม่มีข้อมูลในช่วงนี้</div>';
  } else {
    var fixedDay = dailyFixedCost(db);
    var maxAbs = Math.max.apply(null, days.map(function (k) { return Math.abs(byDay[k]); }).concat([fixedDay])) || 1;
    var bars = days.map(function (k) {
      var v = byDay[k];
      var h = Math.max(3, Math.abs(v) / maxAbs * 130);
      var over = fixedDay > 0 && v >= fixedDay;
      return '<div class="bar-col">' +
        '<span class="bar-val">฿' + baht(v) + '</span>' +
        '<div class="bar' + (v < 0 ? ' is-neg' : (over ? ' is-over' : '')) + '" style="height:' + h + 'px"></div>' +
        '<span class="bar-label">' + fmtDayShort(k) + '</span></div>';
    }).join('');
    /* เส้นอ้างอิงจุดคุ้มทุน — วันไหนแท่งสูงเกินเส้นนี้คือวันที่คุ้มทุน */
    var refLine = fixedDay > 0
      ? '<div class="chart-ref" style="bottom:' + (19 + fixedDay / maxAbs * 130) + 'px">' +
        '<span>คุ้มทุน ฿' + baht(fixedDay) + '</span></div>'
      : '';
    $('#profitChart').innerHTML = '<div class="chart-inner">' + bars + '</div>' + refLine;
  }

  /* ตามช่องทาง */
  var byPlat = {};
  rows.forEach(function (x) {
    var k = x.platformName;
    if (!byPlat[k]) byPlat[k] = { bills: 0, cups: 0, gross: 0, gp: 0, profit: 0 };
    byPlat[k].bills++; byPlat[k].cups += x.cups; byPlat[k].gross += x.gross;
    byPlat[k].gp += x.gpAmount; byPlat[k].profit += x.profit;
  });
  var platKeys = Object.keys(byPlat);
  $('#platformTable').innerHTML = platKeys.length ? '<table><thead><tr>' +
    '<th>ช่องทาง</th><th class="num">บิล</th><th class="num">แก้ว</th><th class="num">ยอดขาย</th><th class="num">ค่า GP</th><th class="num">กำไร</th>' +
    '</tr></thead><tbody>' + platKeys.map(function (k) {
      var v = byPlat[k];
      return '<tr><td>' + esc(k) + '</td><td class="num">' + v.bills + '</td><td class="num">' + v.cups + '</td>' +
        '<td class="num">฿' + baht(v.gross) + '</td><td class="num neg">฿' + baht(v.gp) + '</td>' +
        '<td class="num pos">฿' + baht(v.profit) + '</td></tr>';
    }).join('') + '</tbody></table>' : '<div class="empty">ยังไม่มีข้อมูล</div>';

  /* เมนูขายดี */
  var byMenu = {};
  rows.forEach(function (x) {
    x.items.forEach(function (i) {
      if (!byMenu[i.name]) byMenu[i.name] = { qty: 0, gross: 0, cost: 0 };
      byMenu[i.name].qty += i.qty;
      byMenu[i.name].gross += i.lineTotal;
      byMenu[i.name].cost += i.lineCost;
    });
  });
  var menuKeys = Object.keys(byMenu).sort(function (a, b) { return byMenu[b].qty - byMenu[a].qty; });
  $('#topMenuTable').innerHTML = menuKeys.length ? '<table><thead><tr>' +
    '<th>เมนู</th><th class="num">ขายได้ (แก้ว)</th><th class="num">ยอดขาย</th><th class="num">ต้นทุน</th><th class="num">กำไรขั้นต้น</th>' +
    '</tr></thead><tbody>' + menuKeys.map(function (k) {
      var v = byMenu[k];
      return '<tr><td>' + esc(k) + '</td><td class="num">' + qtyStr(v.qty) + '</td>' +
        '<td class="num">฿' + baht(v.gross) + '</td><td class="num">฿' + baht(v.cost) + '</td>' +
        '<td class="num pos">฿' + baht(v.gross - v.cost) + '</td></tr>';
    }).join('') + '</tbody></table>' : '<div class="empty">ยังไม่มีข้อมูล</div>';
}

/* ============================================================
   VIEW: ตั้งค่า
   ============================================================ */
function renderSettings() {
  $('#setShopName').value = db.shop.name || '';
  $('#setShopTag').value = db.shop.tagline || '';
  $('#setTaxId').value = db.shop.taxId || '';
  $('#setFloat').value = num(db.shop.openingFloat);

  var vat = db.shop.vat || { enabled: false, rate: 7, mode: 'inclusive' };
  $('#vatEnabled').checked = !!vat.enabled;
  $('#vatRate').value = num(vat.rate, 7);
  $('#vatMode').value = vat.mode || 'inclusive';
  $('#vatRate').disabled = !vat.enabled;
  $('#vatMode').disabled = !vat.enabled;
  $('#vatExample').innerHTML = vatExampleText(vat);

  $('#stockMethod').value = stockMethod(db);
  $('#expiryWarn').value = num(db.shop.expiryWarnDays, 14);
  $('#openDays').value = openDays(db);
  renderFixedCosts();

  $('#platformSettings').innerHTML = '<table><thead><tr><th>ชื่อช่องทาง</th><th class="num">%GP</th><th>การรับเงิน</th><th></th></tr></thead><tbody>' +
    db.platforms.map(function (p, i) {
      return '<tr>' +
        '<td><input class="input input-sm" data-pname="' + i + '" value="' + esc(p.name) + '"></td>' +
        '<td class="num"><input class="input input-sm" type="number" min="0" max="100" step="0.5" data-pgp="' + i + '" value="' + p.gp + '" style="max-width:90px"></td>' +
        '<td><select class="input input-sm" data-pset="' + i + '">' +
          '<option value="direct"' + (isDirect(p) ? ' selected' : '') + '>ร้านรับเงินเอง</option>' +
          '<option value="platform"' + (!isDirect(p) ? ' selected' : '') + '>แพลตฟอร์มโอนให้ภายหลัง</option>' +
        '</select></td>' +
        '<td>' + (db.platforms.length > 1 ? '<button class="btn btn-sm btn-danger-ghost" data-pdel="' + i + '">ลบ</button>' : '') + '</td>' +
        '</tr>';
    }).join('') + '</tbody></table>';

  $$('[data-pset]').forEach(function (el) {
    el.addEventListener('change', function () {
      db.platforms[Number(el.getAttribute('data-pset'))].settlement = el.value;
      save(); render(); toast('บันทึกแล้ว');
    });
  });

  $$('[data-pname]').forEach(function (el) {
    el.addEventListener('change', function () {
      db.platforms[Number(el.getAttribute('data-pname'))].name = el.value;
      save(); render(); toast('บันทึกแล้ว');
    });
  });
  $$('[data-pgp]').forEach(function (el) {
    el.addEventListener('change', function () {
      db.platforms[Number(el.getAttribute('data-pgp'))].gp = Math.min(100, Math.max(0, num(el.value)));
      save(); render(); toast('บันทึก %GP แล้ว');
    });
  });
  $$('[data-pdel]').forEach(function (el) {
    el.addEventListener('click', function () {
      var i = Number(el.getAttribute('data-pdel'));
      var removed = db.platforms[i];
      db.platforms.splice(i, 1);
      if (state.platformId === removed.id) state.platformId = db.platforms[0].id;
      save(); render(); toast('ลบช่องทางแล้ว');
    });
  });
}

function renderFixedCosts() {
  var list = db.fixedCosts || [];
  $('#fixedCostTable').innerHTML = '<table><thead><tr>' +
    '<th>รายการ</th><th class="num">จำนวนเงิน</th><th>ต่อ</th><th class="num">เฉลี่ย/วัน</th><th></th>' +
    '</tr></thead><tbody>' + list.map(function (c, i) {
      var perDay = c.period === 'day' ? num(c.amount) : num(c.amount) / openDays(db);
      return '<tr>' +
        '<td><input class="input input-sm" data-fcname="' + i + '" value="' + esc(c.name) + '"></td>' +
        '<td class="num"><input class="input input-sm" type="number" min="0" step="10" data-fcamt="' + i + '" value="' + num(c.amount) + '" style="max-width:110px"></td>' +
        '<td><select class="input input-sm" data-fcper="' + i + '">' +
          '<option value="day"' + (c.period === 'day' ? ' selected' : '') + '>วัน</option>' +
          '<option value="month"' + (c.period !== 'day' ? ' selected' : '') + '>เดือน</option>' +
        '</select></td>' +
        '<td class="num">฿' + baht(perDay) + '</td>' +
        '<td><button class="btn btn-sm btn-danger-ghost" data-fcdel="' + i + '">ลบ</button></td>' +
        '</tr>';
    }).join('') + '</tbody>' +
    '<tfoot><tr><td colspan="3">รวมค่าใช้จ่ายคงที่เฉลี่ยต่อวัน</td>' +
    '<td class="num">฿' + baht(dailyFixedCost(db)) + '</td><td></td></tr></tfoot></table>';

  var be = breakEven(db, db.sales);
  $('#fixedSummary').innerHTML = be.ready
    ? 'ต้องขายวันละประมาณ <strong>' + be.cups + ' แก้ว</strong> (≈ ฿' + baht(be.sales) + ') ถึงจะคุ้มค่าใช้จ่ายคงที่ ' +
      '<br>คิดจากกำไรเฉลี่ย ฿' + baht(be.perCup) + ' ต่อแก้ว — ' + esc(be.basis)
    : 'ยังไม่ได้กรอกค่าใช้จ่ายคงที่ — กรอกแล้วระบบจะคำนวณจุดคุ้มทุนให้อัตโนมัติ';

  $$('[data-fcname]').forEach(function (el) {
    el.addEventListener('change', function () {
      db.fixedCosts[Number(el.getAttribute('data-fcname'))].name = el.value.trim() || 'ค่าใช้จ่าย';
      save(); render();
    });
  });
  $$('[data-fcamt]').forEach(function (el) {
    el.addEventListener('change', function () {
      db.fixedCosts[Number(el.getAttribute('data-fcamt'))].amount = num(el.value);
      save(); render(); toast('บันทึกค่าใช้จ่ายแล้ว');
    });
  });
  $$('[data-fcper]').forEach(function (el) {
    el.addEventListener('change', function () {
      db.fixedCosts[Number(el.getAttribute('data-fcper'))].period = el.value;
      save(); render();
    });
  });
  $$('[data-fcdel]').forEach(function (el) {
    el.addEventListener('click', function () {
      db.fixedCosts.splice(Number(el.getAttribute('data-fcdel')), 1);
      save(); render(); toast('ลบรายการแล้ว');
    });
  });
}

/** ข้อความตัวอย่างใต้การตั้งค่า VAT ให้เห็นภาพว่าตัวเลขจะเปลี่ยนยังไง */
function vatExampleText(vat) {
  var price = 339;
  if (!vat.enabled) return 'ปิดอยู่ — ราคา ฿339 คือยอดที่ลูกค้าจ่าย และไม่มีการแยกภาษี';
  var rate = num(vat.rate, 7) / 100;
  if (vat.mode === 'exclusive') {
    var v = r2(price * rate);
    return 'ราคา ฿339 + VAT ฿' + baht(v) + ' → ลูกค้าจ่าย <strong>฿' + baht(price + v) + '</strong> (ฐานภาษี ฿' + baht(price) + ')';
  }
  var vi = r2(price * rate / (1 + rate));
  return 'ราคา ฿339 รวม VAT แล้ว → ฐานภาษี ฿' + baht(price - vi) + ' + VAT ฿' + baht(vi) +
    ' → ลูกค้าจ่าย <strong>฿' + baht(price) + '</strong> (กำไรคิดจากยอดก่อน VAT)';
}

/* ============================================================
   render router
   ============================================================ */
function render() {
  $('#shopName').textContent = db.shop.name || 'POS';
  $('#shopTag').textContent = (db.shop.tagline || '') + ' · POS';
  document.title = (db.shop.name || 'POS') + ' — ระบบ POS';

  $$('#tabs .tab').forEach(function (t) { t.classList.toggle('is-active', t.getAttribute('data-view') === state.view); });
  $$('.view').forEach(function (v) { v.classList.toggle('is-hidden', v.id !== 'view-' + state.view); });

  if (state.view === 'pos') renderPOS();
  else if (state.view === 'stock') renderStock();
  else if (state.view === 'menu') renderMenus();
  else if (state.view === 'history') renderHistory();
  else if (state.view === 'close') renderClose();
  else if (state.view === 'report') renderReport();
  else if (state.view === 'settings') renderSettings();
}

/* ============================================================
   init
   ============================================================ */
function init() {
  db = migrate(loadStored()) || buildFresh();
  if (!state.platformId) state.platformId = db.platforms[0].id;
  save();

  /* theme */
  var savedTheme = null;
  try { savedTheme = localStorage.getItem(KEY + '_theme'); } catch (e) {}
  if (savedTheme) document.documentElement.setAttribute('data-theme', savedTheme);
  $('#themeBtn').addEventListener('click', function () {
    var cur = document.documentElement.getAttribute('data-theme');
    var next = cur === 'dark' ? 'light' : 'dark';
    document.documentElement.setAttribute('data-theme', next);
    try { localStorage.setItem(KEY + '_theme', next); } catch (e) {}
  });

  /* clock */
  function tick() {
    $('#clock').textContent = new Date().toLocaleString('th-TH', {
      weekday: 'short', day: 'numeric', month: 'short', hour: '2-digit', minute: '2-digit'
    });
  }
  tick(); setInterval(tick, 20000);

  /* tabs */
  $$('#tabs .tab').forEach(function (t) {
    t.addEventListener('click', function () { state.view = t.getAttribute('data-view'); render(); window.scrollTo(0, 0); });
  });

  /* POS */
  $('#menuSearch').addEventListener('input', function () { state.search = this.value; renderPOS(); });
  $('#discount').addEventListener('input', function () { state.discount = num(this.value); renderCart(); });
  $('#received').addEventListener('input', function () {
    state.received = this.value === '' ? null : num(this.value);
    renderCart();
  });
  $('#clearCart').addEventListener('click', function () {
    state.cart = []; state.discount = 0; state.received = null;
    $('#discount').value = 0; $('#received').value = '';
    renderCart();
  });
  $('#checkoutBtn').addEventListener('click', checkout);

  /* stock */
  $('#receiveBtn').addEventListener('click', function () { receiveDialog(db.ingredients[0] ? db.ingredients[0].id : null); });
  $('#addIngBtn').addEventListener('click', function () { ingredientDialog(null); });

  /* menu */
  $('#addMenuBtn').addEventListener('click', function () { menuDialog(null); });

  /* history */
  $('#fromDate').addEventListener('change', function () { state.histFrom = this.value; renderHistory(); });
  $('#toDate').addEventListener('change', function () { state.histTo = this.value; renderHistory(); });
  $('#historyPlatform').addEventListener('change', function () { state.histPlatform = this.value; renderHistory(); });
  $('#exportSalesBtn').addEventListener('click', exportSalesCsv);

  /* settings */
  $('#saveShopBtn').addEventListener('click', function () {
    db.shop.name = $('#setShopName').value.trim() || 'ร้านของฉัน';
    db.shop.tagline = $('#setShopTag').value.trim();
    db.shop.taxId = $('#setTaxId').value.trim();
    db.shop.openingFloat = num($('#setFloat').value);
    save(); render(); toast('บันทึกข้อมูลร้านแล้ว');
  });

  /* VAT */
  function saveVat() {
    db.shop.vat = {
      enabled: $('#vatEnabled').checked,
      rate: num($('#vatRate').value, 7),
      mode: $('#vatMode').value
    };
    save(); render();
    toast(db.shop.vat.enabled ? 'เปิดคิด VAT ' + qtyStr(db.shop.vat.rate) + '% แล้ว' : 'ปิดการคิด VAT แล้ว');
  }
  $('#vatEnabled').addEventListener('change', saveVat);
  $('#vatRate').addEventListener('change', saveVat);
  $('#vatMode').addEventListener('change', saveVat);

  /* วิธีตัดสต็อก & ค่าใช้จ่ายคงที่ */
  $('#stockMethod').addEventListener('change', function () {
    db.shop.stockMethod = this.value;
    save(); render();
    toast(this.value === 'fefo' ? 'เปลี่ยนเป็น FEFO — ใช้ของที่จะหมดอายุก่อน' : 'เปลี่ยนเป็น FIFO — ใช้ของที่รับเข้าก่อน');
  });
  $('#expiryWarn').addEventListener('change', function () {
    db.shop.expiryWarnDays = num(this.value, 14);
    save(); render();
  });
  $('#openDays').addEventListener('change', function () {
    db.shop.openDaysPerMonth = Math.min(31, Math.max(1, num(this.value, 26)));
    save(); render();
  });
  $('#addFixedBtn').addEventListener('click', function () {
    db.fixedCosts.push({ id: uid('fc'), name: 'ค่าใช้จ่ายใหม่', amount: 0, period: 'month' });
    save(); render();
  });

  /* ปิดยอด */
  $('#closeDate').addEventListener('change', function () {
    state.closeDay = this.value;
    state.countedCash = null;
    state.closeFloat = null;
    renderClose();
  });
  $('#addPlatformBtn').addEventListener('click', function () {
    db.platforms.push({ id: uid('pf'), name: 'ช่องทางใหม่', gp: 0 });
    save(); render();
  });
  $('#backupBtn').addEventListener('click', function () {
    download('matcha-pos-backup-' + dayKey(Date.now()) + '.json', JSON.stringify(db, null, 2), 'application/json');
    toast('สำรองข้อมูลแล้ว');
  });
  $('#restoreBtn').addEventListener('click', function () { $('#restoreFile').click(); });
  $('#restoreFile').addEventListener('change', function () {
    var f = this.files[0];
    if (!f) return;
    var reader = new FileReader();
    reader.onload = function () {
      try {
        var parsed = migrate(JSON.parse(reader.result));
        if (!parsed) throw new Error('bad');
        db = parsed; state.cart = []; state.platformId = db.platforms[0].id;
        save(); render(); toast('กู้คืนข้อมูลเรียบร้อย');
      } catch (e) { toast('ไฟล์ไม่ถูกต้อง', true); }
    };
    reader.readAsText(f);
    this.value = '';
  });
  $('#resetSalesBtn').addEventListener('click', function () {
    confirmDialog('ล้างประวัติการขาย', 'ลบบิลทั้งหมดและคืนวัตถุดิบกลับเข้าสต็อก? (สูตรและสต็อกที่รับเข้ายังอยู่)', function () {
      db.moves.forEach(function (m) { if (m.type === 'sale') releaseAllocs(db, m.allocs); });
      db.moves = db.moves.filter(function (m) { return m.type !== 'sale'; });
      db.sales = []; db.billSeq = 0;
      db.closings = []; db.closeSeq = 0;
      save(); render(); toast('ล้างประวัติการขายแล้ว');
    }, 'ล้างข้อมูล');
  });
  $('#resetAllBtn').addEventListener('click', function () {
    confirmDialog('รีเซ็ตทั้งหมด', 'ลบข้อมูลทั้งหมดและกลับไปใช้ค่าเริ่มต้นจากไฟล์ Excel?', function () {
      db = buildFresh(); state.cart = []; state.platformId = db.platforms[0].id;
      save(); render(); toast('รีเซ็ตเรียบร้อย');
    }, 'รีเซ็ต');
  });

  /* modal close */
  $('#modalRoot').addEventListener('click', function (e) {
    if (e.target.getAttribute && e.target.getAttribute('data-close')) closeModal();
  });
  document.addEventListener('keydown', function (e) {
    if (e.key === 'Escape') closeModal();
  });

  render();
}

document.addEventListener('DOMContentLoaded', init);
})();
