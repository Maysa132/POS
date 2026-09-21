/* ------------------------------------------------------------------
   data.js — ข้อมูลตั้งต้น (แปลงมาจาก ร้านมัทฉะในฝัน.xlsx)
   ชีต: Matcha_stock, Other_stock, Recipe_Data, Sale
------------------------------------------------------------------- */

window.SEED = {
  version: 1,
  shop: {
    name: 'ร้านมัทฉะในฝัน',
    tagline: 'Matcha Specialty',
    taxId: '',
    /* VAT: mode 'inclusive' = ราคาที่ตั้งไว้รวม VAT แล้ว, 'exclusive' = บวก VAT เพิ่มจากราคา */
    vat: { enabled: true, rate: 7, mode: 'inclusive' },
    openingFloat: 1000,
    /* 'fifo' = ใช้ของที่รับเข้าก่อน, 'fefo' = ใช้ของที่จะหมดอายุก่อน */
    stockMethod: 'fefo',
    expiryWarnDays: 14,
    openDaysPerMonth: 26
  },

  /* ค่าใช้จ่ายคงที่ — ใส่ตัวเลขจริงของร้านที่แท็บตั้งค่า แล้วระบบจะคำนวณจุดคุ้มทุนให้
     ตั้งค่าเริ่มต้นไว้เป็น 0 ทั้งหมด เพื่อไม่ให้ตัวเลขกำไรเพี้ยนก่อนกรอกจริง */
  fixedCosts: [
    { id: 'rent',     name: 'ค่าเช่าที่',         amount: 0, period: 'month' },
    { id: 'wage',     name: 'ค่าแรงพนักงาน',      amount: 0, period: 'day'   },
    { id: 'utility',  name: 'ค่าไฟ / ค่าน้ำ',      amount: 0, period: 'month' },
    { id: 'internet', name: 'ค่าเน็ต / ค่าโทรศัพท์', amount: 0, period: 'month' },
    { id: 'misc',     name: 'อื่น ๆ',              amount: 0, period: 'month' }
  ],

  /* ช่องทางขาย + %GP (จากไฟล์มีเฉพาะ Grab 35% ที่เหลือแก้ไขได้ในหน้าตั้งค่า)
     settlement: 'direct' = ร้านรับเงินเอง (เข้าลิ้นชัก/บัญชีทันที), 'platform' = รอแพลตฟอร์มโอน */
  platforms: [
    { id: 'walkin', name: 'หน้าร้าน',   gp: 0,  settlement: 'direct' },
    { id: 'grab',   name: 'Grab',        gp: 35, settlement: 'platform' },
    { id: 'lineman',name: 'LINE MAN',    gp: 32, settlement: 'platform' },
    { id: 'shopee', name: 'ShopeeFood',  gp: 30, settlement: 'platform' }
  ],

  /* วิธีชำระเงินสำหรับบิลหน้าร้าน — ใช้กระทบยอดเงินสดตอนปิดยอด */
  payMethods: [
    { id: 'cash',     name: 'เงินสด',  inDrawer: true },
    { id: 'transfer', name: 'โอน/QR',  inDrawer: false },
    { id: 'card',     name: 'บัตร',    inDrawer: false }
  ],

  /* วัตถุดิบทั้งหมด
     costPerUnit    = purchasePrice / qtyPerPurchase  (ตรงกับคอลัมน์ Price/Gram และ Unit/Cup)
     shelfLifeDays  = อายุหลังรับเข้า ใช้เติมวันหมดอายุให้อัตโนมัติตอนรับของ (0 = ไม่มีวันหมดอายุ) */
  ingredients: [
    // ---- ผงมัทฉะ (Matcha_stock) ----
    { id: 'may', name: 'May', group: 'matcha', purchaseUnit: 'ซอง', purchasePrice: 596, qtyPerPurchase: 40, unit: 'กรัม', openingQty: 40, minQty: 10, shelfLifeDays: 180 },
    { id: 'g',   name: 'G',   group: 'matcha', purchaseUnit: 'ซอง', purchasePrice: 1290, qtyPerPurchase: 30, unit: 'กรัม', openingQty: 30, minQty: 10, shelfLifeDays: 180 },
    { id: 'p',   name: 'P',   group: 'matcha', purchaseUnit: 'ซอง', purchasePrice: 789, qtyPerPurchase: 40, unit: 'กรัม', openingQty: 40, minQty: 10, shelfLifeDays: 180 },
    { id: 'pa',  name: 'Pa',  group: 'matcha', purchaseUnit: 'ซอง', purchasePrice: 550, qtyPerPurchase: 20, unit: 'กรัม', openingQty: 20, minQty: 5,  shelfLifeDays: 180 },
    { id: 't',   name: 'T',   group: 'matcha', purchaseUnit: 'ซอง', purchasePrice: 678, qtyPerPurchase: 50, unit: 'กรัม', openingQty: 50, minQty: 10, shelfLifeDays: 180 },

    // ---- วัตถุดิบ/บรรจุภัณฑ์ (Other_stock) ----
    { id: 'milk',  name: 'Milk (นมสด)', group: 'other', purchaseUnit: 'ขวด',  purchasePrice: 102, qtyPerPurchase: 1000, unit: 'มล',  openingQty: 1000, minQty: 200, shelfLifeDays: 7 },
    { id: 'cup',   name: 'แก้ว',        group: 'other', purchaseUnit: 'แพ็ก', purchasePrice: 100, qtyPerPurchase: 50,   unit: 'ใบ',  openingQty: 50,   minQty: 10,  shelfLifeDays: 0 },
    { id: 'lid',   name: 'ฝา',          group: 'other', purchaseUnit: 'แพ็ก', purchasePrice: 100, qtyPerPurchase: 50,   unit: 'ใบ',  openingQty: 50,   minQty: 10,  shelfLifeDays: 0 },
    { id: 'straw', name: 'หลอด',        group: 'other', purchaseUnit: 'แพ็ก', purchasePrice: 78,  qtyPerPurchase: 50,   unit: 'อัน', openingQty: 50,   minQty: 10,  shelfLifeDays: 0 },
    { id: 'bag',   name: 'ถุง',         group: 'other', purchaseUnit: 'แพ็ก', purchasePrice: 150, qtyPerPurchase: 50,   unit: 'ใบ',  openingQty: 50,   minQty: 10,  shelfLifeDays: 0 },
    { id: 'bagsm', name: 'ถุงเล็ก',     group: 'other', purchaseUnit: 'แพ็ก', purchasePrice: 150, qtyPerPurchase: 50,   unit: 'ใบ',  openingQty: 50,   minQty: 10,  shelfLifeDays: 0 },
    { id: 'paper', name: 'กระดาษรอง',   group: 'other', purchaseUnit: 'แพ็ก', purchasePrice: 40,  qtyPerPurchase: 50,   unit: 'ใบ',  openingQty: 50,   minQty: 10,  shelfLifeDays: 0 },
    { id: 'ice',   name: 'น้ำแข็ง',     group: 'other', purchaseUnit: 'ถุง',  purchasePrice: 7.5, qtyPerPurchase: 10,   unit: 'ก้อน',openingQty: 10,   minQty: 5,   shelfLifeDays: 0 },
    { id: 'oat',   name: 'นมโอ๊ต',      group: 'other', purchaseUnit: 'ขวด',  purchasePrice: 99,  qtyPerPurchase: 500,  unit: 'มล',  openingQty: 500,  minQty: 100, shelfLifeDays: 90 },
    { id: 'syrup', name: 'น้ำเชื่อม',   group: 'other', purchaseUnit: 'ขวด',  purchasePrice: 130, qtyPerPurchase: 100,  unit: 'มล',  openingQty: 100,  minQty: 30,  shelfLifeDays: 180 },
    { id: 'water', name: 'น้ำ',         group: 'other', purchaseUnit: 'ขวด',  purchasePrice: 10,  qtyPerPurchase: 1000, unit: 'มล',  openingQty: 1000, minQty: 200, shelfLifeDays: 365 }
  ],

  /* ตัวเลือกมาตรฐาน — ใส่ให้ทั้งสองเมนู (แก้ไข/เพิ่มได้ที่แท็บเมนู & สูตร)
       type 'scale' = คูณปริมาณวัตถุดิบ target ตาม factor
       type 'swap'  = สลับวัตถุดิบ target ไปเป็นตัวที่เลือก
     ทั้งสองแบบอิงจากวัตถุดิบเดิมในสูตร จึงใช้ร่วมกันได้ (เช่น เปลี่ยนผง + เพิ่มความเข้มพร้อมกัน) */
  optionPresets: {
    sweet: {
      id: 'sweet', name: 'ความหวาน', type: 'scale', target: 'syrup', defaultId: 'sw100',
      choices: [
        { id: 'sw0',   label: 'ไม่หวาน',   factor: 0,    priceDelta: 0 },
        { id: 'sw50',  label: 'หวานน้อย',  factor: 0.5,  priceDelta: 0 },
        { id: 'sw75',  label: 'หวาน 75%',  factor: 0.75, priceDelta: 0 },
        { id: 'sw100', label: 'หวานปกติ',  factor: 1,    priceDelta: 0 },
        { id: 'sw150', label: 'หวานมาก',   factor: 1.5,  priceDelta: 0 }
      ]
    },
    /* ราคาบวกเพิ่มตั้งไว้ให้คร่าว ๆ ตามส่วนต่างต้นทุนผงแต่ละตัว — ปรับได้ตามจริง */
    powder: {
      id: 'powder', name: 'ผงมัทฉะ', type: 'swap', target: 'may', defaultId: 'pw-may', alwaysShow: true,
      choices: [
        { id: 'pw-may', label: 'May',           ingId: 'may', factor: 1, priceDelta: 0 },
        { id: 'pw-t',   label: 'T',             ingId: 't',   factor: 1, priceDelta: 0 },
        { id: 'pw-p',   label: 'P',             ingId: 'p',   factor: 1, priceDelta: 30 },
        { id: 'pw-pa',  label: 'Pa',            ingId: 'pa',  factor: 1, priceDelta: 70 },
        { id: 'pw-g',   label: 'G (พรีเมียม)',  ingId: 'g',   factor: 1, priceDelta: 150 }
      ]
    },
    strength: {
      id: 'strength', name: 'ความเข้ม', type: 'scale', target: 'may', defaultId: 'st100',
      choices: [
        { id: 'st70',  label: 'อ่อน',        factor: 0.7, priceDelta: 0 },
        { id: 'st100', label: 'ปกติ',        factor: 1,   priceDelta: 0 },
        { id: 'st140', label: 'เข้มพิเศษ',   factor: 1.4, priceDelta: 20 }
      ]
    }
  },

  /* เมนู + สูตร (Recipe_Data) — optionKeys อ้างถึง optionPresets ข้างบน */
  menus: [
    {
      id: 'matcha-latte',
      name: 'Matcha Latte',
      price: 339,
      color: '#5c8a3a',
      active: true,
      optionKeys: ['sweet', 'powder', 'strength'],
      recipe: [
        { ingId: 'may',   qty: 5 },
        { ingId: 'milk',  qty: 100 },
        { ingId: 'cup',   qty: 1 },
        { ingId: 'lid',   qty: 1 },
        { ingId: 'straw', qty: 1 },
        { ingId: 'bag',   qty: 1 },
        { ingId: 'bagsm', qty: 1 },
        { ingId: 'paper', qty: 1 },
        { ingId: 'ice',   qty: 1 },
        { ingId: 'syrup', qty: 8 },
        { ingId: 'water', qty: 50 }
      ]
    },
    {
      id: 'matcha-latte-oat',
      name: 'Matcha Latte (Oat)',
      price: 341,
      color: '#7a6a3a',
      active: true,
      optionKeys: ['sweet', 'powder', 'strength'],
      recipe: [
        { ingId: 'may',   qty: 5 },
        { ingId: 'oat',   qty: 100 },
        { ingId: 'cup',   qty: 1 },
        { ingId: 'lid',   qty: 1 },
        { ingId: 'straw', qty: 1 },
        { ingId: 'bag',   qty: 1 },
        { ingId: 'bagsm', qty: 1 },
        { ingId: 'paper', qty: 1 },
        { ingId: 'ice',   qty: 1 },
        { ingId: 'syrup', qty: 8 },
        { ingId: 'water', qty: 50 }
      ]
    }
  ],

  /* บิลตั้งต้นจากชีต Sale — ลบทิ้งได้ที่หน้าตั้งค่า */
  seedSales: [
    { platformId: 'grab',   minutesAgo: 320, items: [{ menuId: 'matcha-latte', qty: 1 }] },
    { platformId: 'grab',   minutesAgo: 180, items: [{ menuId: 'matcha-latte', qty: 2 }] },
    { platformId: 'walkin', minutesAgo: 90,  payMethodId: 'cash', items: [{ menuId: 'matcha-latte-oat', qty: 3 }] }
  ]
};
