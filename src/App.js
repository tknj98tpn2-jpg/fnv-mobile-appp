import React, { useState, useMemo, useRef, useEffect } from 'react';
import * as XLSX from 'xlsx';
import { initializeApp } from 'firebase/app';
import {
  getFirestore, collection, doc,
  onSnapshot, setDoc, updateDoc, deleteDoc, writeBatch, getDocs
} from 'firebase/firestore';
import {
  Menu, X, LayoutDashboard, Tag, Scissors, ClipboardList, ShoppingBag,
  PackageCheck, Truck, Truck as TruckIcon, Boxes, Users, Upload, FileSpreadsheet, AlertCircle,
  Trash2, Pencil, Plus, Sprout, ChevronRight, ArrowLeft, Download, Store,
  Search, Layers, IndianRupee, TrendingUp,
} from 'lucide-react';

// ── Firebase — same project as the web admin panel, so data stays in sync ──
const firebaseConfig = {
  apiKey: "AIzaSyDS-QPS9hiBRqIEyiGMTMIO4lWeSSMcY0M",
  authDomain: "fnv-business-app.firebaseapp.com",
  databaseURL: "https://fnv-business-app-default-rtdb.asia-southeast1.firebasedatabase.app",
  projectId: "fnv-business-app",
  storageBucket: "fnv-business-app.firebasestorage.app",
  messagingSenderId: "16781517968",
  appId: "1:16781517968:web:3ea46523739096335a9a415",
  measurementId: "G-END2323XWW",
};
const fbApp = initializeApp(firebaseConfig);
const db = getFirestore(fbApp);

async function seedIfEmpty(colName, rows) {
  const snap = await getDocs(collection(db, colName));
  if (!snap.empty) return;
  const batch = writeBatch(db);
  rows.forEach((r) => batch.set(doc(db, colName, r.id), r));
  await batch.commit();
}
// ────────────────────────────────────────────────────────

const INK = '#20241E';
const LEAF = '#2F5233';
const LEAF_DARK = '#1B2E1D';
const SIDEBAR = '#16241A';
const AMBER = '#C9861F';
const TOMATO = '#D9552C';
const LINE = '#E3DECF';
const MUTED = '#7A7566';
const BG = '#F6F3EA';

const PLATFORMS = ['Blinkit', 'Flipkart'];

// Each indent can be split across dark stores (Flipkart lists one column per store);
// Blinkit has a single store. Orders carry a `store`; older orders without one fall back
// to the platform's default store, if it has one.
const PLATFORM_DEFAULT_STORE = { Blinkit: 'CPC' };
const orderStore = (o) => o.store || PLATFORM_DEFAULT_STORE[o.platform] || '';
// "Jab_103_6_Sarvodya Nagar" -> "Sarvodya Nagar"
const storeLabel = (s) => (s ? (String(s).replace(/^[A-Za-z]{2,5}_\d+_\d+_/, '').replace(/_/g, ' ').trim() || String(s)) : 'No store');
function storeOptionsFor(orders, platform) {
  const stores = new Set();
  orders.forEach((o) => { if (o.platform === platform) stores.add(orderStore(o)); });
  if (PLATFORM_DEFAULT_STORE[platform]) stores.add(PLATFORM_DEFAULT_STORE[platform]);
  return Array.from(stores)
    .sort((a, b) => (a === '' ? 1 : b === '' ? -1 : a.localeCompare(b)))
    .map((s) => ({ value: s || '__none__', store: s, label: storeLabel(s) }));
}
// When a target's short packs are split across its orders (one order per store), keep
// them in whole packs instead of fractions. Returns null when it can't (non-integer input).
function allocateShortPacks(targetOrders, shortQty) {
  const total = targetOrders.reduce((s, o) => s + (Number(o.packQty) || 0), 0);
  if (!(total > 0) || !Number.isInteger(shortQty) || targetOrders.some((o) => !Number.isInteger(Number(o.packQty) || 0))) return null;
  const alloc = {};
  const rem = [];
  let given = 0;
  targetOrders.forEach((o, i) => {
    const p = Number(o.packQty) || 0;
    const exact = (shortQty * p) / total;
    const base = Math.min(Math.floor(exact), p);
    alloc[o.id] = base;
    given += base;
    rem.push({ id: o.id, frac: exact - Math.floor(exact), room: p - base, i });
  });
  let left = shortQty - given;
  rem.sort((a, b) => (b.frac - a.frac) || (a.i - b.i));
  rem.forEach((r) => { if (left > 0 && r.room > 0) { alloc[r.id] += 1; left -= 1; } });
  return alloc;
}

// Phase 1 of multi-city support: each business location gets its own Items and
// Vendors (Orders, Purchases, Dispatch, P&L follow in later phases). Records made
// before this existed have no `city` field — they're treated as belonging to the
// first city here so nothing already in the database disappears.
const CITIES = ['Jabalpur', 'Satna', 'Indore'];

function newAliasId() {
  return `AL-${Date.now().toString(36).toUpperCase().slice(-6)}-${Math.floor(Math.random() * 900 + 100)}`;
}

const SEED_ITEMS = [];

const SEED_ORDERS = [];

const SEED_PURCHASES = [];

const SEED_RECIPES = [];

const SEED_ROLES = [
  { id: 'ROLE-ADMIN', name: 'Admin', permissions: { dashboard: true, items: true, cutprocess: true, orders: true, purchase: true, stockcount: true, pricing: true, profitloss: true, packaging: true, dispatch: true, crates: true, users: true } },
  { id: 'ROLE-WAREHOUSE', name: 'Warehouse Staff', permissions: { dashboard: true, items: false, cutprocess: false, orders: false, purchase: false, stockcount: true, pricing: false, profitloss: false, packaging: true, dispatch: true, crates: true, users: false } },
];

const SEED_USERS = [];

const SEED_VENDORS = [];

const CATEGORY_MAP = { fruit: 'FRUITS', fruits: 'FRUITS', veg: 'VEGETABLES', vegetable: 'VEGETABLES', vegetables: 'VEGETABLES', 'fresh vegetables': 'VEGETABLES', exotic: 'EXOTIC', exotics: 'EXOTIC', flower: 'FLOWER', flowers: 'FLOWER', flowres: 'FLOWER', grain: 'GRAINS', grains: 'GRAINS', cut: 'CUT' };
function normalizeCategory(raw) { return CATEGORY_MAP[String(raw || '').toLowerCase().trim()] || 'VEGETABLES'; }

const CATEGORY_OPTIONS = ['FRUITS', 'VEGETABLES', 'FLOWER', 'EXOTIC', 'GRAINS', 'CUT'];
const UOM_OPTIONS = ['kg', 'dozen', 'bunch', 'piece', 'pack', 'box', 'crate'];

const NAV = [
  { key: 'dashboard', label: 'Dashboard', icon: LayoutDashboard },
  { key: 'items', label: 'Items', icon: Tag },
  { key: 'vendors', label: 'Vendors', icon: Store },
  { key: 'cutprocess', label: 'Cut & Process', icon: Scissors },
  { key: 'orders', label: 'Orders', icon: ClipboardList },
  { key: 'purchase', label: 'Purchases', icon: ShoppingBag },
  { key: 'stockcount', label: 'Stock Count', icon: Layers },
  { key: 'pricing', label: 'Pricing', icon: IndianRupee },
  { key: 'profitloss', label: 'Profit & Loss', icon: TrendingUp },
  { key: 'packaging', label: 'Packaging', icon: PackageCheck },
  { key: 'dispatch', label: 'Dispatch', icon: Truck },
  { key: 'crates', label: 'Crates & boxes', icon: Boxes },
  { key: 'users', label: 'Users & Roles', icon: Users },
];

// Keeps a filter's value in localStorage so it survives leaving the section (or the
// whole page reloading) — it only ever changes when the person picks something new.
function usePersistedState(key, defaultValue) {
  const [state, setState] = useState(() => {
    try {
      const saved = window.localStorage.getItem(key);
      return saved !== null ? JSON.parse(saved) : defaultValue;
    } catch {
      return defaultValue;
    }
  });
  useEffect(() => {
    try {
      window.localStorage.setItem(key, JSON.stringify(state));
    } catch {}
  }, [key, state]);
  return [state, setState];
}

function pickField(rowObj, candidates) {
  const keys = Object.keys(rowObj);
  for (const c of candidates) {
    const found = keys.find((k) => k.toLowerCase().replace(/[^a-z0-9]/g, '').includes(c));
    if (found && String(rowObj[found]).trim() !== '') return rowObj[found];
  }
  return '';
}
const KNOWN_INDENT_HEADERS = new Set(['fsn', 'title', 'category', 'type', 'umo', 'uom', 'unit', 'mrp', 'price', 't100t500fsn', 'eancode', 'shelflifedays', 'shelflife', 'temperaturezone', 'itemcode', 'articlecode', 'productcode', 'sku', 'code', 'productdescription', 'description', 'article', 'product', 'item', 'indent', 'qty', 'quantity', 'orderedqty']);
function sumUnknownNumericColumns(rowObj, headers) {
  // Returns { "<store column header>": qty } for every leftover numeric column with a
  // positive value, or null when there is none.
  const stores = {};
  headers.forEach((h) => {
    if (!h) return;
    const norm = h.toLowerCase().replace(/[^a-z0-9]/g, '');
    if (KNOWN_INDENT_HEADERS.has(norm)) return;
    const v = rowObj[h];
    if (v === '' || v === null || v === undefined) return;
    const num = Number(v);
    if (!isNaN(num) && num > 0) stores[String(h).trim()] = num;
  });
  return Object.keys(stores).length > 0 ? stores : null;
}

function parseIndentRows(json) {
  return json.map((r, idx) => {
    const headers = Object.keys(r);
    const rawName = String(pickField(r, ['title', 'article', 'product', 'item', 'description']) || '').trim();
    const rawCode = String(pickField(r, ['fsn', 'itemcode', 'articlecode', 'productcode', 'sku', 'code']) || '').trim();
    let qty = Number(pickField(r, ['indent', 'qty', 'quantity', 'orderedqty']) || 0);
    let storeQtys = null;
    if (!qty) { const st = sumUnknownNumericColumns(r, headers); if (st) { storeQtys = st; qty = Object.values(st).reduce((s, v) => s + v, 0); } }
    const unit = String(pickField(r, ['umo', 'uom', 'unit']) || '').trim();
    const rawCategory = String(pickField(r, ['type', 'category']) || '').trim();
    return { key: `row-${idx}-${rawName}`, rawName, rawCode, qty, unit, rawCategory, storeQtys };
  }).filter((r) => r.rawName && r.qty > 0);
}

// ---------- shared small UI ----------
function Card({ children, style }) {
  return <div style={{ background: '#fff', border: `1px solid ${LINE}`, borderRadius: 14, padding: 14, ...style }}>{children}</div>;
}
function Field(props) {
  return <input {...props} style={{ width: '100%', boxSizing: 'border-box', padding: '9px 10px', borderRadius: 8, border: `1px solid ${LINE}`, fontSize: 13, marginBottom: 8, ...(props.style || {}) }} />;
}
function PrimaryBtn({ children, onClick, disabled, color }) {
  return (
    <button onClick={onClick} disabled={disabled} style={{ width: '100%', background: disabled ? '#C9C2AE' : (color || LEAF), color: '#fff', border: 'none', borderRadius: 10, padding: '11px 0', fontWeight: 700, fontSize: 13, cursor: disabled ? 'default' : 'pointer' }}>
      {children}
    </button>
  );
}
function Chip({ label, active, onClick }) {
  return (
    <button onClick={onClick} style={{ padding: '6px 11px', borderRadius: 999, border: `1px solid ${active ? LEAF : LINE}`, background: active ? LEAF : '#fff', color: active ? '#fff' : INK, fontSize: 12, fontWeight: 600, marginRight: 6, marginBottom: 6, cursor: 'pointer' }}>
      {label}
    </button>
  );
}
function StatusPill({ status }) {
  const map = { pending: { bg: '#FBEFDC', color: AMBER, label: 'Pending' }, packed: { bg: '#E6F1FB', color: '#1B5E8C', label: 'Packed' }, dispatched: { bg: '#EAF3DE', color: LEAF_DARK, label: 'Dispatched' } };
  const s = map[status] || map.pending;
  return <span style={{ background: s.bg, color: s.color, fontSize: 11, fontWeight: 700, padding: '3px 9px', borderRadius: 999 }}>{s.label}</span>;
}
const sectionTitle = { fontWeight: 700, fontSize: 14, color: INK, marginBottom: 8 };
const hint = { fontSize: 11, color: MUTED, marginBottom: 8 };
const smallLabel = { fontSize: 11, color: MUTED, fontWeight: 700, marginBottom: 4, marginTop: 4 };

export default function FnvMobilePreview() {
  const [drawerOpen, setDrawerOpen] = useState(false);
  const [tab, setTab] = useState('dashboard');

  const [items, setItems] = useState([]);
  const [orders, setOrders] = useState([]);
  const [purchases, setPurchases] = useState([]);
  const [recipes, setRecipes] = useState([]);
  const [roles, setRoles] = useState([]);
  const [users, setUsers] = useState([]);
  const [vendors, setVendors] = useState([]);
  const [vendorLedger, setVendorLedger] = useState([]);  // { id, vendorId, itemName, qty, unit, unitPrice, total, payment, date, note }
  const [placedOrders, setPlacedOrders] = useState([]); // { id, name, date, items: [{itemId, itemName, uom, qty, toBuy}] }
  const [cratesByCity, setCratesByCity] = useState({});
  const [crateLog, setCrateLog] = useState([]);
  const [dispatchLog, setDispatchLog] = useState([]);
  const [indentBatches, setIndentBatches] = useState([]);
  const [packingProgress, setPackingProgress] = useState({}); // { [targetKey]: packedPacks }
  const [stockCounts, setStockCounts] = useState([]); // nightly closing-stock entries, one per item per date
  const [pricingConfig, setPricingConfig] = useState([]); // editable per-article pricing inputs
  const [grnReports, setGrnReports] = useState([]); // uploaded GRN files per channel + day
  const [dbReady, setDbReady] = useState(false);
  const [selectedCity, setSelectedCity] = usePersistedState('fnv_selected_city', CITIES[0]);

  useEffect(() => {
    // Seed collections on first load, then subscribe
    (async () => {
      await Promise.all([
        seedIfEmpty('items',     SEED_ITEMS),
        seedIfEmpty('orders',    SEED_ORDERS),
        seedIfEmpty('purchases', SEED_PURCHASES),
        seedIfEmpty('recipes',   SEED_RECIPES),
        seedIfEmpty('roles',     SEED_ROLES),
        seedIfEmpty('users',     SEED_USERS),
        seedIfEmpty('vendors',   SEED_VENDORS),
      ]);
      setDbReady(true);
    })();

    const cols = ['items','orders','purchases','recipes','roles','users','vendors','vendorLedger','placedOrders','indentBatches','crateLog','dispatchLog','stockCounts','pricingConfig','grnReports'];
    const setters = { items: setItems, orders: setOrders, purchases: setPurchases, recipes: setRecipes, roles: setRoles, users: setUsers, vendors: setVendors, vendorLedger: setVendorLedger, placedOrders: setPlacedOrders, indentBatches: setIndentBatches, crateLog: setCrateLog, dispatchLog: setDispatchLog, stockCounts: setStockCounts, pricingConfig: setPricingConfig, grnReports: setGrnReports };

    const unsubs = cols.map((col) =>
      onSnapshot(collection(db, col), (snap) => {
        const data = snap.docs.map((d) => ({ ...d.data(), id: d.id }));
        setters[col](data);
      })
    );

    // crates — one doc per city (legacy installs had a single flat {crates,boxes} object,
    // which we treat as belonging to the first city so nothing is lost)
    const unsub2 = onSnapshot(doc(db, 'settings', 'crates'), (d) => {
      if (d.exists()) {
        const data = d.data();
        if (typeof data.crates === 'number') {
          setCratesByCity({ [CITIES[0]]: { crates: data.crates, boxes: data.boxes } });
        } else {
          setCratesByCity(data);
        }
      }
    });

    // packing progress — keyed by target id, stored as a map for O(1) lookup
    const unsub3 = onSnapshot(collection(db, 'packingProgress'), (snap) => {
      const map = {};
      snap.docs.forEach((d) => { map[d.id] = { packedQty: d.data().packedQty || 0, shortQty: d.data().shortQty || 0 }; });
      setPackingProgress(map);
    });

    return () => { unsubs.forEach((u) => u()); unsub2(); unsub3(); };
  }, []);

  const fbUpdate = (col, id, patch)  => updateDoc(doc(db, col, id), patch);
  const fbDelete = (col, id)         => deleteDoc(doc(db, col, id));
  const fbSetDoc = (col, id, obj)    => setDoc(doc(db, col, id), obj);

  const addItem = (it) => fbSetDoc('items', it.id, { ...it, city: selectedCity });
  const addItemsBulk = (newItems) => { const b = writeBatch(db); newItems.forEach((it) => b.set(doc(db,'items',it.id), { ...it, city: selectedCity })); b.commit(); };
  const updateItem = (id, u) => fbUpdate('items', id, u);
  const deleteItem = (id) => fbDelete('items', id);
  // Different articles from the same channel can map to the same base item but have
  // their own pack size (e.g. "Baby Banana" 500g vs "Banana 3pc" 600g, both on Blinkit,
  // both = item "Banana"). So each distinct article code gets its OWN alias entry.
  // A single article code must only ever belong to ONE item — if it's already an
  // alias on a different item (e.g. someone picked the wrong item from a long
  // dropdown once), transfer it here instead of letting two items share the same
  // code, which makes future auto-matching pick whichever item happens to come first.
  const ensureAliasForCode = (itemId, channel, code) => {
    const it = items.find((x) => x.id === itemId); if (!it) return;
    if (!code) {
      const exists = (it.aliases || []).some((a) => a.channel === channel && !a.code);
      if (exists) return;
      fbUpdate('items', itemId, { aliases: [...(it.aliases || []), { id: newAliasId(), channel, code: '', packSize: '', packUnit: 'kg' }] });
      return;
    }
    const codeLower = code.toLowerCase();
    const alreadyHere = (it.aliases || []).some((a) => a.channel === channel && a.code && a.code.toLowerCase() === codeLower);
    if (alreadyHere) return;
    items.forEach((other) => {
      if (other.id === itemId) return;
      const hasIt = (other.aliases || []).some((a) => a.channel === channel && a.code && a.code.toLowerCase() === codeLower);
      if (hasIt) {
        fbUpdate('items', other.id, { aliases: other.aliases.filter((a) => !(a.channel === channel && a.code && a.code.toLowerCase() === codeLower)) });
      }
    });
    fbUpdate('items', itemId, { aliases: [...(it.aliases || []), { id: newAliasId(), channel, code, packSize: '', packUnit: 'kg' }] });
  };
  const updateAliasById = (itemId, aliasId, patch) => {
    const it = items.find((x) => x.id === itemId); if (!it) return;
    fbUpdate('items', itemId, { aliases: (it.aliases || []).map((a) => (a.id === aliasId ? { ...a, ...patch } : a)) });
  };
  const importOrder = (o) => fbSetDoc('orders', o.id, { ...o, city: selectedCity });
  const advanceMany = (ids, next) => { const b = writeBatch(db); ids.forEach((id) => b.update(doc(db,'orders',id), { status: next })); b.commit(); };
  const addPurchase = (p) => fbSetDoc('purchases', p.id, { date: new Date().toISOString().split('T')[0], type: 'purchased', ...p, city: selectedCity });
  const addPurchaseRequirements = (rows, dateOverride) => { const b = writeBatch(db); const today = dateOverride || new Date().toISOString().split('T')[0]; rows.forEach((r) => b.set(doc(db,'purchases',r.id), { date: today, type: 'requirement', ...r, city: selectedCity })); b.commit(); };
  const removePurchasesByIds = (ids) => { const b = writeBatch(db); ids.forEach((id) => b.delete(doc(db,'purchases',id))); b.commit(); };
  const recordStockCount = (itemId, itemName, unit, date, closingQty) => {
    fbSetDoc('stockCounts', `${itemId}__${date}`, { id: `${itemId}__${date}`, itemId, itemName, unit, date, closingQty: Number(closingQty) || 0, city: selectedCity });
  };
  const updatePricingConfig = (key, patch) => {
    const existing = pricingConfig.find((p) => p.id === key);
    fbSetDoc('pricingConfig', key, { id: key, ...(existing || {}), ...patch });
  };
  const uploadGrnReport = (channel, date, fileName, rows, batchId) => {
    const id = `GRN-${channel.slice(0, 3).toUpperCase()}-${date}-${Date.now().toString(36).toUpperCase().slice(-6)}`;
    fbSetDoc('grnReports', id, { id, channel, date, fileName, uploadedAt: new Date().toISOString().split('T')[0], rows, batchId: batchId || null });
  };
  const addRecipe = (r) => fbSetDoc('recipes', r.id, r);
  const deleteRecipe = (id) => fbDelete('recipes', id);
  const adjustCrates = async (type, delta, note) => {
    const current = cratesByCity[selectedCity] || { crates: 0, boxes: 0 };
    const next = { ...current, [type]: Math.max(0, current[type] + delta) };
    await setDoc(doc(db, 'settings', 'crates'), { ...cratesByCity, [selectedCity]: next });
    const logId = `CL-${Date.now().toString(36).toUpperCase().slice(-6)}`;
    fbSetDoc('crateLog', logId, { id: logId, type, delta, note: note || '', time: new Date().toLocaleTimeString('en-IN', { hour: '2-digit', minute: '2-digit' }), city: selectedCity });
  };
  // Supports partial dispatch: an order's full qty doesn't have to go out in one trip.
  // Each entry carries how much is actually leaving now (dispatchQty) and how much is
  // permanently short (shortQty) — whatever's left stays "packed" for the next trip.
  const dispatchBatch = ({ items: dispatchItems, vehicleNo, driverName, cratesUsed, boxesUsed }) => {
    const b = writeBatch(db);
    let totalDispatchQty = 0;
    const logItems = [];
    dispatchItems.forEach(({ orderId, dispatchQty, shortQty }) => {
      const o = orders.find((x) => x.id === orderId);
      if (!o) return;
      const dQty = Number(dispatchQty) || 0;
      const sQty = Number(shortQty) || 0;
      if (dQty <= 0 && sQty <= 0) return;
      const prevDispatched = o.dispatchedQty || 0;
      const prevShort = o.shortQty || 0;
      const newDispatched = prevDispatched + dQty;
      const newShort = prevShort + sQty;
      const remaining = Math.round((o.qty - newDispatched - newShort) * 100) / 100;
      const patch = { dispatchedQty: newDispatched, shortQty: newShort };
      patch.status = remaining > 0.01 ? 'packed' : 'dispatched';
      b.update(doc(db, 'orders', orderId), patch);
      totalDispatchQty += dQty;
      logItems.push({
        orderId, product: o.articleName || o.product, unit: o.unit, dispatchQty: dQty, shortQty: sQty, remaining: Math.max(0, remaining),
        platform: o.platform, store: o.store || '', baseProduct: o.product, packSize: o.packSize || null, packUnit: o.packUnit || null, batchId: o.batchId || null,
      });
    });
    b.commit();
    const dispatchDate = new Date().toISOString().split('T')[0];
    if (cratesUsed > 0) adjustCrates('crates', -cratesUsed, `Dispatch ${vehicleNo || ''}`.trim());
    if (boxesUsed > 0) adjustCrates('boxes', -boxesUsed, `Dispatch ${vehicleNo || ''}`.trim());
    const did = `DSP-${Date.now().toString(36).toUpperCase().slice(-6)}`;
    fbSetDoc('dispatchLog', did, { id: did, date: dispatchDate, items: logItems, orderIds: logItems.map((li) => li.orderId), totalDispatchQty, vehicleNo: vehicleNo || '—', driverName: driverName || '—', cratesUsed, boxesUsed, time: new Date().toLocaleTimeString('en-IN', { hour: '2-digit', minute: '2-digit' }), city: selectedCity });
  };
  const createIndentBatch = (b) => fbSetDoc('indentBatches', b.id, { ...b, city: selectedCity });
  const updateIndentBatch = (batchId, patch) => fbUpdate('indentBatches', batchId, patch);
  // An article is only ever fully resolved two ways: fully packed, or packed+short
  // adding up to the full target — there's no partial/unresolved state that reaches
  // Dispatch. Packing progress is tracked per aggregated target (it can combine several
  // orders sharing the same product/platform/pack size/date), so the packed vs short
  // split is distributed across those orders in proportion to each order's own pack
  // count — an order that ends up with zero packed (fully short) never becomes
  // dispatchable; its whole quantity is recorded as short right away instead of sitting
  // in "packed" with nothing to send.
  const updatePackedQty = (key, packedQty, shortQty, orderIds, targetPacks) => {
    fbSetDoc('packingProgress', key, { packedQty, shortQty });
    const resolved = targetPacks > 0 && (packedQty + shortQty) >= targetPacks;
    const targetOrders = orderIds.map((id) => orders.find((o) => o.id === id)).filter(Boolean);
    const totalPacks = targetOrders.reduce((s, o) => s + (Number(o.packQty) || 0), 0) || 1;
    const alloc = resolved ? allocateShortPacks(targetOrders, shortQty) : null;
    targetOrders.forEach((o) => {
      if (o.status === 'dispatched') return;
      if (!resolved) {
        if (o.status !== 'pending') fbUpdate('orders', o.id, { status: 'pending' });
        return;
      }
      const share = (Number(o.packQty) || 0) / totalPacks;
      const myShortPacks = alloc ? alloc[o.id] : Math.round(shortQty * share * 100) / 100;
      const myPackedPacks = alloc ? (Number(o.packQty) || 0) - alloc[o.id] : Math.round(packedQty * share * 100) / 100;
      const packSize = Number(o.packSize) || 1;
      const myShortQty = Math.round(myShortPacks * packSize * 100) / 100;
      if (myPackedPacks <= 0) {
        fbUpdate('orders', o.id, { status: 'dispatched', dispatchedQty: 0, shortQty: myShortQty });
      } else {
        fbUpdate('orders', o.id, { status: 'packed', shortQty: myShortQty });
      }
    });
  };
  const toggleReleaseBatch = (batchId, purchaseDate) => {
    const batch = indentBatches.find((b) => b.id === batchId);
    if (!batch) return;
    if (batch.released) {
      removePurchasesByIds(batch.purchaseRowIds);
      fbUpdate('indentBatches', batchId, { released: false, purchaseRowIds: [] });
    } else {
      const newRows = batch.compiled.map((c, i) => ({ id: `P-REL-${batchId}-${i}`, item: c.itemName, supplier: '', qty: c.qty, unit: c.unit, cost: 0, source: `Released: ${batch.platform} indent (${batch.fileName})` }));
      addPurchaseRequirements(newRows, purchaseDate);
      fbUpdate('indentBatches', batchId, { released: true, purchaseRowIds: newRows.map((r) => r.id), purchaseDate });
    }
  };
  const addUser = (u) => fbSetDoc('users', u.id, u);
  const updateUser = (id, u) => fbUpdate('users', id, u);
  const deleteUser = (id) => fbDelete('users', id);
  const addRole = (r) => fbSetDoc('roles', r.id, r);
  const deleteRole = (id) => fbDelete('roles', id);
  const toggleRolePermission = (roleId, key, val) => {
    const r = roles.find((x) => x.id === roleId); if (!r) return;
    fbUpdate('roles', roleId, { permissions: { ...r.permissions, [key]: val } });
  };
  const addVendor = (v) => fbSetDoc('vendors', v.id, { ...v, city: selectedCity });
  const savePlacedOrder = (order) => fbSetDoc('placedOrders', order.id, order);
  const updatePlacedOrder = (id, itemsList) => fbUpdate('placedOrders', id, { items: itemsList });
  const deletePlacedOrder = (id) => fbDelete('placedOrders', id);
  const addLedgerEntry = (entry) => {
    fbSetDoc('vendorLedger', entry.id, entry);
    addPurchase({ id: `P-${Date.now().toString(36).toUpperCase().slice(-5)}`, item: entry.itemName, supplier: entry.vendorName, qty: entry.qty, unit: entry.unit, cost: entry.total, source: entry.payment === 'credit' ? `Credit — ${entry.vendorName}` : entry.payment, date: entry.date });
  };
  const settleEntries = (ids, paymentMode, note, edits = {}) => {
    const b = writeBatch(db);
    ids.forEach((id) => {
      const e = vendorLedger.find((x) => x.id === id); if (!e) return;
      const d = edits[id] || {};
      const qty = d.qty !== undefined ? Number(d.qty) : e.qty;
      const unitPrice = d.unitPrice !== undefined ? Number(d.unitPrice) : e.unitPrice;
      const total = d.total !== undefined ? Number(d.total) : Math.round(qty * unitPrice * 100) / 100;
      b.update(doc(db, 'vendorLedger', id), { qty, unitPrice, total, settled: true, settledPayment: paymentMode, settledNote: note, settledDate: new Date().toISOString().split('T')[0] });
    });
    b.commit();
  };
  const deleteVendor = (id) => fbDelete('vendors', id);
  const toggleVendorItem = (vendorId, itemId) => {
    const v = vendors.find((x) => x.id === vendorId); if (!v) return;
    const next = v.itemIds.includes(itemId) ? v.itemIds.filter((id) => id !== itemId) : [...v.itemIds, itemId];
    fbUpdate('vendors', vendorId, { itemIds: next });
  };

  const currentNav = NAV.find((n) => n.key === tab);
  const cityItems = items.filter((it) => (it.city || CITIES[0]) === selectedCity);
  const cityVendors = vendors.filter((v) => (v.city || CITIES[0]) === selectedCity);
  const cityOrders = orders.filter((o) => (o.city || CITIES[0]) === selectedCity);
  const cityPurchases = purchases.filter((p) => (p.city || CITIES[0]) === selectedCity);
  const cityIndentBatches = indentBatches.filter((b) => (b.city || CITIES[0]) === selectedCity);
  const cityStockCounts = stockCounts.filter((sc) => (sc.city || CITIES[0]) === selectedCity);
  const cityDispatchLog = dispatchLog.filter((d) => (d.city || CITIES[0]) === selectedCity);
  const cityCrates = cratesByCity[selectedCity] || { crates: 0, boxes: 0 };
  const cityCrateLog = crateLog.filter((l) => (l.city || CITIES[0]) === selectedCity);

  if (!dbReady) return (
    <div style={{ display: 'flex', justifyContent: 'center', padding: '24px 12px', fontFamily: '"Nunito Sans", -apple-system, BlinkMacSystemFont, "Segoe UI", sans-serif' }}>
      <div style={{ width: 390, height: 760, background: BG, borderRadius: 34, border: `8px solid ${INK}`, boxShadow: '0 20px 50px rgba(0,0,0,0.18)', overflow: 'hidden', display: 'flex', alignItems: 'center', justifyContent: 'center', flexDirection: 'column', gap: 16 }}>
        <Sprout size={36} color={LEAF} />
        <div style={{ fontWeight: 700, fontSize: 16, color: INK }}>Connecting to database…</div>
      </div>
    </div>
  );

  return (
    <div style={{ display: 'flex', justifyContent: 'center', padding: '24px 12px', fontFamily: '"Nunito Sans", -apple-system, BlinkMacSystemFont, "Segoe UI", sans-serif' }}>
      <div style={{ width: 390, height: 760, background: BG, borderRadius: 34, border: `8px solid ${INK}`, boxShadow: '0 20px 50px rgba(0,0,0,0.18)', overflow: 'hidden', position: 'relative', display: 'flex', flexDirection: 'column' }}>
        <div style={{ height: 22, background: LEAF_DARK, flexShrink: 0 }} />
        <div style={{ background: LEAF_DARK, color: '#fff', padding: '12px 16px', display: 'flex', alignItems: 'center', gap: 10, flexShrink: 0 }}>
          <button onClick={() => setDrawerOpen(true)} style={{ background: 'none', border: 'none', color: '#fff', cursor: 'pointer', display: 'flex' }}>
            <Menu size={20} />
          </button>
          <Text style={{ fontWeight: 800, fontSize: 16 }}>{currentNav?.label}</Text>
        </div>

        <div style={{ flex: 1, overflowY: 'auto' }}>
          {tab === 'dashboard' && <DashboardTab orders={cityOrders} purchases={cityPurchases} items={cityItems} crates={cityCrates} />}
          {tab === 'items' && <ItemsTab items={cityItems} onAdd={addItem} onAddBulk={addItemsBulk} onUpdate={updateItem} onDelete={deleteItem} />}
          {tab === 'vendors' && (
            <VendorsTab items={cityItems} vendors={cityVendors} vendorLedger={vendorLedger} placedOrders={placedOrders} onAdd={addVendor} onDelete={deleteVendor} onToggleItem={toggleVendorItem} onSettle={settleEntries} onUpdatePlacedOrder={updatePlacedOrder} onDeletePlacedOrder={deletePlacedOrder} />
          )}
          {tab === 'cutprocess' && <CutProcessTab items={items} recipes={recipes} orders={orders} onAddRecipe={addRecipe} onDeleteRecipe={deleteRecipe} onAddPurchaseRequirements={addPurchaseRequirements} />}
          {tab === 'orders' && (
            <OrdersTab
              orders={cityOrders} items={cityItems} indentBatches={cityIndentBatches}
              onImport={importOrder} onAddItem={addItem} onEnsureAlias={ensureAliasForCode} onUpdateAlias={updateAliasById}
              onCreateIndentBatch={createIndentBatch} onToggleReleaseBatch={toggleReleaseBatch}
            />
          )}
          {tab === 'purchase' && <PurchasesTab purchases={cityPurchases} orders={cityOrders} items={cityItems} allItems={items} recipes={recipes} vendors={cityVendors} vendorLedger={vendorLedger} stockCounts={cityStockCounts} onAddLedgerEntry={addLedgerEntry} onSavePlacedOrder={savePlacedOrder} indentBatches={cityIndentBatches} />}
          {tab === 'stockcount' && <StockCountTab items={cityItems} stockCounts={cityStockCounts} onRecord={recordStockCount} />}
          {tab === 'pricing' && <PricingTab orders={orders} items={items} purchases={purchases} pricingConfig={pricingConfig} onUpdate={updatePricingConfig} />}
          {tab === 'profitloss' && <ProfitLossTab orders={orders} items={items} purchases={purchases} pricingConfig={pricingConfig} dispatchLog={dispatchLog} grnReports={grnReports} indentBatches={indentBatches} onUploadGrn={uploadGrnReport} onUpdateIndentBatch={updateIndentBatch} />}
          {tab === 'packaging' && <PackagingTab orders={cityOrders} items={cityItems} onAdvanceMany={advanceMany} packingProgress={packingProgress} onUpdatePackedQty={updatePackedQty} />}
          {tab === 'dispatch' && (
            <DispatchTab orders={cityOrders} crates={cityCrates} dispatchLog={cityDispatchLog} onDispatchBatch={dispatchBatch} />
          )}
          {tab === 'crates' && <CratesTab crates={cityCrates} log={cityCrateLog} onAdjust={adjustCrates} />}
          {tab === 'users' && (
            <UsersRolesTab
              users={users} roles={roles}
              onAddUser={addUser} onUpdateUser={updateUser} onDeleteUser={deleteUser}
              onAddRole={addRole} onDeleteRole={deleteRole} onToggleRolePermission={toggleRolePermission}
            />
          )}
        </div>

        {drawerOpen && (
          <div style={{ position: 'absolute', inset: 0, display: 'flex' }}>
            <div style={{ width: 250, background: SIDEBAR, height: '100%', display: 'flex', flexDirection: 'column' }}>
              <div style={{ padding: '20px 18px 14px', display: 'flex', alignItems: 'center', justifyContent: 'space-between', borderBottom: '1px solid rgba(255,255,255,0.08)' }}>
                <div style={{ display: 'flex', alignItems: 'center', gap: 8 }}>
                  <Sprout size={18} color="#8FBF7A" />
                  <span style={{ color: '#fff', fontWeight: 800, fontSize: 14 }}>FNV Admin</span>
                </div>
                <button onClick={() => setDrawerOpen(false)} style={{ background: 'none', border: 'none', color: '#B7C2B2', cursor: 'pointer' }}><X size={18} /></button>
              </div>
              <div style={{ padding: '12px 18px', borderBottom: '1px solid rgba(255,255,255,0.08)' }}>
                <p style={{ margin: '0 0 6px', fontSize: 10, color: '#8A968A', fontWeight: 700, letterSpacing: 0.5 }}>CITY</p>
                <select
                  value={selectedCity}
                  onChange={(e) => setSelectedCity(e.target.value)}
                  style={{ width: '100%', boxSizing: 'border-box', background: 'rgba(255,255,255,0.08)', color: '#fff', border: '1px solid rgba(255,255,255,0.15)', borderRadius: 8, padding: '8px', fontSize: 13, fontWeight: 700 }}
                >
                  {CITIES.map((c) => <option key={c} value={c} style={{ color: INK }}>{c}</option>)}
                </select>
              </div>
              <div style={{ padding: '10px 8px', flex: 1, overflowY: 'auto' }}>
                {NAV.map((n) => (
                  <button
                    key={n.key}
                    onClick={() => { setTab(n.key); setDrawerOpen(false); }}
                    style={{ width: '100%', display: 'flex', alignItems: 'center', gap: 10, padding: '10px 12px', marginBottom: 2, borderRadius: 8, border: 'none', background: tab === n.key ? 'rgba(255,255,255,0.1)' : 'transparent', color: tab === n.key ? '#fff' : '#B7C2B2', fontSize: 13, fontWeight: tab === n.key ? 700 : 500, cursor: 'pointer', textAlign: 'left' }}
                  >
                    <n.icon size={16} />
                    {n.label}
                  </button>
                ))}
              </div>
            </div>
            <div onClick={() => setDrawerOpen(false)} style={{ flex: 1, background: 'rgba(0,0,0,0.35)' }} />
          </div>
        )}
      </div>
    </div>
  );
}

function Text({ children, style }) {
  return <span style={style}>{children}</span>;
}

// ---------- Dashboard ----------
function DashboardTab({ orders, purchases, items, crates }) {
  const pending = orders.filter((o) => o.status === 'pending').length;
  const dispatched = orders.filter((o) => o.status === 'dispatched').length;
  const totalSpend = purchases.reduce((s, p) => s + p.cost, 0);
  return (
    <div style={{ padding: 16 }}>
      <div style={{ display: 'flex', gap: 10, marginBottom: 10 }}>
        <Card style={{ flex: 1 }}><div style={hint}>Active items</div><div style={{ fontSize: 22, fontWeight: 800 }}>{items.length}</div></Card>
        <Card style={{ flex: 1 }}><div style={hint}>Pending orders</div><div style={{ fontSize: 22, fontWeight: 800, color: AMBER }}>{pending}</div></Card>
      </div>
      <div style={{ display: 'flex', gap: 10, marginBottom: 10 }}>
        <Card style={{ flex: 1 }}><div style={hint}>Dispatched</div><div style={{ fontSize: 22, fontWeight: 800, color: LEAF }}>{dispatched}</div></Card>
        <Card style={{ flex: 1 }}><div style={hint}>Purchase spend</div><div style={{ fontSize: 18, fontWeight: 800, color: TOMATO }}>₹{totalSpend.toLocaleString('en-IN')}</div></Card>
      </div>
      <div style={{ display: 'flex', gap: 10, marginBottom: 14 }}>
        <Card style={{ flex: 1 }}><div style={hint}>Crates</div><div style={{ fontSize: 20, fontWeight: 800 }}>{crates.crates}</div></Card>
        <Card style={{ flex: 1 }}><div style={hint}>Boxes</div><div style={{ fontSize: 20, fontWeight: 800 }}>{crates.boxes}</div></Card>
      </div>
      <Card>
        <div style={sectionTitle}>Recent orders</div>
        {orders.slice(0, 6).map((o) => (
          <div key={o.id} style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', borderTop: `1px solid ${LINE}`, padding: '8px 0' }}>
            <div>
              <div style={{ fontWeight: 700, fontSize: 13 }}>{o.id}</div>
              <div style={{ fontSize: 12, color: MUTED }}>{o.articleName || o.product} · {o.qty}{o.unit}</div>
            </div>
            <StatusPill status={o.status} />
          </div>
        ))}
      </Card>
    </div>
  );
}

// ---------- Items ----------
const ITEM_TEMPLATE_ROWS = [
  {
    'Item Name': 'Tomato',
    UOM: 'kg',
    Category: 'VEGETABLES',
    'Blinkit Code': 'BLK-TOM-240',
    'Blinkit Pack Size': 0.5,
    'Blinkit Pack Unit': 'kg',
    'Flipkart Code': 'FKT-TOM-01',
    'Flipkart Pack Size': 1,
    'Flipkart Pack Unit': 'kg',
  },
];

function downloadItemsTemplate() {
  const wb = XLSX.utils.book_new();
  const ws = XLSX.utils.json_to_sheet(ITEM_TEMPLATE_ROWS);
  XLSX.utils.book_append_sheet(wb, ws, 'Items');
  const wbout = XLSX.write(wb, { bookType: 'xlsx', type: 'array' });
  const blob = new Blob([wbout], { type: 'application/octet-stream' });
  const url = URL.createObjectURL(blob);
  const a = document.createElement('a');
  a.href = url;
  a.download = 'fnv-items-template.xlsx';
  document.body.appendChild(a);
  a.click();
  document.body.removeChild(a);
  URL.revokeObjectURL(url);
}

function parseBulkItemRows(json) {
  const results = { valid: [], skipped: 0 };
  json.forEach((r) => {
    const name = String(pickField(r, ['itemname', 'name', 'article', 'product']) || '').trim();
    if (!name) { results.skipped += 1; return; }
    const uom = String(pickField(r, ['uom', 'unit']) || 'kg').trim() || 'kg';
    const category = normalizeCategory(pickField(r, ['category', 'type']));
    const blinkitCode = String(pickField(r, ['blinkitcode']) || '').trim();
    const blinkitPackSize = String(pickField(r, ['blinkitpacksize']) || '').trim();
    const blinkitPackUnit = String(pickField(r, ['blinkitpackunit']) || 'kg').trim() || 'kg';
    const flipkartCode = String(pickField(r, ['flipkartcode']) || '').trim();
    const flipkartPackSize = String(pickField(r, ['flipkartpacksize']) || '').trim();
    const flipkartPackUnit = String(pickField(r, ['flipkartpackunit']) || 'kg').trim() || 'kg';
    const aliases = [];
    if (blinkitCode || blinkitPackSize) aliases.push({ id: newAliasId(), channel: 'Blinkit', code: blinkitCode, packSize: blinkitPackSize, packUnit: blinkitPackUnit });
    if (flipkartCode || flipkartPackSize) aliases.push({ id: newAliasId(), channel: 'Flipkart', code: flipkartCode, packSize: flipkartPackSize, packUnit: flipkartPackUnit });
    results.valid.push({
      id: `IT-${Date.now().toString(36).toUpperCase().slice(-5)}-${results.valid.length}`,
      name, uom, category, aliases,
    });
  });
  return results;
}

function VendorItemLinkerMobile({ vendorId, vendorItemIds, items, onToggle }) {
  const [search, setSearch] = useState('');
  const linkedItems = items.filter((it) => vendorItemIds.includes(it.id));
  const suggestions = search.trim().length > 0
    ? items.filter((it) => !vendorItemIds.includes(it.id) && it.name.toLowerCase().includes(search.toLowerCase())).slice(0, 6)
    : [];

  return (
    <div>
      <div style={{ display: 'flex', flexWrap: 'wrap', gap: 4, marginBottom: 8 }}>
        {linkedItems.map((it) => (
          <span key={it.id} style={{ display: 'inline-flex', alignItems: 'center', gap: 4, background: '#EAF3DE', color: LEAF_DARK, fontSize: 11, fontWeight: 700, padding: '3px 8px', borderRadius: 999 }}>
            {it.name}
            <button onClick={() => onToggle(vendorId, it.id)} style={{ background: 'none', border: 'none', color: LEAF_DARK, cursor: 'pointer', lineHeight: 1, padding: 0, fontSize: 14, fontWeight: 900 }}>×</button>
          </span>
        ))}
        {linkedItems.length === 0 && <span style={{ fontSize: 11, color: MUTED }}>No items linked yet</span>}
      </div>
      <div style={{ position: 'relative' }}>
        <Field
          placeholder="Search and add item..."
          value={search}
          onChange={(e) => setSearch(e.target.value)}
          style={{ marginBottom: 0, fontSize: 12 }}
        />
        {suggestions.length > 0 && (
          <div style={{ position: 'absolute', left: 0, right: 0, background: '#fff', border: `1px solid ${LINE}`, borderRadius: 8, boxShadow: '0 4px 12px rgba(0,0,0,0.12)', zIndex: 99, maxHeight: 200, overflowY: 'auto' }}>
            {suggestions.map((it) => (
              <div
                key={it.id}
                onClick={() => { onToggle(vendorId, it.id); setSearch(''); }}
                style={{ padding: '10px 12px', cursor: 'pointer', display: 'flex', justifyContent: 'space-between', alignItems: 'center', borderBottom: `1px solid ${LINE}` }}
              >
                <span style={{ fontWeight: 600, fontSize: 13 }}>{it.name}</span>
                <span style={{ fontSize: 11, color: MUTED }}>{it.category} · {it.uom}</span>
              </div>
            ))}
          </div>
        )}
      </div>
    </div>
  );
}

const COMPANY_NAME = 'NILGIRI FNV SUPPLIER COMPANY';

function generateOrderImage(order) {
  const canvas = document.createElement('canvas');
  const ROW_H = 44, HEADER_H = 72, TITLE_H = 52, PAD = 24;
  const cols = [60, 260, 100, 100]; // NO, ITEM NAME, UOM, QTY
  const totalW = cols.reduce((s, c) => s + c, 0) + PAD * 2;
  canvas.width = totalW;
  canvas.height = TITLE_H + HEADER_H + ROW_H * (order.items.length + 1) + PAD;
  const ctx = canvas.getContext('2d');

  // Background
  ctx.fillStyle = '#FFFFFF';
  ctx.fillRect(0, 0, canvas.width, canvas.height);

  // Company name header
  ctx.fillStyle = '#1B2E1D';
  ctx.fillRect(0, 0, canvas.width, TITLE_H);
  ctx.fillStyle = '#FFFFFF';
  ctx.font = 'bold 18px Arial';
  ctx.textAlign = 'center';
  ctx.fillText(COMPANY_NAME, canvas.width / 2, 22);
  ctx.font = '13px Arial';
  ctx.fillStyle = '#B7C2B2';
  ctx.fillText(`${order.name}  ·  ${order.date}`, canvas.width / 2, 42);

  // Column headers
  let x = PAD, y = TITLE_H;
  ctx.fillStyle = '#F0EDE4';
  ctx.fillRect(0, y, canvas.width, HEADER_H);
  const headers = ['NO.', 'ITEM NAME', 'UOM', 'QTY'];
  ctx.fillStyle = '#2F5233';
  ctx.font = 'bold 13px Arial';
  ctx.textAlign = 'left';
  headers.forEach((h, i) => {
    ctx.fillText(h, x + 6, y + 28);
    x += cols[i];
  });

  // Divider
  ctx.strokeStyle = '#D0CBB8';
  ctx.lineWidth = 1;
  ctx.beginPath(); ctx.moveTo(0, y + HEADER_H - 1); ctx.lineTo(canvas.width, y + HEADER_H - 1); ctx.stroke();

  // Rows
  order.items.forEach((it, idx) => {
    const rowY = TITLE_H + HEADER_H + idx * ROW_H;
    ctx.fillStyle = idx % 2 === 0 ? '#FFFFFF' : '#F9F7F0';
    ctx.fillRect(0, rowY, canvas.width, ROW_H);

    ctx.strokeStyle = '#E3DECF';
    ctx.beginPath(); ctx.moveTo(0, rowY + ROW_H); ctx.lineTo(canvas.width, rowY + ROW_H); ctx.stroke();

    let cx = PAD;
    const vals = [String(it.no), it.itemName.toUpperCase(), it.uom, String(it.qty)];
    ctx.fillStyle = '#20241E';
    ctx.font = idx === 0 ? 'bold 13px Arial' : '13px Arial';
    vals.forEach((v, i) => {
      ctx.fillText(v, cx + 6, rowY + ROW_H / 2 + 5);
      cx += cols[i];
    });
  });

  // Column dividers
  ctx.strokeStyle = '#D0CBB8';
  let dx = PAD;
  cols.slice(0, -1).forEach((w) => {
    dx += w;
    ctx.beginPath(); ctx.moveTo(dx, TITLE_H); ctx.lineTo(dx, canvas.height); ctx.stroke();
  });

  return canvas.toDataURL('image/png');
}

function PlacedOrderCard({ order, onUpdate, onDelete }) {
  const [expanded, setExpanded] = useState(false);
  const [editItems, setEditItems] = useState(order.items);
  const [confirmDel, setConfirmDel] = useState(false);

  const updateRow = (idx, field, val) => {
    const updated = editItems.map((it, i) => i === idx ? { ...it, [field]: val } : it);
    setEditItems(updated);
    onUpdate(order.id, updated);
  };

  const downloadImage = () => {
    const dataUrl = generateOrderImage({ ...order, items: editItems });
    const a = document.createElement('a');
    a.href = dataUrl;
    a.download = `${order.name.replace(/\s+/g, '_')}.png`;
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
  };

  return (
    <div style={{ borderTop: `1px solid ${LINE}`, paddingTop: 12, marginBottom: 4 }}>
      <div onClick={() => setExpanded((x) => !x)} style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', cursor: 'pointer', marginBottom: expanded ? 10 : 0 }}>
        <div>
          <div style={{ fontWeight: 700, fontSize: 14 }}>{order.name}</div>
          <div style={{ fontSize: 11, color: MUTED }}>{order.date} · {editItems.length} items · {expanded ? '▲' : '▼'}</div>
        </div>
        <div style={{ display: 'flex', gap: 8 }}>
          <button onClick={(e) => { e.stopPropagation(); downloadImage(); }} style={{ background: LEAF, color: '#fff', border: 'none', borderRadius: 8, padding: '7px 12px', fontSize: 11, fontWeight: 700, cursor: 'pointer' }}>
            ↓ Image
          </button>
          {confirmDel ? (
            <>
              <button onClick={(e) => { e.stopPropagation(); onDelete(order.id); }} style={{ background: TOMATO, color: '#fff', border: 'none', borderRadius: 8, padding: '7px 10px', fontSize: 11, fontWeight: 700, cursor: 'pointer' }}>Yes</button>
              <button onClick={(e) => { e.stopPropagation(); setConfirmDel(false); }} style={{ background: '#fff', color: INK, border: `1px solid ${LINE}`, borderRadius: 8, padding: '7px 10px', fontSize: 11, fontWeight: 700, cursor: 'pointer' }}>No</button>
            </>
          ) : (
            <button onClick={(e) => { e.stopPropagation(); setConfirmDel(true); }} style={{ background: 'none', border: `1px solid ${LINE}`, borderRadius: 8, padding: '7px 10px', fontSize: 11, color: TOMATO, fontWeight: 700, cursor: 'pointer' }}>Delete</button>
          )}
        </div>
      </div>

      {expanded && (
        <div style={{ background: '#F6F3EA', borderRadius: 10, padding: '10px 12px' }}>
          <div style={{ display: 'grid', gridTemplateColumns: '24px 1fr 56px 60px', gap: 6, marginBottom: 6 }}>
            {['#', 'ITEM', 'UOM', 'QTY'].map((h) => <div key={h} style={{ fontSize: 9, fontWeight: 700, color: MUTED }}>{h}</div>)}
          </div>
          {editItems.map((it, idx) => (
            <div key={it.itemId || idx} style={{ display: 'grid', gridTemplateColumns: '24px 1fr 56px 60px', gap: 6, alignItems: 'center', borderTop: `1px solid ${LINE}`, paddingTop: 6, marginTop: 4 }}>
              <div style={{ fontSize: 11, color: MUTED, fontWeight: 700 }}>{it.no}</div>
              <input value={it.itemName} onChange={(e) => updateRow(idx, 'itemName', e.target.value)} style={{ border: `1px solid ${LINE}`, borderRadius: 6, padding: '5px 6px', fontSize: 12, width: '100%', boxSizing: 'border-box' }} />
              <input value={it.uom} onChange={(e) => updateRow(idx, 'uom', e.target.value)} style={{ border: `1px solid ${LINE}`, borderRadius: 6, padding: '5px 6px', fontSize: 12, width: '100%', boxSizing: 'border-box' }} />
              <input type="number" value={it.qty} onChange={(e) => updateRow(idx, 'qty', e.target.value)} style={{ border: `1px solid ${LINE}`, borderRadius: 6, padding: '5px 6px', fontSize: 12, width: '100%', boxSizing: 'border-box' }} />
            </div>
          ))}
          <button onClick={downloadImage} style={{ width: '100%', background: LEAF, color: '#fff', border: 'none', borderRadius: 9, padding: '10px 0', fontWeight: 700, fontSize: 13, cursor: 'pointer', marginTop: 12 }}>
            Download order image
          </button>
        </div>
      )}
    </div>
  );
}

function formatLedgerDate(d, short) {
  if (!d) return 'No date';
  const dt = new Date(`${d}T00:00:00`);
  if (isNaN(dt.getTime())) return String(d);
  return short
    ? dt.toLocaleDateString('en-IN', { day: '2-digit', month: 'short' })
    : dt.toLocaleDateString('en-IN', { weekday: 'short', day: '2-digit', month: 'short', year: 'numeric' });
}
const isDueEntry = (e) => e.payment === 'credit' && !e.settled;
const money = (n) => `₹${(Math.round((Number(n) || 0) * 100) / 100).toLocaleString('en-IN')}`;

function VendorsTab({ items, vendors, vendorLedger, placedOrders, onAdd, onDelete, onToggleItem, onSettle, onUpdatePlacedOrder, onDeletePlacedOrder }) {
  const [name, setName] = useState('');
  const [contact, setContact] = useState('');
  const [vendorSearch, setVendorSearch] = useState('');
  const [openVendorId, setOpenVendorId] = useState(null);
  const [ledgerFilter, setLedgerFilter] = useState('due'); // 'due' | 'all'
  const [selectedDates, setSelectedDates] = useState([]);
  const [showItems, setShowItems] = useState(false);
  const [confirmDeleteId, setConfirmDeleteId] = useState(null);
  const [payModal, setPayModal] = useState(null);
  const [payMode, setPayMode] = useState('cash');
  const [payRef, setPayRef] = useState('');
  const [payNote, setPayNote] = useState('');
  const [expandedGroup, setExpandedGroup] = useState(null);
  const [draftEdits, setDraftEdits] = useState({});

  const submit = () => {
    if (!name.trim()) return;
    onAdd({ id: `VEN-${Date.now().toString(36).toUpperCase().slice(-5)}`, name: name.trim(), contact: contact.trim(), itemIds: [] });
    setName(''); setContact('');
  };

  const updateDraft = (entryId, field, value) => {
    setDraftEdits((prev) => {
      const cur = prev[entryId] || {};
      const updated = { ...cur, [field]: value };
      const q = Number(field === 'qty' ? value : (updated.qty ?? ''));
      const p = Number(field === 'unitPrice' ? value : (updated.unitPrice ?? ''));
      if (q > 0 && p > 0) updated.total = Math.round(q * p * 100) / 100;
      return { ...prev, [entryId]: updated };
    });
  };
  const getEffective = (entry) => {
    const d = draftEdits[entry.id] || {};
    const qty = d.qty !== undefined ? Number(d.qty) : entry.qty;
    const unitPrice = d.unitPrice !== undefined ? Number(d.unitPrice) : entry.unitPrice;
    const total = d.total !== undefined ? Number(d.total) : (qty * unitPrice || entry.total);
    return { qty, unitPrice, total };
  };
  const sumEffective = (entries) => entries.reduce((s, e) => s + getEffective(e).total, 0);
  const groupEffectiveTotal = (g) => sumEffective(g.entries);

  // ---- per-vendor helpers
  const dueEntriesOf = (vendorId) => vendorLedger.filter((e) => e.vendorId === vendorId && isDueEntry(e));
  const dueTotalOf = (vendorId) => sumEffective(dueEntriesOf(vendorId));
  const totalDue = vendors.reduce((s, v) => s + dueTotalOf(v.id), 0);

  // Date-wise ledger for one vendor, newest day first.
  const ledgerGroupsOf = (vendorId) => {
    const map = {};
    vendorLedger.filter((e) => e.vendorId === vendorId).forEach((e) => {
      const d = e.date || '';
      if (!map[d]) map[d] = { date: d, entries: [], due: [] };
      map[d].entries.push(e);
      if (isDueEntry(e)) map[d].due.push(e);
    });
    return Object.values(map).sort((a, b) => b.date.localeCompare(a.date));
  };

  const openPay = (vendor, entries) => {
    if (entries.length === 0) return;
    const sorted = entries.slice().sort((a, b) => (a.date || '').localeCompare(b.date || ''));
    const dates = Array.from(new Set(sorted.map((e) => e.date || '')));
    const label = dates.length > 1
      ? `${dates.length} days · ${formatLedgerDate(dates[0], true)} – ${formatLedgerDate(dates[dates.length - 1], true)}`
      : formatLedgerDate(dates[0]);
    setPayModal({ vendorName: vendor.name, date: label, multi: dates.length > 1, entries: sorted });
    setPayMode('cash'); setPayRef(''); setPayNote('');
  };

  const confirmPayment = () => {
    if (!payModal) return;
    const ids = payModal.entries.map((e) => e.id);
    const note = [payRef.trim() ? `Ref: ${payRef.trim()}` : '', payNote.trim()].filter(Boolean).join(' · ');
    onSettle(ids, payMode, note, draftEdits);
    setDraftEdits((prev) => { const next = { ...prev }; ids.forEach((id) => { delete next[id]; }); return next; });
    setSelectedDates([]);
    setPayModal(null); setPayMode('cash'); setPayRef(''); setPayNote('');
  };

  const openLedger = (id) => {
    setOpenVendorId(id); setLedgerFilter('due'); setSelectedDates([]);
    setExpandedGroup(null); setShowItems(false); setConfirmDeleteId(null);
  };
  const closeLedger = () => { setOpenVendorId(null); setSelectedDates([]); setExpandedGroup(null); setConfirmDeleteId(null); };
  const toggleDate = (d) => setSelectedDates((prev) => (prev.includes(d) ? prev.filter((x) => x !== d) : [...prev, d]));

  // ---- Payment modal (bottom sheet) shared by the list and the ledger view
  const payModalEl = payModal && (
    <div style={{ position: 'fixed', inset: 0, background: 'rgba(0,0,0,0.45)', display: 'flex', alignItems: 'flex-end', justifyContent: 'center', zIndex: 200 }}>
      <div style={{ background: '#fff', borderRadius: '18px 18px 0 0', padding: '24px 20px 32px', width: '100%', maxWidth: 420, maxHeight: '88vh', overflowY: 'auto' }}>
        <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', marginBottom: 18 }}>
          <div style={{ fontWeight: 800, fontSize: 16 }}>Record payment</div>
          <button onClick={() => setPayModal(null)} style={{ background: 'none', border: 'none', fontSize: 20, color: MUTED, cursor: 'pointer' }}>✕</button>
        </div>

        {/* Breakdown */}
        <div style={{ background: '#F6F3EA', borderRadius: 12, padding: '14px 14px', marginBottom: 18 }}>
          <div style={{ fontWeight: 800, fontSize: 14, marginBottom: 2 }}>{payModal.vendorName}</div>
          <div style={{ fontSize: 11, color: MUTED, marginBottom: 10 }}>{payModal.date}</div>
          {payModal.entries.map((e) => {
            const eff = getEffective(e);
            const changed = draftEdits[e.id] !== undefined;
            return (
              <div key={e.id} style={{ display: 'flex', justifyContent: 'space-between', gap: 8, padding: '5px 0', borderTop: `1px solid ${LINE}`, fontSize: 12 }}>
                <span style={{ color: changed ? AMBER : MUTED }}>{payModal.multi ? `${formatLedgerDate(e.date, true)} · ` : ''}{e.itemName} · {eff.qty} {e.unit} @ ₹{eff.unitPrice}/{e.unit}{changed ? ' ✏️' : ''}</span>
                <span style={{ fontWeight: 700, color: changed ? AMBER : INK, flexShrink: 0 }}>₹{eff.total.toLocaleString('en-IN')}</span>
              </div>
            );
          })}
          <div style={{ display: 'flex', justifyContent: 'space-between', borderTop: `2px solid ${LINE}`, marginTop: 8, paddingTop: 10 }}>
            <span style={{ fontWeight: 700 }}>Total to pay</span>
            <span style={{ fontWeight: 900, fontSize: 18, color: LEAF }}>₹{groupEffectiveTotal(payModal).toLocaleString('en-IN')}</span>
          </div>
        </div>

        {/* Payment mode */}
        <div style={smallLabel}>PAYMENT MODE</div>
        <div style={{ display: 'flex', gap: 6, marginBottom: 14 }}>
          {[{ key: 'cash', label: '💵 Cash' }, { key: 'upi', label: '📱 UPI' }, { key: 'bank', label: '🏦 Bank' }, { key: 'cheque', label: '📄 Cheque' }].map((m) => (
            <button key={m.key} onClick={() => setPayMode(m.key)} style={{ flex: 1, padding: '8px 4px', borderRadius: 8, border: `1.5px solid ${payMode === m.key ? LEAF : LINE}`, background: payMode === m.key ? '#EAF3DE' : '#fff', color: payMode === m.key ? LEAF_DARK : INK, fontSize: 10, fontWeight: 700, cursor: 'pointer' }}>
              {m.label}
            </button>
          ))}
        </div>

        {payMode !== 'cash' && (
          <>
            <div style={smallLabel}>{payMode === 'upi' ? 'UPI / TRANSACTION ID' : payMode === 'bank' ? 'NEFT / RTGS REF NO.' : 'CHEQUE NO.'}</div>
            <Field placeholder={payMode === 'cheque' ? 'e.g. 004521' : 'e.g. TXN1234567'} value={payRef} onChange={(e) => setPayRef(e.target.value)} />
          </>
        )}
        <div style={smallLabel}>NOTE (OPTIONAL)</div>
        <Field placeholder="e.g. Full settlement, partial pending..." value={payNote} onChange={(e) => setPayNote(e.target.value)} />

        <PrimaryBtn onClick={confirmPayment} color={LEAF}>
          Confirm payment — ₹{groupEffectiveTotal(payModal).toLocaleString('en-IN')}
        </PrimaryBtn>
      </div>
    </div>
  );

  // =====================  VENDOR LEDGER VIEW  =====================
  const openVendor = openVendorId ? vendors.find((v) => v.id === openVendorId) : null;
  if (openVendor) {
    const vendorDueEntries = dueEntriesOf(openVendor.id);
    const vendorDue = sumEffective(vendorDueEntries);
    const groups = ledgerGroupsOf(openVendor.id).filter((g) => ledgerFilter === 'all' || g.due.length > 0);
    const selectedEntries = groups.filter((g) => selectedDates.includes(g.date)).flatMap((g) => g.due);
    const selectedDays = groups.filter((g) => selectedDates.includes(g.date) && g.due.length > 0).length;

    return (
      <div style={{ padding: 16 }}>
        <button onClick={closeLedger} style={{ display: 'flex', alignItems: 'center', gap: 6, background: 'none', border: 'none', color: LEAF, fontWeight: 700, fontSize: 13, cursor: 'pointer', marginBottom: 12, padding: 0 }}>
          <ArrowLeft size={15} /> Vendors
        </button>

        <Card style={{ marginBottom: 12 }}>
          <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'flex-start', gap: 10 }}>
            <div style={{ minWidth: 0 }}>
              <div style={{ fontWeight: 800, fontSize: 17 }}>{openVendor.name}</div>
              <div style={{ fontSize: 12, color: MUTED, marginTop: 2 }}>
                {openVendor.contact ? <a href={`tel:${openVendor.contact}`} style={{ color: LEAF, fontWeight: 600, textDecoration: 'none' }}>{openVendor.contact}</a> : 'No number'}
              </div>
            </div>
            <div style={{ textAlign: 'right', flexShrink: 0 }}>
              <div style={{ fontSize: 9, color: MUTED, fontWeight: 700 }}>TOTAL OUTSTANDING</div>
              <div style={{ fontWeight: 800, fontSize: 20, color: vendorDue > 0 ? AMBER : LEAF }}>{money(vendorDue)}</div>
            </div>
          </div>
          <div style={{ marginTop: 12 }}>
            <PrimaryBtn onClick={() => openPay(openVendor, vendorDueEntries)} disabled={vendorDueEntries.length === 0}>
              {vendorDueEntries.length === 0 ? 'Nothing due' : `Pay all outstanding — ${money(vendorDue)}`}
            </PrimaryBtn>
          </div>
        </Card>

        <Card style={{ marginBottom: 12 }}>
          <div style={{ ...sectionTitle, marginBottom: 6 }}>Ledger — date wise</div>
          <div style={{ display: 'flex', marginBottom: 6 }}>
            <Chip label="Outstanding" active={ledgerFilter === 'due'} onClick={() => setLedgerFilter('due')} />
            <Chip label="All entries" active={ledgerFilter === 'all'} onClick={() => setLedgerFilter('all')} />
          </div>
          <div style={hint}>Tick the days you want to pay together, or tap a day to view and edit its items.</div>

          {groups.map((g) => {
            const gKey = `${openVendor.id}-${g.date}`;
            const isOpen = expandedGroup === gKey;
            const rows = ledgerFilter === 'due' ? g.due : g.entries;
            const dueAmt = sumEffective(g.due);
            const hasDue = g.due.length > 0;
            const isSelected = selectedDates.includes(g.date);
            return (
              <div key={gKey} style={{ borderTop: `1px solid ${LINE}` }}>
                <div style={{ display: 'flex', alignItems: 'center', gap: 10, padding: '11px 0' }}>
                  {hasDue ? (
                    <div onClick={() => toggleDate(g.date)} style={{ width: 20, height: 20, borderRadius: 5, border: `2px solid ${isSelected ? LEAF : LINE}`, background: isSelected ? LEAF : '#fff', flexShrink: 0, cursor: 'pointer', display: 'flex', alignItems: 'center', justifyContent: 'center' }}>
                      {isSelected && <span style={{ color: '#fff', fontSize: 12, fontWeight: 900 }}>✓</span>}
                    </div>
                  ) : (
                    <div style={{ width: 20, flexShrink: 0 }} />
                  )}
                  <div onClick={() => setExpandedGroup(isOpen ? null : gKey)} style={{ flex: 1, minWidth: 0, cursor: 'pointer' }}>
                    <div style={{ fontWeight: 700, fontSize: 13 }}>{formatLedgerDate(g.date)}</div>
                    <div style={{ fontSize: 11, color: MUTED, marginTop: 2 }}>{g.entries.length} item{g.entries.length !== 1 ? 's' : ''} · {isOpen ? '▲ hide' : '▼ view & edit'}</div>
                  </div>
                  <div onClick={() => setExpandedGroup(isOpen ? null : gKey)} style={{ textAlign: 'right', cursor: 'pointer' }}>
                    {hasDue ? (
                      <div style={{ fontWeight: 800, fontSize: 15, color: AMBER }}>{money(dueAmt)}</div>
                    ) : (
                      <div style={{ fontWeight: 800, fontSize: 13, color: LEAF }}>Paid ✓</div>
                    )}
                  </div>
                  {hasDue && (
                    <button onClick={() => openPay(openVendor, g.due)} style={{ background: LEAF, color: '#fff', border: 'none', borderRadius: 8, padding: '7px 12px', fontWeight: 700, fontSize: 12, cursor: 'pointer', flexShrink: 0 }}>
                      Pay
                    </button>
                  )}
                </div>

                {isOpen && (
                  <div style={{ background: '#F6F3EA', borderRadius: 10, padding: '10px 12px', marginBottom: 8 }}>
                    <div style={{ display: 'grid', gridTemplateColumns: '1fr 60px 52px 72px', gap: 6, marginBottom: 6 }}>
                      {['ITEM', 'QTY', 'UOM', 'RATE (₹)'].map((h) => <div key={h} style={{ fontSize: 9, fontWeight: 700, color: MUTED }}>{h}</div>)}
                    </div>
                    {rows.map((e) => {
                      const due = isDueEntry(e);
                      if (!due) {
                        const modeLabel = String(e.settledPayment || e.payment || '').toUpperCase();
                        return (
                          <div key={e.id} style={{ borderTop: `1px solid ${LINE}`, paddingTop: 8, marginTop: 4 }}>
                            <div style={{ display: 'grid', gridTemplateColumns: '1fr 60px 52px 72px', gap: 6, alignItems: 'center', fontSize: 12 }}>
                              <div style={{ fontWeight: 600 }}>{e.itemName}</div>
                              <div>{e.qty}</div>
                              <div style={{ fontSize: 11, fontWeight: 700, color: MUTED, textAlign: 'center' }}>{e.unit}</div>
                              <div>{e.unitPrice}</div>
                            </div>
                            <div style={{ textAlign: 'right', fontSize: 11, color: LEAF, marginTop: 2 }}>
                              {money(e.total)} · Paid{modeLabel ? ` (${modeLabel})` : ''}{e.settledDate ? ` on ${formatLedgerDate(e.settledDate, true)}` : ''}
                            </div>
                          </div>
                        );
                      }
                      const d = draftEdits[e.id] || {};
                      const effQty = d.qty !== undefined ? d.qty : String(e.qty);
                      const effPrice = d.unitPrice !== undefined ? d.unitPrice : String(e.unitPrice);
                      const eff = getEffective(e);
                      const changed = d.qty !== undefined || d.unitPrice !== undefined;
                      return (
                        <div key={e.id} style={{ borderTop: `1px solid ${LINE}`, paddingTop: 8, marginTop: 4 }}>
                          <div style={{ display: 'grid', gridTemplateColumns: '1fr 60px 52px 72px', gap: 6, alignItems: 'center' }}>
                            <div style={{ fontWeight: 600, fontSize: 12 }}>{e.itemName}</div>
                            <input type="number" value={effQty} onChange={(ev) => updateDraft(e.id, 'qty', ev.target.value)} style={{ border: `1px solid ${changed ? AMBER : LINE}`, borderRadius: 6, padding: '5px 6px', fontSize: 12, background: changed ? '#FFFBF3' : '#fff', width: '100%', boxSizing: 'border-box' }} />
                            <div style={{ fontSize: 11, fontWeight: 700, color: MUTED, textAlign: 'center' }}>{e.unit}</div>
                            <input type="number" value={effPrice} onChange={(ev) => updateDraft(e.id, 'unitPrice', ev.target.value)} style={{ border: `1px solid ${changed ? AMBER : LINE}`, borderRadius: 6, padding: '5px 6px', fontSize: 12, background: changed ? '#FFFBF3' : '#fff', width: '100%', boxSizing: 'border-box' }} />
                          </div>
                          <div style={{ textAlign: 'right', fontSize: 11, color: changed ? AMBER : MUTED, marginTop: 2 }}>
                            Total: {money(eff.total)}{changed ? ' ✏️' : ''}
                          </div>
                        </div>
                      );
                    })}
                    {hasDue && (
                      <div style={{ display: 'flex', justifyContent: 'flex-end', marginTop: 8, paddingTop: 8, borderTop: `1px solid ${LINE}` }}>
                        <span style={{ fontWeight: 800, color: AMBER, fontSize: 13 }}>Revised: {money(dueAmt)}</span>
                      </div>
                    )}
                  </div>
                )}
              </div>
            );
          })}
          {groups.length === 0 && <div style={hint}>{ledgerFilter === 'due' ? 'Nothing outstanding for this vendor.' : 'No ledger entries yet.'}</div>}
        </Card>

        <Card style={{ marginBottom: 12 }}>
          <div onClick={() => setShowItems((x) => !x)} style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', cursor: 'pointer' }}>
            <div style={{ ...sectionTitle, marginBottom: 0 }}>Linked items ({openVendor.itemIds.length})</div>
            <ChevronRight size={16} color={MUTED} style={{ transform: showItems ? 'rotate(90deg)' : 'none', transition: 'transform 0.15s' }} />
          </div>
          {showItems && (
            <div style={{ marginTop: 10 }}>
              <VendorItemLinkerMobile vendorId={openVendor.id} vendorItemIds={openVendor.itemIds} items={items} onToggle={onToggleItem} />
            </div>
          )}
        </Card>

        <div style={{ textAlign: 'center', marginBottom: 12 }}>
          {confirmDeleteId === openVendor.id ? (
            <div>
              <div style={{ fontSize: 12, color: TOMATO, marginBottom: 8 }}>
                Delete {openVendor.name}?{vendorDue > 0 ? ` ${money(vendorDue)} is still unpaid.` : ''}
              </div>
              <div style={{ display: 'flex', gap: 8, justifyContent: 'center' }}>
                <button onClick={() => { onDelete(openVendor.id); closeLedger(); }} style={{ background: TOMATO, color: '#fff', border: 'none', borderRadius: 6, padding: '6px 14px', fontSize: 12, fontWeight: 700, cursor: 'pointer' }}>Yes, delete</button>
                <button onClick={() => setConfirmDeleteId(null)} style={{ background: '#fff', color: INK, border: `1px solid ${LINE}`, borderRadius: 6, padding: '6px 14px', fontSize: 12, fontWeight: 700, cursor: 'pointer' }}>Cancel</button>
              </div>
            </div>
          ) : (
            <button onClick={() => setConfirmDeleteId(openVendor.id)} style={{ background: 'none', border: 'none', color: TOMATO, cursor: 'pointer', fontSize: 12, fontWeight: 700, display: 'inline-flex', alignItems: 'center', gap: 5 }}>
              <Trash2 size={13} /> Delete vendor
            </button>
          )}
        </div>

        {selectedDays > 0 && (
          <div style={{ position: 'sticky', bottom: 0, background: BG, paddingTop: 10, marginLeft: -16, marginRight: -16, paddingLeft: 16, paddingRight: 16, paddingBottom: 8 }}>
            <PrimaryBtn onClick={() => openPay(openVendor, selectedEntries)} color={TOMATO}>
              Pay {selectedDays} day{selectedDays !== 1 ? 's' : ''} together — {money(sumEffective(selectedEntries))}
            </PrimaryBtn>
          </div>
        )}

        {payModalEl}
      </div>
    );
  }

  // =====================  VENDOR LIST  =====================
  const searchTerm = vendorSearch.trim().toLowerCase();
  const shownVendors = vendors
    .filter((v) => !searchTerm || `${v.name} ${v.contact || ''}`.toLowerCase().includes(searchTerm))
    .map((v) => ({ v, due: dueTotalOf(v.id) }))
    .sort((a, b) => (Number(b.due > 0) - Number(a.due > 0)) || (b.due - a.due) || a.v.name.localeCompare(b.v.name));

  return (
    <div style={{ padding: 16 }}>
      <Card style={{ marginBottom: 14 }}>
        <div style={{ ...sectionTitle, display: 'flex', alignItems: 'center', gap: 6 }}><Store size={14} /> Add vendor</div>
        <div style={{ fontSize: 11, color: MUTED, marginBottom: 8 }}>You'll link items to each vendor after adding them.</div>
        <Field placeholder="Vendor name" value={name} onChange={(e) => setName(e.target.value)} />
        <Field placeholder="Phone / contact" value={contact} onChange={(e) => setContact(e.target.value)} />
        <PrimaryBtn onClick={submit}>Add vendor</PrimaryBtn>
      </Card>

      <Card>
        <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'baseline', marginBottom: 8 }}>
          <div style={{ ...sectionTitle, marginBottom: 0 }}>Vendors ({vendors.length})</div>
          {totalDue > 0 && <div style={{ fontSize: 12, fontWeight: 800, color: AMBER }}>Total due {money(totalDue)}</div>}
        </div>
        <Field placeholder="Search vendor…" value={vendorSearch} onChange={(e) => setVendorSearch(e.target.value)} style={{ marginBottom: 4 }} />
        {shownVendors.map(({ v, due }) => (
          <div key={v.id} onClick={() => openLedger(v.id)} style={{ borderTop: `1px solid ${LINE}`, padding: '12px 0', display: 'flex', alignItems: 'center', gap: 10, cursor: 'pointer' }}>
            <div style={{ flex: 1, minWidth: 0 }}>
              <div style={{ fontWeight: 700, fontSize: 14, whiteSpace: 'nowrap', overflow: 'hidden', textOverflow: 'ellipsis' }}>{v.name}</div>
              <div style={{ fontSize: 12, color: MUTED }}>{v.contact || '—'}</div>
            </div>
            <div style={{ textAlign: 'right', flexShrink: 0 }}>
              <div style={{ fontSize: 9, color: MUTED, fontWeight: 700 }}>OUTSTANDING</div>
              <div style={{ fontWeight: 800, fontSize: 15, color: due > 0 ? AMBER : MUTED }}>{money(due)}</div>
            </div>
            <button
              onClick={(e) => { e.stopPropagation(); openPay(v, dueEntriesOf(v.id)); }}
              disabled={due <= 0}
              style={{ background: due > 0 ? LEAF : '#C9C2AE', color: '#fff', border: 'none', borderRadius: 8, padding: '8px 14px', fontWeight: 700, fontSize: 12, cursor: due > 0 ? 'pointer' : 'default', flexShrink: 0 }}
            >
              Pay
            </button>
            <ChevronRight size={15} color={MUTED} style={{ flexShrink: 0 }} />
          </div>
        ))}
        {shownVendors.length === 0 && <div style={hint}>{vendors.length === 0 ? 'No vendors yet.' : 'No vendor matches your search.'}</div>}
      </Card>

      {payModalEl}

      {/* ---- Order Placed ---- */}
      {placedOrders.length > 0 && (
        <Card style={{ marginTop: 14 }}>
          <div style={{ fontWeight: 700, fontSize: 14, marginBottom: 12 }}>Order Placed ({placedOrders.length})</div>
          {placedOrders.map((order) => (
            <PlacedOrderCard key={order.id} order={order} onUpdate={onUpdatePlacedOrder} onDelete={onDeletePlacedOrder} />
          ))}
        </Card>
      )}
    </div>
  );
}

function AliasChip({ alias }) {
  return (
    <span style={{ display: 'inline-block', background: '#EAF3DE', color: LEAF_DARK, fontSize: 11, fontWeight: 700, padding: '3px 8px', borderRadius: 999, marginRight: 4, marginBottom: 4 }}>
      {alias.channel}{alias.code ? `: ${alias.code}` : ''}{alias.packSize ? ` (${alias.packSize}${alias.packUnit || ''})` : ''}
    </span>
  );
}

function AliasRowMobile({ alias, onChange, onRemove }) {
  return (
    <div style={{ border: `1px solid ${LINE}`, borderRadius: 8, padding: 8, marginBottom: 6 }}>
      <div style={{ display: 'flex', gap: 6, marginBottom: 6 }}>
        <Field placeholder="Channel (e.g. Blinkit)" value={alias.channel} onChange={(e) => onChange({ ...alias, channel: e.target.value })} style={{ flex: 1, marginBottom: 0 }} />
        <button onClick={onRemove} style={{ background: 'none', border: 'none', color: TOMATO, cursor: 'pointer', padding: 4 }}><Trash2 size={14} /></button>
      </div>
      <Field placeholder="Item code" value={alias.code} onChange={(e) => onChange({ ...alias, code: e.target.value })} style={{ marginBottom: 6 }} />
      <div style={{ display: 'flex', gap: 6, alignItems: 'center' }}>
        <Field placeholder="Pack size" type="number" value={alias.packSize} onChange={(e) => onChange({ ...alias, packSize: e.target.value })} style={{ flex: 1, marginBottom: 0 }} />
        <select value={alias.packUnit || 'kg'} onChange={(e) => onChange({ ...alias, packUnit: e.target.value })} style={{ borderRadius: 8, border: `1px solid ${LINE}`, fontSize: 13, padding: '8px 6px' }}>
          <option value="kg">kg</option>
          <option value="g">g</option>
          <option value="pieces">pieces</option>
          <option value="pack">pack</option>
        </select>
      </div>
    </div>
  );
}

function ItemForm({ initial, onSave, onCancel }) {
  const isEdit = !!initial;
  const [name, setName] = useState(initial?.name || '');
  const [uom, setUom] = useState(initial?.uom || 'kg');
  const [category, setCategory] = useState(initial?.category || 'VEGETABLES');
  const [aliases, setAliases] = useState((initial?.aliases || []).map((a) => ({ ...a })));

  const addAliasRow = () => setAliases((p) => [...p, { id: newAliasId(), channel: '', code: '', packSize: '', packUnit: 'kg' }]);
  const updateAliasRow = (id, next) => setAliases((p) => p.map((a) => (a.id === id ? next : a)));
  const removeAliasRow = (id) => setAliases((p) => p.filter((a) => a.id !== id));

  const canSave = name.trim();
  const submit = () => {
    if (!canSave) return;
    onSave({
      name: name.trim(),
      uom,
      category,
      aliases: aliases.filter((a) => a.channel.trim()).map((a) => ({ ...a, channel: a.channel.trim(), code: a.code.trim() })),
    });
  };

  return (
    <div style={{ padding: 16 }}>
      <button onClick={onCancel} style={{ display: 'flex', alignItems: 'center', gap: 6, background: 'none', border: 'none', color: LEAF, fontWeight: 700, fontSize: 13, cursor: 'pointer', marginBottom: 12, padding: 0 }}>
        <ArrowLeft size={15} /> Back to items
      </button>
      <Card>
        <div style={{ fontWeight: 800, fontSize: 16 }}>{isEdit ? `Edit ${initial.name}` : 'Create item'}</div>
        <div style={{ ...hint, marginTop: 4 }}>
          Items are generic — add an alias for each channel it's sold on (Blinkit, Flipkart, Zepto, etc.), with that channel's own item name/code and pack size.
        </div>

        <div style={smallLabel}>ITEM NAME</div>
        <Field placeholder="e.g. Tomato" value={name} onChange={(e) => setName(e.target.value)} />

        <div style={smallLabel}>UOM (unit it's purchased in)</div>
        <div style={{ display: 'flex', flexWrap: 'wrap', marginBottom: 6 }}>{UOM_OPTIONS.map((u) => <Chip key={u} label={u} active={uom === u} onClick={() => setUom(u)} />)}</div>

        <div style={smallLabel}>CATEGORY</div>
        <div style={{ display: 'flex', flexWrap: 'wrap', marginBottom: 10 }}>{CATEGORY_OPTIONS.map((c) => <Chip key={c} label={c} active={category === c} onClick={() => setCategory(c)} />)}</div>

        <div style={{ borderTop: `1px solid ${LINE}`, paddingTop: 12, marginTop: 4 }}>
          <div style={sectionTitle}>Channel aliases</div>
          <div style={hint}>e.g. Blinkit's 1kg Tomato pack, Flipkart's 500g pack, Zepto's 350g pack — each with its own channel item code.</div>
          {aliases.map((a) => (
            <AliasRowMobile key={a.id} alias={a} onChange={(next) => updateAliasRow(a.id, next)} onRemove={() => removeAliasRow(a.id)} />
          ))}
          <button onClick={addAliasRow} style={{ display: 'flex', alignItems: 'center', gap: 4, background: 'none', border: `1px dashed ${LINE}`, borderRadius: 8, padding: '8px 10px', fontSize: 12, color: MUTED, cursor: 'pointer', width: '100%', justifyContent: 'center', marginTop: 4 }}>
            <Plus size={12} /> Add alias
          </button>
        </div>

        <div style={{ display: 'flex', gap: 8, marginTop: 16 }}>
          <button
            onClick={submit}
            disabled={!canSave}
            style={{ flex: 1, background: !canSave ? '#C9C2AE' : LEAF, color: '#fff', border: 'none', borderRadius: 10, padding: '11px 0', fontWeight: 700, fontSize: 13, cursor: !canSave ? 'default' : 'pointer' }}
          >
            {isEdit ? 'Save changes' : 'Create item'}
          </button>
          <button onClick={onCancel} style={{ background: '#fff', color: INK, border: `1px solid ${LINE}`, borderRadius: 10, padding: '11px 16px', fontWeight: 700, fontSize: 13, cursor: 'pointer' }}>
            Cancel
          </button>
        </div>
      </Card>
    </div>
  );
}

function ItemsTab({ items, onAdd, onAddBulk, onUpdate, onDelete }) {
  const [view, setView] = useState('list'); // 'list' | 'form'
  const [editingItem, setEditingItem] = useState(null);
  const [search, setSearch] = useState('');
  const [categoryFilter, setCategoryFilter] = useState('ALL');
  const [confirmDeleteId, setConfirmDeleteId] = useState(null);
  const [bulkSummary, setBulkSummary] = useState(null);
  const [bulkError, setBulkError] = useState('');
  const bulkFileRef = useRef(null);

  const handleBulkFile = (e) => {
    const file = e.target.files[0];
    if (!file) return;
    setBulkError('');
    setBulkSummary(null);
    const reader = new FileReader();
    reader.onload = (evt) => {
      try {
        const wb = XLSX.read(evt.target.result, { type: 'array' });
        const sheet = wb.Sheets[wb.SheetNames[0]];
        const json = XLSX.utils.sheet_to_json(sheet, { defval: '' });
        const { valid, skipped } = parseBulkItemRows(json);
        if (valid.length > 0) onAddBulk(valid);
        setBulkSummary({ added: valid.length, skipped });
      } catch (err) {
        setBulkError('Could not read this file. Please upload the template format (.xlsx, .xls, or .csv).');
      }
    };
    reader.readAsArrayBuffer(file);
    e.target.value = '';
  };

  const openCreate = () => { setEditingItem(null); setView('form'); };
  const openEdit = (it) => { setEditingItem(it); setView('form'); };
  const closeForm = () => { setView('list'); setEditingItem(null); };
  const saveItem = (data) => {
    if (editingItem) onUpdate(editingItem.id, data);
    else onAdd({ id: `IT-${Date.now().toString(36).toUpperCase().slice(-5)}`, ...data });
    closeForm();
  };

  const categoryChips = ['ALL', ...CATEGORY_OPTIONS];
  const filteredItems = items.filter((it) => {
    const matchesCategory = categoryFilter === 'ALL' || it.category === categoryFilter;
    const q = search.trim().toLowerCase();
    const matchesSearch = !q || it.name.toLowerCase().includes(q) || it.id.toLowerCase().includes(q);
    return matchesCategory && matchesSearch;
  });

  if (view === 'form') {
    return <ItemForm initial={editingItem} onSave={saveItem} onCancel={closeForm} />;
  }

  return (
    <div style={{ padding: 16 }}>
      <Card style={{ marginBottom: 14 }}>
        <div style={{ display: 'flex', alignItems: 'center', gap: 6, background: BG, border: `1px solid ${LINE}`, borderRadius: 8, padding: '6px 10px', marginBottom: 10 }}>
          <Search size={14} color={MUTED} />
          <input placeholder="Search items..." value={search} onChange={(e) => setSearch(e.target.value)} style={{ border: 'none', background: 'transparent', outline: 'none', fontSize: 12, width: '100%' }} />
        </div>
        <div style={{ display: 'flex', flexWrap: 'wrap', marginBottom: 8 }}>
          {categoryChips.map((c) => <Chip key={c} label={c === 'ALL' ? 'All' : c} active={categoryFilter === c} onClick={() => setCategoryFilter(c)} />)}
        </div>
        <PrimaryBtn onClick={openCreate}>+ Create item</PrimaryBtn>
      </Card>

      <Card style={{ marginBottom: 14 }}>
        <div style={sectionTitle}>Bulk import</div>
        <div style={hint}>Add many items at once from a spreadsheet. Download the format first if you're not sure what columns to use.</div>
        <button
          onClick={downloadItemsTemplate}
          style={{ width: '100%', display: 'flex', alignItems: 'center', justifyContent: 'center', gap: 6, background: '#fff', color: LEAF, border: `1px solid ${LEAF}`, borderRadius: 10, padding: '10px 0', fontWeight: 700, fontSize: 13, cursor: 'pointer', marginBottom: 8 }}
        >
          <Download size={14} /> Download format
        </button>
        <button
          onClick={() => bulkFileRef.current?.click()}
          style={{ width: '100%', display: 'flex', alignItems: 'center', justifyContent: 'center', gap: 6, background: LEAF, color: '#fff', border: 'none', borderRadius: 10, padding: '10px 0', fontWeight: 700, fontSize: 13, cursor: 'pointer' }}
        >
          <Upload size={14} /> Bulk import items
        </button>
        <input ref={bulkFileRef} type="file" accept=".xlsx,.xls,.csv" onChange={handleBulkFile} style={{ display: 'none' }} />
        {bulkSummary && (
          <div style={{ marginTop: 10, fontSize: 12, color: LEAF, fontWeight: 600 }}>
            {bulkSummary.added} item{bulkSummary.added !== 1 ? 's' : ''} added{bulkSummary.skipped > 0 ? `, ${bulkSummary.skipped} skipped (missing name)` : ''}.
          </div>
        )}
        {bulkError && (
          <div style={{ display: 'flex', alignItems: 'center', gap: 6, marginTop: 10, fontSize: 12, color: TOMATO }}>
            <AlertCircle size={13} /> {bulkError}
          </div>
        )}
      </Card>

      {filteredItems.map((it) => (
        <Card key={it.id} style={{ marginBottom: 12 }}>
          <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center' }}>
            <div style={{ fontWeight: 800, fontSize: 15 }}>{it.name}</div>
            {confirmDeleteId === it.id ? (
              <div style={{ display: 'flex', gap: 6 }}>
                <button onClick={() => { onDelete(it.id); setConfirmDeleteId(null); }} style={{ background: TOMATO, color: '#fff', border: 'none', borderRadius: 6, padding: '4px 8px', fontSize: 11, fontWeight: 700, cursor: 'pointer' }}>Yes, delete</button>
                <button onClick={() => setConfirmDeleteId(null)} style={{ background: '#fff', color: INK, border: `1px solid ${LINE}`, borderRadius: 6, padding: '4px 8px', fontSize: 11, fontWeight: 700, cursor: 'pointer' }}>Cancel</button>
              </div>
            ) : (
              <div style={{ display: 'flex', gap: 12 }}>
                <button onClick={() => openEdit(it)} style={{ background: 'none', border: 'none', color: LEAF, cursor: 'pointer', display: 'flex' }}><Pencil size={15} /></button>
                <button onClick={() => setConfirmDeleteId(it.id)} style={{ background: 'none', border: 'none', color: TOMATO, cursor: 'pointer', display: 'flex' }}><Trash2 size={15} /></button>
              </div>
            )}
          </div>
          <div style={{ fontSize: 12, color: MUTED, marginTop: 2 }}>{it.id} · {it.uom} · {it.category}</div>
          <div style={{ marginTop: 8 }}>
            {(it.aliases && it.aliases.length > 0) ? (
              <div style={{ display: 'flex', flexWrap: 'wrap' }}>{it.aliases.map((a) => <AliasChip key={a.id} alias={a} />)}</div>
            ) : (
              <div style={{ fontSize: 11, color: MUTED }}>No aliases yet</div>
            )}
          </div>
        </Card>
      ))}
      {filteredItems.length === 0 && <div style={hint}>No items match this filter/search.</div>}
    </div>
  );
}

// ---------- Cut & Process ----------
function normalizeIngredientQty(qty, unit) { if (unit === 'g') return { value: qty / 1000, unit: 'kg' }; return { value: qty, unit }; }
const UNIT_OPTIONS = ['g', 'kg', 'piece'];

function CutProcessTab({ items, recipes, orders, onAddRecipe, onDeleteRecipe, onAddPurchaseRequirements }) {
  const [name, setName] = useState('');
  const [outputItemId, setOutputItemId] = useState('');
  const [ingredients, setIngredients] = useState([{ key: 'r0', itemId: '', qtyPerUnit: '', unit: 'g' }]);
  const itemName = (id) => items.find((it) => it.id === id)?.name || 'Unknown item';

  const addRow = () => setIngredients((p) => [...p, { key: `r${p.length}-${Date.now()}`, itemId: '', qtyPerUnit: '', unit: 'g' }]);
  const updateRow = (key, field, val) => setIngredients((p) => p.map((r) => (r.key === key ? { ...r, [field]: val } : r)));
  const removeRow = (key) => setIngredients((p) => p.filter((r) => r.key !== key));

  const saveRecipe = () => {
    const valid = ingredients.filter((r) => r.itemId && Number(r.qtyPerUnit) > 0);
    if (!name.trim() || !outputItemId || valid.length === 0) return;
    onAddRecipe({ id: `RCP-${Date.now().toString(36).toUpperCase().slice(-5)}`, name: name.trim(), outputItemId, ingredients: valid.map((r, i) => ({ id: `ing-${i}-${r.key}`, itemId: r.itemId, qtyPerUnit: Number(r.qtyPerUnit), unit: r.unit })) });
    setName(''); setOutputItemId(''); setIngredients([{ key: 'r0', itemId: '', qtyPerUnit: '', unit: 'g' }]);
  };

  const requirements = useMemo(() => recipes.map((recipe) => {
    const outputItem = items.find((it) => it.id === recipe.outputItemId);
    if (!outputItem) return { recipe, outputItem: null, totalQty: 0, rows: [] };
    const totalQty = orders.filter((o) => o.status !== 'dispatched' && o.product === outputItem.name).reduce((s, o) => s + o.qty, 0);
    const rows = recipe.ingredients.map((ing) => {
      const ingItem = items.find((it) => it.id === ing.itemId);
      const norm = normalizeIngredientQty(ing.qtyPerUnit * totalQty, ing.unit);
      return { ingredientName: ingItem?.name || 'Unknown', ...norm };
    });
    return { recipe, outputItem, totalQty, rows };
  }), [recipes, items, orders]);

  const pushToPurchase = (req) => {
    if (!req.totalQty) return;
    onAddPurchaseRequirements(req.rows.map((r, i) => ({ id: `P-REQ-${Date.now().toString(36).toUpperCase().slice(-4)}-${i}`, item: r.ingredientName, supplier: '', qty: r.value, unit: r.unit, cost: 0, source: `Recipe: ${req.recipe.name}` })));
  };

  return (
    <div style={{ padding: 16 }}>
      <Card style={{ marginBottom: 14 }}>
        <div style={sectionTitle}>Create recipe</div>
        <div style={hint}>How much of each item goes into one unit of a processed product.</div>
        <Field placeholder="Recipe name (e.g. Pulao Veggie Mix)" value={name} onChange={(e) => setName(e.target.value)} />
        <div style={smallLabel}>Output item (finished product)</div>
        <div style={{ display: 'flex', flexWrap: 'wrap', marginBottom: 8 }}>{items.map((it) => <Chip key={it.id} label={it.name} active={outputItemId === it.id} onClick={() => setOutputItemId(it.id)} />)}</div>
        <div style={smallLabel}>Ingredients (per 1 output unit)</div>
        {ingredients.map((row) => (
          <div key={row.key} style={{ border: `1px solid ${LINE}`, borderRadius: 10, padding: 10, marginBottom: 8 }}>
            <div style={{ display: 'flex', flexWrap: 'wrap', marginBottom: 6 }}>
              {items.filter((it) => it.id !== outputItemId).map((it) => <Chip key={it.id} label={it.name} active={row.itemId === it.id} onClick={() => updateRow(row.key, 'itemId', it.id)} />)}
            </div>
            <div style={{ display: 'flex', gap: 8, alignItems: 'center' }}>
              <Field placeholder="Qty" type="number" value={row.qtyPerUnit} onChange={(e) => updateRow(row.key, 'qtyPerUnit', e.target.value)} style={{ flex: 1, marginBottom: 0 }} />
              {UNIT_OPTIONS.map((u) => <Chip key={u} label={u} active={row.unit === u} onClick={() => updateRow(row.key, 'unit', u)} />)}
              {ingredients.length > 1 && <button onClick={() => removeRow(row.key)} style={{ background: 'none', border: 'none', color: TOMATO, cursor: 'pointer' }}>✕</button>}
            </div>
          </div>
        ))}
        <button onClick={addRow} style={{ width: '100%', border: `1px dashed ${LINE}`, background: 'none', borderRadius: 8, padding: '8px 0', color: MUTED, fontSize: 12, marginBottom: 10, cursor: 'pointer' }}>+ Add ingredient</button>
        <PrimaryBtn onClick={saveRecipe}>Save recipe</PrimaryBtn>
      </Card>

      <Card style={{ marginBottom: 14 }}>
        <div style={sectionTitle}>Recipes ({recipes.length})</div>
        {recipes.map((r) => (
          <div key={r.id} style={{ borderTop: `1px solid ${LINE}`, padding: '10px 0' }}>
            <div style={{ display: 'flex', justifyContent: 'space-between' }}>
              <div style={{ fontWeight: 700, fontSize: 13 }}>{r.name} → {itemName(r.outputItemId)}</div>
              <button onClick={() => onDeleteRecipe(r.id)} style={{ background: 'none', border: 'none', color: TOMATO, fontWeight: 700, fontSize: 12, cursor: 'pointer' }}>Delete</button>
            </div>
            <div style={{ fontSize: 12, color: MUTED, marginTop: 4 }}>{r.ingredients.map((ing) => `${ing.qtyPerUnit}${ing.unit} ${itemName(ing.itemId)}`).join(', ')}</div>
          </div>
        ))}
        {recipes.length === 0 && <div style={hint}>No recipes yet.</div>}
      </Card>

      <Card>
        <div style={sectionTitle}>Requirements from live orders</div>
        <div style={hint}>Based on pending + packed orders for each recipe's output item.</div>
        {requirements.map((req) => (
          <div key={req.recipe.id} style={{ borderTop: `1px solid ${LINE}`, padding: '10px 0' }}>
            <div style={{ fontWeight: 700, fontSize: 13, marginBottom: 6 }}>{req.recipe.name} — {req.totalQty} {req.outputItem?.uom || ''} ordered</div>
            {req.totalQty === 0 ? <div style={hint}>No open orders for this product right now.</div> : (
              <>
                {req.rows.map((r, i) => <div key={i} style={{ fontSize: 12 }}>{r.ingredientName}: <b style={{ color: LEAF }}>{r.value} {r.unit}</b></div>)}
                <button onClick={() => pushToPurchase(req)} style={{ width: '100%', background: TOMATO, color: '#fff', border: 'none', borderRadius: 8, padding: '8px 0', fontWeight: 700, fontSize: 12, marginTop: 8, cursor: 'pointer' }}>Add to purchase list</button>
              </>
            )}
          </div>
        ))}
      </Card>
    </div>
  );
}

// ---------- Orders ----------
function ReleaseBatchCard({ batch: b, orders, onToggleReleaseBatch }) {
  const [purchaseDate, setPurchaseDate] = useState(b.purchaseDate || '');

  const batchOrders = useMemo(() => orders.filter((o) => o.batchId === b.id), [orders, b.id]);
  const articleCount = new Set(batchOrders.map((o) => o.articleName || o.product)).size;
  const totalQty = batchOrders.reduce((s, o) => s + (Number(o.packQty) || 0), 0);
  const fulfilmentDate = batchOrders[0]?.fulfilmentDate || '';

  return (
    <div style={{ borderTop: `1px solid ${LINE}`, padding: '10px 0' }}>
      <div style={{ fontWeight: 700, fontSize: 13, marginBottom: 8 }}>{b.platform} indent — {b.fileName}</div>
      <div style={{ display: 'flex', gap: 12, marginBottom: 10 }}>
        <div>
          <div style={{ fontSize: 9, color: MUTED, fontWeight: 700 }}>ARTICLE QTY</div>
          <div style={{ fontWeight: 700, fontSize: 13 }}>{articleCount}</div>
        </div>
        <div>
          <div style={{ fontSize: 9, color: MUTED, fontWeight: 700 }}>TOTAL QTY</div>
          <div style={{ fontWeight: 700, fontSize: 13 }}>{totalQty}</div>
        </div>
        <div>
          <div style={{ fontSize: 9, color: MUTED, fontWeight: 700 }}>FULFILMENT DATE</div>
          <div style={{ fontWeight: 700, fontSize: 13 }}>{fulfilmentDate || '—'}</div>
        </div>
      </div>
      {b.released && b.purchaseDate && (
        <div style={{ fontSize: 11, color: LEAF, fontWeight: 700, marginBottom: 6 }}>Purchase date: {b.purchaseDate}</div>
      )}
      {!b.released && (
        <>
          <div style={smallLabel}>Release date (required)</div>
          <Field type="date" value={purchaseDate} onChange={(e) => setPurchaseDate(e.target.value)} style={{ marginBottom: 8 }} />
        </>
      )}
      <button
        onClick={() => onToggleReleaseBatch(b.id, purchaseDate)}
        disabled={!b.released && !purchaseDate}
        style={{ width: '100%', background: b.released ? '#fff' : (!purchaseDate ? '#C9C2AE' : TOMATO), color: b.released ? TOMATO : '#fff', border: b.released ? `1px solid ${TOMATO}` : 'none', borderRadius: 8, padding: '9px 0', fontSize: 12, fontWeight: 700, cursor: (!b.released && !purchaseDate) ? 'default' : 'pointer' }}
      >
        {b.released ? 'Withdraw from Purchase Manager' : 'Release to Purchase Manager'}
      </button>
    </div>
  );
}

function OrdersTab({ orders, items, indentBatches, onImport, onAddItem, onEnsureAlias, onUpdateAlias, onCreateIndentBatch, onToggleReleaseBatch }) {
  const [platform, setPlatform] = useState('Blinkit');
  const [product, setProduct] = useState('');
  const [qty, setQty] = useState('');
  const [unit, setUnit] = useState('kg');
  const [fulfilmentDate, setFulfilmentDate] = useState('');
  const [indentPlatform, setIndentPlatform] = useState('Blinkit');
  const [indentFulfilmentDate, setIndentFulfilmentDate] = useState('');
  const [pendingIndent, setPendingIndent] = useState(null);
  const [selectedRowKeys, setSelectedRowKeys] = useState(new Set());
  const [fileError, setFileError] = useState('');
  const fileInputRef = useRef(null);

  const submit = () => {
    if (!product.trim() || !qty || Number(qty) <= 0 || !fulfilmentDate) return;
    onImport({ id: `${platform.slice(0, 3).toUpperCase()}-${Math.floor(1000 + Math.random() * 9000)}`, platform, product: product.trim(), qty: Number(qty), unit, status: 'pending', fulfilmentDate });
    setProduct(''); setQty(''); setFulfilmentDate('');
  };

  const handleFile = (e) => {
    const file = e.target.files[0];
    if (!file) return;
    if (!indentFulfilmentDate) {
      setFileError('Please set the fulfilment date before uploading an indent.');
      e.target.value = '';
      return;
    }
    setFileError('');
    const reader = new FileReader();
    reader.onload = (evt) => {
      try {
        const wb = XLSX.read(evt.target.result, { type: 'array' });
        const sheet = wb.Sheets[wb.SheetNames[0]];
        const json = XLSX.utils.sheet_to_json(sheet, { defval: '' });
        const rawRows = parseIndentRows(json);
        if (rawRows.length === 0) { setFileError('No article rows with a valid name and quantity were found.'); return; }
        const rows = rawRows.map((r) => {
          const match = items.find((it) => (r.rawCode && (it.aliases || []).some((a) => a.channel === indentPlatform && a.code && a.code.toLowerCase() === r.rawCode.toLowerCase())) || it.name.toLowerCase() === r.rawName.toLowerCase());
          // Each distinct article code gets its own alias — even when it shares a base
          // item with another article on the same channel (e.g. two different pack sizes).
          if (match) onEnsureAlias(match.id, indentPlatform, r.rawCode);
          return { ...r, mappedItemId: match ? match.id : '' };
        });
        setPendingIndent({ platform: indentPlatform, fileName: file.name, rows, fulfilmentDate: indentFulfilmentDate });
        setSelectedRowKeys(new Set(rows.map((r) => r.key))); // select all by default
      } catch (err) {
        setFileError('Could not read this file. Please upload a valid .xlsx, .xls, or .csv indent.');
      }
    };
    reader.readAsArrayBuffer(file);
    e.target.value = '';
  };

  const toggleRowSelected = (key) => {
    setSelectedRowKeys((prev) => {
      const next = new Set(prev);
      if (next.has(key)) next.delete(key);
      else next.add(key);
      return next;
    });
  };
  const selectAllRows = () => setSelectedRowKeys(new Set((pendingIndent?.rows || []).map((r) => r.key)));
  const clearAllRows = () => setSelectedRowKeys(new Set());

  const setRowMapping = (key, value) => {
    setPendingIndent((prev) => ({
      ...prev,
      rows: prev.rows.map((r) => {
        if (r.key !== key) return r;
        if (value === '__new__') {
          const newItem = { id: `IT-${Date.now().toString(36).toUpperCase().slice(-5)}`, name: r.rawName, uom: r.unit || 'kg', category: normalizeCategory(r.rawCategory), aliases: [{ id: newAliasId(), channel: prev.platform, code: r.rawCode || '', packSize: '', packUnit: 'kg' }] };
          onAddItem(newItem);
          return { ...r, mappedItemId: newItem.id };
        }
        onEnsureAlias(value, prev.platform, r.rawCode);
        return { ...r, mappedItemId: value };
      }),
    }));
  };

  const getMappedItem = (id) => items.find((it) => it.id === id);
  // Two different articles (different codes) can map to the same item on the same
  // channel with different pack sizes — so pack size is looked up per-row, matched by
  // this row's own article code, not just by item+channel.
  const getRowAlias = (r) => {
    const item = getMappedItem(r.mappedItemId);
    if (!item) return null;
    const byCode = (item.aliases || []).find((a) => a.channel === pendingIndent?.platform && a.code && r.rawCode && a.code.toLowerCase() === r.rawCode.toLowerCase());
    return byCode || (item.aliases || []).find((a) => a.channel === pendingIndent?.platform) || null;
  };
  const getPackSize = (r) => getRowAlias(r)?.packSize || '';
  const isRowReady = (r) => !!r.mappedItemId && Number(getPackSize(r)) > 0;
  const readyCount = pendingIndent ? pendingIndent.rows.filter(isRowReady).length : 0;
  const importCount = pendingIndent ? pendingIndent.rows.filter((r) => isRowReady(r) && selectedRowKeys.has(r.key)).length : 0;

  const importMapped = () => {
    if (!pendingIndent) return;
    const remaining = [];
    const compiledMap = {};
    const batchId = `BATCH-${Date.now().toString(36).toUpperCase().slice(-6)}`;
    const orderIdBase = Date.now().toString(36).toUpperCase().slice(-5);
    let orderSeq = 0;
    pendingIndent.rows.forEach((r) => {
      if (!isRowReady(r) || !selectedRowKeys.has(r.key)) { remaining.push(r); return; }
      const item = items.find((it) => it.id === r.mappedItemId);
      if (!item) { remaining.push(r); return; }
      const alias = getRowAlias(r);
      const packSize = Number(alias?.packSize) || 1;
      const packUnit = alias?.packUnit || item.uom;
      const key = `${item.name}__${item.uom}`;
      if (!compiledMap[key]) compiledMap[key] = { itemName: item.name, unit: item.uom, qty: 0 };
      // One order per store when the indent has per-store columns (Flipkart); otherwise a
      // single order for the platform's default store.
      const storeEntries = r.storeQtys && Object.keys(r.storeQtys).length > 0
        ? Object.entries(r.storeQtys)
        : [[PLATFORM_DEFAULT_STORE[pendingIndent.platform] || '', r.qty]];
      storeEntries.forEach(([store, storePacks]) => {
        if (!(storePacks > 0)) return;
        const storeQty = Math.round(storePacks * packSize * 100) / 100;
        orderSeq += 1;
        onImport({
          id: `${pendingIndent.platform.slice(0, 3).toUpperCase()}-${orderIdBase}${orderSeq}`,
          platform: pendingIndent.platform,
          store,
          product: item.name,
          articleName: r.rawName,
          qty: storeQty,
          unit: item.uom,
          status: 'pending',
          fulfilmentDate: pendingIndent.fulfilmentDate || '',
          packQty: storePacks,
          packSize,
          packUnit,
          batchId,
        });
        compiledMap[key].qty += storeQty;
      });
    });
    const compiled = Object.values(compiledMap);
    if (compiled.length > 0) onCreateIndentBatch({ id: batchId, platform: pendingIndent.platform, fileName: pendingIndent.fileName, compiled, released: false, purchaseRowIds: [] });
    if (!remaining.length) setIndentFulfilmentDate('');
    setPendingIndent(remaining.length ? { ...pendingIndent, rows: remaining } : null);
    setSelectedRowKeys(new Set(remaining.filter((r) => selectedRowKeys.has(r.key)).map((r) => r.key)));
  };

  return (
    <div style={{ padding: 16 }}>
      {indentBatches.length > 0 && (
        <Card style={{ marginBottom: 14 }}>
          <div style={sectionTitle}>Release to Purchase Manager</div>
          {indentBatches.map((b) => <ReleaseBatchCard key={b.id} batch={b} orders={orders} onToggleReleaseBatch={onToggleReleaseBatch} />)}
        </Card>
      )}

      <Card style={{ marginBottom: 14 }}>
        <div style={{ ...sectionTitle, display: 'flex', alignItems: 'center', gap: 6 }}><FileSpreadsheet size={14} /> Import indent (Excel)</div>
        <div style={hint}>Upload the Blinkit or Flipkart indent file. You can upload another one (e.g. a different date) even while one is still being mapped below — it replaces whatever's unfinished in the table.</div>
        <div style={{ display: 'flex', marginBottom: 8 }}>{PLATFORMS.map((p) => <Chip key={p} label={p} active={indentPlatform === p} onClick={() => setIndentPlatform(p)} />)}</div>
        <div style={smallLabel}>Fulfilment date (required)</div>
        <Field type="date" value={indentFulfilmentDate} onChange={(e) => setIndentFulfilmentDate(e.target.value)} />
        <PrimaryBtn onClick={() => fileInputRef.current?.click()} disabled={!indentFulfilmentDate}>Upload {indentPlatform} indent</PrimaryBtn>
        <input ref={fileInputRef} type="file" accept=".xlsx,.xls,.csv" onChange={handleFile} style={{ display: 'none' }} />
        {!indentFulfilmentDate && <div style={{ display: 'flex', alignItems: 'center', gap: 6, fontSize: 11, color: AMBER, marginTop: 8 }}><AlertCircle size={12} /> Fulfilment date is required before you can upload.</div>}
        {fileError && <div style={{ display: 'flex', alignItems: 'center', gap: 6, fontSize: 12, color: TOMATO, marginTop: 8 }}><AlertCircle size={13} /> {fileError}</div>}

        {pendingIndent && (
          <div style={{ borderTop: `1px solid ${LINE}`, marginTop: 14, paddingTop: 14 }}>
            <div style={{ display: 'flex', justifyContent: 'space-between', marginBottom: 4 }}>
              <div style={{ fontSize: 11, color: MUTED, flex: 1 }}>{pendingIndent.fileName} · {pendingIndent.platform} · {pendingIndent.rows.length} found, {readyCount} ready{pendingIndent.fulfilmentDate ? ` · ${pendingIndent.fulfilmentDate}` : ''}</div>
              <button onClick={() => setPendingIndent(null)} style={{ background: 'none', border: 'none', color: TOMATO, fontSize: 12, fontWeight: 700, cursor: 'pointer' }}>Cancel</button>
            </div>
            <div style={{ display: 'flex', gap: 14, marginBottom: 8 }}>
              <button onClick={selectAllRows} style={{ background: 'none', border: 'none', color: LEAF, fontSize: 12, fontWeight: 700, cursor: 'pointer', padding: 0 }}>Select all</button>
              <button onClick={clearAllRows} style={{ background: 'none', border: 'none', color: MUTED, fontSize: 12, fontWeight: 700, cursor: 'pointer', padding: 0 }}>Clear</button>
              <span style={{ fontSize: 12, color: MUTED }}>{selectedRowKeys.size} selected</span>
            </div>
            {pendingIndent.rows.map((r) => {
              const mappedItem = getMappedItem(r.mappedItemId);
              const rowAlias = getRowAlias(r);
              const packSize = rowAlias?.packSize || '';
              return (
                <div key={r.key} style={{ borderTop: `1px solid ${LINE}`, padding: '10px 0', background: selectedRowKeys.has(r.key) ? '#F6F3EA' : 'transparent' }}>
                  <div style={{ display: 'flex', alignItems: 'flex-start', gap: 8 }}>
                    <input type="checkbox" checked={selectedRowKeys.has(r.key)} onChange={() => toggleRowSelected(r.key)} style={{ marginTop: 3 }} />
                    <div style={{ flex: 1 }}>
                      <div style={{ fontWeight: 700, fontSize: 13 }}>{r.rawName}</div>
                      <div style={{ fontSize: 11, color: MUTED, margin: '2px 0 6px' }}>Qty {r.qty} · UOM {r.unit || '—'} · Code {r.rawCode || '—'} · {r.rawCategory || '—'}</div>
                      <div style={smallLabel}>Map to item</div>
                      <div style={{ display: 'flex', flexWrap: 'wrap' }}>
                        {items.map((it) => <Chip key={it.id} label={it.name} active={r.mappedItemId === it.id} onClick={() => setRowMapping(r.key, it.id)} />)}
                        <Chip label={`+ New "${r.rawName}"`} active={false} onClick={() => setRowMapping(r.key, '__new__')} />
                      </div>
                      {mappedItem && rowAlias && (
                        <>
                          <div style={smallLabel}>Pack size ({mappedItem.uom} per pack)</div>
                          <Field placeholder="e.g. 0.5" type="number" value={packSize} onChange={(e) => onUpdateAlias(mappedItem.id, rowAlias.id, { packSize: e.target.value })} style={{ width: 120 }} />
                        </>
                      )}
                    </div>
                  </div>
                </div>
              );
            })}
            {readyCount < pendingIndent.rows.length && <div style={{ fontSize: 11, color: AMBER, marginTop: 10, marginBottom: 10 }}>{pendingIndent.rows.length - readyCount} article(s) still need mapping and/or pack size.</div>}
            <PrimaryBtn onClick={importMapped} disabled={importCount === 0}>Import {importCount} selected &amp; ready order{importCount !== 1 ? 's' : ''}</PrimaryBtn>
          </div>
        )}
      </Card>

      <Card style={{ marginBottom: 14 }}>
        <div style={sectionTitle}>Add order manually</div>
        <div style={{ display: 'flex', marginBottom: 8 }}>{PLATFORMS.map((p) => <Chip key={p} label={p} active={platform === p} onClick={() => setPlatform(p)} />)}</div>
        <Field placeholder="Product" value={product} onChange={(e) => setProduct(e.target.value)} />
        <div style={smallLabel}>Fulfilment date (required)</div>
        <Field type="date" value={fulfilmentDate} onChange={(e) => setFulfilmentDate(e.target.value)} />
        <Field placeholder="Quantity" type="number" value={qty} onChange={(e) => setQty(e.target.value)} />
        <div style={{ display: 'flex', flexWrap: 'wrap', marginBottom: 10 }}>{['kg', 'dozen', 'bunch', 'crate'].map((u) => <Chip key={u} label={u} active={unit === u} onClick={() => setUnit(u)} />)}</div>
        {!fulfilmentDate && <div style={{ display: 'flex', alignItems: 'center', gap: 6, fontSize: 11, color: AMBER, marginBottom: 8 }}><AlertCircle size={12} /> Fulfilment date is required.</div>}
        <PrimaryBtn onClick={submit} disabled={!product.trim() || !qty || Number(qty) <= 0 || !fulfilmentDate}>Add order</PrimaryBtn>
      </Card>

      <OrdersListCard orders={orders} indentBatches={indentBatches} />
    </div>
  );
}

function OrderBatchGroupMobile({ label, subtitle, badgeText, badgeColor, orders: groupOrders, defaultOpen }) {
  const [open, setOpen] = useState(!!defaultOpen);
  return (
    <div style={{ border: `1px solid ${LINE}`, borderRadius: 10, marginBottom: 10, overflow: 'hidden' }}>
      <div onClick={() => setOpen((x) => !x)} style={{ padding: '10px 12px', cursor: 'pointer', background: open ? '#F6F3EA' : '#fff' }}>
        <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'flex-start', gap: 8 }}>
          <div style={{ flex: 1, minWidth: 0 }}>
            <div style={{ fontWeight: 700, fontSize: 13 }}>{label}</div>
            {subtitle && <div style={{ fontSize: 10.5, color: MUTED, marginTop: 1 }}>{subtitle}</div>}
          </div>
          <ChevronRight size={15} color={MUTED} style={{ transform: open ? 'rotate(90deg)' : 'none', flexShrink: 0, marginTop: 2 }} />
        </div>
        <div style={{ display: 'flex', gap: 6, marginTop: 6 }}>
          {badgeText && <span style={{ background: badgeColor === 'blue' ? '#E6F1FB' : '#FBEFDC', color: badgeColor === 'blue' ? '#1B5E8C' : AMBER, fontSize: 10, fontWeight: 700, padding: '2px 8px', borderRadius: 999 }}>{badgeText}</span>}
          <span style={{ background: '#EAF3DE', color: LEAF_DARK, fontWeight: 800, fontSize: 10, padding: '2px 8px', borderRadius: 999 }}>{groupOrders.length} order{groupOrders.length !== 1 ? 's' : ''}</span>
        </div>
      </div>
      {open && groupOrders.map((o) => (
        <div key={o.id} style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', borderTop: `1px solid ${LINE}`, padding: '8px 12px' }}>
          <div style={{ flex: 1 }}>
            <div style={{ fontWeight: 700, fontSize: 13 }}>{o.id} · {o.platform}{orderStore(o) ? ` · ${storeLabel(orderStore(o))}` : ''}</div>
            <div style={{ fontSize: 12, color: MUTED }}>{o.articleName || o.product} · {o.qty} {o.unit}{o.fulfilmentDate ? ` · due ${o.fulfilmentDate}` : ''}</div>
          </div>
          <StatusPill status={o.status} />
        </div>
      ))}
    </div>
  );
}

function OrdersListCard({ orders, indentBatches }) {
  const grouped = useMemo(() => {
    const byBatch = {};
    const manual = [];
    orders.forEach((o) => {
      if (o.batchId) {
        byBatch[o.batchId] = byBatch[o.batchId] || [];
        byBatch[o.batchId].push(o);
      } else {
        manual.push(o);
      }
    });
    const batchGroups = indentBatches
      .filter((b) => byBatch[b.id]?.length)
      .map((b) => ({ batch: b, orders: byBatch[b.id] }));
    const knownBatchIds = new Set(indentBatches.map((b) => b.id));
    const orphaned = Object.entries(byBatch).filter(([id]) => !knownBatchIds.has(id)).flatMap(([, os]) => os);
    return { batchGroups, manual: [...manual, ...orphaned] };
  }, [orders, indentBatches]);

  return (
    <Card>
      <div style={sectionTitle}>All orders ({orders.length})</div>
      {grouped.batchGroups.map(({ batch, orders: groupOrders }) => (
        <OrderBatchGroupMobile
          key={batch.id}
          label={`${batch.platform} indent — ${batch.fileName}`}
          subtitle={batch.released ? `Released${batch.purchaseDate ? ` · ${batch.purchaseDate}` : ''}` : 'Not yet released'}
          badgeText={batch.released ? 'Released' : 'Not released'}
          badgeColor={batch.released ? 'blue' : 'amber'}
          orders={groupOrders}
        />
      ))}
      {grouped.manual.length > 0 && (
        <OrderBatchGroupMobile label="Manually added orders" orders={grouped.manual} defaultOpen={grouped.batchGroups.length === 0} />
      )}
      {orders.length === 0 && <div style={hint}>No orders yet.</div>}
    </Card>
  );
}

// ---------- Purchases ----------
// CUT is intentionally not listed: processed (CUT) items are never bought directly — only their raw ingredients are.
const PURCHASE_CATEGORY_OPTIONS = ['ALL', 'FRUITS', 'VEGETABLES', 'FLOWER', 'EXOTIC', 'GRAINS'];

function PurchasesTab({ purchases, orders, items, allItems, recipes, vendors, vendorLedger, stockCounts, onAddLedgerEntry, onSavePlacedOrder, indentBatches }) {
  const [categoryFilter, setCategoryFilter] = usePersistedState('fnv_purchase_category', 'ALL');
  const [vendorFilterId, setVendorFilterId] = usePersistedState('fnv_purchase_vendor', '');
  const [qtySort, setQtySort] = usePersistedState('fnv_purchase_qtysort', 'none'); // 'none' | 'asc' | 'desc'
  const [fulfilmentDateFilter, setFulfilmentDateFilter] = usePersistedState('fnv_purchase_fulfilmentdate', 'ALL'); // 'ALL' = All Purchase
  const [bufferPercent, setBufferPercent] = useState('0');
  const [purchasedDate, setPurchasedDate] = useState('');
  const [selectedItemId, setSelectedItemId] = useState(null);
  const [selectedVendorId, setSelectedVendorId] = useState('');
  const [view, setView] = useState('list'); // 'list' | 'purchased'

  // Multi-select / order sharing
  const [selectMode, setSelectMode] = useState(false);
  const [selectedItemIds, setSelectedItemIds] = useState([]);
  const [orderNameDraft, setOrderNameDraft] = useState('');
  const [showOrderNameModal, setShowOrderNameModal] = useState(false);

  const toggleSelectItem = (id) => setSelectedItemIds((p) => p.includes(id) ? p.filter((x) => x !== id) : [...p, id]);
  const exitSelectMode = () => { setSelectMode(false); setSelectedItemIds([]); };

  // Purchase form state
  const [purchaseQty, setPurchaseQty] = useState('');
  const [unitPrice, setUnitPrice] = useState('');
  const [totalInput, setTotalInput] = useState('');
  const [paymentMode, setPaymentMode] = useState('credit');
  const [purchaseNote, setPurchaseNote] = useState('');
  const [purchaseDate] = useState(() => new Date().toISOString().split('T')[0]);
  const [purchaseSuccess, setPurchaseSuccess] = useState(false);

  const [showAllVendorItems, setShowAllVendorItems] = useState(false);
  const openItem = (id, keepVendor = false) => {
    setSelectedItemId(id);
    if (!keepVendor) setSelectedVendorId('');
    setShowAllVendorItems(false);
    setPurchaseQty('');
    setUnitPrice('');
    setTotalInput('');
    setPaymentMode('cash');
    setPurchaseNote('');
    setPurchaseSuccess(false);
  };

  const derivedTotal = purchaseQty && unitPrice ? Math.round(Number(purchaseQty) * Number(unitPrice) * 100) / 100 : null;
  const derivedUnitPrice = purchaseQty && totalInput && !unitPrice ? Math.round(Number(totalInput) / Number(purchaseQty) * 100) / 100 : null;
  const totalPrice = derivedTotal ?? (totalInput ? Number(totalInput) : 0);
  const finalUnitPrice = unitPrice ? Number(unitPrice) : (derivedUnitPrice ?? 0);

  const handleQtyChange = (v) => { setPurchaseQty(v); };
  const handleUnitPriceChange = (v) => { setUnitPrice(v); if (v && purchaseQty) setTotalInput(''); };
  const handleTotalChange = (v) => { setTotalInput(v); if (v && purchaseQty) setUnitPrice(''); };

  const canSubmit = purchaseQty && (unitPrice || (totalInput && purchaseQty)) && selectedVendorId;

  const submitPurchase = () => {
    if (!canSubmit) return;
    const it = items.find((x) => x.id === selectedItemId);
    const vendor = vendors.find((v) => v.id === selectedVendorId);
    const unit = neededByProduct[it?.name]?.unit || it?.uom;
    const entry = {
      id: `LED-${Date.now().toString(36).toUpperCase().slice(-6)}`,
      vendorId: selectedVendorId,
      vendorName: vendor?.name || '',
      itemId: selectedItemId,
      itemName: it?.name || '',
      qty: Number(purchaseQty),
      unit,
      unitPrice: finalUnitPrice,
      total: totalPrice,
      payment: paymentMode,
      date: purchaseDate,
      note: purchaseNote.trim(),
      settled: paymentMode !== 'credit',
    };
    onAddLedgerEntry(entry);
    setPurchaseQty(''); setUnitPrice(''); setTotalInput(''); setPurchaseNote(''); setPaymentMode('cash');
    setPurchaseSuccess(true);
    setTimeout(() => setPurchaseSuccess(false), 3000);
  };

  // Output items (Cut & Process finished products) never get purchased
  // directly — their recipe's raw ingredients do. Demand for an output
  // item is expanded into ingredient demand, compiled across every
  // recipe/order that needs that same ingredient.
  const availableFulfilmentDates = useMemo(() => {
    const releasedBatchIds = new Set(indentBatches.filter((b) => b.released).map((b) => b.id));
    const dates = new Set();
    orders
      .filter((o) => o.status !== 'dispatched')
      .filter((o) => !o.batchId || releasedBatchIds.has(o.batchId))
      .forEach((o) => { if (o.fulfilmentDate) dates.add(o.fulfilmentDate); });
    return Array.from(dates).sort();
  }, [orders, indentBatches]);

  // Demand per raw/buyable item. Orders for a processed (CUT) item are never added directly:
  // they are broken down into their recipe ingredients (recursively, so a recipe can use
  // another CUT item). CUT items that have no recipe yet are collected in `cutWithoutRecipe`
  // so they can be flagged instead of silently landing in the purchase list.
  const neededData = useMemo(() => {
    const map = {};
    const missing = {};
    const nrm = (s) => String(s || '').trim().toUpperCase().replace(/\s+/g, ' ');
    const cityIds = new Set(items.map((it) => it.id));
    // Items are per-city but recipes are shared, so resolve recipe items across ALL cities.
    const itemById = {};
    const catByName = {};
    (allItems || []).forEach((it) => { itemById[it.id] = it; catByName[nrm(it.name)] = it.category; });
    items.forEach((it) => { itemById[it.id] = it; catByName[nrm(it.name)] = it.category; });

    const recipesByOutput = {};
    recipes.forEach((r) => {
      const out = itemById[r.outputItemId];
      if (!out) return;
      const k = nrm(out.name);
      (recipesByOutput[k] = recipesByOutput[k] || []).push(r);
    });
    const pickRecipes = (name) => {
      const all = recipesByOutput[nrm(name)] || [];
      if (all.length === 0) return [];
      // Prefer recipes made for this city's own item; otherwise reuse one other city's recipe (never sum duplicates across cities).
      const local = all.filter((r) => cityIds.has(r.outputItemId));
      if (local.length > 0) return local;
      return all.filter((r) => r.outputItemId === all[0].outputItemId);
    };

    const addDemand = (name, qty, unit) => {
      map[name] = map[name] || { needed: 0, unit };
      map[name].needed += qty;
    };
    const explode = (name, qty, unit, depth) => {
      const recs = depth < 5 ? pickRecipes(name) : [];
      if (recs.length > 0) {
        recs.forEach((recipe) => {
          (recipe.ingredients || []).forEach((ing) => {
            const ingItem = itemById[ing.itemId];
            if (!ingItem) return;
            const norm = normalizeIngredientQty(ing.qtyPerUnit * qty, ing.unit);
            explode(ingItem.name, norm.value, norm.unit, depth + 1);
          });
        });
        return;
      }
      if (catByName[nrm(name)] === 'CUT') {
        missing[name] = missing[name] || { qty: 0, unit };
        missing[name].qty += qty;
        return;
      }
      addDemand(name, qty, unit);
    };

    // An order counts toward "needing purchase" once it's actually been released to
    // Purchase Manager — orders with no batch (added manually) always count, since
    // there's no release step for those.
    const releasedBatchIds = new Set(indentBatches.filter((b) => b.released).map((b) => b.id));
    orders
      .filter((o) => o.status !== 'dispatched')
      .filter((o) => !o.batchId || releasedBatchIds.has(o.batchId))
      .filter((o) => fulfilmentDateFilter === 'ALL' || o.fulfilmentDate === fulfilmentDateFilter)
      .forEach((o) => explode(o.product, o.qty, o.unit, 0));
    return { map, missing };
  }, [orders, recipes, items, allItems, indentBatches, fulfilmentDateFilter]);
  const neededByProduct = neededData.map;
  const cutWithoutRecipe = neededData.missing;

  // "Available stock" = latest nightly stock count (if any) as baseline, plus every
  // actual completed purchase made since — "requirement" rows (from released indents /
  // recipe pushes) are just a to-buy queue, not stock on hand.
  const stockByItem = useMemo(() => {
    const map = {};
    const latestCount = {};
    (stockCounts || []).forEach((sc) => {
      if (!latestCount[sc.itemName] || sc.date > latestCount[sc.itemName].date) {
        latestCount[sc.itemName] = { date: sc.date, qty: sc.closingQty };
      }
    });
    Object.entries(latestCount).forEach(([name, c]) => { map[name] = c.qty; });
    purchases
      .filter((p) => p.type !== 'requirement')
      .forEach((p) => {
        const lc = latestCount[p.item];
        if (!lc || !p.date || p.date > lc.date) {
          map[p.item] = (map[p.item] || 0) + p.qty;
        }
      });
    return map;
  }, [purchases, stockCounts]);

  const filteredItems = useMemo(() => {
    const buffer = Number(bufferPercent) || 0;
    const vendorItemIds = vendorFilterId ? new Set(vendors.find((v) => v.id === vendorFilterId)?.itemIds || []) : null;
    let result = items
      .filter((it) => neededByProduct[it.name])
      .filter((it) => it.category !== 'CUT')
      .filter((it) => categoryFilter === 'ALL' || categoryFilter === 'CUT' || it.category === categoryFilter)
      .filter((it) => !vendorItemIds || vendorItemIds.has(it.id))
      .map((it) => {
        const needed = neededByProduct[it.name].needed;
        const unit = neededByProduct[it.name].unit;
        const stock = stockByItem[it.name] || 0;
        const toBuy = Math.max(0, Math.round((needed - stock) * 100) / 100);
        return { ...it, needed, unit, stock, toBuy };
      })
      // Already sufficiently stocked (stock beats needed by more than the buffer %) — no need to buy.
      .filter((it) => it.stock <= it.needed * (1 + buffer / 100));
    if (qtySort === 'asc') result = result.slice().sort((a, b) => a.toBuy - b.toBuy);
    else if (qtySort === 'desc') result = result.slice().sort((a, b) => b.toBuy - a.toBuy);
    return result;
  }, [items, neededByProduct, categoryFilter, vendorFilterId, vendors, stockByItem, bufferPercent, qtySort]);

  const purchasedList = useMemo(() => {
    return purchases
      .filter((p) => p.type !== 'requirement')
      .filter((p) => !purchasedDate || p.date === purchasedDate)
      .slice()
      .sort((a, b) => (b.date || '').localeCompare(a.date || ''));
  }, [purchases, purchasedDate]);

  if (selectedItemId) {
    const it = items.find((x) => x.id === selectedItemId);
    const needed = neededByProduct[it?.name]?.needed || 0;
    const unit = neededByProduct[it?.name]?.unit || it?.uom;
    const stock = stockByItem[it?.name] || 0;
    const toBuy = Math.max(0, Math.round((needed - stock) * 100) / 100);
    return (
      <div style={{ padding: 16 }}>
        <button onClick={() => setSelectedItemId(null)} style={{ display: 'flex', alignItems: 'center', gap: 6, background: 'none', border: 'none', color: LEAF, fontWeight: 700, fontSize: 13, cursor: 'pointer', marginBottom: 12, padding: 0 }}>
          <ArrowLeft size={15} /> Back to items
        </button>
        <Card>
          <div style={{ fontWeight: 800, fontSize: 16 }}>{it?.name}</div>
          <div style={{ fontSize: 12, color: MUTED, marginTop: 2 }}>{it?.category} · {it?.id}</div>
          <div style={{ display: 'flex', gap: 8, marginTop: 12 }}>
            <div style={{ flex: 1, border: `1px solid ${LINE}`, borderRadius: 8, padding: '8px 10px' }}>
              <div style={{ fontSize: 10, color: MUTED, fontWeight: 700 }}>NEEDED</div>
              <div style={{ fontSize: 14, fontWeight: 800, color: INK }}>{needed}</div>
              <div style={{ fontSize: 10, color: MUTED, marginTop: 1 }}>{unit}</div>
            </div>
            <div style={{ flex: 1, border: `1px solid ${LINE}`, borderRadius: 8, padding: '8px 10px' }}>
              <div style={{ fontSize: 10, color: MUTED, fontWeight: 700 }}>STOCK</div>
              <div style={{ fontSize: 14, fontWeight: 800, color: INK }}>{stock}</div>
              <div style={{ fontSize: 10, color: MUTED, marginTop: 1 }}>{unit}</div>
            </div>
            <div style={{ flex: 1, border: `1px solid ${TOMATO}`, background: '#FBEAE3', borderRadius: 8, padding: '8px 10px' }}>
              <div style={{ fontSize: 10, color: TOMATO, fontWeight: 700 }}>TO BUY</div>
              <div style={{ fontSize: 14, fontWeight: 800, color: TOMATO }}>{toBuy}</div>
              <div style={{ fontSize: 10, color: TOMATO, marginTop: 1 }}>{unit}</div>
            </div>
          </div>
          <div style={{ marginTop: 16 }}>
            <div style={smallLabel}>Select vendor</div>
            <select
              value={selectedVendorId}
              onChange={(e) => { setSelectedVendorId(e.target.value); setShowAllVendorItems(false); }}
              style={{ width: '100%', boxSizing: 'border-box', padding: '9px 8px', borderRadius: 8, border: `1px solid ${LINE}`, fontSize: 13, color: INK, background: '#fff' }}
            >
              <option value="">Choose a vendor</option>
              {vendors.map((v) => (
                <option key={v.id} value={v.id}>{v.name}{v.itemIds.includes(it?.id) ? ' (supplies this item)' : ''}</option>
              ))}
            </select>
          </div>

          {selectedVendorId ? (
            (() => {
              const vendor = vendors.find((v) => v.id === selectedVendorId);
              const allVendorItems = vendor ? items.filter((x) => vendor.itemIds.includes(x.id)) : [];
              // Sort by purchase frequency for this vendor (most purchases first), exclude current item
              const purchaseCount = (itemName) => vendorLedger.filter((e) => e.vendorId === selectedVendorId && e.itemName === itemName).length;
              const sorted = [...allVendorItems].sort((a, b) => {
                if (a.id === it?.id) return -1;
                if (b.id === it?.id) return 1;
                return purchaseCount(b.name) - purchaseCount(a.name);
              });
              const SHOW_DEFAULT = 3;
              const showMore = sorted.length > SHOW_DEFAULT;
              const displayed = showAllVendorItems ? sorted : sorted.slice(0, SHOW_DEFAULT);
              return (
                <div style={{ marginTop: 14 }}>
                  <div style={smallLabel}>{vendor?.name} also supplies</div>
                  {allVendorItems.length === 0 && <div style={hint}>No items linked to this vendor yet — link some in the Vendors section.</div>}
                  {displayed.map((vi) => {
                    const viNeeded = neededByProduct[vi.name]?.needed;
                    const viCount = purchaseCount(vi.name);
                    return (
                      <div key={vi.id} onClick={() => openItem(vi.id, true)} style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', borderTop: `1px solid ${LINE}`, padding: '8px 0', cursor: 'pointer' }}>
                        <div>
                          <div style={{ fontWeight: vi.id === it?.id ? 800 : 600, fontSize: 13, color: vi.id === it?.id ? LEAF : INK }}>{vi.name}{vi.id === it?.id ? ' (current)' : ''}</div>
                          {viCount > 0 && <div style={{ fontSize: 10, color: MUTED }}>{viCount} purchase{viCount !== 1 ? 's' : ''}</div>}
                        </div>
                        <div style={{ fontSize: 11, color: viNeeded ? TOMATO : MUTED, fontWeight: 700, textAlign: 'right' }}>
                          {viNeeded ? `${viNeeded} ${neededByProduct[vi.name].unit} needed` : 'No demand'}
                        </div>
                      </div>
                    );
                  })}
                  {showMore && (
                    <button
                      onClick={() => setShowAllVendorItems((e) => !e)}
                      style={{ background: 'none', border: 'none', color: LEAF, fontSize: 12, fontWeight: 700, padding: '6px 0', cursor: 'pointer', width: '100%', textAlign: 'center', borderTop: `1px solid ${LINE}` }}
                    >
                      {showAllVendorItems ? 'Show less ▲' : `Show ${sorted.length - SHOW_DEFAULT} more ▼`}
                    </button>
                  )}
                </div>
              );
            })()
          ) : null}

          {selectedVendorId && (
            <div style={{ marginTop: 16, borderTop: `1px solid ${LINE}`, paddingTop: 14 }}>
              <div style={{ fontWeight: 700, fontSize: 13, color: INK, marginBottom: 10 }}>Record purchase</div>

              {purchaseSuccess && (
                <div style={{ background: '#EAF3DE', color: LEAF_DARK, borderRadius: 8, padding: '9px 12px', fontSize: 12, fontWeight: 700, marginBottom: 10 }}>
                  ✓ Purchase recorded successfully
                </div>
              )}

              <div style={{ display: 'flex', gap: 8, marginBottom: 8 }}>
                <div style={{ flex: 1 }}>
                  <div style={smallLabel}>QTY ({neededByProduct[it?.name]?.unit || it?.uom})</div>
                  <Field placeholder="e.g. 50" type="number" value={purchaseQty} onChange={(e) => handleQtyChange(e.target.value)} style={{ marginBottom: 0 }} />
                </div>
                <div style={{ flex: 1 }}>
                  <div style={{ ...smallLabel, display: 'flex', justifyContent: 'space-between' }}>
                    <span>UNIT PRICE (₹)</span>
                    {derivedUnitPrice !== null && !unitPrice && <span style={{ color: LEAF, fontSize: 10 }}>auto</span>}
                  </div>
                  <Field
                    placeholder={derivedUnitPrice !== null && !unitPrice ? String(derivedUnitPrice) : 'e.g. 30'}
                    type="number"
                    value={unitPrice}
                    onChange={(e) => handleUnitPriceChange(e.target.value)}
                    style={{ marginBottom: 0, borderColor: derivedUnitPrice !== null && !unitPrice ? LEAF : LINE }}
                  />
                </div>
              </div>

              <div style={{ marginBottom: 10 }}>
                <div style={{ ...smallLabel, display: 'flex', justifyContent: 'space-between' }}>
                  <span>TOTAL AMOUNT (₹)</span>
                  {derivedTotal !== null && !totalInput && <span style={{ color: LEAF, fontSize: 10 }}>auto</span>}
                </div>
                <Field
                  placeholder={derivedTotal !== null ? String(derivedTotal) : 'Enter total or fill unit price'}
                  type="number"
                  value={totalInput}
                  onChange={(e) => handleTotalChange(e.target.value)}
                  style={{ marginBottom: 0, fontWeight: 700, fontSize: 15, borderColor: derivedTotal !== null && !totalInput ? LEAF : LINE }}
                />
                {totalPrice > 0 && (
                  <div style={{ background: BG, borderRadius: 8, padding: '8px 12px', marginTop: 6, display: 'flex', justifyContent: 'space-between', alignItems: 'center' }}>
                    <span style={{ fontSize: 12, color: MUTED }}>Confirmed total</span>
                    <span style={{ fontWeight: 800, fontSize: 15, color: INK }}>₹{totalPrice.toLocaleString('en-IN')}</span>
                  </div>
                )}
              </div>

              <div style={smallLabel}>PAYMENT MODE</div>
              <div style={{ display: 'flex', gap: 6, marginBottom: 10 }}>
                {[{ key: 'cash', label: '💵 Cash' }, { key: 'upi', label: '📱 UPI' }, { key: 'bank', label: '🏦 Bank' }, { key: 'credit', label: '📒 Credit' }].map((m) => (
                  <button
                    key={m.key}
                    onClick={() => setPaymentMode(m.key)}
                    style={{ flex: 1, padding: '7px 4px', borderRadius: 8, border: `1px solid ${paymentMode === m.key ? (m.key === 'credit' ? AMBER : LEAF) : LINE}`, background: paymentMode === m.key ? (m.key === 'credit' ? '#FBEFDC' : '#EAF3DE') : '#fff', color: paymentMode === m.key ? (m.key === 'credit' ? AMBER : LEAF_DARK) : INK, fontSize: 10, fontWeight: 700, cursor: 'pointer', textAlign: 'center' }}
                  >
                    {m.label}
                  </button>
                ))}
              </div>

              {paymentMode === 'credit' && (
                <div style={{ background: '#FBEFDC', borderRadius: 8, padding: '10px 12px', marginBottom: 10 }}>
                  <div style={{ fontWeight: 700, fontSize: 12, color: AMBER, marginBottom: 4 }}>📒 Credit entry</div>
                  <div style={{ fontSize: 11, color: AMBER }}>₹{totalPrice.toLocaleString('en-IN')} will be added to {vendors.find((v) => v.id === selectedVendorId)?.name || 'vendor'}'s account as outstanding credit.</div>
                </div>
              )}

              <Field placeholder="Note (optional)" value={purchaseNote} onChange={(e) => setPurchaseNote(e.target.value)} />

              <PrimaryBtn
                onClick={submitPurchase}
                disabled={!canSubmit}
                color={paymentMode === 'credit' ? AMBER : LEAF}
              >
                {paymentMode === 'credit' ? 'Record on credit' : 'Record purchase'}
              </PrimaryBtn>
            </div>
          )}

          {/* Vendor credit ledger for this item's vendor */}
          {selectedVendorId && (() => {
            const creditEntries = vendorLedger.filter((e) => e.vendorId === selectedVendorId && e.payment === 'credit' && !e.settled);
            const creditTotal = creditEntries.reduce((s, e) => s + e.total, 0);
            if (creditEntries.length === 0) return null;
            return (
              <div style={{ marginTop: 14 }}>
                <div style={{ display: 'flex', justifyContent: 'space-between', marginBottom: 6 }}>
                  <div style={{ fontWeight: 700, fontSize: 13, color: AMBER }}>Outstanding credit</div>
                  <div style={{ fontWeight: 800, color: AMBER }}>₹{creditTotal.toLocaleString('en-IN')}</div>
                </div>
                {creditEntries.slice(0, 5).map((e) => (
                  <div key={e.id} style={{ borderTop: `1px solid ${LINE}`, padding: '6px 0', fontSize: 12 }}>
                    <div style={{ display: 'flex', justifyContent: 'space-between' }}>
                      <span style={{ fontWeight: 600 }}>{e.itemName} · {e.qty} {e.unit}</span>
                      <span style={{ color: AMBER, fontWeight: 700 }}>₹{e.total.toLocaleString('en-IN')}</span>
                    </div>
                    <div style={{ color: MUTED, fontSize: 11 }}>{e.date}{e.note ? ` · ${e.note}` : ''}</div>
                  </div>
                ))}
              </div>
            );
          })()}

        </Card>
      </div>
    );
  }

  const hasActiveFilters = (categoryFilter !== 'ALL' && categoryFilter !== 'CUT') || Number(bufferPercent) !== 0 || !!vendorFilterId || qtySort !== 'none' || fulfilmentDateFilter !== 'ALL';
  const clearFilters = () => { setCategoryFilter('ALL'); setBufferPercent('0'); setVendorFilterId(''); setQtySort('none'); setFulfilmentDateFilter('ALL'); };

  const confirmShareOrder = () => {
    const orderItems = filteredItems
      .filter((it) => selectedItemIds.includes(it.id))
      .map((it, idx) => ({ no: idx + 1, itemId: it.id, itemName: it.name, uom: it.unit, qty: it.toBuy }));
    onSavePlacedOrder({ id: `ORD-${Date.now().toString(36).toUpperCase().slice(-6)}`, name: orderNameDraft.trim() || `Order ${new Date().toLocaleDateString('en-IN')}`, date: new Date().toISOString().split('T')[0], items: orderItems });
    setShowOrderNameModal(false);
    setOrderNameDraft('');
    exitSelectMode();
  };

  const allPurchasedCount = purchases.filter((p) => p.type !== 'requirement').length;

  if (view === 'purchased') {
    return (
      <div style={{ padding: 16 }}>
        <button onClick={() => setView('list')} style={{ display: 'flex', alignItems: 'center', gap: 6, background: 'none', border: 'none', color: LEAF, fontWeight: 700, fontSize: 13, cursor: 'pointer', marginBottom: 12, padding: 0 }}>
          <ArrowLeft size={15} /> Back
        </button>
        <Card>
          <div style={sectionTitle}>Purchased{purchasedDate ? ` on ${purchasedDate}` : ''} ({purchasedList.length})</div>
          <div style={{ display: 'flex', alignItems: 'center', gap: 8, marginTop: 8, marginBottom: 10 }}>
            <Field type="date" value={purchasedDate} onChange={(e) => setPurchasedDate(e.target.value)} style={{ marginBottom: 0, flex: 1 }} />
            {purchasedDate && (
              <button onClick={() => setPurchasedDate('')} style={{ background: 'none', border: 'none', color: TOMATO, fontSize: 12, fontWeight: 700, cursor: 'pointer' }}>Clear</button>
            )}
          </div>
          {purchasedList.map((p) => (
            <div key={p.id} style={{ borderTop: `1px solid ${LINE}`, padding: '8px 0' }}>
              <div style={{ fontWeight: 700, fontSize: 13 }}>{p.item}</div>
              <div style={{ fontSize: 11, color: MUTED }}>{p.date || '—'} · {p.supplier || 'No supplier'} · {p.qty} {p.unit || 'kg'} · ₹{p.cost.toLocaleString('en-IN')}</div>
            </div>
          ))}
          {purchasedList.length === 0 && <div style={hint}>{purchasedDate ? 'Nothing purchased on this date.' : 'No purchases recorded yet.'}</div>}
        </Card>
      </div>
    );
  }

  return (
    <div style={{ padding: 16 }}>
      <Card style={{ marginBottom: 12 }}>
        <div style={{ display: 'flex', gap: 6 }}>
          <select value={vendorFilterId} onChange={(e) => setVendorFilterId(e.target.value)} style={{ flex: '1 1 0', minWidth: 0, boxSizing: 'border-box', padding: '6px 4px', borderRadius: 6, border: `1px solid ${LINE}`, fontSize: 11, color: INK, background: '#fff' }}>
            <option value="">All vendors</option>
            {vendors.map((v) => <option key={v.id} value={v.id}>{v.name}</option>)}
          </select>
          <select value={PURCHASE_CATEGORY_OPTIONS.includes(categoryFilter) ? categoryFilter : 'ALL'} onChange={(e) => setCategoryFilter(e.target.value)} style={{ flex: '1 1 0', minWidth: 0, boxSizing: 'border-box', padding: '6px 4px', borderRadius: 6, border: `1px solid ${LINE}`, fontSize: 11, color: INK, background: '#fff' }}>
            {PURCHASE_CATEGORY_OPTIONS.map((c) => <option key={c} value={c}>{c}</option>)}
          </select>
        </div>
        <div style={{ display: 'flex', gap: 6, marginTop: 6 }}>
          <select value={qtySort} onChange={(e) => setQtySort(e.target.value)} style={{ flex: '1 1 0', minWidth: 0, boxSizing: 'border-box', padding: '6px 4px', borderRadius: 6, border: `1px solid ${LINE}`, fontSize: 11, color: INK, background: '#fff' }}>
            <option value="none">Default sort</option>
            <option value="asc">Qty: Low-High</option>
            <option value="desc">Qty: High-Low</option>
          </select>
          <select value={fulfilmentDateFilter} onChange={(e) => setFulfilmentDateFilter(e.target.value)} style={{ flex: '1 1 0', minWidth: 0, boxSizing: 'border-box', padding: '6px 4px', borderRadius: 6, border: `1px solid ${LINE}`, fontSize: 11, color: INK, background: '#fff' }}>
            <option value="ALL">All Purchase</option>
            {availableFulfilmentDates.map((d) => <option key={d} value={d}>{d}</option>)}
          </select>
        </div>
        <div style={{ display: 'flex', gap: 8, marginTop: 8, alignItems: 'flex-end' }}>
          <div style={{ flex: 1 }}>
            <span style={{ fontSize: 11, color: MUTED }}>Buffer %</span>
            <Field type="number" placeholder="0" value={bufferPercent} onChange={(e) => setBufferPercent(e.target.value)} style={{ marginBottom: 0 }} />
          </div>
        </div>
        {hasActiveFilters && (
          <button onClick={clearFilters} style={{ background: 'none', border: 'none', color: TOMATO, fontSize: 12, fontWeight: 700, cursor: 'pointer', marginTop: 8, padding: 0 }}>Clear filters</button>
        )}
        <button
          onClick={() => setView('purchased')}
          style={{ width: '100%', display: 'flex', alignItems: 'center', justifyContent: 'center', gap: 6, background: LEAF, color: '#fff', border: 'none', borderRadius: 10, padding: '10px 0', fontWeight: 700, fontSize: 13, cursor: 'pointer', marginTop: 10 }}
        >
          Purchased ({allPurchasedCount})
        </button>
      </Card>

      <Card>
        {Object.keys(cutWithoutRecipe).length > 0 && (
          <div style={{ background: '#FFF4D6', border: '1px solid #E8C766', color: '#6B4E00', borderRadius: 8, padding: '8px 10px', fontSize: 11, marginBottom: 10 }}>
            <strong>Recipe missing:</strong> these CUT items have orders but no recipe, so their raw material is not in this list — {Object.entries(cutWithoutRecipe).map(([n, v]) => `${n} (${Math.round(v.qty * 100) / 100} ${v.unit})`).join(', ')}. Add a recipe in Cut &amp; Process.
          </div>
        )}
        <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', marginBottom: 8 }}>
          <div style={sectionTitle}>Items ({filteredItems.length})</div>
          {selectMode ? (
            <button onClick={exitSelectMode} style={{ background: 'none', border: 'none', color: MUTED, fontSize: 12, fontWeight: 700, cursor: 'pointer' }}>Cancel</button>
          ) : (
            <button onClick={() => setSelectMode(true)} style={{ background: 'none', border: 'none', color: LEAF, fontSize: 12, fontWeight: 700, cursor: 'pointer' }}>Select</button>
          )}
        </div>

        {selectMode && selectedItemIds.length > 0 && (
          <button
            onClick={() => { setOrderNameDraft(`Order ${new Date().toLocaleDateString('en-IN')}`); setShowOrderNameModal(true); }}
            style={{ width: '100%', background: LEAF, color: '#fff', border: 'none', borderRadius: 10, padding: '10px 0', fontWeight: 700, fontSize: 13, cursor: 'pointer', marginBottom: 10 }}
          >
            Share order ({selectedItemIds.length} items)
          </button>
        )}

        {filteredItems.map((it) => {
          const isSelected = selectMode && selectedItemIds.includes(it.id);
          return (
            <div
              key={it.id}
              onClick={() => selectMode ? toggleSelectItem(it.id) : openItem(it.id)}
              style={{ display: 'flex', alignItems: 'center', justifyContent: 'space-between', gap: 8, borderTop: `1px solid ${LINE}`, cursor: 'pointer', background: isSelected ? '#EAF3DE' : 'transparent', borderRadius: isSelected ? 8 : 0, padding: '9px 4px' }}
            >
              {selectMode && (
                <div style={{ width: 20, height: 20, borderRadius: 5, border: `2px solid ${isSelected ? LEAF : LINE}`, background: isSelected ? LEAF : '#fff', flexShrink: 0, display: 'flex', alignItems: 'center', justifyContent: 'center' }}>
                  {isSelected && <div style={{ color: '#fff', fontSize: 12, fontWeight: 900 }}>✓</div>}
                </div>
              )}
              <div style={{ flex: 1, minWidth: 0 }}>
                <div style={{ fontWeight: 700, fontSize: 13, whiteSpace: 'nowrap', overflow: 'hidden', textOverflow: 'ellipsis' }}>{it.name}</div>
                <div style={{ fontSize: 10, color: MUTED, marginTop: 1 }}>{it.category}</div>
              </div>
              <div style={{ display: 'flex', alignItems: 'center', gap: 6, flexShrink: 0 }}>
                <div style={{ border: `1px solid ${LINE}`, borderRadius: 6, padding: '4px 6px', textAlign: 'center', minWidth: 52 }}>
                  <div style={{ fontSize: 8, color: MUTED, fontWeight: 700 }}>STOCK</div>
                  <div style={{ fontSize: 11, fontWeight: 700, color: INK }}>{it.stock}</div>
                  <div style={{ fontSize: 8, color: MUTED }}>{it.unit}</div>
                </div>
                <div style={{ border: `1px solid ${TOMATO}`, background: '#FBEAE3', borderRadius: 6, padding: '4px 6px', textAlign: 'center', minWidth: 52 }}>
                  <div style={{ fontSize: 8, color: TOMATO, fontWeight: 700 }}>TO BUY</div>
                  <div style={{ fontSize: 11, fontWeight: 800, color: TOMATO }}>{it.toBuy}</div>
                  <div style={{ fontSize: 8, color: TOMATO }}>{it.unit}</div>
                </div>
                {!selectMode && <ChevronRight size={14} color={MUTED} />}
              </div>
            </div>
          );
        })}
        {filteredItems.length === 0 && (
          <div style={{ textAlign: 'center', padding: '10px 0' }}>
            <div style={{ ...hint, marginBottom: hasActiveFilters ? 8 : 0 }}>No items match these filters.</div>
            {hasActiveFilters && (
              <button onClick={clearFilters} style={{ background: 'none', border: `1px solid ${LINE}`, borderRadius: 8, padding: '7px 14px', color: LEAF, fontSize: 12, fontWeight: 700, cursor: 'pointer' }}>
                Clear filters
              </button>
            )}
          </div>
        )}
      </Card>

      {/* Order name modal */}
      {showOrderNameModal && (
        <div style={{ position: 'fixed', inset: 0, background: 'rgba(0,0,0,0.45)', display: 'flex', alignItems: 'flex-end', justifyContent: 'center', zIndex: 200 }}>
          <div style={{ background: '#fff', borderRadius: '18px 18px 0 0', padding: '24px 20px 32px', width: '100%', maxWidth: 420 }}>
            <div style={{ fontWeight: 800, fontSize: 16, marginBottom: 6 }}>Name this order</div>
            <div style={{ fontSize: 12, color: MUTED, marginBottom: 14 }}>This will be saved to Vendors → Order Placed where you can edit and share it.</div>
            <Field placeholder="Order name (e.g. Morning Order 14 Sep)" value={orderNameDraft} onChange={(e) => setOrderNameDraft(e.target.value)} />
            <div style={{ fontSize: 12, color: MUTED, marginBottom: 14 }}>
              {selectedItemIds.length} item{selectedItemIds.length !== 1 ? 's' : ''}: {filteredItems.filter((it) => selectedItemIds.includes(it.id)).map((it) => `${it.name} (${it.toBuy} ${it.unit})`).join(', ')}
            </div>
            <div style={{ display: 'flex', gap: 8 }}>
              <button onClick={() => setShowOrderNameModal(false)} style={{ flex: 1, background: '#fff', color: INK, border: `1px solid ${LINE}`, borderRadius: 10, padding: '11px 0', fontWeight: 700, fontSize: 13, cursor: 'pointer' }}>Cancel</button>
              <button onClick={confirmShareOrder} style={{ flex: 2, background: LEAF, color: '#fff', border: 'none', borderRadius: 10, padding: '11px 0', fontWeight: 700, fontSize: 13, cursor: 'pointer' }}>Save to Order Placed</button>
            </div>
          </div>
        </div>
      )}
    </div>
  );
}
function StockCountRow({ item, existingCount, lastKnown, unit, onSave }) {
  const [value, setValue] = useState(existingCount !== undefined ? String(existingCount) : '');
  useEffect(() => { setValue(existingCount !== undefined ? String(existingCount) : ''); }, [existingCount]);

  return (
    <div style={{ borderTop: `1px solid ${LINE}`, padding: '10px 0' }}>
      <div style={{ display: 'flex', justifyContent: 'space-between' }}>
        <div style={{ fontWeight: 700, fontSize: 13 }}>{item.name}</div>
        <div style={{ fontSize: 11, color: MUTED }}>{item.category}</div>
      </div>
      <div style={{ fontSize: 11, color: MUTED, marginTop: 2 }}>
        {lastKnown ? `Last known: ${lastKnown.closingQty} ${unit} (${lastKnown.date})` : 'Never counted'}
      </div>
      <div style={{ display: 'flex', alignItems: 'center', gap: 6, marginTop: 6 }}>
        <input
          type="number"
          placeholder="Closing qty"
          value={value}
          onChange={(e) => setValue(e.target.value)}
          style={{ flex: 1, boxSizing: 'border-box', borderRadius: 6, border: `1px solid ${existingCount !== undefined ? LEAF : LINE}`, fontSize: 13, padding: '7px 8px' }}
        />
        <span style={{ fontSize: 12, color: MUTED }}>{unit}</span>
        <button
          onClick={() => onSave(value)}
          disabled={value === ''}
          style={{ background: value === '' ? '#C9C2AE' : LEAF, color: '#fff', border: 'none', borderRadius: 6, padding: '7px 12px', fontSize: 12, fontWeight: 700, cursor: value === '' ? 'default' : 'pointer' }}
        >
          Save
        </button>
      </div>
    </div>
  );
}

function StockCountTab({ items, stockCounts, onRecord }) {
  const [date, setDate] = useState(new Date().toISOString().split('T')[0]);
  const [categoryFilter, setCategoryFilter] = useState('ALL');
  const [search, setSearch] = useState('');

  const countsForDate = useMemo(() => {
    const map = {};
    stockCounts.filter((sc) => sc.date === date).forEach((sc) => { map[sc.itemId] = sc.closingQty; });
    return map;
  }, [stockCounts, date]);

  const latestCountByItem = useMemo(() => {
    const map = {};
    stockCounts.forEach((sc) => { if (!map[sc.itemId] || sc.date > map[sc.itemId].date) map[sc.itemId] = sc; });
    return map;
  }, [stockCounts]);

  const filteredItems = items
    .filter((it) => categoryFilter === 'ALL' || it.category === categoryFilter)
    .filter((it) => !search.trim() || it.name.toLowerCase().includes(search.trim().toLowerCase()));

  const countedToday = filteredItems.filter((it) => countsForDate[it.id] !== undefined).length;

  // Items already counted for this date sink to the bottom; uncounted ones stay on top.
  const sortedItems = [
    ...filteredItems.filter((it) => countsForDate[it.id] === undefined),
    ...filteredItems.filter((it) => countsForDate[it.id] !== undefined),
  ];

  return (
    <div style={{ padding: 16 }}>
      <Card style={{ marginBottom: 12 }}>
        <div style={{ ...sectionTitle, display: 'flex', alignItems: 'center', gap: 6 }}><Layers size={15} /> Nightly stock count</div>
        <div style={hint}>Record the actual closing stock for each item at the end of the day. This becomes the new stock baseline.</div>
        <div style={smallLabel}>DATE</div>
        <Field type="date" value={date} onChange={(e) => setDate(e.target.value)} />
        <div style={smallLabel}>CATEGORY</div>
        <div style={{ display: 'flex', flexWrap: 'wrap', marginBottom: 8 }}>
          {['ALL', ...CATEGORY_OPTIONS].map((c) => <Chip key={c} label={c === 'ALL' ? 'All' : c} active={categoryFilter === c} onClick={() => setCategoryFilter(c)} />)}
        </div>
        <Field placeholder="Search items..." value={search} onChange={(e) => setSearch(e.target.value)} style={{ marginBottom: 0 }} />
        <div style={{ fontSize: 11, color: MUTED, marginTop: 8 }}>{countedToday} / {filteredItems.length} counted for {date}</div>
      </Card>

      <Card>
        {sortedItems.map((it) => (
          <StockCountRow
            key={it.id}
            item={it}
            unit={it.uom}
            existingCount={countsForDate[it.id]}
            lastKnown={latestCountByItem[it.id]}
            onSave={(val) => onRecord(it.id, it.name, it.uom, date, val)}
          />
        ))}
        {filteredItems.length === 0 && <div style={hint}>No items match this filter.</div>}
      </Card>
    </div>
  );
}

function computeFinalPrice(basePrice, config) {
  if (basePrice == null) return null;
  const gradingPercent = config?.gradingPercent ?? 0;
  const vendorMarginPercent = config?.vendorMarginPercent ?? 0;
  const packaging = config?.packaging ?? 0;
  const labour = config?.labour ?? 0;
  const transportation = config?.transportation ?? 0;
  const graded = basePrice * (1 + gradingPercent / 100);
  return Math.round((graded * (1 + vendorMarginPercent / 100) + packaging + labour + transportation) * 100) / 100;
}

function buildLatestUnitPriceByItem(purchases) {
  const map = {};
  purchases
    .filter((p) => p.type !== 'requirement' && p.qty > 0)
    .forEach((p) => {
      if (!map[p.item] || (p.date || '') >= (map[p.item].date || '')) {
        map[p.item] = { date: p.date || '', unitPrice: p.cost / p.qty };
      }
    });
  return map;
}

// One entry per distinct article that has come through an indent — same product can have
// several pack sizes, each priced separately. Shared by Pricing and Profit & Loss tabs.
function buildPricingArticles(orders, items, purchases) {
  const latestUnitPriceByItem = buildLatestUnitPriceByItem(purchases);
  const map = {};
  orders
    .filter((o) => o.packSize && o.packUnit)
    .forEach((o) => {
      const key = `${o.product}__${o.platform}__${o.packSize}__${o.packUnit}`;
      if (map[key]) return;
      const item = items.find((it) => it.name === o.product);
      const unitPriceInfo = latestUnitPriceByItem[o.product];
      const basePrice = unitPriceInfo ? Math.round(unitPriceInfo.unitPrice * o.packSize * 100) / 100 : null;
      const alias = (item?.aliases || []).find((al) => al.channel === o.platform && String(al.packSize) === String(o.packSize) && al.packUnit === o.packUnit);
      map[key] = {
        key,
        articleName: o.articleName || o.product,
        product: o.product,
        category: item?.category || '',
        platform: o.platform,
        code: alias?.code || '',
        packSize: o.packSize,
        packUnit: o.packUnit,
        basePrice,
      };
    });
  return Object.values(map).sort((a, b) => a.articleName.localeCompare(b.articleName));
}

// One indent (batch) may have several articles that don't yet have a purchase price —
// those are simply left out of the running cost until they do (this is what makes the
// batch's total climb from "day one" partial toward a complete figure as purchases happen).
// Quantity marked short at packing time is subtracted from the pack count before costing
// it, so a shortfall we never actually bought or sent out doesn't get counted as spend.
function computeBatchArticleCosts(batch, orders, articlesByKey, configByKey) {
  const batchOrders = orders.filter((o) => o.batchId === batch.id);
  const rows = batchOrders.map((o) => {
    const key = `${o.product}__${o.platform}__${o.packSize}__${o.packUnit}`;
    const article = articlesByKey[key];
    const packSize = Number(o.packSize) || 1;
    const shortPacks = Math.min(Number(o.packQty) || 0, (Number(o.shortQty) || 0) / packSize);
    const effectivePacks = Math.max(0, Math.round(((Number(o.packQty) || 0) - shortPacks) * 100) / 100);
    const finalPricePerPack = article ? computeFinalPrice(article.basePrice, configByKey[key]) : null;
    const cost = finalPricePerPack == null ? null : Math.round(finalPricePerPack * effectivePacks * 100) / 100;
    return {
      orderId: o.id,
      articleName: o.articleName || o.product,
      code: article?.code || '',
      packQty: Number(o.packQty) || 0,
      shortPacks: Math.round(shortPacks * 100) / 100,
      effectivePacks,
      packSize: o.packSize,
      packUnit: o.packUnit,
      finalPricePerPack,
      cost,
    };
  });
  const pricedRows = rows.filter((r) => r.cost != null);
  const totalCost = Math.round(pricedRows.reduce((s, r) => s + r.cost, 0) * 100) / 100;
  return { rows, totalCost, pricedCount: pricedRows.length, totalCount: rows.length };
}

function parseGrnRows(json) {
  // Column-name order matters: Excel exports like Hyperpure's often have BOTH a
  // "PO" and a "GRN" version of quantity/rate (and "Product UPC" alongside
  // "Product Description") — the more specific "...GRN" / "...Description"
  // candidates must be checked before the generic ones, or a generic match
  // (e.g. "quantity") would grab the wrong column ("Quantity - PO") first.
  return json
    .map((r) => {
      const code = String(pickField(r, ['itemcode', 'code', 'sku', 'articlecode', 'fsn']) || '').trim();
      const name = String(pickField(r, ['productdescription', 'itemname', 'name', 'description', 'article', 'product']) || '').trim();
      const qty = Number(pickField(r, ['quantitygrn', 'grnqty', 'receivedqty', 'accepted', 'qty', 'quantity']) || 0);
      const price = Number(pickField(r, ['landingrategrn', 'grnlandingrate', 'rategrn', 'receivedprice', 'unitprice', 'unitrate', 'price', 'rate', 'landingrate']) || 0);
      return { code, name, qty, price };
    })
    .filter((r) => (r.code || r.name) && r.qty > 0);
}

// ── Hyperpure / Blinkit GRN report PDFs — parsed client-side via pdf.js ──
// Each article row in these PDFs follows a fixed column order once all the
// text is flattened onto one line: row# / item code / UPC / description /
// MRP / tax / landing rate (PO avg, then GRN) / qty (PO, then GRN) /
// fill rate% / total GRN amount / GMV loss. We only need the item code,
// description, GRN qty and GRN landing rate.
let pdfJsLoadPromise = null;
function loadPdfJs() {
  if (window.pdfjsLib) return Promise.resolve(window.pdfjsLib);
  if (pdfJsLoadPromise) return pdfJsLoadPromise;
  pdfJsLoadPromise = new Promise((resolve, reject) => {
    const script = document.createElement('script');
    script.src = 'https://cdnjs.cloudflare.com/ajax/libs/pdf.js/3.11.174/pdf.min.js';
    script.onload = () => {
      window.pdfjsLib.GlobalWorkerOptions.workerSrc = 'https://cdnjs.cloudflare.com/ajax/libs/pdf.js/3.11.174/pdf.worker.min.js';
      resolve(window.pdfjsLib);
    };
    script.onerror = () => reject(new Error('Could not load the PDF reader.'));
    document.head.appendChild(script);
  });
  return pdfJsLoadPromise;
}

async function extractPdfText(file) {
  const pdfjsLib = await loadPdfJs();
  const arrayBuffer = await file.arrayBuffer();
  const pdf = await pdfjsLib.getDocument({ data: arrayBuffer }).promise;
  let fullText = '';
  for (let i = 1; i <= pdf.numPages; i++) {
    const page = await pdf.getPage(i);
    const content = await page.getTextContent();
    fullText += content.items.map((item) => item.str).join(' ') + '\n';
  }
  return fullText;
}

function parseGrnPdfText(text) {
  const flat = text.replace(/\s+/g, ' ').trim();
  const pattern = /(\d+) (\d{6,8}) (\d{6,10}) (\d{3,4}) (.*?) (\d+\.\d{2}) (\d+\.\d{2}) (\d+\.\d{2}) (\d+\.\d{2}|-) (\d+) (\d+) (\d+\.\d{2}) (\d+\.\d{2}) (\d+\.\d{2})/g;
  const rows = [];
  let match;
  while ((match = pattern.exec(flat)) !== null) {
    const [, , code, , , desc, , , , rateGrn, , qtyGrn] = match;
    const qty = Number(qtyGrn) || 0;
    const price = rateGrn === '-' ? 0 : Number(rateGrn) || 0;
    if (qty > 0) rows.push({ code: code.trim(), name: desc.trim(), qty, price });
  }
  return rows;
}

function downloadPricingSheet(rows) {
  const sheetRows = rows.map((r) => ({
    'Product Name': r.articleName,
    'Channel Code (SKU)': r.code || '',
    'UOM': `${r.packSize}${r.packUnit}/pack`,
    'Base Price (₹)': r.basePrice ?? '',
    'Grading %': r.gradingPercent ?? 0,
    'Vendor Margin %': r.vendorMarginPercent ?? 0,
    'Packaging (₹)': r.packaging ?? 0,
    'Labour (₹)': r.labour ?? 0,
    'Transportation (₹)': r.transportation ?? 0,
    'Final Price (₹)': r.finalPrice ?? '',
  }));
  const wb = XLSX.utils.book_new();
  const ws = XLSX.utils.json_to_sheet(sheetRows);
  XLSX.utils.book_append_sheet(wb, ws, 'Pricing');
  const wbout = XLSX.write(wb, { bookType: 'xlsx', type: 'array' });
  const blob = new Blob([wbout], { type: 'application/octet-stream' });
  const url = URL.createObjectURL(blob);
  const a = document.createElement('a');
  a.href = url;
  a.download = `fnv-pricing-sheet-${new Date().toISOString().split('T')[0]}.xlsx`;
  document.body.appendChild(a);
  a.click();
  document.body.removeChild(a);
  URL.revokeObjectURL(url);
}

function PricingCard({ article, config, onUpdate }) {
  const [grading, setGrading] = useState(String(config?.gradingPercent ?? 0));
  const [vendorMargin, setVendorMargin] = useState(String(config?.vendorMarginPercent ?? 0));
  const [packaging, setPackaging] = useState(String(config?.packaging ?? 0));
  const [labour, setLabour] = useState(String(config?.labour ?? 0));
  const [transportation, setTransportation] = useState(String(config?.transportation ?? 0));

  useEffect(() => {
    setGrading(String(config?.gradingPercent ?? 0));
    setVendorMargin(String(config?.vendorMarginPercent ?? 0));
    setPackaging(String(config?.packaging ?? 0));
    setLabour(String(config?.labour ?? 0));
    setTransportation(String(config?.transportation ?? 0));
  }, [config]);

  const commit = (field, value) => onUpdate(article.key, { [field]: Number(value) || 0 });
  const basePrice = article.basePrice;
  const finalPrice = computeFinalPrice(basePrice, {
    gradingPercent: Number(grading) || 0, vendorMarginPercent: Number(vendorMargin) || 0,
    packaging: Number(packaging) || 0, labour: Number(labour) || 0, transportation: Number(transportation) || 0,
  });

  const row = (label, value, setValue, field) => (
    <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', marginTop: 6 }}>
      <span style={{ fontSize: 11, color: MUTED }}>{label}</span>
      <input
        type="number" value={value} onChange={(e) => setValue(e.target.value)} onBlur={(e) => commit(field, e.target.value)}
        style={{ width: 80, boxSizing: 'border-box', borderRadius: 6, border: `1px solid ${LINE}`, fontSize: 12, padding: '5px 8px', textAlign: 'right' }}
      />
    </div>
  );

  return (
    <Card style={{ marginBottom: 12 }}>
      <div style={{ fontWeight: 800, fontSize: 14 }}>{article.articleName}</div>
      <div style={{ fontSize: 11, color: MUTED, marginTop: 2 }}>{article.code || 'No code'} · {article.packSize}{article.packUnit}/pack</div>
      <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', marginTop: 8, paddingTop: 8, borderTop: `1px solid ${LINE}` }}>
        <span style={{ fontSize: 11, color: MUTED, fontWeight: 700 }}>BASE PRICE</span>
        <span style={{ fontWeight: 700, color: basePrice == null ? MUTED : LEAF }}>{basePrice == null ? 'No purchase yet' : `₹${basePrice.toFixed(2)}`}</span>
      </div>
      {row('Grading %', grading, setGrading, 'gradingPercent')}
      {row('Vendor margin %', vendorMargin, setVendorMargin, 'vendorMarginPercent')}
      {row('Packaging (₹)', packaging, setPackaging, 'packaging')}
      {row('Labour (₹)', labour, setLabour, 'labour')}
      {row('Transportation (₹)', transportation, setTransportation, 'transportation')}
      <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', marginTop: 8, paddingTop: 8, borderTop: `1px solid ${LINE}` }}>
        <span style={{ fontSize: 12, fontWeight: 700, color: INK }}>FINAL PRICE</span>
        <span style={{ fontWeight: 800, fontSize: 15, color: finalPrice == null ? MUTED : TOMATO }}>{finalPrice == null ? '—' : `₹${finalPrice.toFixed(2)}`}</span>
      </div>
    </Card>
  );
}

function PricingTab({ orders, items, purchases, pricingConfig, onUpdate }) {
  const [search, setSearch] = useState('');
  const [channelFilter, setChannelFilter] = useState('ALL');
  const [categoryFilter, setCategoryFilter] = useState('ALL');

  const articles = useMemo(() => buildPricingArticles(orders, items, purchases), [orders, items, purchases]);
  const configByKey = useMemo(() => { const map = {}; pricingConfig.forEach((c) => { map[c.id] = c; }); return map; }, [pricingConfig]);
  const categoriesPresent = useMemo(() => ['ALL', ...Array.from(new Set(articles.map((a) => a.category).filter(Boolean)))], [articles]);

  const filteredArticles = articles
    .filter((a) => !search.trim() || a.articleName.toLowerCase().includes(search.trim().toLowerCase()))
    .filter((a) => channelFilter === 'ALL' || a.platform === channelFilter)
    .filter((a) => categoryFilter === 'ALL' || a.category === categoryFilter);

  const rowsForExport = filteredArticles.map((a) => {
    const c = configByKey[a.key];
    const gradingPercent = c?.gradingPercent ?? 0, vendorMarginPercent = c?.vendorMarginPercent ?? 0;
    const packaging = c?.packaging ?? 0, labour = c?.labour ?? 0, transportation = c?.transportation ?? 0;
    const finalPrice = computeFinalPrice(a.basePrice, { gradingPercent, vendorMarginPercent, packaging, labour, transportation });
    return { ...a, gradingPercent, vendorMarginPercent, packaging, labour, transportation, finalPrice };
  });

  return (
    <div style={{ padding: 16 }}>
      <Card style={{ marginBottom: 12 }}>
        <div style={{ ...sectionTitle, display: 'flex', alignItems: 'center', gap: 6 }}><IndianRupee size={15} /> Pricing</div>
        <div style={hint}>Base price is fetched from the latest purchase price × pack size. Grading % and vendor margin % apply on the base price; packaging, labour and transportation are flat amounts.</div>
        <Field placeholder="Search article..." value={search} onChange={(e) => setSearch(e.target.value)} style={{ marginBottom: 8 }} />
        <div style={smallLabel}>CHANNEL</div>
        <div style={{ display: 'flex', flexWrap: 'wrap', marginBottom: 4 }}>
          <Chip label="All" active={channelFilter === 'ALL'} onClick={() => setChannelFilter('ALL')} />
          {PLATFORMS.map((p) => <Chip key={p} label={p} active={channelFilter === p} onClick={() => setChannelFilter(p)} />)}
        </div>
        <div style={smallLabel}>CATEGORY</div>
        <div style={{ display: 'flex', flexWrap: 'wrap', marginBottom: 10 }}>
          {categoriesPresent.map((c) => <Chip key={c} label={c === 'ALL' ? 'All' : c} active={categoryFilter === c} onClick={() => setCategoryFilter(c)} />)}
        </div>
        <PrimaryBtn onClick={() => downloadPricingSheet(rowsForExport)}>Download pricing sheet</PrimaryBtn>
      </Card>

      {filteredArticles.map((a) => (
        <PricingCard key={a.key} article={a} config={configByKey[a.key]} onUpdate={onUpdate} />
      ))}
      {filteredArticles.length === 0 && <div style={hint}>No indent-imported articles match this filter.</div>}
    </div>
  );
}

function ProfitLossDayCard({ day, channel, records, grnReportsForDay, onUploadGrn }) {
  const [expanded, setExpanded] = useState(false);
  const [fileError, setFileError] = useState('');
  const fileInputRef = useRef(null);

  const totalDispatchQty = records.reduce((s, r) => s + (r.dispatchQty || 0), 0);
  const totalDispatchValue = records.reduce((s, r) => s + (r.cost || 0), 0);

  const handleGrnFile = (e) => {
    const file = e.target.files[0];
    if (!file) return;
    setFileError('');

    const isPdf = file.type === 'application/pdf' || file.name.toLowerCase().endsWith('.pdf');
    if (isPdf) {
      extractPdfText(file)
        .then((text) => {
          const rows = parseGrnPdfText(text);
          if (rows.length === 0) { setFileError('Could not find any GRN rows in this PDF. Try an Excel/CSV export instead.'); return; }
          onUploadGrn(channel, day.date, file.name, rows);
        })
        .catch(() => setFileError('Could not read this PDF. Please try again or use an Excel/CSV export instead.'));
      e.target.value = '';
      return;
    }

    const reader = new FileReader();
    reader.onload = (evt) => {
      try {
        const wb = XLSX.read(evt.target.result, { type: 'array' });
        const sheet = wb.Sheets[wb.SheetNames[0]];
        const json = XLSX.utils.sheet_to_json(sheet, { defval: '' });
        const rows = parseGrnRows(json);
        if (rows.length === 0) { setFileError('No rows with a valid code/name and received quantity were found.'); return; }
        onUploadGrn(channel, day.date, file.name, rows);
      } catch (err) {
        setFileError('Could not read this file. Please upload a valid .xlsx, .xls, .csv, or .pdf GRN report.');
      }
    };
    reader.readAsArrayBuffer(file);
    e.target.value = '';
  };

  const latestGrn = grnReportsForDay[0];
  const grnComparison = useMemo(() => {
    if (!latestGrn) return [];
    const ours = {};
    records.forEach((r) => {
      const k = (r.code || r.articleName).toLowerCase();
      ours[k] = ours[k] || { articleName: r.articleName, qty: 0, cost: 0, lastPrice: r.finalPricePerPack };
      ours[k].qty += r.packsDispatched;
      ours[k].cost += r.cost || 0;
    });
    return latestGrn.rows.map((g) => {
      const k = (g.code || g.name).toLowerCase();
      const match = ours[k];
      const ourQty = match?.qty || 0;
      const ourPrice = match?.lastPrice ?? null;
      const ourCost = match?.cost || 0;
      const grnCost = g.qty * g.price;
      return {
        key: k, articleName: match?.articleName || g.name || g.code,
        grnQty: g.qty, ourQty: Math.round(ourQty * 100) / 100, qtyDiff: Math.round((g.qty - ourQty) * 100) / 100,
        grnPrice: g.price, ourPrice, priceDiff: ourPrice == null ? null : Math.round((g.price - ourPrice) * 100) / 100,
        costDiff: Math.round((grnCost - ourCost) * 100) / 100,
      };
    });
  }, [latestGrn, records]);

  return (
    <Card style={{ marginBottom: 10, padding: 0, overflow: 'hidden' }}>
      <div onClick={() => setExpanded((x) => !x)} style={{ padding: 12, cursor: 'pointer', background: expanded ? '#F6F3EA' : '#fff' }}>
        <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center' }}>
          <div style={{ fontWeight: 800, fontSize: 14 }}>{day.date}</div>
          <ChevronRight size={16} color={MUTED} style={{ transform: expanded ? 'rotate(90deg)' : 'none' }} />
        </div>
        <div style={{ display: 'flex', gap: 12, marginTop: 8 }}>
          <div>
            <div style={{ fontSize: 9, color: MUTED, fontWeight: 700 }}>INDENT QTY</div>
            <div style={{ fontWeight: 700, fontSize: 13 }}>{day.totalIndentQty}</div>
          </div>
          <div>
            <div style={{ fontSize: 9, color: MUTED, fontWeight: 700 }}>DISPATCHED</div>
            <div style={{ fontWeight: 700, fontSize: 13 }}>{Math.round(totalDispatchQty * 100) / 100}</div>
          </div>
          <div>
            <div style={{ fontSize: 9, color: MUTED, fontWeight: 700 }}>VALUE</div>
            <div style={{ fontWeight: 800, fontSize: 13, color: TOMATO }}>₹{totalDispatchValue.toLocaleString('en-IN', { maximumFractionDigits: 2 })}</div>
          </div>
        </div>
      </div>

      {expanded && (
        <div style={{ borderTop: `1px solid ${LINE}`, padding: 12 }}>
          {records.map((r, i) => (
            <div key={i} style={{ borderTop: i > 0 ? `1px solid ${LINE}` : 'none', padding: '8px 0' }}>
              <div style={{ fontWeight: 700, fontSize: 13 }}>{r.articleName}</div>
              <div style={{ fontSize: 11, color: MUTED }}>{r.dispatchQty} {r.unit} · {r.packsDispatched} packs · {r.finalPricePerPack == null ? 'No price yet' : `₹${r.finalPricePerPack.toFixed(2)}/pack`}</div>
              <div style={{ fontSize: 12, fontWeight: 700, color: r.cost == null ? MUTED : LEAF }}>{r.cost == null ? '—' : `₹${r.cost.toFixed(2)}`}</div>
            </div>
          ))}
          {records.length === 0 && <div style={hint}>No dispatches priced for this day.</div>}

          <div style={{ borderTop: `1px solid ${LINE}`, marginTop: 10, paddingTop: 10 }}>
            <div style={{ fontWeight: 700, fontSize: 13, marginBottom: 4 }}>Upload GRN report — {day.date}</div>
            <div style={hint}>Upload the channel's Goods Received Note for this day — the Blinkit/Hyperpure PDF works directly, or an Excel/CSV export — to compare against our numbers.</div>
            <button
              onClick={() => fileInputRef.current?.click()}
              style={{ display: 'flex', alignItems: 'center', gap: 6, background: LEAF, color: '#fff', border: 'none', borderRadius: 8, padding: '9px 14px', fontSize: 12, fontWeight: 700, cursor: 'pointer', marginBottom: 8 }}
            >
              <Upload size={13} /> Upload GRN report
            </button>
            <input ref={fileInputRef} type="file" accept=".xlsx,.xls,.csv,.pdf" onChange={handleGrnFile} style={{ display: 'none' }} />
            {fileError && <div style={{ display: 'flex', alignItems: 'center', gap: 6, fontSize: 11, color: TOMATO, marginBottom: 8 }}><AlertCircle size={12} /> {fileError}</div>}
            {latestGrn && (
              <>
                <div style={{ fontSize: 11, color: MUTED, marginBottom: 6 }}>Comparing: <strong style={{ color: INK }}>{latestGrn.fileName}</strong></div>
                {grnComparison.map((c) => (
                  <div key={c.key} style={{ borderTop: `1px solid ${LINE}`, padding: '6px 0' }}>
                    <div style={{ fontWeight: 700, fontSize: 12 }}>{c.articleName}</div>
                    <div style={{ fontSize: 11, color: MUTED }}>GRN {c.grnQty} vs Ours {c.ourQty} ({c.qtyDiff > 0 ? '+' : ''}{c.qtyDiff})</div>
                    <div style={{ fontSize: 11, color: c.costDiff !== 0 ? TOMATO : LEAF, fontWeight: 700 }}>Cost diff: {c.costDiff > 0 ? '+' : ''}₹{c.costDiff.toFixed(2)}</div>
                  </div>
                ))}
              </>
            )}
          </div>
        </div>
      )}
    </Card>
  );
}

function IndentBatchCardMobile({ batch, orders, articlesByKey, configByKey, onOpen }) {
  const batchOrders = useMemo(() => orders.filter((o) => o.batchId === batch.id), [orders, batch.id]);
  const { totalCost, pricedCount, totalCount } = useMemo(
    () => computeBatchArticleCosts(batch, orders, articlesByKey, configByKey),
    [batch, orders, articlesByKey, configByKey]
  );
  const totalQty = batchOrders.reduce((s, o) => s + (Number(o.packQty) || 0), 0);
  const fulfilmentDate = batchOrders[0]?.fulfilmentDate || '';

  return (
    <div onClick={onOpen} style={{ background: '#fff', border: `1px solid ${LINE}`, borderRadius: 14, padding: 14, marginBottom: 10, cursor: 'pointer' }}>
      <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center' }}>
        <div style={{ fontWeight: 700, fontSize: 13 }}>{batch.platform} — {batch.fileName}</div>
        <ChevronRight size={15} color={MUTED} />
      </div>
      <div style={{ display: 'flex', gap: 12, marginTop: 8, flexWrap: 'wrap' }}>
        <div>
          <div style={{ fontSize: 9, color: MUTED, fontWeight: 700 }}>FULFILMENT DATE</div>
          <div style={{ fontWeight: 700, fontSize: 12 }}>{fulfilmentDate || '—'}</div>
        </div>
        <div>
          <div style={{ fontSize: 9, color: MUTED, fontWeight: 700 }}>ARTICLES</div>
          <div style={{ fontWeight: 700, fontSize: 12 }}>{totalCount}</div>
        </div>
        <div>
          <div style={{ fontSize: 9, color: MUTED, fontWeight: 700 }}>TOTAL QTY</div>
          <div style={{ fontWeight: 700, fontSize: 12 }}>{totalQty}</div>
        </div>
      </div>
      <div style={{ marginTop: 8, paddingTop: 8, borderTop: `1px solid ${LINE}` }}>
        <div style={{ fontSize: 9, color: MUTED, fontWeight: 700 }}>EXPECTED COST SO FAR</div>
        <div style={{ fontWeight: 800, fontSize: 14, color: TOMATO }}>
          ₹{totalCost.toLocaleString('en-IN', { maximumFractionDigits: 2 })}
          <span style={{ fontWeight: 500, color: MUTED, fontSize: 11 }}> ({pricedCount}/{totalCount} priced)</span>
        </div>
      </div>
    </div>
  );
}

function IndentBatchDetailMobile({ batch, orders, articlesByKey, configByKey, grnReports, onUploadGrn, onUpdateIndentBatch, onBack }) {
  const { rows, totalCost, pricedCount, totalCount } = useMemo(
    () => computeBatchArticleCosts(batch, orders, articlesByKey, configByKey),
    [batch, orders, articlesByKey, configByKey]
  );
  const batchOrders = useMemo(() => orders.filter((o) => o.batchId === batch.id), [orders, batch.id]);
  const totalQty = batchOrders.reduce((s, o) => s + (Number(o.packQty) || 0), 0);
  const fulfilmentDate = batchOrders[0]?.fulfilmentDate || '';

  const [poValue, setPoValue] = useState(batch.poValue != null ? String(batch.poValue) : '');
  const [poFileName, setPoFileName] = useState(batch.poFileName || '');
  const [poFileError, setPoFileError] = useState('');
  const poFileInputRef = useRef(null);

  const [grnFileError, setGrnFileError] = useState('');
  const grnFileInputRef = useRef(null);

  const savePoValue = () => onUpdateIndentBatch(batch.id, { poValue: Number(poValue) || 0, poFileName });

  const handlePoFile = (e) => {
    const file = e.target.files[0];
    if (!file) return;
    setPoFileError('');
    setPoFileName(file.name);
    e.target.value = '';
  };

  const projectedProfit = batch.poValue != null ? Math.round((batch.poValue - totalCost) * 100) / 100 : null;

  const batchGrnReports = useMemo(
    () => grnReports.filter((g) => g.batchId === batch.id).sort((a, b) => (b.uploadedAt || '').localeCompare(a.uploadedAt || '')),
    [grnReports, batch.id]
  );
  const latestGrn = batchGrnReports[0];

  const grnComparison = useMemo(() => {
    if (!latestGrn) return null;
    const ownByCode = {};
    rows.forEach((r) => {
      const k = (r.code || r.articleName).toLowerCase();
      ownByCode[k] = ownByCode[k] || { qty: 0, cost: 0 };
      ownByCode[k].qty += r.packQty;
      ownByCode[k].cost += r.cost || 0;
    });
    let grnValue = 0;
    latestGrn.rows.forEach((g) => {
      const k = (g.code || g.name).toLowerCase();
      if (!ownByCode[k]) return;
      grnValue += g.qty * g.price;
    });
    return { grnValue: Math.round(grnValue * 100) / 100 };
  }, [latestGrn, rows]);

  const finalProfit = grnComparison ? Math.round((grnComparison.grnValue - totalCost) * 100) / 100 : null;

  const handleGrnFile = (e) => {
    const file = e.target.files[0];
    if (!file) return;
    setGrnFileError('');
    const isPdf = file.type === 'application/pdf' || file.name.toLowerCase().endsWith('.pdf');
    if (isPdf) {
      extractPdfText(file)
        .then((text) => {
          const parsedRows = parseGrnPdfText(text);
          if (parsedRows.length === 0) { setGrnFileError('Could not find any GRN rows in this PDF.'); return; }
          onUploadGrn(batch.platform, fulfilmentDate || new Date().toISOString().split('T')[0], file.name, parsedRows, batch.id);
        })
        .catch(() => setGrnFileError('Could not read this PDF. Try an Excel/CSV export instead.'));
      e.target.value = '';
      return;
    }
    const reader = new FileReader();
    reader.onload = (evt) => {
      try {
        const wb = XLSX.read(evt.target.result, { type: 'array' });
        const sheet = wb.Sheets[wb.SheetNames[0]];
        const json = XLSX.utils.sheet_to_json(sheet, { defval: '' });
        const parsedRows = parseGrnRows(json);
        if (parsedRows.length === 0) { setGrnFileError('No rows with a valid code/name and received quantity were found.'); return; }
        onUploadGrn(batch.platform, fulfilmentDate || new Date().toISOString().split('T')[0], file.name, parsedRows, batch.id);
      } catch (err) {
        setGrnFileError('Could not read this file. Please upload a valid .xlsx, .xls, .csv, or .pdf GRN report.');
      }
    };
    reader.readAsArrayBuffer(file);
    e.target.value = '';
  };

  return (
    <div style={{ padding: 16 }}>
      <button onClick={onBack} style={{ display: 'flex', alignItems: 'center', gap: 6, background: 'none', border: 'none', color: LEAF, fontWeight: 700, fontSize: 13, cursor: 'pointer', marginBottom: 12, padding: 0 }}>
        <ArrowLeft size={15} /> Back to indents
      </button>

      <Card style={{ marginBottom: 12 }}>
        <div style={{ fontWeight: 700, fontSize: 14 }}>{batch.platform} — {batch.fileName}</div>
        <div style={{ display: 'flex', gap: 14, marginTop: 8, flexWrap: 'wrap' }}>
          <div><div style={{ fontSize: 9, color: MUTED, fontWeight: 700 }}>FULFILMENT DATE</div><div style={{ fontWeight: 700, fontSize: 12 }}>{fulfilmentDate || '—'}</div></div>
          <div><div style={{ fontSize: 9, color: MUTED, fontWeight: 700 }}>ARTICLES</div><div style={{ fontWeight: 700, fontSize: 12 }}>{totalCount}</div></div>
          <div><div style={{ fontSize: 9, color: MUTED, fontWeight: 700 }}>TOTAL QTY</div><div style={{ fontWeight: 700, fontSize: 12 }}>{totalQty}</div></div>
        </div>
      </Card>

      <Card style={{ marginBottom: 12 }}>
        <div style={sectionTitle}>Expected purchase cost ({pricedCount}/{totalCount} priced)</div>
        {rows.map((r) => (
          <div key={r.orderId} style={{ borderTop: `1px solid ${LINE}`, padding: '8px 0' }}>
            <div style={{ fontWeight: 700, fontSize: 13 }}>{r.articleName}</div>
            <div style={{ fontSize: 11, color: MUTED }}>
              {r.code || '—'} · Ordered {r.packQty}{r.shortPacks > 0 ? <span style={{ color: TOMATO }}> · {r.shortPacks} short</span> : ''} · Costed for {r.effectivePacks} · {r.finalPricePerPack == null ? 'No price yet' : `₹${r.finalPricePerPack.toFixed(2)}/pack`}
            </div>
            <div style={{ fontSize: 12, fontWeight: 700, color: r.cost == null ? MUTED : LEAF }}>{r.cost == null ? '—' : `₹${r.cost.toFixed(2)}`}</div>
          </div>
        ))}
        {rows.length === 0 && <div style={hint}>No articles in this indent.</div>}
        <div style={{ display: 'flex', justifyContent: 'flex-end', marginTop: 10, paddingTop: 10, borderTop: `1px solid ${LINE}` }}>
          <div style={{ fontWeight: 800, fontSize: 15, color: TOMATO }}>Total so far: ₹{totalCost.toLocaleString('en-IN', { maximumFractionDigits: 2 })}</div>
        </div>
      </Card>

      <Card style={{ marginBottom: 12 }}>
        <div style={sectionTitle}>Purchase Order</div>
        <div style={hint}>Attach the channel's PO for this indent (for your records), and enter its final billing value.</div>
        <button
          onClick={() => poFileInputRef.current?.click()}
          style={{ width: '100%', display: 'flex', alignItems: 'center', justifyContent: 'center', gap: 6, background: '#fff', color: LEAF, border: `1px solid ${LEAF}`, borderRadius: 8, padding: '9px 0', fontSize: 12, fontWeight: 700, cursor: 'pointer', marginBottom: 8 }}
        >
          <Upload size={13} /> {poFileName || 'Attach PO file'}
        </button>
        <input ref={poFileInputRef} type="file" accept=".xlsx,.xls,.csv,.pdf" onChange={handlePoFile} style={{ display: 'none' }} />
        {poFileError && <div style={{ display: 'flex', alignItems: 'center', gap: 6, fontSize: 11, color: TOMATO, marginBottom: 8 }}><AlertCircle size={12} /> {poFileError}</div>}
        <div style={smallLabel}>PO value (₹)</div>
        <div style={{ display: 'flex', gap: 8 }}>
          <input
            type="number"
            value={poValue}
            onChange={(e) => setPoValue(e.target.value)}
            style={{ flex: 1, boxSizing: 'border-box', borderRadius: 8, border: `1px solid ${LINE}`, fontSize: 13, padding: '8px 10px' }}
          />
          <button onClick={savePoValue} style={{ background: LEAF, color: '#fff', border: 'none', borderRadius: 8, padding: '8px 16px', fontSize: 12, fontWeight: 700, cursor: 'pointer' }}>Save</button>
        </div>
        {projectedProfit != null && (
          <div style={{ marginTop: 12, padding: '10px 14px', background: projectedProfit >= 0 ? '#EAF3DE' : '#F3E7E2', borderRadius: 10 }}>
            <div style={{ fontSize: 10, color: MUTED, fontWeight: 700 }}>PROJECTED {projectedProfit >= 0 ? 'PROFIT' : 'LOSS'}</div>
            <div style={{ fontWeight: 800, fontSize: 17, color: projectedProfit >= 0 ? LEAF_DARK : TOMATO }}>₹{Math.abs(projectedProfit).toLocaleString('en-IN', { maximumFractionDigits: 2 })}</div>
          </div>
        )}
      </Card>

      <Card>
        <div style={sectionTitle}>GRN report for this indent</div>
        <div style={hint}>Upload the channel's GRN once received — Blinkit/Hyperpure PDF works directly, or Excel/CSV.</div>
        <button
          onClick={() => grnFileInputRef.current?.click()}
          style={{ width: '100%', display: 'flex', alignItems: 'center', justifyContent: 'center', gap: 6, background: LEAF, color: '#fff', border: 'none', borderRadius: 8, padding: '9px 0', fontSize: 12, fontWeight: 700, cursor: 'pointer', marginBottom: 8 }}
        >
          <Upload size={13} /> Upload GRN report
        </button>
        <input ref={grnFileInputRef} type="file" accept=".xlsx,.xls,.csv,.pdf" onChange={handleGrnFile} style={{ display: 'none' }} />
        {grnFileError && <div style={{ display: 'flex', alignItems: 'center', gap: 6, fontSize: 11, color: TOMATO, marginBottom: 8 }}><AlertCircle size={12} /> {grnFileError}</div>}
        {latestGrn && grnComparison && (
          <>
            <div style={{ fontSize: 11, color: MUTED, marginBottom: 8 }}>Latest: <strong style={{ color: INK }}>{latestGrn.fileName}</strong> — matched ₹{grnComparison.grnValue.toLocaleString('en-IN', { maximumFractionDigits: 2 })}</div>
            {finalProfit != null && (
              <div style={{ padding: '10px 14px', background: finalProfit >= 0 ? '#EAF3DE' : '#F3E7E2', borderRadius: 10 }}>
                <div style={{ fontSize: 10, color: MUTED, fontWeight: 700 }}>FINAL {finalProfit >= 0 ? 'PROFIT' : 'LOSS'}</div>
                <div style={{ fontWeight: 800, fontSize: 17, color: finalProfit >= 0 ? LEAF_DARK : TOMATO }}>₹{Math.abs(finalProfit).toLocaleString('en-IN', { maximumFractionDigits: 2 })}</div>
              </div>
            )}
          </>
        )}
      </Card>
    </div>
  );
}

function ProfitLossTab({ orders, items, purchases, pricingConfig, dispatchLog, grnReports, indentBatches, onUploadGrn, onUpdateIndentBatch }) {
  const [channel, setChannel] = useState(PLATFORMS[0]);
  const [plView, setPlView] = useState('dispatch'); // 'dispatch' | 'indent'
  const [selectedBatchId, setSelectedBatchId] = useState(null);

  const articles = useMemo(() => buildPricingArticles(orders, items, purchases), [orders, items, purchases]);
  const articlesByKey = useMemo(() => { const map = {}; articles.forEach((a) => { map[a.key] = a; }); return map; }, [articles]);
  const configByKey = useMemo(() => { const map = {}; pricingConfig.forEach((c) => { map[c.id] = c; }); return map; }, [pricingConfig]);

  const records = useMemo(() => {
    const out = [];
    dispatchLog.forEach((log) => {
      (log.items || []).forEach((it) => {
        const order = orders.find((o) => o.id === it.orderId);
        const platform = it.platform || order?.platform;
        const baseProduct = it.baseProduct || order?.product;
        const packSize = it.packSize || order?.packSize;
        const packUnit = it.packUnit || order?.packUnit;
        if (!platform || !baseProduct || !packSize || !packUnit) return;
        const key = `${baseProduct}__${platform}__${packSize}__${packUnit}`;
        const article = articlesByKey[key];
        const finalPricePerPack = article ? computeFinalPrice(article.basePrice, configByKey[key]) : null;
        const packsDispatched = Math.round((it.dispatchQty / packSize) * 100) / 100;
        const cost = finalPricePerPack == null ? null : Math.round(packsDispatched * finalPricePerPack * 100) / 100;
        out.push({ date: log.date || '—', channel: platform, articleName: it.product, code: article?.code || '', dispatchQty: it.dispatchQty, unit: it.unit, packsDispatched, finalPricePerPack, cost });
      });
    });
    return out;
  }, [dispatchLog, orders, articlesByKey, configByKey]);

  const indentQtyByDate = useMemo(() => {
    const map = {};
    orders.filter((o) => o.platform === channel && o.fulfilmentDate).forEach((o) => { map[o.fulfilmentDate] = (map[o.fulfilmentDate] || 0) + o.qty; });
    return map;
  }, [orders, channel]);

  const channelRecords = records.filter((r) => r.channel === channel);
  const days = useMemo(() => {
    const dateSet = new Set([...channelRecords.map((r) => r.date), ...Object.keys(indentQtyByDate)]);
    return Array.from(dateSet).sort((a, b) => b.localeCompare(a)).map((date) => ({
      date, totalIndentQty: Math.round((indentQtyByDate[date] || 0) * 100) / 100, records: channelRecords.filter((r) => r.date === date),
    }));
  }, [channelRecords, indentQtyByDate]);

  const channelTotalValue = channelRecords.reduce((s, r) => s + (r.cost || 0), 0);

  const channelBatches = useMemo(() => indentBatches.filter((b) => b.platform === channel), [indentBatches, channel]);

  const selectedBatch = selectedBatchId ? indentBatches.find((b) => b.id === selectedBatchId) : null;
  if (selectedBatch) {
    return (
      <IndentBatchDetailMobile
        batch={selectedBatch}
        orders={orders}
        articlesByKey={articlesByKey}
        configByKey={configByKey}
        grnReports={grnReports}
        onUploadGrn={onUploadGrn}
        onUpdateIndentBatch={onUpdateIndentBatch}
        onBack={() => setSelectedBatchId(null)}
      />
    );
  }

  return (
    <div style={{ padding: 16 }}>
      <Card style={{ marginBottom: 12 }}>
        <div style={{ ...sectionTitle, display: 'flex', alignItems: 'center', gap: 6 }}><TrendingUp size={15} /> Profit &amp; Loss</div>
        <div style={hint}>
          {plView === 'dispatch'
            ? "Each day shows total indent qty, total dispatched, and total dispatch value (from Pricing). Tap a day to see the breakdown and upload that day's GRN report."
            : "One card per uploaded indent — expected cost fills in as articles get priced. Tap a card to see the breakdown, add the PO value, and upload its GRN once received."}
        </div>
        <div style={{ display: 'flex', flexWrap: 'wrap', marginBottom: 8 }}>
          {PLATFORMS.map((p) => <Chip key={p} label={p} active={channel === p} onClick={() => setChannel(p)} />)}
        </div>
        <div style={{ display: 'flex', gap: 6, marginBottom: 8 }}>
          <Chip label="By dispatch day" active={plView === 'dispatch'} onClick={() => setPlView('dispatch')} />
          <Chip label="By indent" active={plView === 'indent'} onClick={() => setPlView('indent')} />
        </div>
        {plView === 'dispatch' && (
          <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center' }}>
            <span style={{ fontSize: 11, color: MUTED, fontWeight: 700 }}>{channel.toUpperCase()} TOTAL VALUE</span>
            <span style={{ fontWeight: 800, fontSize: 16, color: TOMATO }}>₹{channelTotalValue.toLocaleString('en-IN', { maximumFractionDigits: 2 })}</span>
          </div>
        )}
      </Card>

      {plView === 'dispatch' ? (
        <>
          {days.map((day) => (
            <ProfitLossDayCard
              key={day.date} day={day} channel={channel} records={day.records}
              grnReportsForDay={grnReports.filter((g) => g.channel === channel && g.date === day.date).sort((a, b) => (b.uploadedAt || '').localeCompare(a.uploadedAt || ''))}
              onUploadGrn={onUploadGrn}
            />
          ))}
          {days.length === 0 && <div style={hint}>No {channel} indents or dispatches yet.</div>}
        </>
      ) : (
        <>
          {channelBatches.map((b) => (
            <IndentBatchCardMobile
              key={b.id}
              batch={b}
              orders={orders}
              articlesByKey={articlesByKey}
              configByKey={configByKey}
              onOpen={() => setSelectedBatchId(b.id)}
            />
          ))}
          {channelBatches.length === 0 && <div style={hint}>No {channel} indents uploaded yet.</div>}
        </>
      )}
    </div>
  );
}

function PackagingTab({ orders, items, onAdvanceMany, packingProgress, onUpdatePackedQty }) {
  const [platformFilter, setPlatformFilter] = usePersistedState('fnv_packaging_platform', 'All');
  const [categoryFilter, setCategoryFilter] = usePersistedState('fnv_packaging_category', 'All');
  const [qtySort, setQtySort] = usePersistedState('fnv_packaging_qtysort', 'none'); // 'none' | 'asc' | 'desc'
  const [dateFilter, setDateFilter] = usePersistedState('fnv_packaging_date', '');
  const [selectedKey, setSelectedKey] = useState(null);
  const [itemSearch, setItemSearch] = useState('');

  const categoryByProduct = useMemo(() => {
    const map = {};
    items.forEach((it) => { map[it.name] = it.category; });
    return map;
  }, [items]);

  const filteredOrders = useMemo(() => {
    return orders
      .filter((o) => o.status !== 'dispatched')
      .filter((o) => platformFilter === 'All' || o.platform === platformFilter)
      .filter((o) => categoryFilter === 'All' || categoryByProduct[o.product] === categoryFilter)
      .filter((o) => !dateFilter || o.fulfilmentDate === dateFilter);
  }, [orders, platformFilter, categoryFilter, categoryByProduct, dateFilter]);

  const groupedByDate = useMemo(() => {
    const map = {};
    filteredOrders.forEach((o) => {
      const dateKey = o.fulfilmentDate || 'No date';
      map[dateKey] = map[dateKey] || {};
      const hasPack = !!(o.packQty && o.packSize);
      const cityKey = o.city || CITIES[0];
      const key = hasPack ? `${cityKey}__${dateKey}__${o.product}__${o.platform}__${o.packSize}__${o.packUnit}` : `${cityKey}__${dateKey}__${o.product}__${o.unit}`;
      map[dateKey][key] = map[dateKey][key] || {
        key, product: o.product, articleName: o.articleName || o.product, unit: o.unit, qty: 0, platforms: new Set(),
        pendingIds: [], orderIds: [], hasPack, packSize: o.packSize, packUnit: o.packUnit, targetPacks: 0,
      };
      map[dateKey][key].qty += o.qty;
      map[dateKey][key].platforms.add(o.platform);
      map[dateKey][key].orderIds.push(o.id);
      if (hasPack) map[dateKey][key].targetPacks += o.packQty;
      if (o.status === 'pending') map[dateKey][key].pendingIds.push(o.id);
    });
    return Object.entries(map)
      .map(([date, targetMap]) => {
        let targets = Object.values(targetMap).map((t) => {
          const progress = packingProgress[t.key] || { packedQty: 0, shortQty: 0 };
          const isComplete = t.hasPack ? (progress.packedQty > 0 && progress.packedQty + progress.shortQty >= t.targetPacks) : (t.pendingIds.length === 0);
          return { ...t, isComplete };
        });
        if (qtySort === 'asc') targets.sort((a, b) => (a.hasPack ? a.targetPacks : a.qty) - (b.hasPack ? b.targetPacks : b.qty));
        else if (qtySort === 'desc') targets.sort((a, b) => (b.hasPack ? b.targetPacks : b.qty) - (a.hasPack ? a.targetPacks : a.qty));
        const term = itemSearch.trim().toLowerCase();
        if (term) targets = targets.filter((t) => `${t.articleName} ${t.product}`.toLowerCase().includes(term));
        // Items still needing work stay on top; fully packed ones sink to the bottom.
        targets.sort((a, b) => (a.isComplete === b.isComplete ? 0 : a.isComplete ? 1 : -1));
        return { date, targets };
      })
      .filter((g) => g.targets.length > 0)
      .sort((a, b) => {
        if (a.date === 'No date') return 1;
        if (b.date === 'No date') return -1;
        return a.date.localeCompare(b.date);
      });
  }, [filteredOrders, qtySort, packingProgress, itemSearch]);

  const categoriesPresent = useMemo(() => ['All', ...Array.from(new Set(items.map((it) => it.category).filter(Boolean)))], [items]);

  const selectedTarget = useMemo(() => {
    for (const g of groupedByDate) {
      const t = g.targets.find((x) => x.key === selectedKey);
      if (t) return { ...t, date: g.date };
    }
    return null;
  }, [groupedByDate, selectedKey]);

  if (selectedTarget) {
    return (
      <PackagingDetail
        target={selectedTarget}
        progress={packingProgress[selectedTarget.key] || { packedQty: 0, shortQty: 0 }}
        onSave={(packedQty, shortQty) => onUpdatePackedQty(selectedTarget.key, packedQty, shortQty, selectedTarget.orderIds, selectedTarget.targetPacks)}
        onBack={() => setSelectedKey(null)}
      />
    );
  }

  return (
    <div style={{ padding: 16 }}>
      <Card style={{ marginBottom: 12, padding: '10px 12px' }}>
        <div style={{ display: 'flex', gap: 6, overflowX: 'auto' }}>
          <select value={platformFilter} onChange={(e) => setPlatformFilter(e.target.value)} style={{ flex: '1 1 0', minWidth: 0, boxSizing: 'border-box', padding: '6px 4px', borderRadius: 6, border: `1px solid ${LINE}`, fontSize: 11, color: INK, background: '#fff' }}>
            <option value="All">All channels</option>
            {PLATFORMS.map((p) => <option key={p} value={p}>{p}</option>)}
          </select>
          <select value={categoryFilter} onChange={(e) => setCategoryFilter(e.target.value)} style={{ flex: '1 1 0', minWidth: 0, boxSizing: 'border-box', padding: '6px 4px', borderRadius: 6, border: `1px solid ${LINE}`, fontSize: 11, color: INK, background: '#fff' }}>
            {categoriesPresent.map((c) => <option key={c} value={c}>{c === 'All' ? 'All categories' : c}</option>)}
          </select>
          <select value={qtySort} onChange={(e) => setQtySort(e.target.value)} style={{ flex: '1 1 0', minWidth: 0, boxSizing: 'border-box', padding: '6px 4px', borderRadius: 6, border: `1px solid ${LINE}`, fontSize: 11, color: INK, background: '#fff' }}>
            <option value="none">Default sort</option>
            <option value="asc">Qty: Low-High</option>
            <option value="desc">Qty: High-Low</option>
          </select>
        </div>
        <div style={{ display: 'flex', gap: 8, alignItems: 'center', marginTop: 6 }}>
          <input type="date" value={dateFilter} onChange={(e) => setDateFilter(e.target.value)} style={{ flex: 1, boxSizing: 'border-box', padding: '6px 8px', borderRadius: 6, border: `1px solid ${LINE}`, fontSize: 11, color: INK, background: '#fff' }} />
          {dateFilter && (
            <button onClick={() => setDateFilter('')} style={{ background: 'none', border: 'none', color: TOMATO, fontSize: 11, fontWeight: 700, cursor: 'pointer', flexShrink: 0 }}>Clear</button>
          )}
        </div>
        <Field placeholder="Search item…" value={itemSearch} onChange={(e) => setItemSearch(e.target.value)} style={{ marginTop: 8, marginBottom: 0 }} />
      </Card>

      {groupedByDate.map(({ date, targets }) => (
        <Card key={date} style={{ marginBottom: 12, padding: 10 }}>
          <div style={{ ...sectionTitle, marginBottom: 4 }}>{date === 'No date' ? 'No fulfilment date' : date}</div>
          {targets.map((t) => (
            <PackagingInlineRow
              key={t.key}
              target={t}
              progress={packingProgress[t.key] || { packedQty: 0, shortQty: 0 }}
              onSave={(packedQty, shortQty) => onUpdatePackedQty(t.key, packedQty, shortQty, t.orderIds, t.targetPacks)}
              onAdvanceMany={onAdvanceMany}
              onOpenDetail={() => setSelectedKey(t.key)}
            />
          ))}
        </Card>
      ))}
      {groupedByDate.length === 0 && (
        <Card>
          <div style={hint}>Nothing to pack right now.</div>
        </Card>
      )}
    </div>
  );
}

function PackagingInlineRow({ target: t, progress, onSave, onAdvanceMany, onOpenDetail }) {
  const [packedValue, setPackedValue] = useState(String(progress.packedQty || ''));
  const [shortValue, setShortValue] = useState(String(progress.shortQty || ''));
  useEffect(() => { setPackedValue(String(progress.packedQty || '')); }, [progress.packedQty]);
  useEffect(() => { setShortValue(String(progress.shortQty || '')); }, [progress.shortQty]);

  const enteredPacked = Number(packedValue) || 0;
  const enteredShort = Number(shortValue) || 0;
  // The only two valid outcomes for an article: fully packed, or packed + short adding
  // up to exactly the target — there's no in-between state that can be saved.
  const isResolved = t.hasPack ? enteredPacked + enteredShort === t.targetPacks : true;
  const changed = enteredPacked !== progress.packedQty || enteredShort !== progress.shortQty;
  const canSave = isResolved && changed;
  const isComplete = t.hasPack ? (progress.packedQty > 0 && progress.packedQty + progress.shortQty >= t.targetPacks) : (t.pendingIds.length === 0);

  return (
    <div style={{ borderTop: `1px solid ${LINE}`, padding: '9px 0', background: isComplete ? '#EAF3DE' : 'transparent' }}>
      <div onClick={() => t.hasPack && onOpenDetail()} style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', gap: 8, cursor: t.hasPack ? 'pointer' : 'default' }}>
        <div style={{ flex: 1, minWidth: 0 }}>
          <div style={{ fontWeight: 700, fontSize: 13 }}>{t.articleName || t.product}</div>
          <div style={{ fontSize: 10.5, color: MUTED }}>
            {[...t.platforms].join(' + ')}{t.hasPack ? ` · ${t.packSize}${t.packUnit}/pack` : ''}
          </div>
        </div>
        {t.hasPack ? (
          <div style={{ color: LEAF, fontWeight: 800, fontSize: 13, flexShrink: 0 }}>{t.targetPacks} packs</div>
        ) : (
          <div style={{ color: LEAF, fontWeight: 800, fontSize: 13, flexShrink: 0 }}>{t.qty} {t.unit}</div>
        )}
        {t.hasPack && <ChevronRight size={15} color={MUTED} style={{ flexShrink: 0 }} />}
      </div>

      {t.hasPack ? (
        <>
          <div style={{ display: 'flex', alignItems: 'flex-end', gap: 6, marginTop: 8 }}>
            <div style={{ flex: 2 }}>
              <div style={{ fontSize: 9, color: MUTED, fontWeight: 700 }}>PACKED</div>
              <input
                type="number"
                placeholder="0"
                value={packedValue}
                onClick={(e) => e.stopPropagation()}
                onChange={(e) => setPackedValue(e.target.value)}
                style={{ width: '100%', boxSizing: 'border-box', borderRadius: 6, border: `1px solid ${LINE}`, fontSize: 13, padding: '6px 8px' }}
              />
            </div>
            <div style={{ flex: 1 }}>
              <div style={{ fontSize: 9, color: enteredShort > 0 ? TOMATO : MUTED, fontWeight: 700 }}>SHORT</div>
              <input
                type="number"
                placeholder="0"
                value={shortValue}
                onClick={(e) => e.stopPropagation()}
                onChange={(e) => setShortValue(e.target.value)}
                style={{ width: '100%', boxSizing: 'border-box', borderRadius: 6, border: `1px solid ${enteredShort > 0 ? TOMATO : LINE}`, fontSize: 13, padding: '6px 8px', color: enteredShort > 0 ? TOMATO : INK }}
              />
            </div>
            <button
              onClick={(e) => { e.stopPropagation(); if (canSave) onSave(enteredPacked, enteredShort); }}
              disabled={!canSave}
              style={{ background: canSave ? LEAF : '#C9C2AE', color: '#fff', border: 'none', borderRadius: 6, padding: '8px 12px', fontSize: 12, fontWeight: 700, cursor: canSave ? 'pointer' : 'default' }}
            >
              Save
            </button>
          </div>
          {!isResolved && (enteredPacked > 0 || enteredShort > 0) && (
            <div style={{ color: TOMATO, fontSize: 10.5, marginTop: 4 }}>Packed + short must total {t.targetPacks}</div>
          )}
        </>
      ) : (
        t.pendingIds.length > 0 ? (
          <button onClick={() => onAdvanceMany(t.pendingIds, 'packed')} style={{ width: '100%', background: '#E6F1FB', color: '#1B5E8C', border: 'none', borderRadius: 8, padding: '6px 0', fontWeight: 700, fontSize: 11, marginTop: 5, cursor: 'pointer' }}>Mark {t.pendingIds.length} packed</button>
        ) : (
          <div style={{ color: LEAF, fontSize: 11, marginTop: 5, fontWeight: 600 }}>✓ All packed</div>
        )
      )}
    </div>
  );
}

function PackagingDetail({ target, progress, onSave, onBack }) {
  const [packedValue, setPackedValue] = useState(String(progress.packedQty || ''));
  const [shortValue, setShortValue] = useState(String(progress.shortQty || ''));
  const enteredPacked = Number(packedValue) || 0;
  const enteredShort = Number(shortValue) || 0;
  const isResolved = enteredPacked + enteredShort === target.targetPacks;
  const changed = enteredPacked !== progress.packedQty || enteredShort !== progress.shortQty;
  const canSave = isResolved && changed;

  const save = () => {
    if (!canSave) return;
    onSave(enteredPacked, enteredShort);
    onBack();
  };

  return (
    <div style={{ padding: 16 }}>
      <button onClick={onBack} style={{ display: 'flex', alignItems: 'center', gap: 6, background: 'none', border: 'none', color: LEAF, fontWeight: 700, fontSize: 13, cursor: 'pointer', marginBottom: 12, padding: 0 }}>
        <ArrowLeft size={15} /> Back to packaging
      </button>
      <Card style={{ marginBottom: 14 }}>
        <div style={{ fontWeight: 800, fontSize: 16 }}>{target.articleName || target.product}</div>
        <div style={{ fontSize: 12, color: MUTED, marginTop: 2 }}>
          {[...target.platforms].join(' + ')} · {target.date === 'No date' ? 'No fulfilment date' : target.date}
        </div>
        <div style={{ fontSize: 12, color: MUTED, marginTop: 2 }}>Pack size: {target.packSize}{target.packUnit} per pack</div>
        <div style={{ display: 'flex', gap: 8, marginTop: 12 }}>
          <div style={{ flex: 1, border: `1px solid ${LINE}`, borderRadius: 8, padding: '8px 10px' }}>
            <div style={{ fontSize: 10, color: MUTED, fontWeight: 700 }}>TARGET</div>
            <div style={{ fontSize: 16, fontWeight: 800, color: INK }}>{target.targetPacks} packs</div>
          </div>
          <div style={{ flex: 1, border: `1px solid ${LINE}`, borderRadius: 8, padding: '8px 10px' }}>
            <div style={{ fontSize: 10, color: MUTED, fontWeight: 700 }}>SAVED SO FAR</div>
            <div style={{ fontSize: 16, fontWeight: 800, color: INK }}>{progress.packedQty} packed, {progress.shortQty} short</div>
          </div>
        </div>
      </Card>

      <Card>
        <div style={sectionTitle}>Update packed / short quantity</div>
        <div style={hint}>An article can only be saved once — either fully packed, or packed plus short adding up to the full target ({target.targetPacks} packs).</div>
        <div style={{ display: 'flex', gap: 10 }}>
          <div style={{ flex: 2 }}>
            <div style={smallLabel}>Packed</div>
            <Field type="number" placeholder="0" value={packedValue} onChange={(e) => setPackedValue(e.target.value)} style={{ fontSize: 16, fontWeight: 700 }} />
          </div>
          <div style={{ flex: 1 }}>
            <div style={smallLabel}>Short</div>
            <Field type="number" placeholder="0" value={shortValue} onChange={(e) => setShortValue(e.target.value)} style={{ fontSize: 16, fontWeight: 700, borderColor: enteredShort > 0 ? TOMATO : undefined, color: enteredShort > 0 ? TOMATO : undefined }} />
          </div>
        </div>
        {!isResolved && (enteredPacked > 0 || enteredShort > 0) && (
          <div style={{ color: TOMATO, fontSize: 12, marginTop: -6, marginBottom: 12 }}>Packed + short must total {target.targetPacks} packs.</div>
        )}
        <PrimaryBtn onClick={save} disabled={!canSave}>Save</PrimaryBtn>
      </Card>
    </div>
  );
}

// ---------- Dispatch ----------
function DispatchModal({ selectedCount, crates, onClose, onConfirm }) {
  const [vehicleNo, setVehicleNo] = useState('');
  const [driverName, setDriverName] = useState('');
  const [cratesUsed, setCratesUsed] = useState('');
  const [boxesUsed, setBoxesUsed] = useState('');

  return (
    <div style={{ position: 'fixed', inset: 0, background: 'rgba(0,0,0,0.45)', display: 'flex', alignItems: 'flex-end', justifyContent: 'center', zIndex: 200 }}>
      <div style={{ background: '#fff', borderRadius: '18px 18px 0 0', padding: '24px 20px 32px', width: '100%', maxWidth: 420 }}>
        <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', marginBottom: 6 }}>
          <div style={{ fontWeight: 800, fontSize: 16, display: 'flex', alignItems: 'center', gap: 6 }}><TruckIcon size={16} /> Dispatch order</div>
          <button onClick={onClose} style={{ background: 'none', border: 'none', fontSize: 20, color: MUTED, cursor: 'pointer', lineHeight: 1 }}>✕</button>
        </div>
        <div style={{ fontSize: 12, color: MUTED, marginBottom: 14 }}>{selectedCount} order(s) selected</div>
        <Field placeholder="Vehicle number" value={vehicleNo} onChange={(e) => setVehicleNo(e.target.value)} />
        <Field placeholder="Driver name" value={driverName} onChange={(e) => setDriverName(e.target.value)} />
        <Field placeholder={`Crates (${crates.crates} in stock)`} type="number" value={cratesUsed} onChange={(e) => setCratesUsed(e.target.value)} />
        <Field placeholder={`Boxes (${crates.boxes} in stock)`} type="number" value={boxesUsed} onChange={(e) => setBoxesUsed(e.target.value)} />
        <div style={{ display: 'flex', gap: 8, marginTop: 4 }}>
          <button onClick={onClose} style={{ flex: 1, background: '#fff', color: INK, border: `1px solid ${LINE}`, borderRadius: 10, padding: '11px 0', fontWeight: 700, fontSize: 13, cursor: 'pointer' }}>Cancel</button>
          <button
            onClick={() => onConfirm({ vehicleNo: vehicleNo.trim(), driverName: driverName.trim(), cratesUsed: Number(cratesUsed) || 0, boxesUsed: Number(boxesUsed) || 0 })}
            style={{ flex: 2, background: LEAF, color: '#fff', border: 'none', borderRadius: 10, padding: '11px 0', fontWeight: 700, fontSize: 13, cursor: 'pointer' }}
          >
            Confirm dispatch
          </button>
        </div>
      </div>
    </div>
  );
}

function DispatchTab({ orders, crates, dispatchLog, onDispatchBatch }) {
  const packed = useMemo(() => orders
    .filter((o) => o.status === 'packed')
    .map((o) => ({ ...o, remaining: Math.max(0, Math.round((o.qty - (o.dispatchedQty || 0) - (o.shortQty || 0)) * 100) / 100) })),
  [orders]);

  const [view, setView] = useState('dispatch'); // 'dispatch' | 'history'
  const [channel, setChannel] = usePersistedState('fnv_dispatch_channel', PLATFORMS[0]);
  const [storeSel, setStoreSel] = usePersistedState('fnv_dispatch_store', '');
  const [selected, setSelected] = useState([]);
  const [showModal, setShowModal] = useState(false);

  // Channel -> its stores (Blinkit has one, Flipkart has one per dark store).
  const activeChannel = PLATFORMS.includes(channel) ? channel : PLATFORMS[0];
  const storeOptions = useMemo(() => storeOptionsFor(orders, activeChannel), [orders, activeChannel]);
  const activeStore = storeOptions.find((s) => s.value === storeSel) || storeOptions[0] || null;
  const visiblePacked = useMemo(
    () => packed.filter((o) => o.platform === activeChannel && activeStore && orderStore(o) === activeStore.store),
    [packed, activeChannel, activeStore],
  );
  const changeChannel = (c) => { setChannel(c); setStoreSel(''); setSelected([]); };
  const changeStore = (v) => { setStoreSel(v); setSelected([]); };

  const toggleSelect = (id) => setSelected((p) => (p.includes(id) ? p.filter((x) => x !== id) : [...p, id]));

  const submitDispatch = ({ vehicleNo, driverName, cratesUsed, boxesUsed }) => {
    if (selected.length === 0) return;
    const dispatchItems = selected.map((id) => {
      const o = packed.find((x) => x.id === id);
      return { orderId: id, dispatchQty: o?.remaining || 0, shortQty: 0 };
    });
    onDispatchBatch({ items: dispatchItems, vehicleNo, driverName, cratesUsed, boxesUsed });
    setSelected([]); setShowModal(false);
  };

  return (
    <div style={{ padding: 16 }}>
      <div style={{ display: 'flex', gap: 6, marginBottom: 14, flexWrap: 'wrap' }}>
        <Chip label="Dispatch" active={view === 'dispatch'} onClick={() => setView('dispatch')} />
        <Chip label={`History (${dispatchLog.length})`} active={view === 'history'} onClick={() => setView('history')} />
      </div>

      {view === 'dispatch' && (
        <div style={{ display: 'flex', gap: 8, marginBottom: 12 }}>
          <select value={activeChannel} onChange={(e) => changeChannel(e.target.value)} style={{ flex: '1 1 0', minWidth: 0, boxSizing: 'border-box', padding: '8px 6px', borderRadius: 8, border: `1px solid ${LINE}`, fontSize: 12, fontWeight: 700, color: INK, background: '#fff' }}>
            {PLATFORMS.map((p) => <option key={p} value={p}>{p}</option>)}
          </select>
          <select value={activeStore ? activeStore.value : ''} onChange={(e) => changeStore(e.target.value)} disabled={storeOptions.length === 0} style={{ flex: '1.4 1 0', minWidth: 0, boxSizing: 'border-box', padding: '8px 6px', borderRadius: 8, border: `1px solid ${LINE}`, fontSize: 12, fontWeight: 700, color: INK, background: '#fff' }}>
            {storeOptions.length === 0 && <option value="">No stores yet</option>}
            {storeOptions.map((s) => <option key={s.value} value={s.value}>{s.label}</option>)}
          </select>
        </div>
      )}

      {view === 'dispatch' && (
        <>
          <Card style={{ marginBottom: 14 }}>
            <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'flex-start', gap: 8 }}>
              <div style={{ flex: 1 }}>
                <div style={sectionTitle}>Packed — ready ({visiblePacked.length})</div>
                <div style={hint}>An article only shows up here once it's been fully resolved in Packaging — either fully packed, or packed with the rest marked short. Quantities aren't editable here; go back to Packaging to change them.</div>
              </div>
            </div>
            {visiblePacked.map((o) => (
              <div key={o.id} style={{ borderTop: `1px solid ${LINE}`, padding: '9px 0' }}>
                <div style={{ display: 'flex', alignItems: 'center', gap: 10 }}>
                  <div onClick={() => toggleSelect(o.id)} style={{ width: 18, height: 18, borderRadius: 4, border: `1.5px solid ${LINE}`, background: selected.includes(o.id) ? LEAF : '#fff', cursor: 'pointer', flexShrink: 0 }} />
                  <div style={{ flex: 1, minWidth: 0 }}>
                    <div style={{ fontWeight: 700, fontSize: 13 }}>{o.articleName || o.product}</div>
                    <div style={{ fontSize: 11, color: MUTED }}>{o.id}</div>
                  </div>
                  <div style={{ fontWeight: 800, fontSize: 14, color: LEAF, flexShrink: 0 }}>{o.remaining} {o.unit}</div>
                </div>
              </div>
            ))}
            {visiblePacked.length === 0 && <div style={hint}>Nothing packed yet for this store — resolve articles in Packaging first.</div>}
          </Card>
        </>
      )}
      {view === 'dispatch' && (
        <div style={{ position: 'sticky', bottom: 0, background: BG, paddingTop: 10, marginTop: -10, marginLeft: -16, marginRight: -16, paddingLeft: 16, paddingRight: 16, paddingBottom: 4 }}>
          <button
            onClick={() => setShowModal(true)}
            disabled={selected.length === 0}
            style={{ width: '100%', display: 'flex', alignItems: 'center', justifyContent: 'center', gap: 6, background: selected.length === 0 ? '#C9C2AE' : TOMATO, color: '#fff', border: 'none', borderRadius: 10, padding: '13px 0', fontWeight: 700, fontSize: 14, cursor: selected.length === 0 ? 'default' : 'pointer', boxShadow: '0 -4px 10px rgba(0,0,0,0.06)' }}
          >
            <TruckIcon size={15} /> Dispatch order{selected.length > 0 ? ` (${selected.length})` : ''}
          </button>
        </div>
      )}

      {view === 'history' && (
        <Card>
          <div style={sectionTitle}>Dispatch history ({dispatchLog.length})</div>
          {dispatchLog.map((d) => (
            <div key={d.id} style={{ borderTop: `1px solid ${LINE}`, padding: '8px 0' }}>
              <div style={{ fontWeight: 700, fontSize: 13 }}>{d.id} · {d.vehicleNo}</div>
              <div style={{ fontSize: 12, color: MUTED }}>{d.driverName} · {(d.orderIds || []).length} orders · {d.totalDispatchQty || 0} dispatched · {d.cratesUsed} crates, {d.boxesUsed} boxes · {d.time}</div>
            </div>
          ))}
          {dispatchLog.length === 0 && <div style={hint}>No dispatches yet.</div>}
        </Card>
      )}

      {showModal && (
        <DispatchModal selectedCount={selected.length} crates={crates} onClose={() => setShowModal(false)} onConfirm={submitDispatch} />
      )}
    </div>
  );
}

// ---------- Crates ----------
function CountBlock({ label, value, color, onMinus, onPlus }) {
  return (
    <Card style={{ flex: 1 }}>
      <div style={{ fontSize: 12, color: MUTED, fontWeight: 700 }}>{label}</div>
      <div style={{ fontSize: 26, fontWeight: 800, color, marginTop: 4 }}>{value}</div>
      <div style={{ display: 'flex', gap: 8, marginTop: 8 }}>
        <button onClick={onMinus} style={{ flex: 1, border: `1px solid ${LINE}`, background: '#fff', borderRadius: 8, padding: '8px 0', fontWeight: 700, cursor: 'pointer' }}>−</button>
        <button onClick={onPlus} style={{ flex: 1, border: 'none', background: color, color: '#fff', borderRadius: 8, padding: '8px 0', fontWeight: 700, cursor: 'pointer' }}>+</button>
      </div>
    </Card>
  );
}
function CratesTab({ crates, log, onAdjust }) {
  return (
    <div style={{ padding: 16 }}>
      <div style={{ display: 'flex', gap: 12, marginBottom: 14 }}>
        <CountBlock label="Crates" value={crates.crates} color={LEAF} onMinus={() => onAdjust('crates', -1)} onPlus={() => onAdjust('crates', 1)} />
        <CountBlock label="Boxes" value={crates.boxes} color={AMBER} onMinus={() => onAdjust('boxes', -1)} onPlus={() => onAdjust('boxes', 1)} />
      </div>
      <Card>
        <div style={sectionTitle}>Recent activity</div>
        {log.map((l) => (
          <div key={l.id} style={{ display: 'flex', justifyContent: 'space-between', borderTop: `1px solid ${LINE}`, padding: '8px 0' }}>
            <div style={{ fontSize: 12, flex: 1 }}>{l.delta > 0 ? 'Added' : 'Removed'} {Math.abs(l.delta)} {l.type}{l.note ? ` · ${l.note}` : ''}</div>
            <div style={{ fontSize: 11, color: MUTED }}>{l.time}</div>
          </div>
        ))}
        {log.length === 0 && <div style={hint}>No activity yet — use +/− above.</div>}
      </Card>
    </div>
  );
}

// ---------- Users & Roles ----------
function UsersRolesTab({ users, roles, onAddUser, onUpdateUser, onDeleteUser, onAddRole, onDeleteRole, onToggleRolePermission }) {
  const [name, setName] = useState('');
  const [contact, setContact] = useState('');
  const [username, setUsername] = useState('');
  const [password, setPassword] = useState('');
  const [usernameError, setUsernameError] = useState('');
  const [roleId, setRoleId] = useState(roles[0]?.id || '');
  const [newRoleName, setNewRoleName] = useState('');
  const [visiblePasswordId, setVisiblePasswordId] = useState(null);
  const PERMISSION_SECTIONS = NAV.map((n) => ({ key: n.key, label: n.label }));

  const submitUser = () => {
    if (!name.trim() || !roleId || !username.trim() || !password.trim()) return;
    const uname = username.trim().toLowerCase();
    if (users.some((u) => (u.username || '').toLowerCase() === uname)) {
      setUsernameError('This username is already taken.');
      return;
    }
    setUsernameError('');
    onAddUser({ id: `U-${Date.now().toString(36).toUpperCase().slice(-5)}`, name: name.trim(), contact: contact.trim(), roleId, status: 'active', username: uname, password: password.trim() });
    setName(''); setContact(''); setUsername(''); setPassword('');
  };
  const submitRole = () => {
    if (!newRoleName.trim()) return;
    const perms = {}; PERMISSION_SECTIONS.forEach((s) => { perms[s.key] = false; });
    onAddRole({ id: `ROLE-${Date.now().toString(36).toUpperCase().slice(-5)}`, name: newRoleName.trim(), permissions: perms });
    setNewRoleName('');
  };

  return (
    <div style={{ padding: 16 }}>
      <Card style={{ marginBottom: 14 }}>
        <div style={sectionTitle}>Add employee</div>
        <Field placeholder="Name" value={name} onChange={(e) => setName(e.target.value)} />
        <Field placeholder="Phone / email" value={contact} onChange={(e) => setContact(e.target.value)} />
        <div style={smallLabel}>Role</div>
        <div style={{ display: 'flex', flexWrap: 'wrap', marginBottom: 10 }}>{roles.map((r) => <Chip key={r.id} label={r.name} active={roleId === r.id} onClick={() => setRoleId(r.id)} />)}</div>
        <div style={smallLabel}>Login credentials</div>
        <Field placeholder="Username" value={username} onChange={(e) => { setUsername(e.target.value); setUsernameError(''); }} />
        {usernameError && <div style={{ fontSize: 11, color: TOMATO, marginTop: -6, marginBottom: 8 }}>{usernameError}</div>}
        <Field placeholder="Password" value={password} onChange={(e) => setPassword(e.target.value)} />
        <PrimaryBtn onClick={submitUser}>Add employee</PrimaryBtn>
      </Card>

      <Card style={{ marginBottom: 14 }}>
        <div style={sectionTitle}>Employees ({users.length})</div>
        {users.map((u) => (
          <div key={u.id} style={{ borderTop: `1px solid ${LINE}`, padding: '10px 0' }}>
            <div style={{ display: 'flex', justifyContent: 'space-between' }}>
              <div style={{ fontWeight: 700, fontSize: 14 }}>{u.name}</div>
              <button onClick={() => onDeleteUser(u.id)} style={{ background: 'none', border: 'none', color: TOMATO, fontWeight: 700, fontSize: 12, cursor: 'pointer' }}>Remove</button>
            </div>
            <div style={{ fontSize: 12, color: MUTED }}>{u.contact || '—'}</div>
            <div style={{ fontSize: 12, color: MUTED, marginTop: 4 }}>
              Username: <strong style={{ color: INK }}>{u.username || '—'}</strong>
              {u.password && (
                <>
                  {' · Password: '}
                  <span style={{ fontFamily: 'monospace' }}>{visiblePasswordId === u.id ? u.password : '••••••••'}</span>
                  {' '}
                  <button onClick={() => setVisiblePasswordId(visiblePasswordId === u.id ? null : u.id)} style={{ background: 'none', border: 'none', color: LEAF, fontWeight: 700, fontSize: 11, cursor: 'pointer' }}>
                    {visiblePasswordId === u.id ? 'Hide' : 'Show'}
                  </button>
                </>
              )}
            </div>
            <div style={{ display: 'flex', flexWrap: 'wrap', marginTop: 6 }}>{roles.map((r) => <Chip key={r.id} label={r.name} active={u.roleId === r.id} onClick={() => onUpdateUser(u.id, { roleId: r.id })} />)}</div>
            <button onClick={() => onUpdateUser(u.id, { status: u.status === 'active' ? 'inactive' : 'active' })} style={{ background: u.status === 'active' ? '#EAF3DE' : '#F3E7E2', color: u.status === 'active' ? LEAF_DARK : TOMATO, border: 'none', borderRadius: 999, padding: '4px 10px', fontSize: 11, fontWeight: 700, marginTop: 8, cursor: 'pointer' }}>
              {u.status === 'active' ? 'Active' : 'Inactive'}
            </button>
          </div>
        ))}
      </Card>

      <Card>
        <div style={sectionTitle}>Roles & permissions</div>
        <div style={hint}>Tap a section to toggle access for that role.</div>
        {roles.map((r) => {
          const inUse = users.some((u) => u.roleId === r.id);
          return (
            <div key={r.id} style={{ borderTop: `1px solid ${LINE}`, padding: '10px 0' }}>
              <div style={{ display: 'flex', justifyContent: 'space-between' }}>
                <div style={{ fontWeight: 700, fontSize: 14 }}>{r.name}</div>
                {!inUse && <button onClick={() => onDeleteRole(r.id)} style={{ background: 'none', border: 'none', color: TOMATO, fontWeight: 700, fontSize: 12, cursor: 'pointer' }}>Delete</button>}
              </div>
              <div style={{ display: 'flex', flexWrap: 'wrap', marginTop: 6 }}>
                {PERMISSION_SECTIONS.map((s) => <Chip key={s.key} label={s.label} active={!!r.permissions[s.key]} onClick={() => onToggleRolePermission(r.id, s.key, !r.permissions[s.key])} />)}
              </div>
            </div>
          );
        })}
        <div style={{ display: 'flex', gap: 8, marginTop: 10 }}>
          <Field placeholder="New role name" value={newRoleName} onChange={(e) => setNewRoleName(e.target.value)} style={{ flex: 1, marginBottom: 0 }} />
          <button onClick={submitRole} style={{ background: LEAF, color: '#fff', border: 'none', borderRadius: 8, padding: '0 14px', fontWeight: 700, fontSize: 12, cursor: 'pointer' }}>+ Add</button>
        </div>
      </Card>
    </div>
  );
}
