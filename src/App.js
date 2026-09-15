import React, { useState, useMemo, useRef, useEffect } from 'react';
import * as XLSX from 'xlsx';
import { initializeApp } from 'firebase/app';
import {
  getFirestore, collection, doc,
  onSnapshot, setDoc, updateDoc, deleteDoc, writeBatch, getDocs
} from 'firebase/firestore';
import {
  Menu, X, LayoutDashboard, Tag, Scissors, ClipboardList, ShoppingBag,
  PackageCheck, Truck, Boxes, Users, Upload, FileSpreadsheet, AlertCircle,
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

function newAliasId() {
  return `AL-${Date.now().toString(36).toUpperCase().slice(-6)}-${Math.floor(Math.random() * 900 + 100)}`;
}

const SEED_ITEMS = [
  { id: 'IT-001', name: 'Tomato', uom: 'kg', category: 'VEGETABLES', aliases: [
    { id: 'AL-1001', channel: 'Blinkit', code: 'BLK-TOM-240', packSize: '0.5', packUnit: 'kg' },
    { id: 'AL-1002', channel: 'Flipkart', code: 'FKT-TOM-01', packSize: '1', packUnit: 'kg' },
  ] },
  { id: 'IT-002', name: 'Onion', uom: 'kg', category: 'VEGETABLES', aliases: [
    { id: 'AL-1003', channel: 'Flipkart', code: 'FKT-ONI-500', packSize: '0.5', packUnit: 'kg' },
  ] },
  { id: 'IT-003', name: 'Banana', uom: 'dozen', category: 'FRUITS', aliases: [
    { id: 'AL-1004', channel: 'Blinkit', code: 'BLK-BAN-DZ', packSize: '1', packUnit: 'pieces' },
  ] },
  { id: 'IT-004', name: 'Cauliflower', uom: 'piece', category: 'VEGETABLES', aliases: [] },
  { id: 'IT-005', name: 'Carrot', uom: 'kg', category: 'VEGETABLES', aliases: [] },
  { id: 'IT-006', name: 'Green Beans', uom: 'kg', category: 'VEGETABLES', aliases: [] },
  { id: 'IT-007', name: 'Green Pea', uom: 'kg', category: 'VEGETABLES', aliases: [] },
  { id: 'IT-008', name: 'Pulao Veggie Mix', uom: 'pack', category: 'VEGETABLES', aliases: [
    { id: 'AL-1005', channel: 'Blinkit', code: 'BLK-PVM-01', packSize: '1', packUnit: 'pack' },
  ] },
];

const SEED_ORDERS = [
  { id: 'BLK-1042', platform: 'Blinkit', product: 'Tomato', articleName: 'Tomato Hybrid(Pack)', qty: 240, unit: 'kg', status: 'pending', packQty: 480, packSize: 0.5, packUnit: 'kg', fulfilmentDate: '' },
  { id: 'FKT-3391', platform: 'Flipkart', product: 'Onion', qty: 500, unit: 'kg', status: 'pending', fulfilmentDate: '' },
  { id: 'BLK-1043', platform: 'Blinkit', product: 'Banana', qty: 120, unit: 'dozen', status: 'packed', fulfilmentDate: '' },
  { id: 'FKT-3402', platform: 'Flipkart', product: 'Potato', qty: 350, unit: 'kg', status: 'dispatched', fulfilmentDate: '' },
  { id: 'BLK-1050', platform: 'Blinkit', product: 'Pulao Veggie Mix', qty: 10, unit: 'pack', status: 'pending', fulfilmentDate: '' },
];

const SEED_PURCHASES = [
  { id: 'P-01', item: 'Tomato', supplier: 'Ramesh Farms', qty: 300, unit: 'kg', cost: 9000, source: 'Manual', date: '2026-09-10', type: 'purchased' },
  { id: 'P-02', item: 'Onion', supplier: 'Patil Traders', qty: 500, unit: 'kg', cost: 12500, source: 'Manual', date: '2026-09-11', type: 'purchased' },
];

const SEED_RECIPES = [
  { id: 'RCP-001', name: 'Pulao Veggie Mix', outputItemId: 'IT-008', ingredients: [
    { id: 'ing-1', itemId: 'IT-004', qtyPerUnit: 0.5, unit: 'piece' },
    { id: 'ing-2', itemId: 'IT-005', qtyPerUnit: 100, unit: 'g' },
    { id: 'ing-3', itemId: 'IT-006', qtyPerUnit: 100, unit: 'g' },
    { id: 'ing-4', itemId: 'IT-007', qtyPerUnit: 100, unit: 'g' },
  ] },
];

const SEED_ROLES = [
  { id: 'ROLE-ADMIN', name: 'Admin', permissions: { dashboard: true, items: true, cutprocess: true, orders: true, purchase: true, stockcount: true, pricing: true, profitloss: true, packaging: true, dispatch: true, crates: true, users: true } },
  { id: 'ROLE-WAREHOUSE', name: 'Warehouse Staff', permissions: { dashboard: true, items: false, cutprocess: false, orders: false, purchase: false, stockcount: true, pricing: false, profitloss: false, packaging: true, dispatch: true, crates: true, users: false } },
];

const SEED_USERS = [
  { id: 'U-001', name: 'Rohit Sharma', contact: '98765 43210', roleId: 'ROLE-ADMIN', status: 'active', username: 'rohit', password: 'admin123' },
  { id: 'U-002', name: 'Suresh Patil', contact: '91234 56780', roleId: 'ROLE-WAREHOUSE', status: 'active', username: 'suresh', password: 'warehouse123' },
];

const SEED_VENDORS = [
  { id: 'VEN-001', name: 'Ramesh Farms', contact: '98765 11111', itemIds: ['IT-001', 'IT-005'] },
  { id: 'VEN-002', name: 'Patil Traders', contact: '98765 22222', itemIds: ['IT-002', 'IT-006', 'IT-007'] },
  { id: 'VEN-003', name: 'Kadam Orchards', contact: '98765 33333', itemIds: ['IT-003'] },
];

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
  let total = 0, found = false;
  headers.forEach((h) => {
    if (!h) return;
    const norm = h.toLowerCase().replace(/[^a-z0-9]/g, '');
    if (KNOWN_INDENT_HEADERS.has(norm)) return;
    const v = rowObj[h];
    if (v === '' || v === null || v === undefined) return;
    const num = Number(v);
    if (!isNaN(num)) { total += num; found = true; }
  });
  return found ? total : null;
}
function parseIndentRows(json) {
  return json.map((r, idx) => {
    const headers = Object.keys(r);
    const rawName = String(pickField(r, ['title', 'article', 'product', 'item', 'description']) || '').trim();
    const rawCode = String(pickField(r, ['fsn', 'itemcode', 'articlecode', 'productcode', 'sku', 'code']) || '').trim();
    let qty = Number(pickField(r, ['indent', 'qty', 'quantity', 'orderedqty']) || 0);
    if (!qty) { const st = sumUnknownNumericColumns(r, headers); if (st) qty = st; }
    const unit = String(pickField(r, ['umo', 'uom', 'unit']) || '').trim();
    const rawCategory = String(pickField(r, ['type', 'category']) || '').trim();
    return { key: `row-${idx}-${rawName}`, rawName, rawCode, qty, unit, rawCategory };
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

function LoginScreenMobile({ onLogin, error }) {
  const [username, setUsername] = useState('');
  const [password, setPassword] = useState('');
  const [showPassword, setShowPassword] = useState(false);

  const submit = () => {
    if (!username.trim() || !password.trim()) return;
    onLogin(username, password);
  };

  return (
    <div style={{ display: 'flex', justifyContent: 'center', padding: '24px 12px', fontFamily: '"Nunito Sans", -apple-system, BlinkMacSystemFont, "Segoe UI", sans-serif' }}>
      <div style={{ width: 390, height: 760, background: BG, borderRadius: 34, border: `8px solid ${INK}`, boxShadow: '0 20px 50px rgba(0,0,0,0.18)', overflow: 'hidden', display: 'flex', flexDirection: 'column', alignItems: 'center', justifyContent: 'center', padding: 24 }}>
        <Sprout size={32} color={LEAF} />
        <div style={{ fontWeight: 800, fontSize: 18, color: INK, marginTop: 10 }}>FNV Business App</div>
        <div style={{ fontSize: 12, color: MUTED, marginBottom: 20 }}>Sign in to continue</div>
        <div style={{ width: '100%' }}>
          <Field placeholder="Username" value={username} onChange={(e) => setUsername(e.target.value)} />
          <div style={{ position: 'relative' }}>
            <Field placeholder="Password" type={showPassword ? 'text' : 'password'} value={password} onChange={(e) => setPassword(e.target.value)} style={{ paddingRight: 56 }} />
            <button onClick={() => setShowPassword((s) => !s)} style={{ position: 'absolute', right: 10, top: 9, background: 'none', border: 'none', color: LEAF, fontSize: 11, fontWeight: 700, cursor: 'pointer' }}>
              {showPassword ? 'Hide' : 'Show'}
            </button>
          </div>
          {error && (
            <div style={{ display: 'flex', alignItems: 'center', gap: 6, fontSize: 11, color: TOMATO, marginTop: -4, marginBottom: 10 }}>
              <AlertCircle size={12} /> {error}
            </div>
          )}
          <PrimaryBtn onClick={submit}>Sign in</PrimaryBtn>
        </div>
      </div>
    </div>
  );
}

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
  const [crates, setCrates] = useState({ crates: 180, boxes: 260 });
  const [crateLog, setCrateLog] = useState([]);
  const [dispatchLog, setDispatchLog] = useState([]);
  const [indentBatches, setIndentBatches] = useState([]);
  const [packingProgress, setPackingProgress] = useState({}); // { [targetKey]: packedPacks }
  const [stockCounts, setStockCounts] = useState([]); // nightly closing-stock entries, one per item per date
  const [pricingConfig, setPricingConfig] = useState([]); // editable per-article pricing inputs
  const [grnReports, setGrnReports] = useState([]); // uploaded GRN files per channel + day
  const [dbReady, setDbReady] = useState(false);
  const [currentUser, setCurrentUser] = useState(null);
  const [loginError, setLoginError] = useState('');

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

    // crates is a single doc
    const unsub2 = onSnapshot(doc(db, 'settings', 'crates'), (d) => {
      if (d.exists()) setCrates(d.data());
    });

    // packing progress — keyed by target id, stored as a map for O(1) lookup
    const unsub3 = onSnapshot(collection(db, 'packingProgress'), (snap) => {
      const map = {};
      snap.docs.forEach((d) => { map[d.id] = d.data().packedQty || 0; });
      setPackingProgress(map);
    });

    return () => { unsubs.forEach((u) => u()); unsub2(); unsub3(); };
  }, []);

  // ── Session — restore a saved login once the users list has loaded ──
  useEffect(() => {
    if (!dbReady || currentUser) return;
    const savedId = window.localStorage.getItem('fnv_current_user_id');
    if (!savedId) return;
    const u = users.find((x) => x.id === savedId && x.status === 'active');
    if (u) setCurrentUser(u);
  }, [dbReady, users, currentUser]);

  const handleLogin = (usernameInput, passwordInput) => {
    const uname = usernameInput.trim().toLowerCase();
    const match = users.find((u) => (u.username || '').toLowerCase() === uname && u.password === passwordInput && u.status === 'active');
    if (!match) { setLoginError('Incorrect username or password, or this account is inactive.'); return; }
    setLoginError('');
    setCurrentUser(match);
    window.localStorage.setItem('fnv_current_user_id', match.id);
  };
  const handleLogout = () => {
    setCurrentUser(null);
    window.localStorage.removeItem('fnv_current_user_id');
  };

  const fbUpdate = (col, id, patch)  => updateDoc(doc(db, col, id), patch);
  const fbDelete = (col, id)         => deleteDoc(doc(db, col, id));
  const fbSetDoc = (col, id, obj)    => setDoc(doc(db, col, id), obj);

  const addItem = (it) => fbSetDoc('items', it.id, it);
  const addItemsBulk = (newItems) => { const b = writeBatch(db); newItems.forEach((it) => b.set(doc(db,'items',it.id), it)); b.commit(); };
  const updateItem = (id, u) => fbUpdate('items', id, u);
  const deleteItem = (id) => fbDelete('items', id);
  // Different articles from the same channel can map to the same base item but have
  // their own pack size (e.g. "Baby Banana" 500g vs "Banana 3pc" 600g, both on Blinkit,
  // both = item "Banana"). So each distinct article code gets its OWN alias entry.
  const ensureAliasForCode = (itemId, channel, code) => {
    const it = items.find((x) => x.id === itemId); if (!it) return;
    const exists = (it.aliases || []).some((a) => a.channel === channel && a.code && code && a.code.toLowerCase() === code.toLowerCase());
    if (exists) return;
    fbUpdate('items', itemId, { aliases: [...(it.aliases || []), { id: newAliasId(), channel, code: code || '', packSize: '', packUnit: 'kg' }] });
  };
  const updateAliasById = (itemId, aliasId, patch) => {
    const it = items.find((x) => x.id === itemId); if (!it) return;
    fbUpdate('items', itemId, { aliases: (it.aliases || []).map((a) => (a.id === aliasId ? { ...a, ...patch } : a)) });
  };
  const importOrder = (o) => fbSetDoc('orders', o.id, o);
  const advanceStatus = (id, next) => fbUpdate('orders', id, { status: next });
  const advanceMany = (ids, next) => { const b = writeBatch(db); ids.forEach((id) => b.update(doc(db,'orders',id), { status: next })); b.commit(); };
  const addPurchase = (p) => fbSetDoc('purchases', p.id, { date: new Date().toISOString().split('T')[0], type: 'purchased', ...p });
  const addPurchaseRequirements = (rows, dateOverride) => { const b = writeBatch(db); const today = dateOverride || new Date().toISOString().split('T')[0]; rows.forEach((r) => b.set(doc(db,'purchases',r.id), { date: today, type: 'requirement', ...r })); b.commit(); };
  const removePurchasesByIds = (ids) => { const b = writeBatch(db); ids.forEach((id) => b.delete(doc(db,'purchases',id))); b.commit(); };
  const recordStockCount = (itemId, itemName, unit, date, closingQty) => {
    fbSetDoc('stockCounts', `${itemId}__${date}`, { id: `${itemId}__${date}`, itemId, itemName, unit, date, closingQty: Number(closingQty) || 0 });
  };
  const updatePricingConfig = (key, patch) => {
    const existing = pricingConfig.find((p) => p.id === key);
    fbSetDoc('pricingConfig', key, { id: key, ...(existing || {}), ...patch });
  };
  const uploadGrnReport = (channel, date, fileName, rows) => {
    const id = `GRN-${channel.slice(0, 3).toUpperCase()}-${date}-${Date.now().toString(36).toUpperCase().slice(-6)}`;
    fbSetDoc('grnReports', id, { id, channel, date, fileName, uploadedAt: new Date().toISOString().split('T')[0], rows });
  };
  const addRecipe = (r) => fbSetDoc('recipes', r.id, r);
  const deleteRecipe = (id) => fbDelete('recipes', id);
  const adjustCrates = async (type, delta, note) => {
    const next = { ...crates, [type]: Math.max(0, crates[type] + delta) };
    await setDoc(doc(db, 'settings', 'crates'), next);
    const logId = `CL-${Date.now().toString(36).toUpperCase().slice(-6)}`;
    fbSetDoc('crateLog', logId, { id: logId, type, delta, note: note || '', time: new Date().toLocaleTimeString('en-IN', { hour: '2-digit', minute: '2-digit' }) });
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
        platform: o.platform, baseProduct: o.product, packSize: o.packSize || null, packUnit: o.packUnit || null,
      });
    });
    b.commit();
    const dispatchDate = new Date().toISOString().split('T')[0];
    if (cratesUsed > 0) adjustCrates('crates', -cratesUsed, `Dispatch ${vehicleNo || ''}`.trim());
    if (boxesUsed > 0) adjustCrates('boxes', -boxesUsed, `Dispatch ${vehicleNo || ''}`.trim());
    const did = `DSP-${Date.now().toString(36).toUpperCase().slice(-6)}`;
    fbSetDoc('dispatchLog', did, { id: did, date: dispatchDate, items: logItems, orderIds: logItems.map((li) => li.orderId), totalDispatchQty, vehicleNo: vehicleNo || '—', driverName: driverName || '—', cratesUsed, boxesUsed, time: new Date().toLocaleTimeString('en-IN', { hour: '2-digit', minute: '2-digit' }) });
  };
  const createIndentBatch = (b) => fbSetDoc('indentBatches', b.id, b);
  const updatePackedQty = (key, packedQty, orderIds, targetPacks) => {
    fbSetDoc('packingProgress', key, { packedQty });
    const complete = targetPacks > 0 && packedQty >= targetPacks;
    orderIds.forEach((id) => {
      const o = orders.find((x) => x.id === id);
      if (o && o.status !== 'dispatched') fbUpdate('orders', id, { status: complete ? 'packed' : 'pending' });
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
  const addVendor = (v) => fbSetDoc('vendors', v.id, v);
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

  if (!dbReady) return (
    <div style={{ display: 'flex', justifyContent: 'center', padding: '24px 12px', fontFamily: '"Nunito Sans", -apple-system, BlinkMacSystemFont, "Segoe UI", sans-serif' }}>
      <div style={{ width: 390, height: 760, background: BG, borderRadius: 34, border: `8px solid ${INK}`, boxShadow: '0 20px 50px rgba(0,0,0,0.18)', overflow: 'hidden', display: 'flex', alignItems: 'center', justifyContent: 'center', flexDirection: 'column', gap: 16 }}>
        <Sprout size={36} color={LEAF} />
        <div style={{ fontWeight: 700, fontSize: 16, color: INK }}>Connecting to database…</div>
      </div>
    </div>
  );

  if (!currentUser) return <LoginScreenMobile onLogin={handleLogin} error={loginError} />;

  const currentRole = roles.find((r) => r.id === currentUser.roleId);
  const visibleNavItems = NAV.filter((n) => !currentRole || currentRole.permissions[n.key] !== false);

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
          {tab === 'dashboard' && <DashboardTab orders={orders} purchases={purchases} items={items} crates={crates} />}
          {tab === 'items' && <ItemsTab items={items} onAdd={addItem} onAddBulk={addItemsBulk} onUpdate={updateItem} onDelete={deleteItem} />}
          {tab === 'vendors' && (
            <VendorsTab items={items} vendors={vendors} vendorLedger={vendorLedger} placedOrders={placedOrders} onAdd={addVendor} onDelete={deleteVendor} onToggleItem={toggleVendorItem} onSettle={settleEntries} onUpdatePlacedOrder={updatePlacedOrder} onDeletePlacedOrder={deletePlacedOrder} />
          )}
          {tab === 'cutprocess' && <CutProcessTab items={items} recipes={recipes} orders={orders} onAddRecipe={addRecipe} onDeleteRecipe={deleteRecipe} onAddPurchaseRequirements={addPurchaseRequirements} />}
          {tab === 'orders' && (
            <OrdersTab
              orders={orders} items={items} indentBatches={indentBatches}
              onImport={importOrder} onAddItem={addItem} onEnsureAlias={ensureAliasForCode} onUpdateAlias={updateAliasById}
              onCreateIndentBatch={createIndentBatch} onToggleReleaseBatch={toggleReleaseBatch}
            />
          )}
          {tab === 'purchase' && <PurchasesTab purchases={purchases} orders={orders} items={items} recipes={recipes} vendors={vendors} vendorLedger={vendorLedger} stockCounts={stockCounts} onAddLedgerEntry={addLedgerEntry} onSavePlacedOrder={savePlacedOrder} indentBatches={indentBatches} />}
          {tab === 'stockcount' && <StockCountTab items={items} stockCounts={stockCounts} onRecord={recordStockCount} />}
          {tab === 'pricing' && <PricingTab orders={orders} items={items} purchases={purchases} pricingConfig={pricingConfig} onUpdate={updatePricingConfig} />}
          {tab === 'profitloss' && <ProfitLossTab orders={orders} items={items} purchases={purchases} pricingConfig={pricingConfig} dispatchLog={dispatchLog} grnReports={grnReports} onUploadGrn={uploadGrnReport} />}
          {tab === 'packaging' && <PackagingTab orders={orders} onAdvanceMany={advanceMany} packingProgress={packingProgress} onUpdatePackedQty={updatePackedQty} />}
          {tab === 'dispatch' && (
            <DispatchTab orders={orders} crates={crates} dispatchLog={dispatchLog} onAdvance={advanceStatus} onDispatchBatch={dispatchBatch} />
          )}
          {tab === 'crates' && <CratesTab crates={crates} log={crateLog} onAdjust={adjustCrates} />}
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
              <div style={{ padding: '10px 8px', flex: 1, overflowY: 'auto' }}>
                {visibleNavItems.map((n) => (
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
              <div style={{ padding: '10px 14px 16px', borderTop: '1px solid rgba(255,255,255,0.08)' }}>
                <div style={{ fontSize: 11, color: '#8A968A', marginBottom: 8 }}>Signed in as <strong style={{ color: '#fff' }}>{currentUser.name}</strong></div>
                <button onClick={() => { handleLogout(); setDrawerOpen(false); }} style={{ display: 'flex', alignItems: 'center', gap: 8, background: 'transparent', border: 'none', color: '#B7C2B2', fontSize: 12, cursor: 'pointer', padding: 0 }}>
                  <ArrowLeft size={14} /> Log out
                </button>
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

function VendorsTab({ items, vendors, vendorLedger, placedOrders, onAdd, onDelete, onToggleItem, onSettle, onUpdatePlacedOrder, onDeletePlacedOrder }) {
  const [name, setName] = useState('');
  const [contact, setContact] = useState('');
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

  const allCredit = vendorLedger.filter((e) => e.payment === 'credit' && !e.settled);
  const totalCredit = allCredit.reduce((s, e) => s + e.total, 0);

  const grouped = useMemo(() => {
    const map = {};
    allCredit.forEach((e) => {
      const key = `${e.vendorId}__${e.date}`;
      if (!map[key]) map[key] = { vendorId: e.vendorId, vendorName: e.vendorName, date: e.date, entries: [], total: 0 };
      map[key].entries.push(e);
      map[key].total += e.total;
    });
    return Object.values(map).sort((a, b) => a.date.localeCompare(b.date) || a.vendorName.localeCompare(b.vendorName));
  }, [allCredit]);

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
  const groupEffectiveTotal = (g) => g.entries.reduce((s, e) => s + getEffective(e).total, 0);

  const confirmPayment = () => {
    if (!payModal) return;
    onSettle(payModal.entries.map((e) => e.id), payMode, payNote.trim(), draftEdits);
    setPayModal(null); setPayMode('cash'); setPayRef(''); setPayNote('');
  };

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
        <div style={sectionTitle}>Vendors ({vendors.length})</div>
        {vendors.map((v) => (
          <div key={v.id} style={{ borderTop: `1px solid ${LINE}`, padding: '12px 0' }}>
            <div style={{ display: 'flex', justifyContent: 'space-between', marginBottom: 8 }}>
              <div>
                <div style={{ fontWeight: 700, fontSize: 14 }}>{v.name}</div>
                <div style={{ fontSize: 12, color: MUTED }}>{v.contact || '—'} · {v.itemIds.length} item{v.itemIds.length !== 1 ? 's' : ''} linked</div>
              </div>
              <div style={{ textAlign: 'right' }}>
                {(() => {
                  const credit = vendorLedger.filter((e) => e.vendorId === v.id && e.payment === 'credit' && !e.settled).reduce((s, e) => s + e.total, 0);
                  return credit > 0 ? (
                    <div style={{ background: '#FBEFDC', color: AMBER, borderRadius: 6, padding: '3px 8px', fontSize: 11, fontWeight: 800 }}>
                      ₹{credit.toLocaleString('en-IN')} credit due
                    </div>
                  ) : null;
                })()}
                {confirmDeleteId === v.id ? (
                  <div style={{ display: 'flex', gap: 6, marginTop: 4 }}>
                    <button onClick={() => { onDelete(v.id); setConfirmDeleteId(null); }} style={{ background: TOMATO, color: '#fff', border: 'none', borderRadius: 6, padding: '4px 8px', fontSize: 11, fontWeight: 700, cursor: 'pointer' }}>Yes</button>
                    <button onClick={() => setConfirmDeleteId(null)} style={{ background: '#fff', color: INK, border: `1px solid ${LINE}`, borderRadius: 6, padding: '4px 8px', fontSize: 11, fontWeight: 700, cursor: 'pointer' }}>No</button>
                  </div>
                ) : (
                  <button onClick={() => setConfirmDeleteId(v.id)} style={{ background: 'none', border: 'none', color: TOMATO, cursor: 'pointer', marginTop: 4 }}><Trash2 size={14} /></button>
                )}
              </div>
            </div>
            <VendorItemLinkerMobile vendorId={v.id} vendorItemIds={v.itemIds} items={items} onToggle={onToggleItem} />
          </div>
        ))}
        {vendors.length === 0 && <div style={hint}>No vendors yet.</div>}
      </Card>

      {allCredit.length > 0 && (
        <Card style={{ marginTop: 14 }}>
          <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', marginBottom: 12 }}>
            <div style={{ fontWeight: 700, fontSize: 14 }}>Outstanding credit</div>
            <div style={{ fontWeight: 800, color: AMBER }}>₹{totalCredit.toLocaleString('en-IN')}</div>
          </div>
          {grouped.map((g) => {
              const gKey = `${g.vendorId}-${g.date}`;
              const isOpen = expandedGroup === gKey;
              const effTotal = groupEffectiveTotal(g);
              return (
                <div key={gKey} style={{ borderTop: `1px solid ${LINE}` }}>
                  {/* Header row — tap to expand */}
                  <div onClick={() => setExpandedGroup(isOpen ? null : gKey)} style={{ padding: '12px 0', display: 'flex', alignItems: 'center', justifyContent: 'space-between', cursor: 'pointer' }}>
                    <div>
                      <div style={{ fontWeight: 700, fontSize: 14 }}>{g.vendorName}</div>
                      <div style={{ fontSize: 11, color: MUTED, marginTop: 2 }}>{g.date} · {g.entries.length} item{g.entries.length !== 1 ? 's' : ''} · {isOpen ? '▲ collapse' : '▼ view & edit'}</div>
                    </div>
                    <div style={{ display: 'flex', alignItems: 'center', gap: 10 }}>
                      <div style={{ fontWeight: 800, color: AMBER, fontSize: 16 }}>₹{effTotal.toLocaleString('en-IN')}</div>
                      <button
                        onClick={(e) => { e.stopPropagation(); setPayModal(g); setPayMode('cash'); setPayRef(''); setPayNote(''); }}
                        style={{ background: LEAF, color: '#fff', border: 'none', borderRadius: 8, padding: '7px 14px', fontWeight: 700, fontSize: 12, cursor: 'pointer' }}
                      >
                        Pay
                      </button>
                    </div>
                  </div>

                  {/* Expanded editable items */}
                  {isOpen && (
                    <div style={{ background: '#F6F3EA', borderRadius: 10, padding: '10px 12px', marginBottom: 8 }}>
                      <div style={{ display: 'grid', gridTemplateColumns: '1fr 60px 52px 72px', gap: 6, marginBottom: 6 }}>
                        {['ITEM', 'QTY', 'UOM', 'RATE (₹)'].map((h) => <div key={h} style={{ fontSize: 9, fontWeight: 700, color: MUTED }}>{h}</div>)}
                      </div>
                      {g.entries.map((e) => {
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
                              Total: ₹{eff.total.toLocaleString('en-IN')}{changed ? ' ✏️' : ''}
                            </div>
                          </div>
                        );
                      })}
                      <div style={{ display: 'flex', justifyContent: 'flex-end', marginTop: 8, paddingTop: 8, borderTop: `1px solid ${LINE}` }}>
                        <span style={{ fontWeight: 800, color: AMBER, fontSize: 13 }}>Revised: ₹{effTotal.toLocaleString('en-IN')}</span>
                      </div>
                    </div>
                  )}
                </div>
              );
            })}
        </Card>
      )}

      {/* Payment modal — bottom sheet style */}
      {payModal && (
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
                  <div key={e.id} style={{ display: 'flex', justifyContent: 'space-between', padding: '5px 0', borderTop: `1px solid ${LINE}`, fontSize: 12 }}>
                    <span style={{ color: changed ? AMBER : MUTED }}>{e.itemName} · {eff.qty} {e.unit} @ ₹{eff.unitPrice}/{e.unit}{changed ? ' ✏️' : ''}</span>
                    <span style={{ fontWeight: 700, color: changed ? AMBER : INK }}>₹{eff.total.toLocaleString('en-IN')}</span>
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
      )}

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
function ReleaseBatchCard({ batch: b, onToggleReleaseBatch }) {
  const [purchaseDate, setPurchaseDate] = useState(b.purchaseDate || '');

  return (
    <div style={{ borderTop: `1px solid ${LINE}`, padding: '10px 0' }}>
      <div style={{ fontWeight: 700, fontSize: 13 }}>{b.platform} indent — {b.fileName}</div>
      <div style={{ fontSize: 11, color: MUTED, margin: '2px 0 8px' }}>{b.compiled.map((c) => `${c.qty} ${c.unit} ${c.itemName}`).join(', ')}</div>
      {b.released && b.purchaseDate && (
        <div style={{ fontSize: 11, color: LEAF, fontWeight: 700, marginBottom: 6 }}>Purchase date: {b.purchaseDate}</div>
      )}
      {!b.released && (
        <>
          <div style={smallLabel}>Purchase date (required)</div>
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
    pendingIndent.rows.forEach((r) => {
      if (!isRowReady(r) || !selectedRowKeys.has(r.key)) { remaining.push(r); return; }
      const item = items.find((it) => it.id === r.mappedItemId);
      if (!item) { remaining.push(r); return; }
      const alias = getRowAlias(r);
      const packSize = Number(alias?.packSize) || 1;
      const packUnit = alias?.packUnit || item.uom;
      const finalQty = Math.round(r.qty * packSize * 100) / 100;
      onImport({
        id: `${pendingIndent.platform.slice(0, 3).toUpperCase()}-${Math.floor(1000 + Math.random() * 9000)}`,
        platform: pendingIndent.platform,
        product: item.name,
        articleName: r.rawName,
        qty: finalQty,
        unit: item.uom,
        status: 'pending',
        fulfilmentDate: pendingIndent.fulfilmentDate || '',
        packQty: r.qty,
        packSize,
        packUnit,
      });
      const key = `${item.name}__${item.uom}`;
      if (!compiledMap[key]) compiledMap[key] = { itemName: item.name, unit: item.uom, qty: 0 };
      compiledMap[key].qty += finalQty;
    });
    const compiled = Object.values(compiledMap);
    if (compiled.length > 0) onCreateIndentBatch({ id: `BATCH-${Date.now().toString(36).toUpperCase().slice(-6)}`, platform: pendingIndent.platform, fileName: pendingIndent.fileName, compiled, released: false, purchaseRowIds: [] });
    if (!remaining.length) setIndentFulfilmentDate('');
    setPendingIndent(remaining.length ? { ...pendingIndent, rows: remaining } : null);
    setSelectedRowKeys(new Set(remaining.filter((r) => selectedRowKeys.has(r.key)).map((r) => r.key)));
  };

  return (
    <div style={{ padding: 16 }}>
      {indentBatches.length > 0 && (
        <Card style={{ marginBottom: 14 }}>
          <div style={sectionTitle}>Release to Purchase Manager</div>
          {indentBatches.map((b) => <ReleaseBatchCard key={b.id} batch={b} onToggleReleaseBatch={onToggleReleaseBatch} />)}
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

      <Card>
        <div style={sectionTitle}>All orders ({orders.length})</div>
        {orders.map((o) => (
          <div key={o.id} style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', borderTop: `1px solid ${LINE}`, padding: '8px 0' }}>
            <div style={{ flex: 1 }}>
              <div style={{ fontWeight: 700, fontSize: 13 }}>{o.id} · {o.platform}</div>
              <div style={{ fontSize: 12, color: MUTED }}>{o.articleName || o.product} · {o.qty} {o.unit}{o.fulfilmentDate ? ` · due ${o.fulfilmentDate}` : ''}</div>
            </div>
            <StatusPill status={o.status} />
          </div>
        ))}
      </Card>
    </div>
  );
}

// ---------- Purchases ----------
const PURCHASE_CATEGORY_OPTIONS = ['ALL', 'FRUITS', 'VEGETABLES', 'FLOWER', 'EXOTIC', 'GRAINS', 'CUT'];

function PurchasesTab({ purchases, orders, items, recipes, vendors, vendorLedger, stockCounts, onAddLedgerEntry, onSavePlacedOrder, indentBatches }) {
  const [categoryFilter, setCategoryFilter] = useState('ALL');
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
  const [paymentMode, setPaymentMode] = useState('cash');
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
  const neededByProduct = useMemo(() => {
    const map = {};
    const addDemand = (name, qty, unit) => {
      map[name] = map[name] || { needed: 0, unit };
      map[name].needed += qty;
    };
    orders
      .filter((o) => o.status !== 'dispatched')
      .forEach((o) => {
        const matchingRecipes = recipes.filter((r) => items.find((it) => it.id === r.outputItemId)?.name === o.product);
        if (matchingRecipes.length > 0) {
          matchingRecipes.forEach((recipe) => {
            recipe.ingredients.forEach((ing) => {
              const ingItem = items.find((it) => it.id === ing.itemId);
              if (!ingItem) return;
              const norm = normalizeIngredientQty(ing.qtyPerUnit * o.qty, ing.unit);
              addDemand(ingItem.name, norm.value, norm.unit);
            });
          });
        } else {
          addDemand(o.product, o.qty, o.unit);
        }
      });
    return map;
  }, [orders, recipes, items]);

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
    return items
      .filter((it) => neededByProduct[it.name])
      .filter((it) => categoryFilter === 'ALL' || it.category === categoryFilter)
      .map((it) => {
        const needed = neededByProduct[it.name].needed;
        const unit = neededByProduct[it.name].unit;
        const stock = stockByItem[it.name] || 0;
        const toBuy = Math.max(0, Math.round((needed - stock) * 100) / 100);
        return { ...it, needed, unit, stock, toBuy };
      })
      // Already sufficiently stocked (stock beats needed by more than the buffer %) — no need to buy.
      .filter((it) => it.stock <= it.needed * (1 + buffer / 100));
  }, [items, neededByProduct, categoryFilter, stockByItem, bufferPercent]);

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

  const hasActiveFilters = categoryFilter !== 'ALL' || Number(bufferPercent) !== 0;
  const clearFilters = () => { setCategoryFilter('ALL'); setBufferPercent('0'); };

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
        <div style={{ display: 'flex', gap: 8 }}>
          <div style={{ flex: 1 }}>
            <span style={{ fontSize: 11, color: MUTED }}>Category</span>
            <select
              value={categoryFilter}
              onChange={(e) => setCategoryFilter(e.target.value)}
              style={{ width: '100%', boxSizing: 'border-box', marginTop: 4, padding: '9px 8px', borderRadius: 8, border: `1px solid ${LINE}`, fontSize: 12, color: INK, background: '#fff' }}
            >
              {PURCHASE_CATEGORY_OPTIONS.map((c) => <option key={c} value={c}>{c}</option>)}
            </select>
          </div>
          <div style={{ width: 90 }}>
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
        {filteredItems.map((it) => (
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

function parseGrnRows(json) {
  return json
    .map((r) => {
      const code = String(pickField(r, ['code', 'sku', 'itemcode', 'articlecode', 'fsn']) || '').trim();
      const name = String(pickField(r, ['itemname', 'name', 'article', 'product', 'description']) || '').trim();
      const qty = Number(pickField(r, ['receivedqty', 'qty', 'quantity', 'accepted']) || 0);
      const price = Number(pickField(r, ['price', 'rate', 'unitprice', 'unitrate']) || 0);
      return { code, name, qty, price };
    })
    .filter((r) => (r.code || r.name) && r.qty > 0);
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
        setFileError('Could not read this file. Please upload a valid .xlsx, .xls, or .csv GRN report.');
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
            <div style={hint}>Upload the channel's Goods Received Note for this day to compare against our numbers.</div>
            <button
              onClick={() => fileInputRef.current?.click()}
              style={{ display: 'flex', alignItems: 'center', gap: 6, background: LEAF, color: '#fff', border: 'none', borderRadius: 8, padding: '9px 14px', fontSize: 12, fontWeight: 700, cursor: 'pointer', marginBottom: 8 }}
            >
              <Upload size={13} /> Upload GRN report
            </button>
            <input ref={fileInputRef} type="file" accept=".xlsx,.xls,.csv" onChange={handleGrnFile} style={{ display: 'none' }} />
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

function ProfitLossTab({ orders, items, purchases, pricingConfig, dispatchLog, grnReports, onUploadGrn }) {
  const [channel, setChannel] = useState(PLATFORMS[0]);

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

  return (
    <div style={{ padding: 16 }}>
      <Card style={{ marginBottom: 12 }}>
        <div style={{ ...sectionTitle, display: 'flex', alignItems: 'center', gap: 6 }}><TrendingUp size={15} /> Profit &amp; Loss</div>
        <div style={hint}>Each day shows total indent qty, total dispatched, and total dispatch value (from Pricing). Tap a day to see the breakdown and upload that day's GRN report.</div>
        <div style={{ display: 'flex', flexWrap: 'wrap', marginBottom: 8 }}>
          {PLATFORMS.map((p) => <Chip key={p} label={p} active={channel === p} onClick={() => setChannel(p)} />)}
        </div>
        <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center' }}>
          <span style={{ fontSize: 11, color: MUTED, fontWeight: 700 }}>{channel.toUpperCase()} TOTAL VALUE</span>
          <span style={{ fontWeight: 800, fontSize: 16, color: TOMATO }}>₹{channelTotalValue.toLocaleString('en-IN', { maximumFractionDigits: 2 })}</span>
        </div>
      </Card>

      {days.map((day) => (
        <ProfitLossDayCard
          key={day.date} day={day} channel={channel} records={day.records}
          grnReportsForDay={grnReports.filter((g) => g.channel === channel && g.date === day.date).sort((a, b) => (b.uploadedAt || '').localeCompare(a.uploadedAt || ''))}
          onUploadGrn={onUploadGrn}
        />
      ))}
      {days.length === 0 && <div style={hint}>No {channel} indents or dispatches yet.</div>}
    </div>
  );
}

function PackagingTab({ orders, onAdvanceMany, packingProgress, onUpdatePackedQty }) {
  const [platformFilter, setPlatformFilter] = useState('All');

  const [dateFilter, setDateFilter] = useState('');
  const [selectedKey, setSelectedKey] = useState(null);

  const filteredOrders = useMemo(() => {
    return orders
      .filter((o) => o.status !== 'dispatched')
      .filter((o) => platformFilter === 'All' || o.platform === platformFilter)
      .filter((o) => !dateFilter || o.fulfilmentDate === dateFilter);
  }, [orders, platformFilter, dateFilter]);

  const groupedByDate = useMemo(() => {
    const map = {};
    filteredOrders.forEach((o) => {
      const dateKey = o.fulfilmentDate || 'No date';
      map[dateKey] = map[dateKey] || {};
      const hasPack = !!(o.packQty && o.packSize);
      const key = hasPack ? `${dateKey}__${o.product}__${o.platform}__${o.packSize}__${o.packUnit}` : `${dateKey}__${o.product}__${o.unit}`;
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
      .map(([date, targetMap]) => ({ date, targets: Object.values(targetMap) }))
      .sort((a, b) => {
        if (a.date === 'No date') return 1;
        if (b.date === 'No date') return -1;
        return a.date.localeCompare(b.date);
      });
  }, [filteredOrders]);

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
        packedQty={packingProgress[selectedTarget.key] || 0}
        onSave={(packedQty) => onUpdatePackedQty(selectedTarget.key, packedQty, selectedTarget.orderIds, selectedTarget.targetPacks)}
        onBack={() => setSelectedKey(null)}
      />
    );
  }

  return (
    <div style={{ padding: 16 }}>
      <Card style={{ marginBottom: 12 }}>
        <div style={smallLabel}>Channel</div>
        <div style={{ display: 'flex', flexWrap: 'wrap' }}>
          <Chip label="All" active={platformFilter === 'All'} onClick={() => setPlatformFilter('All')} />
          {PLATFORMS.map((p) => (
            <Chip key={p} label={p} active={platformFilter === p} onClick={() => setPlatformFilter(p)} />
          ))}
        </div>
        <div style={smallLabel}>Fulfilment date</div>
        <div style={{ display: 'flex', gap: 8, alignItems: 'center' }}>
          <Field type="date" value={dateFilter} onChange={(e) => setDateFilter(e.target.value)} style={{ marginBottom: 0, flex: 1 }} />
          {dateFilter && (
            <button onClick={() => setDateFilter('')} style={{ background: 'none', border: 'none', color: TOMATO, fontSize: 12, fontWeight: 700, cursor: 'pointer' }}>Clear</button>
          )}
        </div>
      </Card>

      {groupedByDate.map(({ date, targets }) => (
        <Card key={date} style={{ marginBottom: 12, padding: 10 }}>
          <div style={{ ...sectionTitle, marginBottom: 4 }}>{date === 'No date' ? 'No fulfilment date' : date}</div>
          {targets.map((t) => (
            <PackagingInlineRow
              key={t.key}
              target={t}
              packedQty={packingProgress[t.key] || 0}
              onSave={(packedQty) => onUpdatePackedQty(t.key, packedQty, t.orderIds, t.targetPacks)}
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

function PackagingInlineRow({ target: t, packedQty, onSave, onAdvanceMany, onOpenDetail }) {
  const [value, setValue] = useState(String(packedQty || ''));
  useEffect(() => { setValue(String(packedQty || '')); }, [packedQty]);

  const entered = Number(value) || 0;
  const shortfall = t.hasPack ? Math.max(0, t.targetPacks - entered) : 0;
  const changed = entered !== packedQty;

  return (
    <div style={{ borderTop: `1px solid ${LINE}`, padding: '9px 0' }}>
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
        <div style={{ display: 'flex', alignItems: 'flex-end', gap: 8, marginTop: 8 }}>
          <div style={{ flex: 1 }}>
            <div style={{ fontSize: 9, color: MUTED, fontWeight: 700 }}>PACKED</div>
            <input
              type="number"
              placeholder="0"
              value={value}
              onClick={(e) => e.stopPropagation()}
              onChange={(e) => setValue(e.target.value)}
              style={{ width: '100%', boxSizing: 'border-box', borderRadius: 6, border: `1px solid ${LINE}`, fontSize: 13, padding: '6px 8px' }}
            />
          </div>
          <div style={{ flex: 1, textAlign: 'right' }}>
            <div style={{ fontSize: 9, color: shortfall > 0 ? TOMATO : MUTED, fontWeight: 700 }}>SHORTFALL</div>
            <div style={{ fontSize: 13, fontWeight: 800, color: shortfall > 0 ? TOMATO : LEAF }}>{shortfall > 0 ? `${shortfall} short` : '✓'}</div>
          </div>
          <button
            onClick={(e) => { e.stopPropagation(); onSave(Math.max(0, entered)); }}
            disabled={!changed}
            style={{ background: changed ? LEAF : '#C9C2AE', color: '#fff', border: 'none', borderRadius: 6, padding: '8px 14px', fontSize: 12, fontWeight: 700, cursor: changed ? 'pointer' : 'default' }}
          >
            Save
          </button>
        </div>
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

function PackagingDetail({ target, packedQty, onSave, onBack }) {
  const [value, setValue] = useState(String(packedQty || ''));
  const entered = Number(value) || 0;
  const shortfall = Math.max(0, target.targetPacks - entered);
  const excess = Math.max(0, entered - target.targetPacks);

  const save = () => {
    onSave(Math.max(0, entered));
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
            <div style={{ fontSize: 10, color: MUTED, fontWeight: 700 }}>PACKED SO FAR</div>
            <div style={{ fontSize: 16, fontWeight: 800, color: INK }}>{packedQty} packs</div>
          </div>
        </div>
      </Card>

      <Card>
        <div style={sectionTitle}>Update packed quantity</div>
        <div style={hint}>Enter how many packs are actually packed — the shortfall is calculated automatically.</div>
        <Field
          type="number"
          placeholder="Packed packs"
          value={value}
          onChange={(e) => setValue(e.target.value)}
          style={{ fontSize: 16, fontWeight: 700 }}
        />
        <div style={{ display: 'flex', gap: 8, marginTop: 4, marginBottom: 14 }}>
          <div style={{ flex: 1, border: `1px solid ${shortfall > 0 ? TOMATO : LINE}`, background: shortfall > 0 ? '#FBEAE3' : '#fff', borderRadius: 8, padding: '8px 10px' }}>
            <div style={{ fontSize: 10, color: shortfall > 0 ? TOMATO : MUTED, fontWeight: 700 }}>SHORTFALL</div>
            <div style={{ fontSize: 15, fontWeight: 800, color: shortfall > 0 ? TOMATO : LEAF }}>{shortfall} pack{shortfall !== 1 ? 's' : ''}</div>
          </div>
          {excess > 0 && (
            <div style={{ flex: 1, border: `1px solid ${AMBER}`, background: '#FBEFDC', borderRadius: 8, padding: '8px 10px' }}>
              <div style={{ fontSize: 10, color: AMBER, fontWeight: 700 }}>EXTRA</div>
              <div style={{ fontSize: 15, fontWeight: 800, color: AMBER }}>{excess} pack{excess !== 1 ? 's' : ''}</div>
            </div>
          )}
        </div>
        <PrimaryBtn onClick={save}>Save packed quantity</PrimaryBtn>
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

function DispatchTab({ orders, crates, dispatchLog, onAdvance, onDispatchBatch }) {
  const pending = orders.filter((o) => o.status === 'pending');
  const packed = useMemo(() => orders
    .filter((o) => o.status === 'packed')
    .map((o) => ({ ...o, remaining: Math.max(0, Math.round((o.qty - (o.dispatchedQty || 0) - (o.shortQty || 0)) * 100) / 100) })),
  [orders]);
  const dispatched = orders.filter((o) => o.status === 'dispatched');

  const [view, setView] = useState('dispatch'); // 'dispatch' | 'history' | 'all'
  const [selected, setSelected] = useState([]);
  const [dispatchQtyById, setDispatchQtyById] = useState({});
  const [shortQtyById, setShortQtyById] = useState({});
  const [showModal, setShowModal] = useState(false);
  const [awaitingOpen, setAwaitingOpen] = useState(false);

  const dispatchQtyFor = (o) => dispatchQtyById[o.id] !== undefined ? dispatchQtyById[o.id] : String(o.remaining);
  const shortQtyFor = (o) => shortQtyById[o.id] !== undefined ? shortQtyById[o.id] : '';

  const toggleSelect = (id) => setSelected((p) => (p.includes(id) ? p.filter((x) => x !== id) : [...p, id]));

  const submitDispatch = ({ vehicleNo, driverName, cratesUsed, boxesUsed }) => {
    if (selected.length === 0) return;
    const dispatchItems = selected.map((id) => {
      const o = packed.find((x) => x.id === id);
      return {
        orderId: id,
        dispatchQty: dispatchQtyById[id] !== undefined ? dispatchQtyById[id] : o?.remaining,
        shortQty: shortQtyById[id] || 0,
      };
    });
    onDispatchBatch({ items: dispatchItems, vehicleNo, driverName, cratesUsed, boxesUsed });
    setSelected([]); setDispatchQtyById({}); setShortQtyById({}); setShowModal(false);
  };

  return (
    <div style={{ padding: 16 }}>
      <div style={{ display: 'flex', gap: 6, marginBottom: 14, flexWrap: 'wrap' }}>
        <Chip label="Dispatch" active={view === 'dispatch'} onClick={() => setView('dispatch')} />
        <Chip label={`History (${dispatchLog.length})`} active={view === 'history'} onClick={() => setView('history')} />
        <Chip label={`All dispatched (${dispatched.length})`} active={view === 'all'} onClick={() => setView('all')} />
      </div>

      {view === 'dispatch' && (
        <>
          {pending.length > 0 && (
            <Card style={{ marginBottom: 14 }}>
              <div onClick={() => setAwaitingOpen((x) => !x)} style={{ display: 'flex', alignItems: 'center', justifyContent: 'space-between', cursor: 'pointer' }}>
                <div style={sectionTitle}>Awaiting packing</div>
                <div style={{ display: 'flex', alignItems: 'center', gap: 8 }}>
                  <span style={{ background: '#FBEFDC', color: AMBER, fontWeight: 800, fontSize: 12, padding: '3px 9px', borderRadius: 999 }}>{pending.length}</span>
                  <ChevronRight size={15} color={MUTED} style={{ transform: awaitingOpen ? 'rotate(90deg)' : 'none' }} />
                </div>
              </div>
              {awaitingOpen && pending.map((o) => (
                <div key={o.id} style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', borderTop: `1px solid ${LINE}`, padding: '8px 0' }}>
                  <div><div style={{ fontWeight: 700, fontSize: 13 }}>{o.articleName || o.product}</div><div style={{ fontSize: 11, color: MUTED }}>{o.id} · {o.qty} {o.unit}</div></div>
                  <button onClick={() => onAdvance(o.id, 'packed')} style={{ background: '#E6F1FB', color: '#1B5E8C', border: 'none', borderRadius: 8, padding: '6px 10px', fontSize: 11, fontWeight: 700, cursor: 'pointer' }}>Mark packed</button>
                </div>
              ))}
            </Card>
          )}

          <Card style={{ marginBottom: 14 }}>
            <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'flex-start', gap: 8 }}>
              <div style={{ flex: 1 }}>
                <div style={sectionTitle}>Packed — ready ({packed.length})</div>
                <div style={hint}>Dispatch qty defaults to what's left — lower it if only part is going now. Whatever isn't dispatched stays "packed" for next trip, unless marked short.</div>
              </div>
            </div>
            {packed.map((o) => (
              <div key={o.id} style={{ borderTop: `1px solid ${LINE}`, padding: '9px 0' }}>
                <div style={{ display: 'flex', alignItems: 'center', gap: 10 }}>
                  <div onClick={() => toggleSelect(o.id)} style={{ width: 18, height: 18, borderRadius: 4, border: `1.5px solid ${LINE}`, background: selected.includes(o.id) ? LEAF : '#fff', cursor: 'pointer', flexShrink: 0 }} />
                  <div style={{ flex: 1, minWidth: 0 }}>
                    <div style={{ fontWeight: 700, fontSize: 13 }}>{o.articleName || o.product}</div>
                    <div style={{ fontSize: 11, color: MUTED }}>{o.id} · Remaining {o.remaining} {o.unit}</div>
                  </div>
                </div>
                <div style={{ display: 'flex', gap: 8, marginTop: 6, marginLeft: 28 }}>
                  <div style={{ flex: 1 }}>
                    <div style={{ fontSize: 9, color: MUTED, fontWeight: 700 }}>DISPATCH QTY</div>
                    <input
                      type="number"
                      value={dispatchQtyFor(o)}
                      onChange={(e) => setDispatchQtyById((p) => ({ ...p, [o.id]: e.target.value }))}
                      style={{ width: '100%', boxSizing: 'border-box', borderRadius: 6, border: `1px solid ${LINE}`, fontSize: 12, padding: '6px 8px' }}
                    />
                  </div>
                  <div style={{ flex: 1 }}>
                    <div style={{ fontSize: 9, color: MUTED, fontWeight: 700 }}>SHORT QTY</div>
                    <input
                      type="number"
                      placeholder="0"
                      value={shortQtyFor(o)}
                      onChange={(e) => setShortQtyById((p) => ({ ...p, [o.id]: e.target.value }))}
                      style={{ width: '100%', boxSizing: 'border-box', borderRadius: 6, border: `1px solid ${Number(shortQtyFor(o)) > 0 ? TOMATO : LINE}`, fontSize: 12, padding: '6px 8px', color: Number(shortQtyFor(o)) > 0 ? TOMATO : INK }}
                    />
                  </div>
                </div>
              </div>
            ))}
            {packed.length === 0 && <div style={hint}>Nothing packed yet.</div>}
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

      {view === 'all' && (
        <Card>
          <div style={sectionTitle}>All dispatched ({dispatched.length})</div>
          {dispatched.map((o) => (
            <div key={o.id} style={{ borderTop: `1px solid ${LINE}`, padding: '8px 0' }}>
              <div style={{ fontWeight: 700, fontSize: 13 }}>{o.articleName || o.product}</div>
              <div style={{ fontSize: 11, color: MUTED }}>{o.qty} {o.unit} · {o.id}{o.shortQty > 0 ? ` · ${o.shortQty} ${o.unit} short` : ''}</div>
            </div>
          ))}
          {dispatched.length === 0 && <div style={hint}>No dispatched orders yet.</div>}
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
