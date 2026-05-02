// ============================================================
// TAZWEED HUB — app.js
// Loads product data from GitHub-hosted Price.xlsx (via SheetJS)
// Falls back to demo data if file not available yet
// ============================================================

/* ───────── CONFIG ───────── */
// 🔧 UPDATE THIS URL after uploading Price.xlsx to GitHub:
// Example: 'https://raw.githubusercontent.com/YOUR_USER/YOUR_REPO/main/Price.xlsx'
const EXCEL_URL = 'https://raw.githubusercontent.com/YOUR_USERNAME/YOUR_REPO/main/Price.xlsx';

// Column mapping from your Excel (0-indexed):
// A=0: S (category code)
// B=1: Code
// C=2: Name
// D=3: Price (SAR)
// E=4: Offer 20%
// F=5: Image URL

/* ───────── STATE ───────── */
let allProducts = [];
let cart = [];
let currentCategory = 'all';
let ymmFilter = null;

/* ───────── CATEGORY MAP ───────── */
// Maps the "S" column value to internal category keys
const CATEGORY_MAP = {
  'E': 'engine',
  'C': 'cooling',
  'S': 'suspension',
  'P': 'power',
  'A': 'aero',
  'M': 'maintenance',
  // Add more mappings based on your actual data
};

const CATEGORY_ICONS = {
  engine: '⚙️',
  cooling: '🌡️',
  suspension: '🔩',
  power: '💨',
  aero: '🚗',
  maintenance: '🔧',
  default: '📦',
};

/* ───────── DEMO DATA (fallback) ───────── */
const DEMO_PRODUCTS = [
  { S: 'E', Code: 'TH-CA18-TK01', Name: 'CA18DET Trigger Kit — نظام حساسات المحرك', Price: 850, Offer: 680, Image: '', category: 'engine' },
  { S: 'E', Code: 'TH-RB25-TK01', Name: 'RB25DET Trigger Kit — طقم الحساسات الكامل', Price: 950, Offer: 760, Image: '', category: 'engine' },
  { S: 'C', Code: 'TH-RAD-SR01', Name: 'Aluminium Racing Radiator — رادياتير ألمونيوم رياضي', Price: 1200, Offer: 960, Image: '', category: 'cooling' },
  { S: 'C', Code: 'TH-INT-CA01', Name: 'Front Mount Intercooler Kit — مبرد هواء أمامي', Price: 1800, Offer: 1440, Image: '', category: 'cooling' },
  { S: 'C', Code: 'TH-OC-RB01', Name: 'Oil Cooler Kit RB — طقم تبريد الزيت', Price: 650, Offer: 520, Image: '', category: 'cooling' },
  { S: 'S', Code: 'TH-AK-S13-01', Name: 'GKTech Angle Kit S13/S14 — طقم زوايا التوجيه', Price: 1650, Offer: 1320, Image: '', category: 'suspension' },
  { S: 'S', Code: 'TH-HDK-RB01', Name: 'Head Drain Kit RB — طقم تصريف زيت الرأس', Price: 420, Offer: 336, Image: '', category: 'engine' },
  { S: 'P', Code: 'TH-TURBO-SR01', Name: 'GT2871 Turbocharger — شاحن توربيني', Price: 4500, Offer: 3600, Image: '', category: 'power' },
  { S: 'P', Code: 'TH-BOOST-CTL01', Name: 'Electronic Boost Controller — وحدة تحكم البوست', Price: 380, Offer: 304, Image: '', category: 'power' },
  { S: 'A', Code: 'TH-BKIT-S14-01', Name: 'S14 Silvia Vertex Body Kit — طقم هيكل خارجي', Price: 3800, Offer: 3040, Image: '', category: 'aero' },
  { S: 'A', Code: 'TH-WHEEL-TE37-01', Name: 'TE37 Style Wheels 17" — جنوط رياضية', Price: 2800, Offer: 2240, Image: '', category: 'aero' },
  { S: 'M', Code: 'TH-FLT-KN01', Name: 'K&N High Flow Air Filter — فلتر هواء رياضي', Price: 280, Offer: 224, Image: '', category: 'maintenance' },
  { S: 'M', Code: 'TH-BBK-S13-01', Name: 'ICOOH Big Brake Kit S13 — طقم مكابح كبيرة', Price: 2200, Offer: 1760, Image: '', category: 'maintenance' },
  { S: 'M', Code: 'TH-PLUG-NGK01', Name: 'NGK Iridium Spark Plugs — شمعات احتراق إيريديوم', Price: 180, Offer: 144, Image: '', category: 'maintenance' },
];

/* ───────── INIT ───────── */
window.addEventListener('DOMContentLoaded', () => {
  initNavbar();
  loadProducts();
  loadCart();
});

/* ───────── NAVBAR SCROLL ───────── */
function initNavbar() {
  const nav = document.getElementById('navbar');
  window.addEventListener('scroll', () => {
    nav.classList.toggle('scrolled', window.scrollY > 50);
  });
  document.getElementById('hamburger').addEventListener('click', () => {
    document.getElementById('mobile-menu').classList.toggle('open');
  });
}

/* ───────── LOAD PRODUCTS FROM EXCEL ───────── */
async function loadProducts() {
  const loadingEl = document.getElementById('loading-products');

  // Try to load SheetJS from CDN, then fetch Excel
  try {
    await loadSheetJS();
    const products = await fetchExcel();
    if (products && products.length > 0) {
      allProducts = products;
      renderProducts(allProducts);
      loadingEl.style.display = 'none';
      return;
    }
  } catch (e) {
    console.warn('Could not load Excel from GitHub, using demo data:', e.message);
  }

  // Fallback to demo data
  allProducts = DEMO_PRODUCTS.map(p => ({ ...p }));
  renderProducts(allProducts);
  loadingEl.style.display = 'none';

  // Show note if on demo
  showDemoBanner();
}

function loadSheetJS() {
  return new Promise((resolve, reject) => {
    if (window.XLSX) { resolve(); return; }
    const s = document.createElement('script');
    s.src = 'https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.18.5/xlsx.full.min.js';
    s.onload = resolve;
    s.onerror = () => reject(new Error('SheetJS CDN failed'));
    document.head.appendChild(s);
  });
}

async function fetchExcel() {
  if (EXCEL_URL.includes('YOUR_USERNAME')) {
    throw new Error('GitHub URL not configured yet');
  }
  const res = await fetch(EXCEL_URL);
  if (!res.ok) throw new Error(`HTTP ${res.status}`);
  const ab = await res.arrayBuffer();
  const wb = XLSX.read(ab, { type: 'array' });
  const sheet = wb.Sheets[wb.SheetNames[0]];
  const rows = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: '' });

  // rows[0] = headers, rows[1+] = data
  const products = [];
  for (let i = 1; i < rows.length; i++) {
    const r = rows[i];
    if (!r[2]) continue; // skip empty rows (no name)
    const catCode = String(r[0] || '').trim().toUpperCase();
    products.push({
      S: catCode,
      Code: String(r[1] || '').trim(),
      Name: String(r[2] || '').trim(),
      Price: parseFloat(r[3]) || 0,
      Offer: parseFloat(r[4]) || 0,
      Image: String(r[5] || '').trim(),
      category: CATEGORY_MAP[catCode] || 'maintenance',
    });
  }
  return products;
}

function showDemoBanner() {
  const banner = document.createElement('div');
  banner.className = 'demo-banner';
  banner.innerHTML = `
    <span>⚠️ يتم عرض بيانات تجريبية — قم برفع ملف Price.xlsx على GitHub وتحديث رابط EXCEL_URL في app.js</span>
  `;
  banner.style.cssText = `
    background: rgba(255,184,0,0.1);
    border: 1px solid #FFB800;
    color: #FFB800;
    text-align: center;
    padding: 10px 20px;
    font-size: 0.82rem;
    margin-bottom: 24px;
    border-radius: 8px;
  `;
  document.getElementById('products-grid').parentElement.insertBefore(
    banner,
    document.getElementById('products-grid')
  );
}

/* ───────── RENDER PRODUCTS ───────── */
function renderProducts(products) {
  const grid = document.getElementById('products-grid');
  const noEl = document.getElementById('no-products');

  const filtered = products.filter(p => {
    if (currentCategory !== 'all' && p.category !== currentCategory) return false;
    return true;
  });

  if (filtered.length === 0) {
    grid.innerHTML = '';
    noEl.style.display = 'block';
    return;
  }
  noEl.style.display = 'none';

  grid.innerHTML = filtered.map((p, idx) => {
    const icon = CATEGORY_ICONS[p.category] || CATEGORY_ICONS.default;
    const originalPrice = p.Price;
    const offerPrice = p.Offer || (p.Price * 0.8);
    const hasImage = p.Image && p.Image.startsWith('http');

    return `
      <div class="product-card" data-category="${p.category}" data-idx="${idx}">
        <div class="product-image-wrap">
          ${hasImage
            ? `<img src="${p.Image}" alt="${escHtml(p.Name)}" loading="lazy" onerror="this.style.display='none';this.nextElementSibling.style.display='flex'" /><div class="product-placeholder" style="display:none">${icon}</div>`
            : `<div class="product-placeholder">${icon}</div>`
          }
          <span class="offer-tag">خصم 20%</span>
        </div>
        <div class="product-body">
          <div class="product-code">${escHtml(p.Code || p.S + '-' + String(idx + 1).padStart(3, '0'))}</div>
          <div class="product-name">${escHtml(p.Name)}</div>
          <div class="product-pricing">
            <span class="price-offer">${formatNum(offerPrice)}</span>
            <span class="price-unit">ر.س</span>
            <span class="price-original">${formatNum(originalPrice)} ر.س</span>
          </div>
          <button class="add-to-cart-btn" onclick='addToCart(${JSON.stringify({ ...p, offerPrice })})'>
            🛒 أضف للسلة
          </button>
        </div>
      </div>
    `;
  }).join('');
}

/* ───────── FILTER ───────── */
function filterCategory(cat) {
  currentCategory = cat;
  // Update active button
  document.querySelectorAll('.filter-btn').forEach(b => b.classList.remove('active'));
  event?.target?.classList.add('active');
  renderProducts(allProducts);
}

/* ───────── YMM FILTER ───────── */
function filterYMM() {
  // For now just store the selection; full fitment DB needs a real backend
}

function applyYMM() {
  const year = document.getElementById('ymm-year').value;
  const make = document.getElementById('ymm-make').value;
  const model = document.getElementById('ymm-model').value;

  document.getElementById('ymm-modal').style.display = 'none';

  if (!year && !make && !model) return;

  // Simple filter demo: show all products and scroll to section
  // In production, filter by fitment database
  const label = [year, make, model].filter(Boolean).join(' ');
  const banner = document.createElement('div');
  banner.style.cssText = `
    background: rgba(255,69,0,0.1);
    border: 1px solid var(--orange);
    color: #FF4500;
    text-align: center;
    padding: 10px 20px;
    font-size: 0.9rem;
    margin-bottom: 20px;
    border-radius: 8px;
    font-weight: 600;
  `;
  banner.innerHTML = `🔧 عرض القطع المتوافقة مع: <strong>${escHtml(label)}</strong> — <a href="#" onclick="clearYMM()" style="color:inherit">مسح الفلتر ×</a>`;
  banner.id = 'ymm-banner';
  const existing = document.getElementById('ymm-banner');
  if (existing) existing.remove();
  document.getElementById('products').querySelector('.filter-bar').after(banner);

  document.getElementById('products').scrollIntoView({ behavior: 'smooth' });
}

function clearYMM() {
  document.getElementById('ymm-banner')?.remove();
}

/* ───────── CART ───────── */
function addToCart(product) {
  const existing = cart.find(i => i.Code === product.Code);
  if (existing) {
    existing.qty = (existing.qty || 1) + 1;
  } else {
    cart.push({ ...product, qty: 1 });
  }
  saveCart();
  renderCart();
  openCart();

  // Button feedback
  event?.target && animateButton(event.target);
}

function animateButton(btn) {
  const orig = btn.textContent;
  btn.textContent = '✅ تمت الإضافة!';
  btn.style.background = '#22c55e';
  btn.style.borderColor = '#22c55e';
  btn.style.color = 'white';
  setTimeout(() => {
    btn.textContent = orig;
    btn.style.background = '';
    btn.style.borderColor = '';
    btn.style.color = '';
  }, 1500);
}

function removeFromCart(idx) {
  cart.splice(idx, 1);
  saveCart();
  renderCart();
}

function renderCart() {
  const itemsEl = document.getElementById('cart-items');
  const footerEl = document.getElementById('cart-footer');
  const countEl = document.getElementById('cart-count');
  const totalEl = document.getElementById('cart-total');

  const total = cart.reduce((sum, i) => sum + (i.offerPrice * (i.qty || 1)), 0);
  const count = cart.reduce((sum, i) => sum + (i.qty || 1), 0);

  countEl.textContent = count;

  if (cart.length === 0) {
    itemsEl.innerHTML = '<p class="cart-empty">سلتك فارغة</p>';
    footerEl.style.display = 'none';
    return;
  }

  footerEl.style.display = 'block';
  totalEl.textContent = formatNum(total);

  itemsEl.innerHTML = cart.map((item, idx) => `
    <div class="cart-item">
      <div class="cart-item-info">
        <div class="cart-item-name">${escHtml(item.Name)}</div>
        <div class="cart-item-price">${formatNum(item.offerPrice)} ر.س × ${item.qty || 1}</div>
      </div>
      <button class="cart-item-remove" onclick="removeFromCart(${idx})">🗑️</button>
    </div>
  `).join('');
}

function openCart() {
  document.getElementById('cart-sidebar').classList.add('open');
  document.getElementById('cart-overlay').classList.add('open');
  document.body.style.overflow = 'hidden';
}

function toggleCart() {
  const sidebar = document.getElementById('cart-sidebar');
  const overlay = document.getElementById('cart-overlay');
  const isOpen = sidebar.classList.contains('open');
  sidebar.classList.toggle('open');
  overlay.classList.toggle('open');
  document.body.style.overflow = isOpen ? '' : 'hidden';
}

function checkout() {
  if (cart.length === 0) return;
  // Build order summary for WhatsApp
  const lines = cart.map(i =>
    `• ${i.Name} (${i.Code}) × ${i.qty} = ${formatNum(i.offerPrice * (i.qty||1))} ر.س`
  ).join('\n');
  const total = cart.reduce((s, i) => s + i.offerPrice * (i.qty||1), 0);
  const msg = `طلب جديد من موقع Tazweed Hub 🔧\n\n${lines}\n\nالإجمالي: ${formatNum(total)} ر.س\n\nيرجى التواصل لتأكيد الطلب ✅`;

  // Open WhatsApp (update number)
  const waUrl = `https://wa.me/966500000000?text=${encodeURIComponent(msg)}`;
  window.open(waUrl, '_blank');

  // Also show confirmation modal
  document.getElementById('cart-sidebar').classList.remove('open');
  document.getElementById('cart-overlay').classList.remove('open');
  document.body.style.overflow = '';
  document.getElementById('order-modal').style.display = 'flex';
  cart = [];
  saveCart();
  renderCart();
}

function saveCart() {
  try { localStorage.setItem('tazweed_cart', JSON.stringify(cart)); } catch(e) {}
}

function loadCart() {
  try {
    const saved = localStorage.getItem('tazweed_cart');
    if (saved) { cart = JSON.parse(saved); renderCart(); }
  } catch(e) {}
}

/* ───────── CONTACT FORM ───────── */
function submitInquiry(e) {
  e.preventDefault();
  document.getElementById('inquiry-modal').style.display = 'flex';
  e.target.reset();
}

/* ───────── MODAL CLOSE ───────── */
function closeModal(e) {
  if (e.target.classList.contains('modal-overlay')) {
    e.target.style.display = 'none';
  }
}

/* ───────── UTILS ───────── */
function formatNum(n) {
  return Number(n).toLocaleString('ar-SA', { minimumFractionDigits: 0, maximumFractionDigits: 0 });
}

function escHtml(str) {
  return String(str)
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;');
}
