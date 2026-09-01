import { useState, useEffect, useRef, useCallback } from "react";

// ── Export Excel Helper ──────────────────────────────────────────────────────
function loadSheetJS(callback) {
  if (window.XLSX) { callback(window.XLSX); return; }
  const script = document.createElement("script");
  script.src = "https://cdn.sheetjs.com/xlsx-0.20.1/package/dist/xlsx.full.min.js";
  script.onload = () => callback(window.XLSX);
  script.onerror = () => alert("Gagal load library Excel. Cek koneksi internet.");
  document.head.appendChild(script);
}

function exportTransaksiExcel(transactions) {
  loadSheetJS(XLSX => {
    const rows = [];
    transactions.forEach(t => {
      t.items.forEach(item => {
        rows.push({
          "ID Transaksi": t.id,
          "Tanggal": new Date(t.date).toLocaleString("id-ID"),
          "Kasir": t.cashier || "-",
          "Nama Produk": item.name,
          "Qty": item.qty,
          "Harga Satuan": item.price,
          "Subtotal Item": item.qty * item.price,
          "Total Transaksi": t.total,
          "Diskon": t.discount || 0,
          "Pajak": t.tax || 0,
          "Metode Bayar": t.payment,
          "Uang Diterima": t.cashPaid,
          "Kembalian": t.change || 0,
          "Catatan": t.note || "",
        });
      });
    });

    const ws = XLSX.utils.json_to_sheet(rows);
    // Auto column width
    const colWidths = Object.keys(rows[0]||{}).map(k => ({wch: Math.max(k.length, 14)}));
    ws["!cols"] = colWidths;

    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "Transaksi");

    // Rekap sheet
    const rekapRows = transactions.map(t => ({
      "ID": t.id,
      "Tanggal": new Date(t.date).toLocaleString("id-ID"),
      "Kasir": t.cashier || "-",
      "Jumlah Item": t.items.reduce((s,x)=>s+x.qty,0),
      "Subtotal": t.subtotal,
      "Diskon": t.discount || 0,
      "Pajak": t.tax || 0,
      "Total": t.total,
      "Metode": t.payment,
    }));
    if (rekapRows.length > 0) {
      const ws2 = XLSX.utils.json_to_sheet(rekapRows);
      XLSX.utils.book_append_sheet(wb, ws2, "Rekap Transaksi");
    }

    const tgl = new Date().toISOString().slice(0,10);
    XLSX.writeFile(wb, `transaksi-kasir-${tgl}.xlsx`);
  });
}

function exportProdukExcel(products) {
  loadSheetJS(XLSX => {
    const rows = products.map(p => ({
      "SKU": p.sku,
      "Barcode": p.barcode,
      "Nama Produk": p.name,
      "Kategori": p.category,
      "Harga": p.price,
      "Stok": p.stock,
    }));

    const ws = XLSX.utils.json_to_sheet(rows);
    ws["!cols"] = [{wch:10},{wch:16},{wch:24},{wch:12},{wch:12},{wch:8}];
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "Produk");
    XLSX.writeFile(wb, `produk-kasir-${new Date().toISOString().slice(0,10)}.xlsx`);
  });
}

function exportPembelianExcel(pembelian) {
  loadSheetJS(XLSX => {
    const rows = pembelian.map(p => ({
      "Tanggal": p.tanggal,
      "Jenis": p.jenis,
      "Jumlah": p.jumlah,
      "Satuan": p.satuan,
      "Harga Satuan": p.harga,
      "Total": p.harga * p.jumlah,
      "Tempat Beli": p.tempat,
      "Catatan": p.catatan || "",
    }));

    const ws = XLSX.utils.json_to_sheet(rows.length > 0 ? rows : [{"Info":"Belum ada data pembelian"}]);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "Pembelian");
    XLSX.writeFile(wb, `pembelian-${new Date().toISOString().slice(0,10)}.xlsx`);
  });
}

function exportPenjualanExcel(penjualan) {
  loadSheetJS(XLSX => {
    const rows = penjualan.map(p => ({
      "Tanggal": p.tanggal,
      "Produk": p.produk,
      "Jumlah": p.jumlah,
      "Satuan": p.satuan,
      "Harga Satuan": p.harga,
      "Total": p.harga * p.jumlah,
      "Catatan": p.catatan || "",
    }));

    const ws = XLSX.utils.json_to_sheet(rows.length > 0 ? rows : [{"Info":"Belum ada data penjualan"}]);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "Penjualan");
    XLSX.writeFile(wb, `penjualan-${new Date().toISOString().slice(0,10)}.xlsx`);
  });
}

// ── Data ────────────────────────────────────────────────────────────────────
const CATEGORY_ICONS = { Makanan:"Makanan", Minuman:"Minuman", Snack:"Snack", Lainnya:"Lainnya", Semua:"Semua" };
const CATEGORY_MAT_ICONS = { Makanan:"restaurant", Minuman:"local_cafe", Snack:"fastfood", Lainnya:"category", Semua:"apps" };
const CATEGORY_COLORS = { Makanan:"#FF6B35", Minuman:"#0EA5E9", Snack:"#A855F7", Lainnya:"#6B7280" };

const INITIAL_PRODUCTS = [
  { id:1, name:"Nasi Goreng Spesial", price:25000, category:"Makanan", stock:50, sku:"MKN001", barcode:"8991234000001" },
  { id:2, name:"Mie Ayam Bakso", price:18000, category:"Makanan", stock:45, sku:"MKN002", barcode:"8991234000002" },
  { id:3, name:"Ayam Bakar", price:32000, category:"Makanan", stock:30, sku:"MKN003", barcode:"8991234000003" },
  { id:4, name:"Soto Ayam", price:15000, category:"Makanan", stock:40, sku:"MKN004", barcode:"8991234000004" },
  { id:5, name:"Gado-gado", price:14000, category:"Makanan", stock:35, sku:"MKN005", barcode:"8991234000005" },
  { id:6, name:"Es Teh Manis", price:5000, category:"Minuman", stock:100, sku:"MNM001", barcode:"8991234000006" },
  { id:7, name:"Es Jeruk", price:8000, category:"Minuman", stock:80, sku:"MNM002", barcode:"8991234000007" },
  { id:8, name:"Jus Alpukat", price:15000, category:"Minuman", stock:60, sku:"MNM003", barcode:"8991234000008" },
  { id:9, name:"Kopi Hitam", price:6000, category:"Minuman", stock:90, sku:"MNM004", barcode:"8991234000009" },
  { id:10, name:"Air Mineral", price:3000, category:"Minuman", stock:200, sku:"MNM005", barcode:"8991234000010" },
  { id:11, name:"Kerupuk", price:2000, category:"Snack", stock:150, sku:"SNK001", barcode:"8991234000011" },
  { id:12, name:"Pisang Goreng", price:8000, category:"Snack", stock:7, sku:"SNK002", barcode:"8991234000012" },
  { id:13, name:"Tempe Mendoan", price:6000, category:"Snack", stock:80, sku:"SNK003", barcode:"8991234000013" },
  { id:14, name:"Teh Botol", price:5000, category:"Minuman", stock:120, sku:"MNM006", barcode:"8991234000014" },
  { id:15, name:"Nasi Putih", price:4000, category:"Makanan", stock:200, sku:"MKN006", barcode:"8991234000015" },
  { id:16, name:"Bakso Kuah", price:20000, category:"Makanan", stock:0, sku:"MKN007", barcode:"8991234000016" },
];

const INITIAL_TRANSACTIONS = [
  { id:"TRX-20240618-001", date:"2024-06-18T08:30:00", items:[{name:"Nasi Goreng Spesial",qty:2,price:25000},{name:"Es Teh Manis",qty:2,price:5000}], subtotal:60000, discount:0, tax:6000, total:66000, payment:"cash", cashPaid:70000, change:4000, cashier:"Budi" },
  { id:"TRX-20240618-002", date:"2024-06-18T09:15:00", items:[{name:"Mie Ayam Bakso",qty:1,price:18000},{name:"Kopi Hitam",qty:1,price:6000}], subtotal:24000, discount:2400, tax:2160, total:23760, payment:"qris", cashPaid:23760, change:0, cashier:"Sari" },
  { id:"TRX-20240618-003", date:"2024-06-18T10:00:00", items:[{name:"Ayam Bakar",qty:3,price:32000},{name:"Nasi Putih",qty:3,price:4000},{name:"Es Jeruk",qty:3,price:8000}], subtotal:132000, discount:13200, tax:11880, total:130680, payment:"debit", cashPaid:130680, change:0, cashier:"Budi" },
];

const CATEGORIES = ["Semua","Makanan","Minuman","Snack","Lainnya"];
const PAYMENT_METHODS = ["cash","debit","kredit","qris","transfer"];

const formatRp = (n) => new Intl.NumberFormat("id-ID",{style:"currency",currency:"IDR",minimumFractionDigits:0}).format(n);
const formatDate = (d) => new Date(d).toLocaleString("id-ID",{dateStyle:"short",timeStyle:"short"});
const genId = () => `TRX-${new Date().toISOString().slice(0,10).replace(/-/g,"")}-${String(Math.floor(Math.random()*900)+100)}`;

// ── Material Design 3 CSS ────────────────────────────────────────────────────
const css = `
  @import url('https://fonts.googleapis.com/css2?family=Nunito:wght@400;500;600;700;800&family=Roboto:wght@300;400;500;700&family=JetBrains+Mono:wght@400;500&display=swap');
  @import url('https://fonts.googleapis.com/icon?family=Material+Icons+Round');

  *, *::before, *::after { box-sizing:border-box; margin:0; padding:0; }

  /* ── Light Mode ── */
  :root[data-theme="light"] {
    --md-primary:#0061A4;
    --md-on-primary:#ffffff;
    --md-primary-container:#D3E3FD;
    --md-on-primary-container:#001D35;
    --md-secondary:#535F70;
    --md-secondary-container:#D7E3F7;
    --md-on-secondary-container:#101C2B;
    --md-tertiary:#6B5778;
    --md-tertiary-container:#F2DAFF;
    --md-surface:#F8F9FA;
    --md-surface-1:#EEF1F8;
    --md-surface-2:#E8ECF4;
    --md-surface-3:#E2E7F0;
    --md-surface-4:#E0E5EF;
    --md-surface-5:#DCE2EC;
    --md-on-surface:#1A1C1E;
    --md-on-surface-variant:#44474F;
    --md-outline:#74777F;
    --md-outline-variant:#C4C6D0;
    --md-error:#BA1A1A;
    --md-error-container:#FFDAD6;
    --md-on-error-container:#410002;
    --md-success:#006E46;
    --md-success-container:#91F4C3;
    --md-scrim:rgba(0,0,0,.4);
    --elevation-1:0 1px 2px rgba(0,0,0,.15),0 1px 3px 1px rgba(0,0,0,.1);
    --elevation-2:0 1px 2px rgba(0,0,0,.15),0 2px 6px 2px rgba(0,0,0,.1);
    --elevation-3:0 4px 8px 3px rgba(0,0,0,.1),0 1px 3px rgba(0,0,0,.15);
    --elevation-4:0 6px 10px 4px rgba(0,0,0,.1),0 2px 3px rgba(0,0,0,.15);
  }

  :root {
    /* M3 Color Tokens - Dark scheme */
    --md-primary:#C2E7FF;
    --md-on-primary:#003352;
    --md-primary-container:#004A77;
    --md-on-primary-container:#C2E7FF;
    --md-secondary:#B8C8DA;
    --md-secondary-container:#3A4857;
    --md-on-secondary-container:#D4E4F6;
    --md-tertiary:#C8BFF0;
    --md-tertiary-container:#4A4178;
    --md-surface:#111315;
    --md-surface-1:#1A1D20;
    --md-surface-2:#1F2327;
    --md-surface-3:#252A2E;
    --md-surface-4:#272C30;
    --md-surface-5:#2A2F34;
    --md-on-surface:#E2E8F0;
    --md-on-surface-variant:#9AABB8;
    --md-outline:#4A5568;
    --md-outline-variant:#2D3748;
    --md-error:#FFB4AB;
    --md-error-container:#93000A;
    --md-on-error-container:#FFDAD6;
    --md-success:#6EE7B7;
    --md-success-container:#065F46;
    --md-scrim:rgba(0,0,0,.6);

    /* Elevation shadows M3 */
    --elevation-1:0 1px 2px rgba(0,0,0,.3),0 1px 3px 1px rgba(0,0,0,.15);
    --elevation-2:0 1px 2px rgba(0,0,0,.3),0 2px 6px 2px rgba(0,0,0,.15);
    --elevation-3:0 4px 8px 3px rgba(0,0,0,.15),0 1px 3px rgba(0,0,0,.3);
    --elevation-4:0 6px 10px 4px rgba(0,0,0,.15),0 2px 3px rgba(0,0,0,.3);

    /* Shape */
    --shape-xs:4px;
    --shape-sm:8px;
    --shape-md:12px;
    --shape-lg:16px;
    --shape-xl:28px;
    --shape-full:100px;

    /* Typography */
    --font-brand:'Nunito','Roboto',sans-serif;
    --font-body:'Roboto',sans-serif;
    --font-mono:'JetBrains Mono',monospace;

    /* State layers */
    --state-hover:rgba(194,231,255,.08);
    --state-pressed:rgba(194,231,255,.12);
    --state-focused:rgba(194,231,255,.12);
  }

  html,body,#root { height:100%; font-family:var(--font-body); background:var(--md-surface); color:var(--md-on-surface); overflow:hidden; -webkit-font-smoothing:antialiased; transition:background-color .3s, color .3s; }

  /* ════════════════════════════════════════
     APP SHELL
  ════════════════════════════════════════ */
  .app { display:flex; flex-direction:column; height:100vh; height:100dvh; overflow:hidden; background:var(--md-surface); transition:background .25s; }
  .app, .top-app-bar, .search-section, .filter-chips-wrap, .m3-bottom-nav, .pos-right, .cart-top, .note-wrap, .cart-items-list, .cart-totals-wrap, .cart-actions-wrap, .m3-settings-section, .m3-card, .m3-stat-card, .page-content, .m3-product-card-wrap { transition:background-color .25s, border-color .25s; }
  .content { flex:1; overflow:hidden; position:relative; background:var(--md-surface); transition:background .25s; }

  /* ════════════════════════════════════════
     M3 TOP APP BAR
  ════════════════════════════════════════ */
  .top-app-bar {
    display:flex; align-items:center; justify-content:space-between;
    padding:0 4px 0 16px; height:64px;
    background:var(--md-surface-2);
    border-bottom:1px solid var(--md-outline-variant);
    flex-shrink:0;
    position:relative;
    z-index:10;
  }
  .top-app-bar .brand { display:flex; align-items:center; gap:12px; }
  .top-app-bar .brand img { width:40px; height:40px; border-radius:50%; object-fit:cover; }
  .top-app-bar .brand-text { display:flex; flex-direction:column; }
  .top-app-bar .brand-title { font-family:var(--font-brand); font-size:18px; font-weight:600; color:var(--md-on-surface); line-height:1.2; }
  .top-app-bar .brand-sub { font-size:11px; color:var(--md-on-surface-variant); }
  .top-app-bar .actions { display:flex; align-items:center; gap:4px; }
  .icon-btn-top { width:48px; height:48px; border-radius:50%; border:none; background:transparent; color:var(--md-on-surface-variant); cursor:pointer; display:flex; align-items:center; justify-content:center; font-size:20px; transition:background .2s; -webkit-tap-highlight-color:transparent; }
  .icon-btn-top:hover { background:var(--state-hover); }
  .icon-btn-top:active { background:var(--state-pressed); }
  .cashier-chip { display:flex; align-items:center; gap:6px; padding:6px 12px 6px 8px; border-radius:var(--shape-full); background:var(--md-primary-container); border:none; cursor:pointer; }
  .cashier-avatar { width:24px; height:24px; border-radius:50%; background:var(--md-primary); color:var(--md-on-primary); font-size:11px; font-weight:700; display:flex; align-items:center; justify-content:center; font-family:var(--font-brand); }
  .cashier-name { font-size:12px; font-weight:500; color:var(--md-on-primary-container); font-family:var(--font-brand); }

  /* Desktop nav pills */
  .desktop-nav { display:flex; gap:2px; }
  .dnav-btn { padding:8px 16px; border-radius:var(--shape-full); border:none; cursor:pointer; font-size:13px; font-weight:500; font-family:var(--font-brand); transition:all .2s; background:transparent; color:var(--md-on-surface-variant); white-space:nowrap; }
  .dnav-btn:hover { background:var(--state-hover); color:var(--md-on-surface); }
  .dnav-btn.active { background:var(--md-secondary-container); color:var(--md-on-secondary-container); }

  /* ════════════════════════════════════════
     M3 SEARCH BAR
  ════════════════════════════════════════ */
  .search-section { padding:10px 12px 8px; background:var(--md-surface-1); border-bottom:1px solid var(--md-outline-variant); flex-shrink:0; }
  .m3-search-bar {
    display:flex; align-items:center; gap:8px;
    padding:0 16px 0 12px;
    height:56px;
    background:var(--md-surface-4);
    border-radius:var(--shape-full);
    box-shadow:var(--elevation-1);
    transition:box-shadow .2s;
  }
  .m3-search-bar:focus-within { box-shadow:var(--elevation-2); }
  .search-icon { color:var(--md-on-surface-variant); font-size:20px; flex-shrink:0; }
  .m3-search-input { flex:1; background:transparent; border:none; outline:none; font-size:16px; color:var(--md-on-surface); font-family:var(--font-body); caret-color:var(--md-primary); }
  .m3-search-input::placeholder { color:var(--md-on-surface-variant); }
  .scan-icon-btn {
    width:40px; height:40px; border-radius:50%; border:none;
    background:var(--md-primary-container);
    color:var(--md-on-primary-container);
    cursor:pointer; font-size:18px;
    display:flex; align-items:center; justify-content:center;
    flex-shrink:0; transition:all .2s;
    -webkit-tap-highlight-color:transparent;
  }
  .scan-icon-btn:active { transform:scale(.92); }

  /* ════════════════════════════════════════
     M3 FILTER CHIPS
  ════════════════════════════════════════ */
  .filter-chips-wrap { display:flex; gap:8px; padding:8px 12px; overflow-x:auto; flex-shrink:0; background:var(--md-surface-1); border-bottom:1px solid var(--md-outline-variant); }
  .filter-chips-wrap::-webkit-scrollbar { display:none; }
  .m3-chip {
    display:flex; align-items:center; gap:6px;
    padding:0 16px; height:32px;
    border-radius:var(--shape-sm);
    border:1px solid var(--md-outline);
    background:transparent;
    color:var(--md-on-surface-variant);
    font-size:14px; font-weight:500;
    font-family:var(--font-brand);
    cursor:pointer; white-space:nowrap; flex-shrink:0;
    transition:all .15s;
    -webkit-tap-highlight-color:transparent;
  }
  .m3-chip:hover { background:var(--state-hover); }
  .m3-chip.active {
    background:var(--md-secondary-container);
    border-color:var(--md-secondary-container);
    color:var(--md-on-secondary-container);
    padding-left:8px;
  }
  .chip-check { font-size:16px; line-height:1; }
  .chip-icon { font-size:14px; }

  /* ════════════════════════════════════════
     POS LAYOUT
  ════════════════════════════════════════ */
  .pos { display:flex; height:100%; overflow:hidden; background:var(--md-surface); transition:background .25s; }
  .pos-left { flex:1; display:flex; flex-direction:column; overflow:hidden; min-width:0; min-height:0; background:var(--md-surface); transition:background .25s; }
  .pos-right { width:360px; display:flex; flex-direction:column; background:var(--md-surface-1); border-left:1px solid var(--md-outline-variant); flex-shrink:0; }

  /* ════════════════════════════════════════
     M3 PRODUCT GRID
  ════════════════════════════════════════ */
  .product-grid {
    flex:1;
    overflow-y:auto;
    overflow-x:hidden;
    -webkit-overflow-scrolling:touch;
    padding:10px;
    display:grid;
    grid-template-columns:repeat(2,1fr);
    gap:8px;
    align-content:start;
    min-height:0;
  }
  .product-grid::-webkit-scrollbar { width:4px; }
  .product-grid::-webkit-scrollbar-thumb { background:var(--md-outline-variant); border-radius:2px; }

  /* M3 Elevated Card */
  .m3-product-card {
    background:var(--md-surface-2);
    border-radius:var(--shape-lg);
    overflow:hidden;
    cursor:pointer;
    transition:box-shadow .2s, transform .15s;
    box-shadow:var(--elevation-1);
    position:relative;
    -webkit-tap-highlight-color:transparent;
    border:1px solid var(--md-outline-variant);
    display:flex;
    flex-direction:column;
  }
  .m3-product-card:hover { box-shadow:var(--elevation-3); transform:translateY(-1px); }
  .m3-product-card:active { box-shadow:var(--elevation-1); transform:scale(.98); }
  .m3-product-card.out-of-stock { opacity:.5; cursor:not-allowed; }
  .m3-product-card.out-of-stock:hover { transform:none; box-shadow:var(--elevation-1); }

  /* Card avatar/image header */
  .card-avatar {
    width:100%;
    height:90px;
    display:flex; align-items:center; justify-content:center;
    font-size:42px; position:relative;
    background:linear-gradient(135deg,var(--md-surface-3),var(--md-surface-4));
    overflow:hidden;
    flex-shrink:0;
  }
  .card-avatar-emoji { 
    font-size:42px; 
    line-height:1;
    display:block;
    user-select:none;
  }

  /* Card body */
  .card-body { padding:8px 10px 4px; }
  .card-name {
    font-size:13px; font-weight:600; font-family:var(--font-brand);
    color:var(--md-on-surface); line-height:1.35;
    overflow:hidden;
    display:-webkit-box; -webkit-line-clamp:2; -webkit-box-orient:vertical;
    min-height:34px; margin-bottom:4px;
    word-break:break-word;
  }
  .card-price {
    font-size:14px; font-weight:700; font-family:var(--font-brand);
    color:#90CAF9; letter-spacing:-.2px;
  }

  /* Card footer */
  .card-footer { padding:4px 10px 8px; display:flex; align-items:center; justify-content:space-between; min-height:40px; }
  .stock-label { font-size:11px; font-weight:500; color:var(--md-on-surface-variant); }
  .stock-label.low { color:var(--md-error); background:rgba(255,180,171,.12); padding:2px 6px; border-radius:var(--shape-full); }
  .stock-label.out { color:var(--md-error); font-weight:600; }

  /* Quick-add button inside card */
  .card-add-btn {
    width:36px; height:36px; border-radius:50%;
    background:var(--md-primary-container);
    color:var(--md-on-primary-container);
    border:none; cursor:pointer; font-size:22px; font-weight:300;
    display:flex; align-items:center; justify-content:center;
    transition:all .15s; flex-shrink:0;
    -webkit-tap-highlight-color:transparent;
    box-shadow:var(--elevation-1);
  }
  .card-add-btn:active { transform:scale(.88); background:var(--md-primary); color:var(--md-on-primary); }

  /* Cart qty badge on card */
  .card-qty-badge {
    position:absolute; top:8px; right:8px;
    background:var(--md-primary); color:var(--md-on-primary);
    border-radius:var(--shape-full); min-width:24px; height:24px;
    font-size:12px; font-weight:700; font-family:var(--font-brand);
    display:flex; align-items:center; justify-content:center;
    padding:0 6px; box-shadow:var(--elevation-2);
  }

  /* Out of stock overlay - only on avatar area */
  .card-out-badge {
    position:absolute; top:0; left:0; right:0; height:90px;
    background:rgba(17,19,21,.65);
    display:flex; align-items:center; justify-content:center;
    border-radius:0;
  }
  .card-out-text { background:var(--md-error-container); color:#FFDAD6; font-size:11px; font-weight:700; padding:4px 10px; border-radius:var(--shape-full); letter-spacing:.5px; }

  /* ════════════════════════════════════════
     FLOATING CART SUMMARY BAR (PRD Component 5)
  ════════════════════════════════════════ */
  .cart-fab-bar {
    position:absolute; bottom:0; left:0; right:0;
    padding:10px 12px;
    background:linear-gradient(to top, var(--md-surface) 85%, transparent);
    z-index:20;
    transition:transform .3s cubic-bezier(.4,0,.2,1), opacity .3s;
    pointer-events:none;
    flex-shrink:0;
  }
  .cart-fab-bar.visible { pointer-events:all; }
  .cart-fab-inner {
    display:flex; align-items:center;
    background:var(--md-primary-container);
    border-radius:var(--shape-xl);
    padding:12px 12px 12px 20px;
    box-shadow:var(--elevation-4);
    gap:12px;
    cursor:pointer;
    -webkit-tap-highlight-color:transparent;
    transition:transform .15s;
  }
  .cart-fab-inner:active { transform:scale(.97); }
  .cart-fab-info { flex:1; }
  .cart-fab-count { font-size:12px; color:var(--md-primary); font-weight:500; font-family:var(--font-brand); }
  .cart-fab-total { font-size:18px; font-weight:700; color:var(--md-on-primary-container); font-family:var(--font-brand); }
  .cart-fab-btn {
    background:var(--md-primary); color:var(--md-on-primary);
    border:none; border-radius:var(--shape-lg);
    padding:12px 20px; font-size:14px; font-weight:600;
    font-family:var(--font-brand); cursor:pointer; white-space:nowrap;
    display:flex; align-items:center; gap:6px;
  }

  /* ════════════════════════════════════════
     M3 CART PANEL (Desktop right / Mobile bottom sheet)
  ════════════════════════════════════════ */
  .cart-top { padding:16px 16px 12px; border-bottom:1px solid var(--md-outline-variant); flex-shrink:0; }
  .cart-top h3 { font-size:18px; font-weight:600; font-family:var(--font-brand); color:var(--md-on-surface); }
  .cart-top h3 span { font-size:13px; font-weight:400; color:var(--md-on-surface-variant); margin-left:8px; }

  /* Note input at top */
  .note-wrap { padding:8px 12px; border-bottom:1px solid var(--md-outline-variant); flex-shrink:0; }
  .m3-note-input {
    width:100%; padding:10px 14px;
    background:var(--md-surface-3);
    border:1px solid var(--md-outline-variant);
    border-radius:var(--shape-md);
    color:var(--md-on-surface); font-size:14px;
    outline:none; font-family:var(--font-body);
    transition:border-color .2s;
  }
  .m3-note-input:focus { border-color:var(--md-primary); }
  .m3-note-input::placeholder { color:var(--md-on-surface-variant); }

  .cart-items-list { flex:1; overflow-y:auto; padding:8px; -webkit-overflow-scrolling:touch; }
  .cart-items-list::-webkit-scrollbar { width:3px; }
  .cart-items-list::-webkit-scrollbar-thumb { background:var(--md-outline-variant); border-radius:2px; }
  .cart-empty-state { display:flex; flex-direction:column; align-items:center; justify-content:center; height:120px; color:var(--md-on-surface-variant); font-size:13px; gap:8px; font-family:var(--font-brand); }
  .cart-empty-icon { font-size:40px; opacity:.3; }

  /* M3 Cart Item */
  .m3-cart-item {
    display:flex; align-items:center; gap:10px;
    padding:10px; border-radius:var(--shape-md);
    margin-bottom:4px; transition:background .15s;
  }
  .m3-cart-item:hover { background:var(--md-surface-3); }
  .cart-item-avatar {
    width:40px; height:40px; border-radius:var(--shape-sm);
    background:var(--md-surface-3);
    display:flex; align-items:center; justify-content:center;
    font-size:20px; flex-shrink:0;
  }
  .cart-item-info { flex:1; min-width:0; }
  .cart-item-name { font-size:13px; font-weight:500; font-family:var(--font-brand); color:var(--md-on-surface); white-space:nowrap; overflow:hidden; text-overflow:ellipsis; }
  .cart-item-price { font-size:11px; color:var(--md-on-surface-variant); margin-top:1px; }

  /* M3 Qty stepper */
  .m3-qty { display:flex; align-items:center; gap:4px; }
  .m3-qty-btn {
    width:32px; height:32px; border-radius:50%;
    border:1px solid var(--md-outline);
    background:transparent; color:var(--md-on-surface);
    cursor:pointer; font-size:18px;
    display:flex; align-items:center; justify-content:center;
    transition:all .15s; -webkit-tap-highlight-color:transparent;
  }
  .m3-qty-btn:hover { background:var(--state-hover); }
  .m3-qty-btn:active { background:var(--state-pressed); border-color:var(--md-primary); }
  .m3-qty-val { font-size:14px; font-weight:600; font-family:var(--font-mono); min-width:24px; text-align:center; color:var(--md-on-surface); }
  .cart-item-subtotal { font-size:13px; font-weight:600; font-family:var(--font-brand); color:var(--md-on-surface); min-width:70px; text-align:right; }
  .m3-del-btn { width:32px; height:32px; border-radius:50%; border:none; background:transparent; color:var(--md-on-surface-variant); cursor:pointer; font-size:14px; display:flex; align-items:center; justify-content:center; transition:all .15s; }
  .m3-del-btn:hover { background:rgba(255,180,171,.15); color:var(--md-error); }

  /* ── Totals ── */
  .cart-totals-wrap { padding:10px 16px; border-top:1px solid var(--md-outline-variant); flex-shrink:0; }
  .total-row { display:flex; justify-content:space-between; font-size:13px; color:var(--md-on-surface-variant); margin-bottom:5px; }
  .total-row .val { font-family:var(--font-mono); }
  .total-row.discount .val { color:var(--md-success); }
  .total-row.tax .val { color:#93C5FD; }
  .total-divider { border:none; border-top:1px solid var(--md-outline-variant); margin:8px 0; }
  .total-row.grand { font-size:17px; font-weight:700; font-family:var(--font-brand); color:var(--md-on-surface); margin-bottom:0; }
  .total-row.grand .val { color:var(--md-primary); }

  /* ── Cart actions ── */
  .cart-actions-wrap { padding:10px 16px 12px; display:flex; flex-direction:column; gap:8px; border-top:1px solid var(--md-outline-variant); flex-shrink:0; }
  .discount-row { display:flex; gap:8px; }
  .m3-text-field {
    flex:1; padding:10px 14px;
    background:var(--md-surface-3);
    border:1px solid var(--md-outline-variant);
    border-radius:var(--shape-md);
    color:var(--md-on-surface); font-size:14px;
    outline:none; font-family:var(--font-body);
    transition:border-color .2s;
  }
  .m3-text-field:focus { border-color:var(--md-primary); }
  .m3-select {
    padding:10px 12px;
    background:var(--md-surface-3);
    border:1px solid var(--md-outline-variant);
    border-radius:var(--shape-md);
    color:var(--md-on-surface); font-size:14px;
    outline:none; font-family:var(--font-body); cursor:pointer;
    transition:border-color .2s;
  }
  .m3-select:focus { border-color:var(--md-primary); }

  /* M3 Filled Button */
  .m3-btn-filled {
    width:100%; padding:14px;
    background:var(--md-primary); color:var(--md-on-primary);
    border:none; border-radius:var(--shape-xl);
    font-size:15px; font-weight:600; font-family:var(--font-brand);
    cursor:pointer; transition:all .2s; letter-spacing:.1px;
    -webkit-tap-highlight-color:transparent;
    box-shadow:var(--elevation-1);
  }
  .m3-btn-filled:hover { box-shadow:var(--elevation-2); }
  .m3-btn-filled:active { box-shadow:none; opacity:.9; }
  .m3-btn-filled:disabled { opacity:.38; cursor:not-allowed; box-shadow:none; }

  /* M3 Text Button */
  .m3-btn-text {
    width:100%; padding:10px;
    background:transparent; color:var(--md-error);
    border:none; border-radius:var(--shape-xl);
    font-size:13px; font-weight:500; font-family:var(--font-brand);
    cursor:pointer; transition:background .15s;
  }
  .m3-btn-text:hover { background:rgba(255,180,171,.08); }

  /* ════════════════════════════════════════
     M3 BOTTOM NAVIGATION BAR
  ════════════════════════════════════════ */
  .m3-bottom-nav {
    display:none;
    background:var(--md-surface-2);
    border-top:1px solid var(--md-outline-variant);
    flex-shrink:0;
    padding-bottom:env(safe-area-inset-bottom,0);
  }
  .m3-nav-inner { display:flex; height:80px; }
  .m3-nav-item {
    flex:1; display:flex; flex-direction:column; align-items:center; justify-content:center;
    gap:4px; border:none; background:transparent; cursor:pointer;
    color:var(--md-on-surface-variant); font-size:10px; font-weight:500;
    font-family:var(--font-brand); transition:color .2s;
    -webkit-tap-highlight-color:transparent; position:relative; padding:0;
  }
  .m3-nav-item.active { color:var(--md-on-secondary-container); }
  .nav-indicator {
    width:64px; height:32px; border-radius:var(--shape-full);
    display:flex; align-items:center; justify-content:center;
    transition:background .2s; position:relative;
  }
  .m3-nav-item.active .nav-indicator { background:var(--md-secondary-container); }
  .nav-icon { font-size:22px; line-height:1; }
  .nav-badge {
    position:absolute; top:-2px; right:6px;
    background:var(--md-primary); color:var(--md-on-primary);
    border-radius:var(--shape-full); min-width:16px; height:16px;
    font-size:10px; font-weight:700; display:flex; align-items:center; justify-content:center; padding:0 4px;
  }
  .nav-label { font-size:10px; font-weight:500; }

  /* ════════════════════════════════════════
     M3 MODAL (Bottom Sheet on mobile)
  ════════════════════════════════════════ */
  .m3-overlay {
    position:fixed; inset:0; background:var(--md-scrim); z-index:100;
    display:flex; align-items:flex-end; justify-content:center;
    backdrop-filter:blur(2px);
  }
  .m3-sheet {
    background:var(--md-surface-2);
    border-radius:28px 28px 0 0;
    width:100%; max-width:560px;
    max-height:92dvh;
    display:flex; flex-direction:column;
    box-shadow:var(--elevation-4);
    animation:sheetUp .3s cubic-bezier(.4,0,.2,1);
    overflow:hidden;
  }
  @keyframes sheetUp { from{transform:translateY(100%);opacity:0} to{transform:translateY(0);opacity:1} }
  .sheet-handle { width:32px; height:4px; background:var(--md-outline); border-radius:2px; margin:12px auto 8px; flex-shrink:0; }
  .sheet-header { padding:0 16px 14px; display:flex; align-items:center; justify-content:space-between; border-bottom:1px solid var(--md-outline-variant); flex-shrink:0; }
  .sheet-title { font-size:20px; font-weight:600; font-family:var(--font-brand); color:var(--md-on-surface); }
  .sheet-close { width:40px; height:40px; border-radius:50%; border:none; background:var(--md-surface-3); color:var(--md-on-surface); cursor:pointer; font-size:18px; display:flex; align-items:center; justify-content:center; }
  .sheet-body { padding:16px; overflow-y:auto; flex:1; -webkit-overflow-scrolling:touch; }
  .sheet-footer { padding:12px 16px; border-top:1px solid var(--md-outline-variant); display:flex; gap:10px; flex-shrink:0; padding-bottom:max(12px,env(safe-area-inset-bottom)); }

  /* ════════════════════════════════════════
     PAYMENT SHEET
  ════════════════════════════════════════ */
  .pay-summary-card {
    background:var(--md-surface-3); border-radius:var(--shape-lg);
    padding:16px; margin-bottom:16px;
  }
  .pay-row { display:flex; justify-content:space-between; font-size:14px; color:var(--md-on-surface-variant); margin-bottom:6px; }
  .pay-row:last-child { margin-bottom:0; }
  .pay-row.grand { font-size:20px; font-weight:700; font-family:var(--font-brand); color:var(--md-on-surface); border-top:1px dashed var(--md-outline); padding-top:10px; margin-top:4px; }
  .pay-row.grand .pay-val { color:var(--md-primary); }
  .pay-row .pay-val { font-family:var(--font-mono); font-weight:500; }

  .pay-method-label { font-size:13px; font-weight:500; color:var(--md-on-surface-variant); margin-bottom:10px; font-family:var(--font-brand); }
  .pay-method-grid { display:grid; grid-template-columns:repeat(5,1fr); gap:8px; margin-bottom:16px; }
  .pay-method-btn {
    padding:10px 4px; border-radius:var(--shape-md);
    border:1px solid var(--md-outline);
    background:transparent; color:var(--md-on-surface-variant);
    font-size:10px; font-weight:600; font-family:var(--font-brand);
    cursor:pointer; text-align:center; text-transform:uppercase;
    transition:all .15s; -webkit-tap-highlight-color:transparent;
  }
  .pay-method-btn.active { background:var(--md-primary-container); border-color:var(--md-primary); color:var(--md-on-primary-container); }
  .pay-field-label { font-size:13px; color:var(--md-on-surface-variant); margin-bottom:6px; font-family:var(--font-brand); }
  .pay-amount-input {
    width:100%; padding:14px 16px;
    background:var(--md-surface-3); border:2px solid var(--md-outline);
    border-radius:var(--shape-md); color:var(--md-on-surface);
    font-size:22px; font-weight:700; font-family:var(--font-mono);
    outline:none; margin-bottom:12px; transition:border-color .2s;
  }
  .pay-amount-input:focus { border-color:var(--md-primary); }
  .quick-amounts { display:flex; gap:8px; margin-bottom:16px; flex-wrap:wrap; }
  .quick-chip {
    padding:8px 14px; border-radius:var(--shape-full);
    border:1px solid var(--md-outline);
    background:transparent; color:var(--md-on-surface);
    font-size:12px; font-weight:500; cursor:pointer;
    font-family:var(--font-mono); transition:all .15s;
    -webkit-tap-highlight-color:transparent;
  }
  .quick-chip:active { background:var(--md-secondary-container); border-color:var(--md-secondary-container); }
  .change-card {
    background:rgba(110,231,183,.1); border:1px solid rgba(110,231,183,.3);
    border-radius:var(--shape-lg); padding:14px 16px;
    display:flex; justify-content:space-between; align-items:center;
  }
  .change-card.neg { background:rgba(255,180,171,.1); border-color:rgba(255,180,171,.3); }
  .change-lbl { font-size:13px; font-weight:600; font-family:var(--font-brand); color:var(--md-success); }
  .change-val { font-size:22px; font-weight:700; font-family:var(--font-mono); color:var(--md-success); }
  .change-card.neg .change-lbl,.change-card.neg .change-val { color:var(--md-error); }

  /* M3 Buttons */
  .btn { padding:12px 20px; border-radius:var(--shape-xl); border:none; font-size:14px; font-weight:600; font-family:var(--font-brand); cursor:pointer; transition:all .2s; -webkit-tap-highlight-color:transparent; }
  .btn-filled { background:var(--md-primary); color:var(--md-on-primary); flex:1; }
  .btn-filled:disabled { opacity:.38; cursor:not-allowed; }
  .btn-outlined { background:transparent; color:var(--md-on-surface); border:1px solid var(--md-outline); }

  /* ════════════════════════════════════════
     RECEIPT SHEET
  ════════════════════════════════════════ */
  .receipt {
    background:white; color:#111;
    border-radius:var(--shape-md); padding:20px;
    font-family:var(--font-mono); font-size:12px; line-height:1.7;
  }
  .receipt-header { text-align:center; margin-bottom:12px; }
  .receipt-header h2 { font-size:15px; font-weight:700; }
  .receipt-header p { font-size:11px; color:#555; }
  .receipt hr { border:none; border-top:1px dashed #ccc; margin:8px 0; }
  .receipt table { width:100%; border-collapse:collapse; }
  .receipt td { padding:2px 0; vertical-align:top; }
  .receipt td:last-child { text-align:right; }
  .receipt .grand td { font-size:14px; font-weight:700; border-top:1px dashed #ccc; padding-top:6px; }
  .receipt-footer { text-align:center; margin-top:10px; font-size:10px; color:#777; }

  /* ════════════════════════════════════════
     DASHBOARD
  ════════════════════════════════════════ */
  .page-content { padding:16px; overflow-y:auto; height:100%; -webkit-overflow-scrolling:touch; }
  .page-title { font-size:22px; font-weight:600; font-family:var(--font-brand); color:var(--md-on-surface); margin-bottom:16px; }
  .m3-stats-grid { display:grid; grid-template-columns:repeat(2,1fr); gap:12px; margin-bottom:16px; }
  .m3-stat-card { background:var(--md-surface-2); border-radius:var(--shape-lg); padding:16px; border:1px solid var(--md-outline-variant); }
  .stat-icon { font-size:24px; margin-bottom:8px; }
  .stat-label { font-size:11px; font-weight:500; color:var(--md-on-surface-variant); text-transform:uppercase; letter-spacing:.6px; margin-bottom:4px; font-family:var(--font-brand); }
  .stat-val { font-size:20px; font-weight:700; font-family:var(--font-brand); color:var(--md-on-surface); }
  .stat-val.primary { color:var(--md-primary); }
  .stat-val.success { color:var(--md-success); }
  .stat-sub { font-size:11px; color:var(--md-on-surface-variant); margin-top:3px; }
  .m3-card { background:var(--md-surface-2); border-radius:var(--shape-lg); padding:16px; border:1px solid var(--md-outline-variant); margin-bottom:12px; }
  .m3-card h4 { font-size:14px; font-weight:600; font-family:var(--font-brand); color:var(--md-on-surface); margin-bottom:12px; }
  .list-row { display:flex; align-items:center; gap:10px; padding:8px 0; border-bottom:1px solid var(--md-outline-variant); font-size:12px; }
  .list-row:last-child { border-bottom:none; }
  .list-row .lr-id { font-family:var(--font-mono); color:#93C5FD; font-size:10px; flex:1; overflow:hidden; text-overflow:ellipsis; white-space:nowrap; }
  .list-row .lr-badge { font-size:9px; padding:2px 7px; border-radius:var(--shape-full); background:var(--md-surface-3); color:var(--md-on-surface-variant); text-transform:uppercase; font-weight:600; white-space:nowrap; }
  .list-row .lr-total { font-family:var(--font-mono); font-weight:700; color:var(--md-primary); white-space:nowrap; }
  .list-row .lr-name { flex:1; font-weight:500; overflow:hidden; text-overflow:ellipsis; white-space:nowrap; }
  .list-row .lr-qty { font-family:var(--font-mono); font-size:10px; color:#93C5FD; white-space:nowrap; }
  .list-row .lr-rev { font-family:var(--font-mono); font-size:10px; color:var(--md-primary); white-space:nowrap; }

  /* ════════════════════════════════════════
     PRODUCTS / TRANSACTIONS PAGE
  ════════════════════════════════════════ */
  .page-toolbar { display:flex; gap:8px; align-items:center; margin-bottom:14px; flex-wrap:nowrap; }
  .table-wrap { overflow-x:auto; border-radius:var(--shape-md); border:1px solid var(--md-outline-variant); }
  .m3-table { width:100%; border-collapse:collapse; font-size:12px; min-width:500px; }
  .m3-table th { text-align:left; padding:10px 12px; font-size:10px; font-weight:600; text-transform:uppercase; letter-spacing:.5px; color:var(--md-on-surface-variant); background:var(--md-surface-2); border-bottom:1px solid var(--md-outline-variant); font-family:var(--font-brand); }
  .m3-table td { padding:10px 12px; border-bottom:1px solid var(--md-outline-variant); color:var(--md-on-surface); vertical-align:middle; }
  .m3-table tr:last-child td { border-bottom:none; }
  .m3-table tr:hover td { background:var(--md-surface-3); }
  .m3-badge { display:inline-flex; align-items:center; padding:3px 8px; border-radius:var(--shape-full); font-size:9px; font-weight:600; text-transform:uppercase; letter-spacing:.4px; }
  .badge-makanan { background:rgba(255,107,53,.15); color:#FF6B35; }
  .badge-minuman { background:rgba(14,165,233,.15); color:#0EA5E9; }
  .badge-snack { background:rgba(168,85,247,.15); color:#A855F7; }
  .badge-lainnya,.badge-other { background:var(--md-surface-3); color:var(--md-on-surface-variant); }
  .m3-icon-btn { width:32px; height:32px; border-radius:50%; border:none; background:transparent; cursor:pointer; font-size:14px; display:flex; align-items:center; justify-content:center; color:var(--md-on-surface-variant); transition:background .15s; }
  .m3-icon-btn:hover { background:var(--state-hover); }
  .m3-icon-btn.danger:hover { background:rgba(255,180,171,.12); color:var(--md-error); }

  /* ════════════════════════════════════════
     SETTINGS
  ════════════════════════════════════════ */
  .m3-settings-section { background:var(--md-surface-2); border-radius:var(--shape-lg); border:1px solid var(--md-outline-variant); overflow:hidden; margin-bottom:12px; }
  .settings-section-title { font-size:12px; font-weight:600; color:var(--md-primary); font-family:var(--font-brand); padding:14px 16px 8px; text-transform:uppercase; letter-spacing:.6px; }
  .settings-row { display:flex; justify-content:space-between; align-items:center; padding:14px 16px; border-bottom:1px solid var(--md-outline-variant); gap:12px; }
  .settings-row:last-child { border-bottom:none; }
  .settings-row .slbl { font-size:14px; color:var(--md-on-surface); font-family:var(--font-brand); }
  .settings-row .sval { font-size:12px; color:var(--md-on-surface-variant); margin-top:2px; }
  .m3-toggle { width:52px; height:32px; border-radius:var(--shape-full); cursor:pointer; position:relative; transition:background .2s; border:none; flex-shrink:0; background:var(--md-outline); }
  .m3-toggle.on { background:var(--md-primary); }
  .m3-toggle::after { content:''; position:absolute; top:4px; left:4px; width:24px; height:24px; background:white; border-radius:50%; transition:left .2s; box-shadow:0 1px 3px rgba(0,0,0,.3); }
  .m3-toggle.on::after { left:24px; }
  .m3-settings-input { padding:8px 12px; background:var(--md-surface-3); border:1px solid var(--md-outline-variant); border-radius:var(--shape-md); color:var(--md-on-surface); font-size:14px; outline:none; width:160px; font-family:var(--font-body); }
  .m3-settings-input:focus { border-color:var(--md-primary); }
  .btn-danger-tonal { padding:8px 16px; background:rgba(255,180,171,.12); color:var(--md-error); border:none; border-radius:var(--shape-full); font-size:13px; font-weight:500; cursor:pointer; font-family:var(--font-brand); transition:background .15s; }
  .btn-danger-tonal:hover { background:rgba(255,180,171,.2); }

  /* ════════════════════════════════════════
     PRODUCT FORM SHEET
  ════════════════════════════════════════ */
  .form-group { margin-bottom:14px; }
  .form-label { font-size:12px; font-weight:500; color:var(--md-on-surface-variant); margin-bottom:5px; display:block; font-family:var(--font-brand); }
  .m3-form-input { width:100%; padding:12px 14px; background:var(--md-surface-3); border:1px solid var(--md-outline-variant); border-radius:var(--shape-md); color:var(--md-on-surface); font-size:14px; outline:none; font-family:var(--font-body); transition:border-color .2s; }
  .m3-form-input:focus { border-color:var(--md-primary); }
  .m3-form-select { width:100%; padding:12px 14px; background:var(--md-surface-3); border:1px solid var(--md-outline-variant); border-radius:var(--shape-md); color:var(--md-on-surface); font-size:14px; outline:none; font-family:var(--font-body); cursor:pointer; }
  .form-row-2 { display:grid; grid-template-columns:1fr 1fr; gap:12px; }

  /* ════════════════════════════════════════
     BARCODE SCANNER
  ════════════════════════════════════════ */
  .scan-modal { position:fixed; inset:0; background:rgba(0,0,0,.88); z-index:200; display:flex; flex-direction:column; align-items:center; justify-content:flex-start; padding-top:32px; }
  .scan-box-wrap { background:var(--md-surface-2); border-radius:var(--shape-xl); width:min(92vw,380px); overflow:hidden; }
  .scan-header { padding:14px 16px 12px; border-bottom:1px solid var(--md-outline-variant); display:flex; align-items:center; justify-content:space-between; }
  .scan-header span { font-weight:600; font-size:15px; font-family:var(--font-brand); color:var(--md-on-surface); }
  .scan-close { width:36px; height:36px; border-radius:50%; border:none; background:var(--md-surface-3); color:var(--md-on-surface); cursor:pointer; font-size:18px; display:flex; align-items:center; justify-content:center; }
  .scan-video-wrap { position:relative; background:#000; height:200px; }
  .scan-video-wrap video { width:100%; height:100%; object-fit:cover; }
  .scan-reticle { position:absolute; inset:0; display:flex; align-items:center; justify-content:center; pointer-events:none; }
  .scan-reticle-box { width:210px; height:95px; border:2px solid var(--md-primary); border-radius:var(--shape-md); box-shadow:0 0 0 9999px rgba(0,0,0,.45); position:relative; }
  .scan-line { position:absolute; left:0; right:0; height:2px; background:var(--md-primary); opacity:.9; animation:scanline 1.5s ease-in-out infinite; }
  @keyframes scanline { 0%,100%{top:8%} 50%{top:82%} }
  .scan-hint-text { position:absolute; bottom:8px; left:0; right:0; text-align:center; font-size:11px; color:rgba(255,255,255,.6); }
  .scan-body { padding:14px 16px; }
  .scan-warn { background:rgba(245,158,11,.1); border:1px solid rgba(245,158,11,.3); border-radius:var(--shape-md); padding:10px 12px; margin-bottom:10px; font-size:12px; color:#FCD34D; }
  .scan-error { background:rgba(255,180,171,.1); border:1px solid rgba(255,180,171,.3); border-radius:var(--shape-md); padding:10px 12px; margin-bottom:10px; font-size:12px; color:var(--md-error); }
  .scan-manual-row { display:flex; gap:8px; }
  .scan-manual-input { flex:1; padding:11px 14px; background:var(--md-surface-3); border:1px solid var(--md-outline-variant); border-radius:var(--shape-md); color:var(--md-on-surface); font-size:15px; outline:none; font-family:monospace; }
  .scan-manual-input:focus { border-color:var(--md-primary); }
  .scan-ok-btn { padding:11px 18px; background:var(--md-primary); color:var(--md-on-primary); border:none; border-radius:var(--shape-md); font-weight:600; font-size:13px; cursor:pointer; font-family:var(--font-brand); }
  .scan-tip { margin-top:8px; font-size:10px; color:var(--md-on-surface-variant); }

  /* ════════════════════════════════════════
     TOAST
  ════════════════════════════════════════ */
  .toast-wrap { position:fixed; bottom:max(90px,calc(env(safe-area-inset-bottom)+90px)); right:12px; z-index:300; display:flex; flex-direction:column; gap:8px; }
  .m3-toast { padding:12px 16px; border-radius:var(--shape-md); font-size:13px; font-weight:500; font-family:var(--font-brand); box-shadow:var(--elevation-4); animation:toastIn .2s ease; display:flex; align-items:center; gap:8px; max-width:280px; }
  @keyframes toastIn { from{opacity:0;transform:translateX(20px)} to{opacity:1;transform:translateX(0)} }
  .toast-success { background:#065F46; color:#6EE7B7; border:1px solid rgba(110,231,183,.2); }
  .toast-error { background:#7F1D1D; color:#FCA5A5; border:1px solid rgba(252,165,165,.2); }
  .toast-info { background:var(--md-primary-container); color:var(--md-on-primary-container); }

  select option { background:var(--md-surface-3); }

  /* Light mode hardcoded color overrides */
  :root[data-theme="light"] .m3-product-card-wrap { background:#E8ECF4 !important; border-color:#C4C6D0 !important; }
  :root[data-theme="light"] .card-category-label { color:#0061A4 !important; }
  :root[data-theme="light"] .card-name-text { color:#1A1C1E !important; }
  :root[data-theme="light"] .card-price-text { color:#0061A4 !important; }
  :root[data-theme="light"] .card-stock-text { color:#44474F !important; }
  :root[data-theme="light"] .card-add-btn-inner { background:#D3E3FD !important; color:#001D35 !important; }
  :root[data-theme="light"] { color-scheme: light; }
  :root[data-theme="light"] html,
  :root[data-theme="light"] body,
  :root[data-theme="light"] #root { background:#F8F9FA; }
  :root[data-theme="light"] .pos { background:#F8F9FA; }
  :root[data-theme="light"] .content { background:#F8F9FA; }
  :root[data-theme="light"] .app { background:#F8F9FA; }
  :root[data-theme="dark"] { color-scheme: dark; }

  /* Material Icons */
  .material-icons-round {
    font-family: 'Material Icons Round';
    font-weight: normal;
    font-style: normal;
    font-size: 24px;
    line-height: 1;
    letter-spacing: normal;
    text-transform: none;
    display: inline-block;
    white-space: nowrap;
    word-wrap: normal;
    direction: ltr;
    -webkit-font-smoothing: antialiased;
    user-select: none;
  }
  .m3-nav-item .material-icons-round {
    font-size: 24px;
    transition: color .2s;
  }

  /* ════════════════════════════════════════
     MOBILE RESPONSIVE
  ════════════════════════════════════════ */
  @media (max-width:640px) {
    .product-grid { grid-template-columns:repeat(2,1fr); padding:10px; gap:8px; }
    .top-app-bar { height:56px; padding:0 4px 0 12px; }
    .desktop-nav { display:none; }
    .m3-bottom-nav { display:block; }
    .pos { flex-direction:column; }
    .pos-right { display:none; }
    .pos-left { height:100%; position:relative; }
    .m3-stats-grid { grid-template-columns:1fr 1fr; }
    .form-row-2 { grid-template-columns:1fr; }
    .m3-settings-input { width:120px; }
    .toast-wrap { right:8px; }
    .top-app-bar .brand-title { font-size:16px; }
  }
  @media (min-width:641px) {
    .product-grid { grid-template-columns:repeat(auto-fill,minmax(160px,1fr)); }
    .m3-bottom-nav { display:none !important; }
    .cart-fab-bar { display:none !important; }
    .m3-stats-grid { grid-template-columns:repeat(4,1fr); }
    .m3-sheet { border-radius:28px; align-self:center; max-height:88vh; }
    @keyframes sheetUp { from{opacity:0;transform:scale(.95) translateY(10px)} to{opacity:1;transform:scale(1) translateY(0)} }
  }

  /* Mobile cart page */
  .cart-page-mobile { display:flex; flex-direction:column; height:100%; background:var(--md-surface-1); overflow:hidden; transition:background .25s; }
  @media (max-width:640px) {
    .cart-page-mobile { overflow-y:auto; -webkit-overflow-scrolling:touch; }
    .cart-page-mobile .cart-items-list { overflow-y:visible; flex:none; }
  }
`;

// ── Helper: Category Avatar ─────────────────────────────────────────────────
function CategoryAvatar({ category, size=48 }) {
  const icon = CATEGORY_ICONS[category] || "📦";
  const color = CATEGORY_COLORS[category] || "#6B7280";
  return (
    <div style={{ width:size, height:size, borderRadius:size*0.25, background:`${color}22`, display:"flex", alignItems:"center", justifyContent:"center", fontSize:size*0.5, flexShrink:0 }}>
      {icon}
    </div>
  );
}

// ── ESC/POS Bluetooth Printer (EPPOS RPP02) ─────────────────────────────────
const ESC  = 0x1B;
const GS   = 0x1D;

function encodeText(text) {
  return new TextEncoder().encode(text);
}

function buildReceipt(receipt, settings, formatRp, formatDate) {
  const lines = [];

  const center = (text, width=32) => {
    const pad = Math.max(0, Math.floor((width - text.length) / 2));
    return " ".repeat(pad) + text;
  };
  const leftRight = (left, right, width=32) => {
    const space = Math.max(1, width - left.length - right.length);
    return left + " ".repeat(space) + right;
  };
  const line = (char="-", width=32) => char.repeat(width);

  // ESC/POS commands
  const INIT        = [ESC, 0x40];                    // Initialize
  const ALIGN_CTR   = [ESC, 0x61, 0x01];              // Center
  const ALIGN_L     = [ESC, 0x61, 0x00];              // Left
  const BOLD_ON     = [ESC, 0x45, 0x01];
  const BOLD_OFF    = [ESC, 0x45, 0x00];
  const FONT_LARGE  = [GS,  0x21, 0x11];              // Double size
  const FONT_NORMAL = [GS,  0x21, 0x00];
  const CUT         = [GS,  0x56, 0x41, 0x10];        // Partial cut
  const FEED        = [ESC, 0x64, 0x04];              // Feed 4 lines

  let bytes = [];
  const add = (arr) => bytes.push(...arr);
  const text = (str) => bytes.push(...encodeText(str));
  const nl   = () => bytes.push(0x0A);

  add(INIT);

  // Header
  add(ALIGN_CTR);
  add(FONT_LARGE);
  add(BOLD_ON);
  text(settings.storeName); nl();
  add(FONT_NORMAL);
  add(BOLD_OFF);
  text(settings.address); nl();
  text("Tel: 021-12345678"); nl();
  add(ALIGN_L);
  text(line()); nl();

  // Info transaksi
  text("No  : " + receipt.id); nl();
  text("Tgl : " + formatDate(receipt.date)); nl();
  text("Kasir: " + (receipt.cashier || settings.cashier)); nl();
  text(line()); nl();

  // Items
  receipt.items.forEach(item => {
    const nama = item.name.length > 20 ? item.name.substring(0,20)+"." : item.name;
    text(nama); nl();
    const detail = `  ${item.qty} x ${formatRp(item.price)}`;
    const subtotal = formatRp(item.qty * item.price);
    text(leftRight(detail, subtotal)); nl();
  });

  text(line()); nl();

  // Totals
  text(leftRight("Subtotal", formatRp(receipt.subtotal))); nl();
  if (receipt.discount > 0) {
    text(leftRight("Diskon", "-" + formatRp(receipt.discount))); nl();
  }
  if (receipt.tax > 0) {
    text(leftRight("Pajak", formatRp(receipt.tax))); nl();
  }
  text(line()); nl();

  add(BOLD_ON);
  text(leftRight("TOTAL", formatRp(receipt.total))); nl();
  add(BOLD_OFF);

  text(leftRight("Bayar (" + receipt.payment + ")", formatRp(receipt.cashPaid))); nl();
  if (receipt.change > 0) {
    text(leftRight("Kembalian", formatRp(receipt.change))); nl();
  }

  if (receipt.note) {
    text(line()); nl();
    text("Catatan: " + receipt.note); nl();
  }

  text(line()); nl();
  add(ALIGN_CTR);
  text("Terima kasih!"); nl();
  text("Powered by Kasir-Pro"); nl();

  add(FEED);
  add(CUT);

  return new Uint8Array(bytes);
}

// Hook untuk Bluetooth Printer
function useBluetoothPrinter() {
  const [device, setDevice]   = useState(null);
  const [status, setStatus]   = useState("disconnected"); // disconnected | connecting | connected | printing | error
  const [errMsg, setErrMsg]   = useState("");
  const charRef = useRef(null);

  const isSupported = typeof navigator !== "undefined" && "bluetooth" in navigator;

  const connect = async () => {
    if (!isSupported) { setErrMsg("Web Bluetooth tidak didukung browser ini. Gunakan Chrome Android."); setStatus("error"); return; }
    try {
      setStatus("connecting"); setErrMsg("");
      const dev = await navigator.bluetooth.requestDevice({
        // EPPOS RPP02 service UUID — thermal printer generic
        filters: [
          { namePrefix: "RPP" },
          { namePrefix: "EPPOS" },
          { namePrefix: "Printer" },
          { namePrefix: "MPT" },
          { services: ["000018f0-0000-1000-8000-00805f9b34fb"] },
        ],
        optionalServices: [
          "000018f0-0000-1000-8000-00805f9b34fb",
          "e7810a71-73ae-499d-8c15-faa9aef0c3f2",
          "49535343-fe7d-4ae5-8fa9-9fafd205e455",
          "000018f0-0000-1000-8000-00805f9b34fb",
        ],
      });
      setDevice(dev);
      dev.addEventListener("gattserverdisconnected", () => { setStatus("disconnected"); charRef.current=null; });
      const server = await dev.gatt.connect();

      // Coba berbagai service UUID thermal printer
      let char = null;
      const serviceUUIDs = [
        "000018f0-0000-1000-8000-00805f9b34fb",
        "e7810a71-73ae-499d-8c15-faa9aef0c3f2",
        "49535343-fe7d-4ae5-8fa9-9fafd205e455",
      ];
      const charUUIDs = [
        "00002af1-0000-1000-8000-00805f9b34fb",
        "bef8d6c9-9c21-4c9e-b632-bd58c1009f9f",
        "49535343-8841-43f4-a8d4-ecbe34729bb3",
      ];

      for (const svcUUID of serviceUUIDs) {
        try {
          const svc = await server.getPrimaryService(svcUUID);
          for (const cUUID of charUUIDs) {
            try { char = await svc.getCharacteristic(cUUID); break; } catch {}
          }
          // Coba semua characteristics jika UUID spesifik gagal
          if (!char) {
            const chars = await svc.getCharacteristics();
            for (const c of chars) {
              if (c.properties.write || c.properties.writeWithoutResponse) { char = c; break; }
            }
          }
          if (char) break;
        } catch {}
      }

      if (!char) throw new Error("Karakteristik printer tidak ditemukan. Pastikan printer EPPOS RPP02 menyala dan dalam mode pairing.");
      charRef.current = char;
      setStatus("connected");
    } catch (e) {
      setStatus("error");
      setErrMsg(e.message || "Gagal terhubung ke printer.");
    }
  };

  const disconnect = async () => {
    if (device?.gatt?.connected) { device.gatt.disconnect(); }
    setDevice(null); charRef.current=null; setStatus("disconnected");
  };

  const print = async (data) => {
    if (!charRef.current) { setErrMsg("Printer belum terhubung."); setStatus("error"); return false; }
    try {
      setStatus("printing");
      // Kirim data dalam chunk 512 bytes (batas BLE MTU)
      const CHUNK = 512;
      for (let i = 0; i < data.length; i += CHUNK) {
        const chunk = data.slice(i, i + CHUNK);
        await charRef.current.writeValueWithoutResponse(chunk);
        await new Promise(r => setTimeout(r, 60)); // delay antar chunk
      }
      setStatus("connected");
      return true;
    } catch (e) {
      setStatus("error");
      setErrMsg("Gagal mencetak: " + e.message);
      return false;
    }
  };

  return { device, status, errMsg, isSupported, connect, disconnect, print };
}

// ── CartPanel (standalone — no remount on parent re-render) ─────────────────
function CartPanel({ cart, note, setNote, discount, setDiscount, discountType, setDiscountType,
  subtotal, discountAmt, taxAmt, total, settings, updateQty, removeFromCart, clearCart, setShowPayment }) {
  return (
    <>
      <div className="cart-top">
        <h3>Pesanan <span>{cart.reduce((s,x)=>s+x.qty,0)} item</span></h3>
      </div>
      <div className="note-wrap">
        <input type="text" className="m3-note-input" placeholder="📝 Catatan pesanan..." value={note} onChange={e=>setNote(e.target.value)} />
      </div>
      <div className="cart-items-list">
        {cart.length === 0 && (
          <div className="cart-empty-state">
            <div className="cart-empty-icon">🛒</div>
            Keranjang masih kosong
          </div>
        )}
        {cart.map(item => (
          <div key={item.id} className="m3-cart-item">
            <CategoryAvatar category={item.category} size={40} />
            <div className="cart-item-info">
              <div className="cart-item-name">{item.name}</div>
              <div className="cart-item-price" style={{color:"var(--md-on-surface-variant)",fontSize:11}}>{formatRp(item.price)}</div>
            </div>
            <div className="m3-qty">
              <button className="m3-qty-btn" onClick={()=>item.qty===1?removeFromCart(item.id):updateQty(item.id,-1)}>−</button>
              <span className="m3-qty-val">{item.qty}</span>
              <button className="m3-qty-btn" onClick={()=>updateQty(item.id,1)}>+</button>
            </div>
            <div className="cart-item-subtotal">{formatRp(item.price*item.qty)}</div>
            <button className="m3-del-btn" onClick={()=>removeFromCart(item.id)}>✕</button>
          </div>
        ))}
      </div>
      <div className="cart-totals-wrap">
        <div className="total-row"><span>Subtotal</span><span className="val">{formatRp(subtotal)}</span></div>
        {discountAmt>0 && <div className="total-row discount"><span>Diskon</span><span className="val">-{formatRp(discountAmt)}</span></div>}
        {settings.taxEnabled && <div className="total-row tax"><span>Pajak {settings.taxRate}%</span><span className="val">{formatRp(taxAmt)}</span></div>}
        <hr className="total-divider" />
        <div className="total-row grand"><span>TOTAL</span><span className="val">{formatRp(total)}</span></div>
      </div>
      <div className="cart-actions-wrap">
        <div className="discount-row">
          <input className="m3-text-field" placeholder="Diskon" type="number" value={discount} onChange={e=>setDiscount(e.target.value)} />
          <select className="m3-select" value={discountType} onChange={e=>setDiscountType(e.target.value)}>
            <option value="pct">%</option>
            <option value="fixed">Rp</option>
          </select>
        </div>
        <button className="m3-btn-filled" disabled={cart.length===0} onClick={()=>setShowPayment(true)}>
          💳 Bayar {cart.length>0 && formatRp(total)}
        </button>
        {cart.length>0 && <button className="m3-btn-text" onClick={clearCart}>Hapus Semua</button>}
      </div>
    </>
  );
}

// ── ProductFormModal ────────────────────────────────────────────────────────
function ProductFormModal({ product, onClose, onSave }) {
  const [form, setForm] = useState({sku:"",barcode:"",name:"",category:"Makanan",price:"",stock:"",...product});
  const set = (k,v) => setForm(f=>({...f,[k]:v}));
  const handleSave = () => {
    if (!form.name||!form.price||!form.stock) { alert("Isi semua field wajib!"); return; }
    onSave({...form,price:parseFloat(form.price),stock:parseInt(form.stock)});
  };
  return (
    <div className="m3-overlay">
      <div className="m3-sheet">
        <div className="sheet-handle" />
        <div className="sheet-header">
          <span className="sheet-title">{form.id?"✏️ Edit Produk":"➕ Tambah Produk"}</span>
          <button className="sheet-close" onClick={onClose}>✕</button>
        </div>
        <div className="sheet-body">
          <div className="form-row-2">
            <div className="form-group"><label className="form-label">SKU</label><input className="m3-form-input" value={form.sku} onChange={e=>set("sku",e.target.value)} placeholder="MKN001" /></div>
            <div className="form-group">
              <label className="form-label">Barcode</label>
              <div style={{display:"flex",gap:6}}>
                <input className="m3-form-input" value={form.barcode} onChange={e=>set("barcode",e.target.value)} placeholder="8991234..." style={{flex:1}} />
                <BarcodeInlineButton onResult={code=>set("barcode",code)} />
              </div>
            </div>
          </div>
          <div className="form-group"><label className="form-label">Nama Produk *</label><input className="m3-form-input" value={form.name} onChange={e=>set("name",e.target.value)} placeholder="Nama produk" /></div>
          <div className="form-row-2">
            <div className="form-group"><label className="form-label">Kategori</label><select className="m3-form-select" value={form.category} onChange={e=>set("category",e.target.value)}><option>Makanan</option><option>Minuman</option><option>Snack</option><option>Lainnya</option></select></div>
            <div className="form-group"><label className="form-label">Harga Jual (Rp) *</label><input className="m3-form-input" type="number" inputMode="numeric" value={form.price} onChange={e=>set("price",e.target.value)} placeholder="0" /></div>
          </div>
          <div className="form-group"><label className="form-label">Stok *</label><input className="m3-form-input" type="number" inputMode="numeric" value={form.stock} onChange={e=>set("stock",e.target.value)} placeholder="0" /></div>
        </div>
        <div className="sheet-footer">
          <button className="btn btn-outlined" onClick={onClose}>Batal</button>
          <button className="btn btn-filled" onClick={handleSave}>💾 Simpan</button>
        </div>
      </div>
    </div>
  );
}

// ── BarcodeInlineButton ─────────────────────────────────────────────────────
function BarcodeInlineButton({ onResult }) {
  const [open, setOpen] = useState(false);
  return (
    <>
      <button style={{width:46,minHeight:46,borderRadius:"var(--shape-md)",border:"1px solid var(--md-outline-variant)",background:"var(--md-primary-container)",color:"var(--md-on-primary-container)",cursor:"pointer",fontSize:20,display:"flex",alignItems:"center",justifyContent:"center",flexShrink:0}} onClick={()=>setOpen(true)}>📷</button>
      {open && <BarcodeModal onClose={()=>setOpen(false)} onResult={code=>{onResult(code);setOpen(false);}} />}
    </>
  );
}

// ── BarcodeModal ────────────────────────────────────────────────────────────
function BarcodeModal({ onResult, onClose }) {
  const [error, setError] = useState("");
  const [manual, setManual] = useState("");
  const videoRef = useRef(null);
  const streamRef = useRef(null);
  const rafRef = useRef(null);
  const canvasRef = useRef(document.createElement("canvas"));
  const hasDetector = "BarcodeDetector" in window;

  useEffect(() => { startCamera(); return () => stopCamera(); }, []);

  const stopCamera = () => {
    if (streamRef.current) { streamRef.current.getTracks().forEach(t=>t.stop()); streamRef.current=null; }
    if (rafRef.current) { cancelAnimationFrame(rafRef.current); rafRef.current=null; }
  };
  const startCamera = async () => {
    setError("");
    try {
      const stream = await navigator.mediaDevices.getUserMedia({video:{facingMode:"environment",width:{ideal:1280},height:{ideal:720}}});
      streamRef.current=stream;
      if (videoRef.current) { videoRef.current.srcObject=stream; await videoRef.current.play(); }
      if (hasDetector) scanLoop();
    } catch { setError("Kamera tidak dapat diakses. Berikan izin kamera di browser."); }
  };
  const scanLoop = () => {
    if (!videoRef.current||!streamRef.current) return;
    const v=videoRef.current;
    if (v.readyState===v.HAVE_ENOUGH_DATA) {
      const c=canvasRef.current; c.width=v.videoWidth; c.height=v.videoHeight;
      c.getContext("2d").drawImage(v,0,0);
      try {
        new window.BarcodeDetector({formats:["ean_13","ean_8","code_128","code_39","qr_code","upc_a","upc_e"]})
          .detect(c).then(codes=>{ if(codes.length>0){stopCamera();onResult(codes[0].rawValue);} }).catch(()=>{});
      } catch {}
    }
    rafRef.current=requestAnimationFrame(scanLoop);
  };
  const handleManual = () => { if(manual.trim()){stopCamera();onResult(manual.trim());} };
  const handleClose = () => { stopCamera(); onClose(); };

  return (
    <div className="scan-modal" onClick={e=>e.target===e.currentTarget&&handleClose()}>
      <div className="scan-box-wrap">
        <div className="scan-header">
          <span>📷 Scanner Barcode</span>
          <button className="scan-close" onClick={handleClose}>✕</button>
        </div>
        <div className="scan-video-wrap">
          <video ref={videoRef} style={{width:"100%",height:"100%",objectFit:"cover"}} muted playsInline />
          <div className="scan-reticle"><div className="scan-reticle-box"><div className="scan-line"/></div></div>
          <div className="scan-hint-text">Arahkan barcode ke dalam kotak</div>
        </div>
        <div className="scan-body">
          {!hasDetector&&!error&&<div className="scan-warn">⚠️ Gunakan Chrome terbaru untuk scan otomatis.</div>}
          {error&&<div className="scan-error">⚠️ {error}</div>}
          <div style={{fontSize:12,color:"var(--md-on-surface-variant)",marginBottom:6,fontWeight:500}}>Input Manual / Scanner USB</div>
          <div className="scan-manual-row">
            <input className="scan-manual-input" placeholder="Ketik atau scan barcode..." value={manual} onChange={e=>setManual(e.target.value)} onKeyDown={e=>e.key==="Enter"&&handleManual()} autoFocus={!!error||!hasDetector} />
            <button className="scan-ok-btn" onClick={handleManual}>✓ OK</button>
          </div>
          <div className="scan-tip">💡 Scanner USB: scan ke field di atas → Enter</div>
        </div>
      </div>
    </div>
  );
}

// ── Thermal Printer (Adaptasi dari modulprinter.db — FreeKasir) ──────────────
// Protokol: Web Bluetooth GATT → ESC/POS
// Kompatibel: EPPOS RPP02, printer thermal Bluetooth 58mm China

// ── UUID Service Bluetooth (SPP / BLE hybrid umum printer China) ──
const PRINTER_SERVICE_UUIDS = [
  "000018f0-0000-1000-8000-00805f9b34fb",  // SPP standar printer thermal
  "00001101-0000-1000-8000-00805f9b34fb",  // Serial Port Profile klasik
  "e7810a71-73ae-499d-8c15-faa9aef0c3f2",  // Epson BLE
  "49535343-fe7d-4ae5-8fa9-9fafd205e455",  // Seiko Instruments
  "0000ff00-0000-1000-8000-00805f9b34fb",  // Generic printer China
  "0000ffe0-0000-1000-8000-00805f9b34fb",  // HM-10 / JDY-08 module
];

// ── ESC/POS Commands (dari modulprinter.db) ──────────────────────
const ESC_INIT         = [0x1B, 0x40];        // ESC @ — reset printer
const ESC_CENTER       = [0x1B, 0x61, 0x01];  // ESC a 1 — rata tengah
const ESC_LEFT         = [0x1B, 0x61, 0x00];  // ESC a 0 — rata kiri
const ESC_BOLD_ON      = [0x1B, 0x45, 0x01];  // ESC E 1 — bold on
const ESC_BOLD_OFF     = [0x1B, 0x45, 0x00];  // ESC E 0 — bold off
const ESC_DOUBLE_ON    = [0x1B, 0x21, 0x10];  // ESC ! 16 — double height (dari getKitchenTicketESCPOSData)
const ESC_DOUBLE_OFF   = [0x1B, 0x21, 0x00];  // ESC ! 0 — reset size
const GS_CUT           = [0x1D, 0x56, 0x41, 0x00]; // GS V A 0 — cut paper
const LF               = [0x0A];              // Line feed

// ── Helper ───────────────────────────────────────────────────────
const rp = n => "Rp " + new Intl.NumberFormat("id-ID").format(n);

// Persis seperti formatRow di modulprinter.db
const formatRow = (left, right, width=32) => {
  const spaceCount = width - left.length - right.length;
  return left + " ".repeat(Math.max(1, spaceCount)) + right + "\n";
};

function strBytes(str) {
  return new TextEncoder().encode(str);
}

function mergeUint8(...arrays) {
  const total = arrays.reduce((s, a) => s + a.length, 0);
  const out = new Uint8Array(total);
  let offset = 0;
  for (const a of arrays) {
    const arr = a instanceof Uint8Array ? a : new Uint8Array(a);
    out.set(arr, offset);
    offset += arr.length;
  }
  return out;
}

// ── getESCPOSData — adaptasi 1:1 dari modulprinter.db ────────────
function getESCPOSData(r, settings) {
  const lines = [];

  lines.push("\x1B\x61\x01"); // Center align
  lines.push((settings.storeName || "Toko") + "\n");
  if (settings.address) lines.push(settings.address + "\n");
  lines.push("--------------------------------\n");
  
  lines.push("\x1B\x61\x00"); // Left align

  lines.push("No: " + r.id + "\n");
  lines.push(new Date(r.date).toLocaleString("id-ID", {
    day:"2-digit", month:"2-digit", year:"numeric",
    hour:"2-digit", minute:"2-digit"
  }) + "\n");
  if (r.cashier || settings.cashier)
    lines.push("Kasir: " + (r.cashier || settings.cashier) + "\n");
  if (r.note)
    lines.push("Catatan: " + r.note + "\n");
  lines.push("--------------------------------\n");

  lines.push("\x1B\x61\x00"); // Left align

  for (const item of r.items) {
    lines.push(item.name + "\n");
    lines.push(formatRow(
      "  " + item.qty + " x " + rp(item.price),
      rp(item.qty * item.price)
    ));
  }

  lines.push("--------------------------------\n");
  lines.push(formatRow("Subtotal:", rp(r.subtotal)));
  if (r.discount > 0)
  lines.push(formatRow("Diskon:", "-" + rp(r.discount)));
  lines.push(formatRow("Pajak", rp(r.tax)));
  lines.push(formatRow("TOTAL:", rp(r.total)));
  lines.push(formatRow("Bayar:", rp(r.cashPaid)));
  if (r.change > 0)
    lines.push(formatRow("Kembali:", rp(r.change)));
  lines.push("--------------------------------\n");

  lines.push("\x1B\x61\x01"); // Center
  lines.push("Aplikasi Kasir:\n");
  lines.push("https://kasir-pro-v5.vercel.app/");
  lines.push("\n\n\n"); // Feed paper (persis seperti modulprinter.db)

  return lines.join("");
}

// ── Web Bluetooth GATT (karena browser tidak bisa akses bluetoothSerial) ──
async function findWriteCharacteristic(server) {
  // Coba UUID satu per satu
  for (const svcUUID of PRINTER_SERVICE_UUIDS) {
    try {
      const svc = await server.getPrimaryService(svcUUID);
      const chars = await svc.getCharacteristics();
      for (const c of chars) {
        if (c.properties.write || c.properties.writeWithoutResponse) {
          return c;
        }
      }
    } catch { continue; }
  }
  // Fallback: scan semua service yang tersedia
  try {
    const services = await server.getPrimaryServices();
    for (const svc of services) {
      const chars = await svc.getCharacteristics().catch(() => []);
      for (const c of chars) {
        if (c.properties.write || c.properties.writeWithoutResponse) {
          return c;
        }
      }
    }
  } catch {}
  return null;
}

async function writeInChunks(characteristic, data, chunkSize=100) {
  // Chunk kecil (100 byte) agar buffer printer tidak overflow
  // Mirip dengan window.bluetoothSerial.write di modulprinter.db
  for (let i = 0; i < data.length; i += chunkSize) {
    const chunk = data.slice(i, i + chunkSize);
    try {
      if (characteristic.properties.writeWithoutResponse) {
        await characteristic.writeValueWithoutResponse(chunk);
      } else {
        await characteristic.writeValue(chunk);
      }
    } catch (e) {
      throw new Error("Gagal kirim data chunk ke-" + Math.floor(i/chunkSize) + ": " + e.message);
    }
    // Jeda 50ms antar chunk — sama seperti prinsip write di plugin native
    await new Promise(res => setTimeout(res, 50));
  }
}

// ── printRawWebBluetooth — adaptasi printRawNativeBluetooth dari modulprinter.db
async function printRawWebBluetooth(rawText, onStatus) {
  // Cek HTTPS dulu
  if (!window.isSecureContext) {
    throw new Error("Butuh HTTPS! Buka via link Vercel (https://...) bukan IP lokal. Web Bluetooth hanya aktif di HTTPS.");
  }
  if (!navigator.bluetooth) {
    throw new Error("Web Bluetooth tidak tersedia. Pastikan: Chrome Android 85+ dan buka via HTTPS (Vercel).");
  }

  // TAHAP 1: Pemindaian & Otorisasi (persis seperti listPairedBluetoothDevices + connect)
  onStatus("🔍 Pilih printer EPPOS RPP02...");

  let device;
  try {
    device = await navigator.bluetooth.requestDevice({
      acceptAllDevices: true,
      optionalServices: PRINTER_SERVICE_UUIDS,
    });
  } catch(e) {
    if (e.name === "NotFoundError") throw new Error("Tidak ada printer dipilih");
    throw new Error("Scan gagal: " + e.message);
  }

  // TAHAP 2: Koneksi GATT (menggantikan bluetoothSerial.connect)
  onStatus("🔗 Menghubungkan ke " + (device.name || "printer") + "...");
  let server;
  try {
    server = await device.gatt.connect();
  } catch(e) {
    throw new Error("Koneksi gagal: " + e.message);
  }

  // TAHAP 3: Cari write channel (menggantikan bluetoothSerial.write target)
  onStatus("📡 Mencari jalur tulis...");
  const writeChar = await findWriteCharacteristic(server);
  if (!writeChar) {
    server.disconnect();
    throw new Error("Write characteristic tidak ditemukan di printer ini");
  }

  // TAHAP 4: Encode & kirim (persis seperti bluetoothSerial.write di modulprinter.db)
  onStatus("📤 Mencetak...");
  const textBytes = typeof rawText === "string"
    ? new TextEncoder().encode(rawText)
    : rawText;

  try {
    // ESC init dulu sebelum data (dari ESC_INIT di modulprinter.db)
    const initBytes = new Uint8Array([0x1B, 0x40]);
    await writeChar.writeValueWithoutResponse
      ? writeChar.writeValueWithoutResponse(initBytes).catch(() => writeChar.writeValue(initBytes))
      : writeChar.writeValue(initBytes);
    await new Promise(r => setTimeout(r, 100));

    // Kirim data dalam chunk kecil
    await writeInChunks(writeChar, textBytes, 100);

    // Potong kertas (GS V A 0)
    const cutBytes = new Uint8Array([0x1D, 0x56, 0x41, 0x00]);
    await new Promise(r => setTimeout(r, 200));
    if (writeChar.properties.writeWithoutResponse) {
      await writeChar.writeValueWithoutResponse(cutBytes).catch(() => {});
    } else {
      await writeChar.writeValue(cutBytes).catch(() => {});
    }
  } catch(e) {
    server.disconnect();
    throw e;
  }

  // TAHAP 5: Disconnect (persis seperti bluetoothSerial.disconnect)
  onStatus("✅ Berhasil dicetak!");
  setTimeout(() => { try { server.disconnect(); } catch {} }, 1000);
  return true;
}

// ── printWebBluetooth — adaptasi printNativeBluetooth dari modulprinter.db ──
async function printWebBluetooth(r, settings, onStatus) {
  // Generate ESC/POS string sama persis seperti getESCPOSData di modulprinter.db
  const rawText = getESCPOSData(r, settings);
  return printRawWebBluetooth(rawText, onStatus);
}

// ── HTML fallback untuk browser print dialog ─────────────────────
function buildReceiptHTML(r, settings) {
  const fRp = n => "Rp " + new Intl.NumberFormat("id-ID").format(n);
  const rows = r.items.map(it =>
    `<tr><td>${it.name}</td><td style="text-align:center">${it.qty}x</td><td style="text-align:right">${fRp(it.price)}</td><td style="text-align:right">${fRp(it.qty*it.price)}</td></tr>`
  ).join("");
  return `<!DOCTYPE html><html><head>
<meta charset="UTF-8"/><meta name="viewport" content="width=device-width,initial-scale=1"/>
<title>Struk ${r.id}</title>
<style>
  @page{size:58mm auto;margin:2mm}
  *{margin:0;padding:0;box-sizing:border-box}
  body{font-family:'Courier New',monospace;font-size:11px;width:54mm;padding:2mm;color:#000;background:#fff}
  h2{font-size:13px;text-align:center;font-weight:bold;margin-bottom:1px}
  .sub{font-size:10px;text-align:center;color:#333;margin-bottom:4px}
  .divider{border:none;border-top:1px dashed #000;margin:4px 0}
  table{width:100%;border-collapse:collapse;margin:2px 0}
  td{padding:1px;vertical-align:top;font-size:10px}
  .grand td{font-weight:bold;font-size:12px;border-top:1px dashed #000;padding-top:3px}
  .footer{text-align:center;margin-top:6px;font-size:10px;color:#555}
  @media print{button{display:none!important}}
</style></head><body>
<h2>${settings.storeName}</h2><div class="sub">${settings.address}</div>
<hr class="divider"/>
<p>No: ${r.id}</p>
<p>Tgl: ${new Date(r.date).toLocaleString("id-ID",{dateStyle:"short",timeStyle:"short"})}</p>
<p>Kasir: ${r.cashier||settings.cashier}</p>
<hr class="divider"/>
<table><tbody>${rows}</tbody></table>
<hr class="divider"/>
<table><tbody>
<tr><td>Subtotal</td><td style="text-align:right">${fRp(r.subtotal)}</td></tr>
${r.discount>0?`<tr><td>Diskon</td><td style="text-align:right">-${fRp(r.discount)}</td></tr>`:""}
${r.tax>0?`<tr><td>Pajak</td><td style="text-align:right">${fRp(r.tax)}</td></tr>`:""}
<tr class="grand"><td>TOTAL</td><td style="text-align:right">${fRp(r.total)}</td></tr>
<tr><td>Bayar(${r.payment})</td><td style="text-align:right">${fRp(r.cashPaid)}</td></tr>
${r.change>0?`<tr><td>Kembalian</td><td style="text-align:right">${fRp(r.change)}</td></tr>`:""}
</tbody></table>
${r.note?`<hr class="divider"/><p>Catatan: ${r.note}</p>`:""}
<hr class="divider"/>
<div class="footer"><p>Terima kasih!</p><p style="font-size:9px;margin-top:2px">Kasir-Pro</p></div>
<br/><br/>
<script>window.onload=()=>{setTimeout(()=>window.print(),300)}<\/script>
</body></html>`;
}

function openPrintWindow(r, settings) {
  const html = buildReceiptHTML(r, settings);
  const w = window.open("","_blank","width=320,height=600,menubar=no,toolbar=no,scrollbars=yes");
  if (w) { w.document.open(); w.document.write(html); w.document.close(); return true; }
  return false;
}




export default function App() {
  const [page, setPage] = useState("pos");
  const [products, setProducts] = useState(()=>{try{const s=localStorage.getItem("kasir_products");return s?JSON.parse(s):INITIAL_PRODUCTS;}catch{return INITIAL_PRODUCTS;}});
  const [transactions, setTransactions] = useState(()=>{try{const s=localStorage.getItem("kasir_transactions");return s?JSON.parse(s):INITIAL_TRANSACTIONS;}catch{return INITIAL_TRANSACTIONS;}});
  const [cart, setCart] = useState([]);
  const [searchQ, setSearchQ] = useState("");
  const [category, setCategory] = useState("Semua");
  const [discount, setDiscount] = useState("");
  const [discountType, setDiscountType] = useState("pct");
  const [note, setNote] = useState("");
  const [showPayment, setShowPayment] = useState(false);
  const [showReceipt, setShowReceipt] = useState(null);
  const [showProductForm, setShowProductForm] = useState(null);
  const [showPosScanner, setShowPosScanner] = useState(false);
  const [payMethod, setPayMethod] = useState("cash");
  const [cashPaid, setCashPaid] = useState("");
  const [toasts, setToasts] = useState([]);
  const DEFAULT_SETTINGS = {storeName:"Kasir-Pro",address:"Jl. Merdeka No. 17",taxRate:10,taxEnabled:true,receiptPrint:true,lowStockAlert:10,cashier:"Budi"};
  const [settings, setSettings] = useState(()=>{try{const s=localStorage.getItem("kasir_settings");return s?{...DEFAULT_SETTINGS,...JSON.parse(s)}:DEFAULT_SETTINGS;}catch{return DEFAULT_SETTINGS;}});
  const [trxSearch, setTrxSearch] = useState("");
  const [prodSearch, setProdSearch] = useState("");
  const [printerStatus, setPrinterStatus] = useState("");
  const [isPrinting, setIsPrinting] = useState(false);

  // ── Catatan Ternak ──
  const [pembelian, setPembelian] = useState(()=>{try{const s=localStorage.getItem("kasir_pembelian");return s?JSON.parse(s):[];}catch{return [];}});
  const [penjualan, setPenjualan] = useState(()=>{try{const s=localStorage.getItem("kasir_penjualan");return s?JSON.parse(s):[];}catch{return [];}});
  const [showFormPembelian, setShowFormPembelian] = useState(false);
  const [showFormPenjualan, setShowFormPenjualan] = useState(false);
  const [formPembelian, setFormPembelian] = useState({barcode:"",produk:"",satuan:"",hargaSatuan:"",jumlahBarang:"",supplier:"",tanggal:new Date().toISOString().split("T")[0],keterangan:""});
  const [editPembelianId, setEditPembelianId] = useState(null);
  const [formPenjualan, setFormPenjualan] = useState({produk:"",satuan:"",hargaSatuan:"",jumlahBarang:"",tanggal:new Date().toISOString().split("T")[0],keterangan:""});
  const [editPenjualanId, setEditPenjualanId] = useState(null);
  const [tabTernak, setTabTernak] = useState("pembelian"); // pembelian | penjualan | rekap
  const [darkMode, setDarkMode] = useState(()=>{
    try{ return localStorage.getItem("kasir_theme")==="light"?false:true; }catch{ return true; }
  });

  useEffect(()=>{
    document.documentElement.setAttribute("data-theme", darkMode?"dark":"light");
    try{ localStorage.setItem("kasir_theme", darkMode?"dark":"light"); }catch{}
  },[darkMode]);
  const [now, setNow] = useState(new Date());
  const clockRef = useRef(null);

  useEffect(()=>{clockRef.current=setInterval(()=>setNow(new Date()),1000);return()=>clearInterval(clockRef.current);},[]);
  useEffect(()=>{try{localStorage.setItem("kasir_products",JSON.stringify(products));}catch{}},[products]);
  useEffect(()=>{try{localStorage.setItem("kasir_transactions",JSON.stringify(transactions));}catch{}},[transactions]);
  useEffect(()=>{try{localStorage.setItem("kasir_settings",JSON.stringify(settings));}catch{}},[settings]);
  useEffect(()=>{try{localStorage.setItem("kasir_pembelian",JSON.stringify(pembelian));}catch{}},[pembelian]);
  useEffect(()=>{try{localStorage.setItem("kasir_penjualan",JSON.stringify(penjualan));}catch{}},[penjualan]);

  const addToast = useCallback((msg,type="success")=>{
    const id=Date.now();
    setToasts(t=>[...t,{id,msg,type}]);
    setTimeout(()=>setToasts(t=>t.filter(x=>x.id!==id)),3000);
  },[]);

  const addToCart = (product) => {
    if (product.stock<=0){addToast("Stok habis!","error");return;}
    setCart(c=>{
      const ex=c.find(x=>x.id===product.id);
      if(ex){if(ex.qty>=product.stock){addToast("Stok tidak cukup!","error");return c;}return c.map(x=>x.id===product.id?{...x,qty:x.qty+1}:x);}
      return [...c,{...product,qty:1}];
    });
  };
  const updateQty = (id,delta) => setCart(c=>c.map(x=>x.id===id?{...x,qty:Math.max(1,x.qty+delta)}:x));
  const removeFromCart = (id) => setCart(c=>c.filter(x=>x.id!==id));
  const clearCart = () => {setCart([]);setDiscount("");setNote("");};

  const subtotal = cart.reduce((s,x)=>s+x.price*x.qty,0);
  const discountAmt = discountType==="pct"?Math.round(subtotal*(parseFloat(discount)||0)/100):(parseFloat(discount)||0);
  const taxable = subtotal-discountAmt;
  const taxAmt = settings.taxEnabled?Math.round(taxable*settings.taxRate/100):0;
  const total = taxable+taxAmt;
  const cashPaidNum = parseFloat(cashPaid)||0;
  const change = cashPaidNum-total;
  const cartTotal = cart.reduce((s,x)=>s+x.qty,0);
  const cartQtyMap = {};
  cart.forEach(x=>{cartQtyMap[x.id]=x.qty;});

  const filteredProducts = products.filter(p=>{
    const matchCat=category==="Semua"||p.category===category;
    const matchQ=!searchQ||p.name.toLowerCase().includes(searchQ.toLowerCase())||p.barcode.includes(searchQ)||p.sku.toLowerCase().includes(searchQ.toLowerCase());
    return matchCat&&matchQ;
  });

  const handleCheckout = () => {
    if(cart.length===0){addToast("Keranjang kosong!","error");return;}
    if(payMethod==="cash"&&cashPaidNum<total){addToast("Uang kurang!","error");return;}
    const trx={id:genId(),date:new Date().toISOString(),items:cart.map(x=>({name:x.name,qty:x.qty,price:x.price})),subtotal,discount:discountAmt,tax:taxAmt,total,payment:payMethod,cashPaid:payMethod==="cash"?cashPaidNum:total,change:payMethod==="cash"?Math.max(0,change):0,cashier:settings.cashier,note};
    setProducts(ps=>ps.map(p=>{const ci=cart.find(x=>x.id===p.id);return ci?{...p,stock:p.stock-ci.qty}:p;}));
    setTransactions(t=>[trx,...t]);
    setShowPayment(false); setShowReceipt(trx); clearCart(); setCashPaid("");
    addToast(`Transaksi berhasil! ${formatRp(trx.total)}`,"success");
  };

  const todayStr = new Date().toISOString().slice(0,10);
  const todayTrx = transactions.filter(t=>t.date.startsWith(todayStr));
  const todayRevenue = todayTrx.reduce((s,t)=>s+t.total,0);
  const allRevenue = transactions.reduce((s,t)=>s+t.total,0);
  const topProducts = (()=>{const map={};transactions.forEach(t=>t.items.forEach(it=>{if(!map[it.name])map[it.name]={qty:0,rev:0};map[it.name].qty+=it.qty;map[it.name].rev+=it.qty*it.price;}));return Object.entries(map).sort((a,b)=>b[1].rev-a[1].rev).slice(0,5);})();
  const quickAmounts = [total,Math.ceil(total/10000)*10000,Math.ceil(total/50000)*50000,100000,200000].filter((v,i,a)=>a.indexOf(v)===i).slice(0,4);

  const cartProps = {cart,note,setNote,discount,setDiscount,discountType,setDiscountType,subtotal,discountAmt,taxAmt,total,settings,updateQty,removeFromCart,clearCart,setShowPayment};

  const handleTambahPembelian = () => {
    if(!formPembelian.produk || !formPembelian.satuan || !formPembelian.hargaSatuan || !formPembelian.jumlahBarang){
      addToast("Isi field: Produk, Satuan, Harga Satuan, Jumlah Barang!","error");
      return;
    }
    const hargaSatuan = parseFloat(formPembelian.hargaSatuan) || 0;
    const jumlahBarang = parseInt(formPembelian.jumlahBarang) || 0;
    if(jumlahBarang <= 0){addToast("Jumlah Barang harus lebih dari 0!","error");return;}
    if(hargaSatuan <= 0){addToast("Harga Satuan harus lebih dari 0!","error");return;}
    const jumlahRp = jumlahBarang * hargaSatuan;
    
    // Cari produk berdasarkan barcode atau nama
    let matchedProduct = null;
    if(formPembelian.barcode){
      matchedProduct = products.find(p=>p.barcode===formPembelian.barcode);
    }
    if(!matchedProduct && formPembelian.produk){
      matchedProduct = products.find(p=>p.name.toLowerCase()===formPembelian.produk.toLowerCase());
    }
    
    // Tambah stok jika produk ditemukan
    if(matchedProduct){
      setProducts(ps=>ps.map(p=>p.id===matchedProduct.id?{...p,stock:p.stock+jumlahBarang}:p));
      addToast(`✅ Stok ${matchedProduct.name} ditambah ${jumlahBarang}!`,"success");
    }
    
    // Tambah catatan pembelian
    const newPembelian = {
      id:editPembelianId||genId(),
      barcode:formPembelian.barcode,
      jenis:formPembelian.produk,
      satuan:formPembelian.satuan,
      harga:hargaSatuan,
      jumlah:jumlahBarang,
      total:jumlahRp,
      tempat:formPembelian.supplier,
      tanggal:formPembelian.tanggal,
      catatan:formPembelian.keterangan
    };
    
    if(editPembelianId){
      // Update pembelian
      setPembelian(p=>p.map(x=>x.id===editPembelianId?newPembelian:x));
      addToast("Pembelian berhasil diperbarui!");
      setEditPembelianId(null);
    } else {
      // Tambah pembelian baru
      setPembelian(p=>[newPembelian,...p]);
      addToast("Pembelian berhasil dicatat!");
    }
    
    setFormPembelian({barcode:"",produk:"",satuan:"",hargaSatuan:"",jumlahBarang:"",supplier:"",tanggal:new Date().toISOString().split("T")[0],keterangan:""});
    setShowFormPembelian(false);
  };

  const handleTambahPenjualan = () => {
    if(!formPenjualan.produk || !formPenjualan.satuan || !formPenjualan.hargaSatuan || !formPenjualan.jumlahBarang){
      addToast("Isi field: Produk, Satuan, Harga Satuan, Jumlah Barang!","error");
      return;
    }
    const hargaSatuan = parseFloat(formPenjualan.hargaSatuan) || 0;
    const jumlahBarang = parseInt(formPenjualan.jumlahBarang) || 0;
    if(jumlahBarang <= 0){addToast("Jumlah Barang harus lebih dari 0!","error");return;}
    if(hargaSatuan <= 0){addToast("Harga Satuan harus lebih dari 0!","error");return;}
    const jumlahRp = jumlahBarang * hargaSatuan;
    
    const newPenjualan = {
      id:editPenjualanId||genId(),
      produk:formPenjualan.produk,
      satuan:formPenjualan.satuan,
      harga:hargaSatuan,
      jumlah:jumlahBarang,
      total:jumlahRp,
      tanggal:formPenjualan.tanggal,
      catatan:formPenjualan.keterangan,
      type:"manual"
    };
    
    if(editPenjualanId){
      setPenjualan(p=>p.map(x=>x.id===editPenjualanId?newPenjualan:x));
      addToast("Penjualan berhasil diperbarui!");
      setEditPenjualanId(null);
    } else {
      setPenjualan(p=>[newPenjualan,...p]);
      addToast("Penjualan berhasil dicatat!");
    }
    
    setFormPenjualan({produk:"",satuan:"",hargaSatuan:"",jumlahBarang:"",tanggal:new Date().toISOString().split("T")[0],keterangan:""});
    setShowFormPenjualan(false);
  };

  const MOBILE_NAVS = [
    ["pos","storefront","Kasir"],
    ["products","inventory_2","Produk"],
    ["dashboard","analytics","Dasbor"],
    ["ternak","book","Catatan"],
    ["settings","settings","Setting"]
  ];
  const DESKTOP_NAVS = [
    ["pos","Kasir"],["dashboard","Dasbor"],["products","Produk"],
    ["transactions","Transaksi"],["ternak","Catatan"],["settings","Setting"]
  ];

  // Try to load logo from public
  const logoSrc = "/icon-192.png";

  return (
    <>
      <style>{css}</style>
      <div className="app">

        {/* ── M3 Top App Bar ── */}
        <header className="top-app-bar">
          <div className="brand">
            <img src={logoSrc} alt="logo" onError={e=>{e.target.style.display="none";}} />
            <div className="brand-text">
              <div className="brand-title">{settings.storeName}</div>
              <div className="brand-sub">{now.toLocaleTimeString("id-ID")}</div>
            </div>
          </div>
          <nav className="desktop-nav">
            {DESKTOP_NAVS.map(([k,label])=>(
              <button key={k} className={`dnav-btn${page===k?" active":""}`} onClick={()=>setPage(k)}>{label}</button>
            ))}
          </nav>
          <div className="actions">
            <div className="cashier-chip">
              <span className="material-icons-round" style={{fontSize:18,color:"var(--md-on-primary-container)"}}>person</span>
              <span className="cashier-name">{settings.cashier}</span>
            </div>
          </div>
        </header>

        <main className="content">

          {/* ── CART (mobile page) ── */}
          {page==="cart" && (
            <div className="cart-page-mobile">
              <CartPanel {...cartProps} />
            </div>
          )}

          {/* ── POS ── */}
          {page==="pos" && (
            <div className="pos">
              <div className="pos-left" style={{flex:1,display:"flex",flexDirection:"column",overflow:"hidden",minHeight:0,minWidth:0,background:"var(--md-surface)"}}>
                {/* M3 Search Bar */}
                <div className="search-section">
                  <div className="m3-search-bar">
                    <span className="material-icons-round" style={{fontSize:22,color:"var(--md-on-surface-variant)"}}>search</span>
                    <input className="m3-search-input" placeholder="Cari produk, SKU, atau barcode..." value={searchQ} onChange={e=>setSearchQ(e.target.value)} />
                    <button className="scan-icon-btn" onClick={()=>setShowPosScanner(true)}><span className="material-icons-round" style={{fontSize:20}}>qr_code_scanner</span></button>
                  </div>
                </div>

                {/* M3 Filter Chips */}
                <div className="filter-chips-wrap">
                  {CATEGORIES.map(c=>(
                    <button key={c} className={`m3-chip${category===c?" active":""}`} onClick={()=>setCategory(c)}>
                      {category===c
                        ? <span className="material-icons-round" style={{fontSize:16}}>check</span>
                        : <span className="material-icons-round" style={{fontSize:16}}>{CATEGORY_MAT_ICONS[c]}</span>
                      }
                      {c}
                    </button>
                  ))}
                </div>

                {/* Product Grid */}
                <div className="product-grid" style={{
                  flex:1,
                  overflowY:"auto",
                  overflowX:"hidden",
                  WebkitOverflowScrolling:"touch",
                  padding:"10px",
                  display:"grid",
                  gridTemplateColumns:"repeat(2,1fr)",
                  gap:"8px",
                  alignContent:"start",
                  minHeight:0,
                  background:"var(--md-surface)",
                }}>
                  {filteredProducts.length===0 && (
                    <div style={{gridColumn:"1/-1",textAlign:"center",color:"var(--md-on-surface-variant)",padding:"48px 0",fontSize:14,fontFamily:"var(--font-brand)"}}>
                      <div style={{fontSize:48,marginBottom:8}}>🔍</div>
                      Produk tidak ditemukan
                    </div>
                  )}
                  {filteredProducts.map(p=>(
                    <div
                      key={p.id}
                      onClick={()=>addToCart(p)}
                      className="m3-product-card-wrap"
                      style={{
                        background:"var(--md-surface-2)",
                        borderRadius:12,
                        border:"1px solid var(--md-outline-variant)",
                        cursor:p.stock===0?"not-allowed":"pointer",
                        opacity:p.stock===0?0.55:1,
                        position:"relative",
                        WebkitTapHighlightColor:"transparent",
                        padding:"10px",
                        boxSizing:"border-box",
                      }}
                    >
                      {/* Qty badge */}
                      {cartQtyMap[p.id] && (
                        <div style={{
                          position:"absolute", top:6, right:6,
                          background:"#C2E7FF", color:"#003352",
                          borderRadius:99, minWidth:20, height:20,
                          fontSize:10, fontWeight:700,
                          display:"flex", alignItems:"center", justifyContent:"center",
                          padding:"0 4px",
                        }}>{cartQtyMap[p.id]}</div>
                      )}

                      {/* Kategori */}
                      <div className="card-category-label" style={{
                        fontSize:10, fontWeight:700,
                        color:CATEGORY_COLORS[p.category]||"var(--md-on-surface-variant)",
                        textTransform:"uppercase",
                        letterSpacing:"0.5px",
                        marginBottom:4,
                      }}>{p.category}</div>

                      {/* Nama */}
                      <div className="card-name-text" style={{
                        fontSize:13, fontWeight:700,
                        color:"var(--md-on-surface)",
                        lineHeight:"1.4",
                        marginBottom:6,
                        whiteSpace:"normal",
                        wordBreak:"break-word",
                      }}>{p.name}</div>

                      {/* Harga */}
                      <div className="card-price-text" style={{
                        fontSize:14, fontWeight:800,
                        color:"var(--md-primary)",
                        marginBottom:8,
                      }}>{formatRp(p.price)}</div>

                      {/* Stok + tombol */}
                      <div style={{display:"flex",alignItems:"center",justifyContent:"space-between"}}>
                        <div className="card-stock-text" style={{
                          fontSize:11, fontWeight:500,
                          color:p.stock===0?"var(--md-error)":p.stock<=settings.lowStockAlert?"#F59E0B":"var(--md-on-surface-variant)",
                        }}>
                          {p.stock===0?"HABIS":p.stock<=settings.lowStockAlert?`⚠ Sisa ${p.stock}`:`Stok ${p.stock}`}
                        </div>
                        <button
                          className="card-add-btn-inner"
                          onClick={e=>{e.stopPropagation();addToCart(p);}}
                          style={{
                            width:34, height:34, borderRadius:"50%",
                            background:"var(--md-primary-container)",
                            color:"var(--md-on-primary-container)",
                            border:"none", cursor:"pointer",
                            fontSize:22, lineHeight:1,
                            display:"flex", alignItems:"center", justifyContent:"center",
                            flexShrink:0,
                          }}
                        >+</button>
                      </div>
                    </div>
                  ))}
                </div>

                {/* Floating Cart Summary Bar (mobile) */}
                <div className={`cart-fab-bar${cart.length>0?" visible":""}`} style={{transform:cart.length>0?"translateY(0)":"translateY(120%)",opacity:cart.length>0?1:0}}>
                  <div className="cart-fab-inner" onClick={()=>setPage("cart")}>
                    <div className="cart-fab-info">
                      <div className="cart-fab-count">{cartTotal} item dipilih</div>
                      <div className="cart-fab-total">{formatRp(total)}</div>
                    </div>
                    <div className="cart-fab-btn">🛒 Pesanan</div>
                  </div>
                </div>
              </div>

              {/* Desktop Cart */}
              <div className="pos-right">
                <CartPanel {...cartProps} />
              </div>
            </div>
          )}

          {/* ── DASHBOARD ── */}
          {page==="dashboard" && (
            <div className="page-content">
              <div className="page-title">Dashboard</div>
              <div className="m3-stats-grid">
                <div className="m3-stat-card"><div className="stat-icon">💰</div><div className="stat-label">Hari Ini</div><div className="stat-val primary">{formatRp(todayRevenue)}</div><div className="stat-sub">{todayTrx.length} transaksi</div></div>
                <div className="m3-stat-card"><div className="stat-icon">📈</div><div className="stat-label">Total</div><div className="stat-val success">{formatRp(allRevenue)}</div><div className="stat-sub">{transactions.length} transaksi</div></div>
                <div className="m3-stat-card"><div className="stat-icon">📦</div><div className="stat-label">Produk</div><div className="stat-val">{products.length}</div><div className="stat-sub">{products.filter(p=>p.stock<=settings.lowStockAlert).length} stok rendah</div></div>
                <div className="m3-stat-card"><div className="stat-icon">🧾</div><div className="stat-label">Rata-rata</div><div className="stat-val">{transactions.length>0?formatRp(Math.round(allRevenue/transactions.length)):"Rp 0"}</div><div className="stat-sub">per transaksi</div></div>
              </div>
              <div className="m3-card"><h4>Produk Terlaris</h4>{topProducts.map(([name,data])=>(<div key={name} className="list-row"><span className="lr-name">{name}</span><span className="lr-qty">{data.qty}x</span><span className="lr-rev">{formatRp(data.rev)}</span></div>))}</div>
              <div className="m3-card"><h4>Stok Rendah</h4>{products.filter(p=>p.stock<=settings.lowStockAlert).length===0?<div style={{color:"var(--md-on-surface-variant)",fontSize:13,padding:"8px 0"}}>✅ Semua stok aman</div>:products.filter(p=>p.stock<=settings.lowStockAlert).map(p=>(<div key={p.id} className="list-row"><span className="lr-name">{p.name}</span><span style={{fontFamily:"var(--font-mono)",fontSize:11,color:p.stock===0?"var(--md-error)":"#FCD34D"}}>{p.stock===0?"HABIS":`Sisa ${p.stock}`}</span></div>))}</div>
              <div style={{marginTop:24}}>
                <div className="page-toolbar">
                  <div className="page-title" style={{marginBottom:0,flex:1}}>Riwayat Transaksi</div>
                  <input className="m3-text-field" placeholder="🔍 Cari..." value={trxSearch} onChange={e=>setTrxSearch(e.target.value)} style={{maxWidth:140}} />
                  <button className="btn btn-outlined" style={{padding:"8px 12px",fontSize:12,whiteSpace:"nowrap",display:"flex",alignItems:"center",gap:4}}
                    onClick={()=>exportTransaksiExcel(transactions)}>
                    <span className="material-icons-round" style={{fontSize:16}}>download</span>Excel
                  </button>
                </div>
                <div className="table-wrap">
                  <table className="m3-table">
                    <thead><tr><th>ID</th><th>Waktu</th><th>Kasir</th><th>Metode</th><th>Total</th><th>Aksi</th></tr></thead>
                    <tbody>
                      {transactions.filter(t=>!trxSearch||t.id.toLowerCase().includes(trxSearch.toLowerCase())||t.cashier?.toLowerCase().includes(trxSearch.toLowerCase())).map(t=>(
                        <tr key={t.id}>
                          <td style={{fontFamily:"var(--font-mono)",fontSize:10,color:"#93C5FD"}}>{t.id}</td>
                          <td style={{fontSize:11,color:"var(--md-on-surface-variant)",whiteSpace:"nowrap"}}>{formatDate(t.date)}</td>
                          <td style={{fontSize:12,fontFamily:"var(--font-brand)"}}>{t.cashier||"-"}</td>
                          <td><span className="m3-badge" style={{background:"var(--md-surface-3)",color:"var(--md-on-surface-variant)",textTransform:"uppercase"}}>{t.payment}</span></td>
                          <td style={{fontFamily:"var(--font-mono)",fontWeight:700,color:"var(--md-primary)",whiteSpace:"nowrap"}}>{formatRp(t.total)}</td>
                          <td><div style={{display:"flex",gap:4}}><button className="m3-icon-btn" onClick={()=>setShowReceipt(t)}>🧾</button><button className="m3-icon-btn danger" onClick={()=>{if(window.confirm("Hapus transaksi ini?")){setTransactions(t2=>t2.filter(x=>x.id!==t.id));addToast("Transaksi dihapus");}}}>🗑️</button></div></td>
                        </tr>
                      ))}
                    </tbody>
                  </table>
                </div>
              </div>
            </div>
          )}

          {/* ── PRODUCTS ── */}
          {page==="products" && (
            <div className="page-content">
              <div className="page-toolbar">
                <div className="page-title" style={{marginBottom:0,flex:1}}>Produk</div>
                <input className="m3-text-field" placeholder="🔍 Cari..." value={prodSearch} onChange={e=>setProdSearch(e.target.value)} style={{maxWidth:120}} />
                <button className="btn btn-outlined" style={{padding:"8px 10px",fontSize:12,whiteSpace:"nowrap",display:"flex",alignItems:"center",gap:4}}
                  onClick={()=>exportProdukExcel(products)}>
                  <span className="material-icons-round" style={{fontSize:16}}>download</span>Excel
                </button>
                <button className="btn btn-filled" style={{whiteSpace:"nowrap",padding:"8px 12px",fontSize:13}} onClick={()=>setShowProductForm({id:null,name:"",price:"",category:"Makanan",stock:"",sku:"",barcode:""})}>+ Tambah</button>
              </div>
              <div className="table-wrap">
                <table className="m3-table">
                  <thead><tr><th>SKU</th><th>Nama</th><th>Kategori</th><th>Harga</th><th>Stok</th><th>Aksi</th></tr></thead>
                  <tbody>
                    {products.filter(p=>!prodSearch||p.name.toLowerCase().includes(prodSearch.toLowerCase())||p.sku.toLowerCase().includes(prodSearch.toLowerCase())).map(p=>(
                      <tr key={p.id}>
                        <td style={{fontFamily:"var(--font-mono)",fontSize:10,color:"#93C5FD"}}>{p.sku}</td>
                        <td style={{fontWeight:500,fontFamily:"var(--font-brand)"}}>{p.name}</td>
                        <td><span className={`m3-badge badge-${p.category.toLowerCase()}`}>{p.category}</span></td>
                        <td style={{fontFamily:"var(--font-mono)",color:"var(--md-primary)",fontWeight:600}}>{formatRp(p.price)}</td>
                        <td style={{fontFamily:"var(--font-mono)",color:p.stock===0?"var(--md-error)":p.stock<=settings.lowStockAlert?"#FCD34D":"var(--md-on-surface)"}}>{p.stock}</td>
                        <td><div style={{display:"flex",gap:4}}><button className="m3-icon-btn" onClick={()=>setShowProductForm({...p})}>✏️</button><button className="m3-icon-btn danger" onClick={()=>{if(window.confirm("Hapus produk ini?")){setProducts(ps=>ps.filter(x=>x.id!==p.id));addToast("Produk dihapus");}}}>🗑️</button></div></td>
                      </tr>
                    ))}
                  </tbody>
                </table>
              </div>
            </div>
          )}

          {/* ── SETTINGS ── */}
          {/* ── CATATAN TERNAK ── */}
          {page==="ternak" && (
            <div className="page-content">
              <div className="page-title">📋 Catatan Usaha</div>

              {/* Tab */}
              <div style={{display:"flex",gap:8,marginBottom:16,background:"var(--md-surface-2)",borderRadius:12,padding:4}}>
                {[["pembelian","Pembelian","shopping_cart"],["penjualan","Penjualan","sell"],["rekap","Rekap","summarize"]].map(([k,label,icon])=>(
                  <button key={k} onClick={()=>setTabTernak(k)} style={{
                    flex:1,display:"flex",alignItems:"center",justifyContent:"center",gap:6,
                    padding:"10px 8px",borderRadius:10,border:"none",cursor:"pointer",
                    background:tabTernak===k?"var(--md-primary-container)":"transparent",
                    color:tabTernak===k?"var(--md-on-primary-container)":"var(--md-on-surface-variant)",
                    fontWeight:600,fontSize:12,fontFamily:"var(--font-brand)",transition:"all .2s",
                  }}>
                    <span className="material-icons-round" style={{fontSize:16}}>{icon}</span>
                    {label}
                  </button>
                ))}
              </div>

              {/* ── TAB PEMBELIAN ── */}
              {tabTernak==="pembelian" && (
                <div>
                  <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:12,gap:8}}>
                    <div style={{fontSize:14,fontWeight:600,color:"var(--md-on-surface-variant)"}}>
                      {pembelian.length} catatan
                    </div>
                    <div style={{display:"flex",gap:8}}>
                      <button className="btn btn-outlined" style={{padding:"8px 10px",fontSize:12,display:"flex",alignItems:"center",gap:4}}
                        onClick={()=>exportPembelianExcel(pembelian)}>
                        <span className="material-icons-round" style={{fontSize:15}}>download</span>Excel
                      </button>
                      <button className="btn btn-filled" style={{padding:"8px 14px",fontSize:13}} onClick={()=>{setFormPembelian({barcode:"",produk:"",satuan:"",hargaSatuan:"",jumlahBarang:"",supplier:"",tanggal:new Date().toISOString().split("T")[0],keterangan:""});setEditPembelianId(null);setShowFormPembelian(true);}}>
                        + Tambah
                      </button>
                    </div>
                  </div>

                  {pembelian.length===0 && (
                    <div style={{textAlign:"center",padding:"40px 0",color:"var(--md-on-surface-variant)"}}>
                      <span className="material-icons-round" style={{fontSize:48,display:"block",marginBottom:8,opacity:.4}}>shopping_cart</span>
                      Belum ada catatan pembelian
                    </div>
                  )}

                  {pembelian.slice().reverse().map((item,i)=>(
                    <div key={item.id} style={{
                      background:"var(--md-surface-2)",borderRadius:12,
                      border:"1px solid var(--md-outline-variant)",
                      padding:"12px 14px",marginBottom:8,
                    }}>
                      <div style={{display:"flex",justifyContent:"space-between",alignItems:"flex-start",marginBottom:6}}>
                        <div style={{flex:1}}>
                          <div style={{fontSize:14,fontWeight:700,color:"var(--md-on-surface)",marginBottom:2}}>{item.jenis}</div>
                          <div style={{fontSize:11,color:"var(--md-on-surface-variant)",marginBottom:4}}>{item.tanggal} · {item.tempat||"-"}</div>
                          <div style={{fontSize:11,color:"var(--md-on-surface-variant)"}}>{item.jumlah} {item.satuan} × {formatRp(item.harga)}</div>
                        </div>
                        <div style={{textAlign:"right"}}>
                          <div style={{fontSize:14,fontWeight:700,color:"var(--md-error)"}}>-{formatRp(item.harga*item.jumlah)}</div>
                        </div>
                      </div>
                      {item.catatan && <div style={{fontSize:12,color:"var(--md-on-surface-variant)",fontStyle:"italic",marginBottom:8}}>📝 {item.catatan}</div>}
                      <div style={{display:"flex",gap:6,justifyContent:"flex-end"}}>
                        <button onClick={()=>{
                          setFormPembelian({
                            barcode:item.barcode,
                            produk:item.jenis,
                            satuan:item.satuan,
                            hargaSatuan:item.harga.toString(),
                            jumlahBarang:item.jumlah.toString(),
                            supplier:item.tempat,
                            tanggal:item.tanggal,
                            keterangan:item.catatan
                          });
                          setEditPembelianId(item.id);
                          setShowFormPembelian(true);
                        }}
                          style={{background:"var(--md-secondary-container)",border:"none",cursor:"pointer",color:"var(--md-on-secondary-container)",padding:"6px 10px",borderRadius:8,fontSize:12,fontWeight:600,display:"flex",alignItems:"center",gap:4}}>
                          <span className="material-icons-round" style={{fontSize:16}}>edit</span>Edit
                        </button>
                        <button onClick={()=>{if(window.confirm("Hapus catatan ini?")){setPembelian(p=>p.filter(x=>x.id!==item.id));}}}
                          style={{background:"rgba(255,180,171,.1)",border:"1px solid rgba(255,180,171,.3)",cursor:"pointer",color:"var(--md-error)",padding:"6px 10px",borderRadius:8,fontSize:12,fontWeight:600,display:"flex",alignItems:"center",gap:4}}>
                          <span className="material-icons-round" style={{fontSize:16}}>delete_outline</span>Hapus
                        </button>
                      </div>
                    </div>
                  ))}
                </div>
              )}

              {/* ── TAB PENJUALAN ── */}
              {tabTernak==="penjualan" && (
                <div>
                  <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:12,gap:8}}>
                    <div style={{fontSize:14,fontWeight:600,color:"var(--md-on-surface-variant)"}}>
                      {(penjualan.length + transactions.length)} catatan
                    </div>
                    <div style={{display:"flex",gap:8}}>
                      <button className="btn btn-outlined" style={{padding:"8px 10px",fontSize:12,display:"flex",alignItems:"center",gap:4}}
                        onClick={()=>{
                          const allData = [...transactions.map(t=>({produk:"Multiple Items",satuan:"transaksi",harga:0,jumlah:1,tanggal:formatDate(t.date),catatan:"",type:"pos"})),...penjualan];
                          exportPenjualanExcel(allData);
                        }}>
                        <span className="material-icons-round" style={{fontSize:15}}>download</span>Excel
                      </button>
                      <button className="btn btn-filled" style={{padding:"8px 14px",fontSize:13}} onClick={()=>{setFormPenjualan({produk:"",satuan:"",hargaSatuan:"",jumlahBarang:"",tanggal:new Date().toISOString().split("T")[0],keterangan:""});setEditPenjualanId(null);setShowFormPenjualan(true);}}>
                        + Tambah
                      </button>
                    </div>
                  </div>

                  {(penjualan.length + transactions.length)===0 && (
                    <div style={{textAlign:"center",padding:"40px 0",color:"var(--md-on-surface-variant)"}}>
                      <span className="material-icons-round" style={{fontSize:48,display:"block",marginBottom:8,opacity:.4}}>sell</span>
                      Belum ada catatan penjualan
                    </div>
                  )}

                  {/* Transaksi POS */}
                  {transactions.slice().reverse().map((t,i)=>(
                    <div key={t.id} style={{
                      background:"var(--md-surface-2)",borderRadius:12,
                      border:"1px solid var(--md-outline-variant)",
                      padding:"12px 14px",marginBottom:8,
                    }}>
                      <div style={{display:"flex",justifyContent:"space-between",alignItems:"flex-start",marginBottom:6}}>
                        <div style={{flex:1}}>
                          <div style={{fontSize:14,fontWeight:700,color:"var(--md-on-surface)",marginBottom:2}}>Transaksi POS #{t.id}</div>
                          <div style={{fontSize:11,color:"var(--md-on-surface-variant)",marginBottom:4}}>{formatDate(t.date)} · {t.cashier||"-"}</div>
                          <div style={{fontSize:11,color:"var(--md-on-surface-variant)"}}>{t.items.length} item</div>
                        </div>
                        <div style={{textAlign:"right"}}>
                          <div style={{fontSize:14,fontWeight:700,color:"var(--md-success)"}}>+{formatRp(t.total)}</div>
                        </div>
                      </div>
                      <div style={{fontSize:11,color:"var(--md-on-surface-variant)",maxHeight:100,overflowY:"auto"}}>
                        {t.items.map((item,idx)=>(
                          <div key={idx} style={{padding:"4px 0"}}>• {item.qty}x {item.name}</div>
                        ))}
                      </div>
                    </div>
                  ))}

                  {/* Penjualan Manual */}
                  {penjualan.slice().reverse().map((item,i)=>(
                    <div key={item.id} style={{
                      background:"var(--md-surface-2)",borderRadius:12,
                      border:"1px solid var(--md-outline-variant)",
                      padding:"12px 14px",marginBottom:8,
                    }}>
                      <div style={{display:"flex",justifyContent:"space-between",alignItems:"flex-start",marginBottom:6}}>
                        <div style={{flex:1}}>
                          <div style={{fontSize:14,fontWeight:700,color:"var(--md-on-surface)",marginBottom:2}}>{item.produk}</div>
                          <div style={{fontSize:11,color:"var(--md-on-surface-variant)",marginBottom:4}}>{item.tanggal}</div>
                          <div style={{fontSize:11,color:"var(--md-on-surface-variant)"}}>{item.jumlah} {item.satuan} × {formatRp(item.harga)}</div>
                        </div>
                        <div style={{textAlign:"right"}}>
                          <div style={{fontSize:14,fontWeight:700,color:"var(--md-success)"}}>+{formatRp(item.harga*item.jumlah)}</div>
                        </div>
                      </div>
                      {item.catatan && <div style={{fontSize:12,color:"var(--md-on-surface-variant)",fontStyle:"italic",marginBottom:8}}>📝 {item.catatan}</div>}
                      <div style={{display:"flex",gap:6,justifyContent:"flex-end"}}>
                        <button onClick={()=>{
                          setFormPenjualan({
                            produk:item.produk,
                            satuan:item.satuan,
                            hargaSatuan:item.harga.toString(),
                            jumlahBarang:item.jumlah.toString(),
                            tanggal:item.tanggal,
                            keterangan:item.catatan
                          });
                          setEditPenjualanId(item.id);
                          setShowFormPenjualan(true);
                        }}
                          style={{background:"var(--md-secondary-container)",border:"none",cursor:"pointer",color:"var(--md-on-secondary-container)",padding:"6px 10px",borderRadius:8,fontSize:12,fontWeight:600,display:"flex",alignItems:"center",gap:4}}>
                          <span className="material-icons-round" style={{fontSize:16}}>edit</span>Edit
                        </button>
                        <button onClick={()=>{if(window.confirm("Hapus catatan ini?")){setPenjualan(p=>p.filter(x=>x.id!==item.id));}}}
                          style={{background:"rgba(255,180,171,.1)",border:"1px solid rgba(255,180,171,.3)",cursor:"pointer",color:"var(--md-error)",padding:"6px 10px",borderRadius:8,fontSize:12,fontWeight:600,display:"flex",alignItems:"center",gap:4}}>
                          <span className="material-icons-round" style={{fontSize:16}}>delete_outline</span>Hapus
                        </button>
                      </div>
                    </div>
                  ))}
                </div>
              )}

              {/* ── TAB REKAP ── */}
              {tabTernak==="rekap" && (()=>{
                const totalBeli = pembelian.reduce((s,x)=>s+(x.harga*x.jumlah),0);
                const totalJualPOS = transactions.reduce((s,x)=>s+x.total,0);
                const totalJualManual = penjualan.reduce((s,x)=>s+(x.harga*x.jumlah),0);
                const totalJual = totalJualPOS + totalJualManual;
                const laba = totalJual - totalBeli;

                // Rekap pembelian per jenis
                const rekapBeli = {};
                pembelian.forEach(x=>{
                  if(!rekapBeli[x.jenis]) rekapBeli[x.jenis]={total:0,count:0};
                  rekapBeli[x.jenis].total += x.harga*x.jumlah;
                  rekapBeli[x.jenis].count += 1;
                });

                // Rekap penjualan per produk
                const rekapJual = {};
                penjualan.forEach(x=>{
                  if(!rekapJual[x.produk]) rekapJual[x.produk]={total:0,jumlah:0};
                  rekapJual[x.produk].total += x.harga*x.jumlah;
                  rekapJual[x.produk].jumlah += Number(x.jumlah);
                });

                return (
                  <div>
                    {/* Ringkasan */}
                    <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:10,marginBottom:16}}>
                      <div style={{background:"rgba(255,180,171,.1)",border:"1px solid rgba(255,180,171,.3)",borderRadius:12,padding:"14px 12px"}}>
                        <div style={{fontSize:11,fontWeight:600,color:"var(--md-error)",textTransform:"uppercase",letterSpacing:".5px",marginBottom:6}}>Total Pembelian</div>
                        <div style={{fontSize:18,fontWeight:800,color:"var(--md-error)",fontFamily:"var(--font-mono)"}}>{formatRp(totalBeli)}</div>
                        <div style={{fontSize:11,color:"var(--md-on-surface-variant)",marginTop:4}}>{pembelian.length} transaksi</div>
                      </div>
                      <div style={{background:"rgba(110,231,183,.1)",border:"1px solid rgba(110,231,183,.3)",borderRadius:12,padding:"14px 12px"}}>
                        <div style={{fontSize:11,fontWeight:600,color:"var(--md-success)",textTransform:"uppercase",letterSpacing:".5px",marginBottom:6}}>Total Penjualan</div>
                        <div style={{fontSize:18,fontWeight:800,color:"var(--md-success)",fontFamily:"var(--font-mono)"}}>{formatRp(totalJual)}</div>
                        <div style={{fontSize:11,color:"var(--md-on-surface-variant)",marginTop:4}}>{transactions.length + penjualan.length} transaksi</div>
                        <div style={{fontSize:10,color:"var(--md-on-surface-variant)",marginTop:2}}>POS: {transactions.length} + Manual: {penjualan.length}</div>
                      </div>
                    </div>

                    {/* Laba/Rugi */}
                    <div style={{
                      background:laba>=0?"rgba(110,231,183,.1)":"rgba(255,180,171,.1)",
                      border:`1px solid ${laba>=0?"rgba(110,231,183,.3)":"rgba(255,180,171,.3)"}`,
                      borderRadius:12,padding:"14px 16px",marginBottom:16,
                      display:"flex",justifyContent:"space-between",alignItems:"center",
                    }}>
                      <div>
                        <div style={{fontSize:13,fontWeight:600,color:"var(--md-on-surface-variant)"}}>{laba>=0?"Laba Bersih":"Rugi"}</div>
                        <div style={{fontSize:11,color:"var(--md-on-surface-variant)",marginTop:2}}>Penjualan - Pembelian</div>
                      </div>
                      <div style={{fontSize:22,fontWeight:800,color:laba>=0?"var(--md-success)":"var(--md-error)",fontFamily:"var(--font-mono)"}}>
                        {laba>=0?"+":""}{formatRp(laba)}
                      </div>
                    </div>

                    {/* Rekap Pembelian per Jenis */}
                    <div className="m3-card">
                      <h4>📥 Rekap Pembelian per Jenis</h4>
                      {Object.keys(rekapBeli).length===0 && <div style={{color:"var(--md-on-surface-variant)",fontSize:13}}>Belum ada data</div>}
                      {Object.entries(rekapBeli).sort((a,b)=>b[1].total-a[1].total).map(([jenis,data])=>(
                        <div key={jenis} className="list-row">
                          <span className="lr-name">{jenis}</span>
                          <span style={{fontSize:10,color:"var(--md-on-surface-variant)"}}>{data.count}x</span>
                          <span style={{fontFamily:"var(--font-mono)",fontSize:12,color:"var(--md-error)",fontWeight:700}}>{formatRp(data.total)}</span>
                        </div>
                      ))}
                    </div>

                    {/* Rekap Penjualan per Produk */}
                    <div className="m3-card">
                      <h4>📤 Rekap Penjualan per Produk</h4>
                      {Object.keys(rekapJual).length===0 && <div style={{color:"var(--md-on-surface-variant)",fontSize:13}}>Belum ada data</div>}
                      {Object.entries(rekapJual).sort((a,b)=>b[1].total-a[1].total).map(([produk,data])=>(
                        <div key={produk} className="list-row">
                          <span className="lr-name">{produk}</span>
                          <span style={{fontSize:10,color:"var(--md-on-surface-variant)"}}>{data.jumlah} pcs</span>
                          <span style={{fontFamily:"var(--font-mono)",fontSize:12,color:"var(--md-success)",fontWeight:700}}>{formatRp(data.total)}</span>
                        </div>
                      ))}
                    </div>
                  </div>
                );
              })()}

            </div>
          )}

          {page==="settings" && (
            <div className="page-content">
              <div className="page-title">Pengaturan</div>

              {/* 1. Informasi Toko */}
              <div className="m3-settings-section">
                <div className="settings-section-title">🏪 Informasi Toko</div>
                <div className="settings-row"><div><div className="slbl">Nama Toko</div></div><input className="m3-settings-input" value={settings.storeName} onChange={e=>setSettings(s=>({...s,storeName:e.target.value}))} /></div>
                <div className="settings-row"><div><div className="slbl">Alamat</div></div><input className="m3-settings-input" value={settings.address} onChange={e=>setSettings(s=>({...s,address:e.target.value}))} /></div>
                <div className="settings-row"><div><div className="slbl">Nama Kasir</div></div><input className="m3-settings-input" value={settings.cashier} onChange={e=>setSettings(s=>({...s,cashier:e.target.value}))} /></div>
              </div>

              {/* 2. Reset Data */}
              <div className="m3-settings-section">
                <div className="settings-section-title">🗑️ Reset Data</div>
                <div className="settings-row"><div><div className="slbl">Hapus Semua Transaksi</div><div className="sval">Data produk tidak ikut terhapus</div></div><button className="btn-danger-tonal" onClick={()=>{if(window.confirm("Hapus semua transaksi?")){setTransactions([]);addToast("Semua transaksi dihapus");}}}>Hapus</button></div>
                <div className="settings-row"><div><div className="slbl">Delete Semua Produk</div><div className="sval">Hapus semua data produk</div></div><button className="btn-danger-tonal" onClick={()=>{if(window.confirm("Delete semua produk?")){setProducts([]);addToast("Semua produk dihapus");}}}>Delete</button></div>
              </div>

              {/* 3. Tampilan */}
              <div className="m3-settings-section">
                <div className="settings-section-title">🎨 Tampilan</div>
                <div className="settings-row">
                  <div>
                    <div className="slbl">Mode Tampilan</div>
                    <div className="sval">{darkMode?"Mode Gelap aktif":"Mode Terang aktif"}</div>
                  </div>
                  <div style={{display:"flex",alignItems:"center",gap:10}}>
                    <button
                      onClick={()=>setDarkMode(false)}
                      style={{
                        display:"flex",flexDirection:"column",alignItems:"center",gap:4,
                        padding:"10px 16px",borderRadius:12,border:"2px solid",
                        borderColor:!darkMode?"var(--md-primary)":"var(--md-outline-variant)",
                        background:!darkMode?"var(--md-primary-container)":"transparent",
                        color:!darkMode?"var(--md-on-primary-container)":"var(--md-on-surface-variant)",
                        cursor:"pointer",transition:"all .2s",minWidth:70,
                      }}
                    >
                      <span className="material-icons-round" style={{fontSize:22}}>light_mode</span>
                      <span style={{fontSize:11,fontWeight:600,fontFamily:"var(--font-brand)"}}>Terang</span>
                    </button>
                    <button
                      onClick={()=>setDarkMode(true)}
                      style={{
                        display:"flex",flexDirection:"column",alignItems:"center",gap:4,
                        padding:"10px 16px",borderRadius:12,border:"2px solid",
                        borderColor:darkMode?"var(--md-primary)":"var(--md-outline-variant)",
                        background:darkMode?"var(--md-primary-container)":"transparent",
                        color:darkMode?"var(--md-on-primary-container)":"var(--md-on-surface-variant)",
                        cursor:"pointer",transition:"all .2s",minWidth:70,
                      }}
                    >
                      <span className="material-icons-round" style={{fontSize:22}}>dark_mode</span>
                      <span style={{fontSize:11,fontWeight:600,fontFamily:"var(--font-brand)"}}>Gelap</span>
                    </button>
                  </div>
                </div>
              </div>

              {/* 4. Pajak */}
              <div className="m3-settings-section">
                <div className="settings-section-title">💰 Pajak</div>
                <div className="settings-row"><div><div className="slbl">Aktifkan Pajak</div><div className="sval">PPN / Pajak penjualan</div></div><button className={`m3-toggle${settings.taxEnabled?" on":""}`} onClick={()=>setSettings(s=>({...s,taxEnabled:!s.taxEnabled}))} /></div>
                <div className="settings-row"><div><div className="slbl">Persentase Pajak (%)</div></div><input className="m3-settings-input" type="number" value={settings.taxRate} onChange={e=>setSettings(s=>({...s,taxRate:parseFloat(e.target.value)||0}))} /></div>
              </div>

              {/* 5. Stok */}
              <div className="m3-settings-section">
                <div className="settings-section-title">📦 Stok</div>
                <div className="settings-row"><div><div className="slbl">Batas Stok Rendah</div></div><input className="m3-settings-input" type="number" value={settings.lowStockAlert} onChange={e=>setSettings(s=>({...s,lowStockAlert:parseInt(e.target.value)||0}))} /></div>
              </div>

              {/* 6. Manajemen Produk */}
              <div className="m3-settings-section">
                <div className="settings-section-title">📦 Manajemen Produk</div>
                <div className="settings-row">
                  <div><div className="slbl">Kelola Produk</div><div className="sval">Tambah, edit, hapus produk & stok</div></div>
                  <button className="btn btn-filled" style={{padding:"8px 16px",fontSize:13}} onClick={()=>setPage("products")}>Buka →</button>
                </div>
              </div>

              {/* 7. Printer Bluetooth GATT */}
              <div className="m3-settings-section">
                <div className="settings-section-title">🖨️ Printer Bluetooth GATT</div>

                <div style={{margin:"0 16px 12px",padding:"12px",background:"rgba(110,231,183,.08)",border:"1px solid rgba(110,231,183,.2)",borderRadius:10}}>
                  <div style={{fontSize:12,fontWeight:700,color:"var(--md-success)",marginBottom:6}}>
                    ✅ Web Bluetooth GATT — EPPOS RPP02
                  </div>
                  <div style={{fontSize:12,color:"var(--md-on-surface-variant)",lineHeight:"1.8"}}>
                    <strong style={{color:"var(--md-on-surface)"}}>Syarat:</strong><br/>
                    • Browser Chrome Android 85+ (bukan Firefox/Safari)<br/>
                    • Akses dari <strong style={{color:"var(--md-primary)"}}>HTTPS</strong> (Vercel ✅ sudah HTTPS)<br/>
                    • Bluetooth HP aktif<br/>
                    • Printer RPP02 menyala<br/>
                    <br/>
                    <strong style={{color:"var(--md-on-surface)"}}>Cara pakai:</strong><br/>
                    1. Buka struk transaksi<br/>
                    2. Tekan <strong style={{color:"var(--md-on-surface)"}}>"Cetak via Bluetooth"</strong><br/>
                    3. Pilih <strong style={{color:"var(--md-on-surface)"}}>EPPOS RPP02</strong> dari dialog<br/>
                    4. Printer langsung cetak ✅
                  </div>
                </div>

                {/* Test print */}
                <div className="settings-row">
                  <div>
                    <div className="slbl">Test Print ke RPP02</div>
                    <div className="sval">Kirim struk test via Bluetooth GATT</div>
                  </div>
                  <button
                    className="btn btn-outlined"
                    style={{padding:"8px 12px",fontSize:12,whiteSpace:"nowrap"}}
                    disabled={isPrinting}
                    onClick={async()=>{
                      setIsPrinting(true);
                      setPrinterStatus("");
                      const testR = {
                        id:"TEST-"+Date.now().toString().slice(-6),
                        date:new Date().toISOString(),
                        items:[{name:"TEST CETAK RPP02",qty:1,price:0}],
                        subtotal:0,discount:0,tax:0,total:0,
                        payment:"test",cashPaid:0,change:0,
                        cashier:settings.cashier,note:"Printer berhasil konek!"
                      };
                      try {
                        await printWebBluetooth(testR, settings, setPrinterStatus);
                        addToast("✅ Test berhasil!");
                      } catch(e) {
                        setPrinterStatus("❌ "+e.message);
                        addToast("❌ "+e.message,"error");
                      } finally { setIsPrinting(false); }
                    }}>
                    {isPrinting?"⏳":"🖨️"} Test
                  </button>
                </div>

                {printerStatus && (
                  <div style={{
                    margin:"0 16px 12px",padding:"10px 12px",borderRadius:8,fontSize:12,
                    background:printerStatus.startsWith("✅")?"rgba(110,231,183,.1)":"rgba(255,180,171,.1)",
                    color:printerStatus.startsWith("✅")?"var(--md-success)":"var(--md-error)",
                    border:"1px solid",
                    borderColor:printerStatus.startsWith("✅")?"rgba(110,231,183,.3)":"rgba(255,180,171,.3)",
                  }}>{printerStatus}</div>
                )}

                <div className="settings-row">
                  <div><div className="slbl">Lebar Kertas</div><div className="sval">58mm — 32 karakter per baris</div></div>
                  <span style={{fontSize:12,color:"var(--md-on-surface-variant)",fontWeight:600}}>58mm</span>
                </div>

                <div className="settings-row">
                  <div><div className="slbl">Protocol</div><div className="sval">Web Bluetooth GATT / ESC-POS</div></div>
                  <span style={{fontSize:11,color:"var(--md-success)",fontWeight:600}}>GATT</span>
                </div>
              </div>
            </div>
          )}
        </main>

        {/* ── M3 Bottom Navigation Bar ── */}
        <nav className="m3-bottom-nav">
          <div className="m3-nav-inner">
            {MOBILE_NAVS.map(([k,icon,label])=>(
              <button key={k} className={`m3-nav-item${page===k?" active":""}`} onClick={()=>setPage(k)}>
                <div className="nav-indicator" style={{position:"relative"}}>
                  <span className="material-icons-round">{icon}</span>
                  {k==="cart"&&cart.length>0&&(
                    <span className="nav-badge">{cartTotal}</span>
                  )}
                </div>
                <span className="nav-label">{label}</span>
              </button>
            ))}
          </div>
        </nav>
      </div>

      {/* ── POS Scanner ── */}
      {showPosScanner && (
        <BarcodeModal
          onClose={()=>setShowPosScanner(false)}
          onResult={code=>{
            const found=products.find(p=>p.barcode===code||p.sku===code);
            if(found){addToCart(found);addToast(`✅ ${found.name} ditambahkan`);setShowPosScanner(false);}
            else{setSearchQ(code);setShowPosScanner(false);addToast("Barcode tidak ditemukan","error");}
          }}
        />
      )}

      {/* ── Payment Sheet ── */}
      {showPayment && (
        <div className="m3-overlay" onClick={e=>e.target===e.currentTarget&&setShowPayment(false)}>
          <div className="m3-sheet">
            <div className="sheet-handle" />
            <div className="sheet-header">
              <span className="sheet-title">💳 Pembayaran</span>
              <button className="sheet-close" onClick={()=>setShowPayment(false)}>✕</button>
            </div>
            <div className="sheet-body">
              <div className="pay-summary-card">
                <div className="pay-row"><span>Subtotal</span><span className="pay-val">{formatRp(subtotal)}</span></div>
                {discountAmt>0&&<div className="pay-row"><span>Diskon</span><span className="pay-val" style={{color:"var(--md-success)"}}>-{formatRp(discountAmt)}</span></div>}
                {settings.taxEnabled&&<div className="pay-row"><span>Pajak {settings.taxRate}%</span><span className="pay-val" style={{color:"#93C5FD"}}>{formatRp(taxAmt)}</span></div>}
                {note&&<div className="pay-row"><span>📝 {note}</span></div>}
                <div className="pay-row grand"><span>TOTAL</span><span className="pay-val">{formatRp(total)}</span></div>
              </div>

              <div className="pay-method-label">Metode Pembayaran</div>
              <div className="pay-method-grid">
                {PAYMENT_METHODS.map(m=>(
                  <button key={m} className={`pay-method-btn${payMethod===m?" active":""}`} onClick={()=>setPayMethod(m)}>{m}</button>
                ))}
              </div>

              {payMethod==="cash" && (
                <>
                  <div className="pay-field-label">Uang Diterima</div>
                  <input className="pay-amount-input" type="number" inputMode="numeric" placeholder="0" value={cashPaid} onChange={e=>setCashPaid(e.target.value)} />
                  <div className="quick-amounts">
                    {quickAmounts.map(a=>(
                      <button key={a} className="quick-chip" onClick={()=>setCashPaid(String(a))}>{formatRp(a)}</button>
                    ))}
                  </div>
                  <div className={`change-card${change<0?" neg":""}`}>
                    <span className="change-lbl">{change>=0?"Kembalian":"Kurang"}</span>
                    <span className="change-val">{cashPaid?formatRp(Math.abs(change)):"-"}</span>
                  </div>
                </>
              )}
              {payMethod!=="cash" && (
                <div style={{textAlign:"center",padding:"20px 0",color:"var(--md-on-surface-variant)",fontFamily:"var(--font-brand)"}}>
                  <div style={{fontSize:48,marginBottom:8}}>{payMethod==="qris"?"📱":payMethod==="transfer"?"🏦":"💳"}</div>
                  <div style={{fontSize:14,marginBottom:6}}>Pembayaran {payMethod.toUpperCase()}</div>
                  <strong style={{color:"var(--md-primary)",fontFamily:"var(--font-mono)",fontSize:24,display:"block"}}>{formatRp(total)}</strong>
                </div>
              )}
            </div>
            <div className="sheet-footer">
              <button className="btn btn-outlined" onClick={()=>setShowPayment(false)}>Batal</button>
              <button className="btn btn-filled" onClick={handleCheckout} disabled={payMethod==="cash"&&cashPaidNum<total}>✅ Proses</button>
            </div>
          </div>
        </div>
      )}

      {/* ── Receipt Sheet ── */}
      {showReceipt && (
        <div className="m3-overlay" onClick={e=>e.target===e.currentTarget&&setShowReceipt(null)}>
          <div className="m3-sheet">
            <div className="sheet-handle" />
            <div className="sheet-header">
              <span className="sheet-title">🧾 Struk</span>
              <button className="sheet-close" onClick={()=>setShowReceipt(null)}>✕</button>
            </div>
            <div className="sheet-body">
              <div className="receipt">
                <div className="receipt-header"><h2>{settings.storeName}</h2><p>{settings.address}</p></div>
                <hr />
                <p>No: {showReceipt.id}</p><p>Tgl: {formatDate(showReceipt.date)}</p><p>Kasir: {showReceipt.cashier||settings.cashier}</p>
                <hr />
                <table><tbody>
                  {showReceipt.items.map((it,i)=>(<tr key={i}><td style={{paddingRight:6}}>{it.name}</td><td style={{textAlign:"center",minWidth:24}}>{it.qty}x</td><td style={{textAlign:"right",minWidth:56}}>{formatRp(it.price)}</td><td style={{textAlign:"right",minWidth:64}}>{formatRp(it.qty*it.price)}</td></tr>))}
                </tbody></table>
                <hr />
                <table><tbody>
                  <tr><td>Subtotal</td><td style={{textAlign:"right"}}>{formatRp(showReceipt.subtotal)}</td></tr>
                  {showReceipt.discount>0&&<tr><td>Diskon</td><td style={{textAlign:"right"}}>-{formatRp(showReceipt.discount)}</td></tr>}
                  {showReceipt.tax>0&&<tr><td>Pajak</td><td style={{textAlign:"right"}}>{formatRp(showReceipt.tax)}</td></tr>}
                  <tr className="grand"><td>TOTAL</td><td style={{textAlign:"right"}}>{formatRp(showReceipt.total)}</td></tr>
                  <tr><td>Bayar ({showReceipt.payment})</td><td style={{textAlign:"right"}}>{formatRp(showReceipt.cashPaid)}</td></tr>
                  {showReceipt.change>0&&<tr><td>Kembalian</td><td style={{textAlign:"right"}}>{formatRp(showReceipt.change)}</td></tr>}
                </tbody></table>
                {showReceipt.note&&(<><hr /><p>📝 {showReceipt.note}</p></>)}
                <hr />
                <div className="receipt-footer"><p>Terima kasih!</p><p style={{marginTop:4,fontSize:9}}>Powered by Kasir-Pro</p></div>
              </div>
            </div>
            <div className="sheet-footer">
              {!window.isSecureContext && (
                <div style={{
                  padding:"8px 10px",borderRadius:8,fontSize:11,
                  background:"rgba(252,211,77,.12)",border:"1px solid rgba(252,211,77,.3)",
                  color:"#FCD34D",lineHeight:1.5,width:"100%",marginBottom:4,
                }}>
                  ⚠️ <strong>Buka via Vercel (HTTPS)</strong> agar Bluetooth aktif.<br/>
                  Sekarang: HTTP lokal = Bluetooth diblokir browser.
                </div>
              )}
              <button className="btn btn-outlined" onClick={()=>setShowReceipt(null)}>Tutup</button>
              {/* Status printer */}
              {printerStatus && (
                <div style={{
                  padding:"8px 12px",borderRadius:8,marginBottom:4,fontSize:12,
                  textAlign:"center",fontFamily:"var(--font-brand)",
                  background:printerStatus.startsWith("✅")?"rgba(110,231,183,.12)":
                    printerStatus.startsWith("❌")?"rgba(255,180,171,.12)":"rgba(194,231,255,.08)",
                  color:printerStatus.startsWith("✅")?"var(--md-success)":
                    printerStatus.startsWith("❌")?"var(--md-error)":"var(--md-primary)",
                  border:"1px solid",
                  borderColor:printerStatus.startsWith("✅")?"rgba(110,231,183,.3)":
                    printerStatus.startsWith("❌")?"rgba(255,180,171,.3)":"rgba(194,231,255,.2)",
                }}>
                  {printerStatus}
                </div>
              )}

              <div style={{display:"flex",flexDirection:"column",gap:8,width:"100%"}}>
                {/* Tombol utama: Web Bluetooth GATT */}
                <button
                  className="btn btn-filled"
                  disabled={isPrinting}
                  onClick={async()=>{
                    setIsPrinting(true);
                    setPrinterStatus("");
                    try {
                      // Gunakan printWebBluetooth — adaptasi dari modulprinter.db
                      await printWebBluetooth(showReceipt, settings, setPrinterStatus);
                      addToast("✅ Struk berhasil dicetak!");
                    } catch(e) {
                      setPrinterStatus("❌ " + e.message);
                      addToast("Gagal: " + e.message, "error");
                    } finally {
                      setIsPrinting(false);
                    }
                  }}
                  style={{display:"flex",alignItems:"center",justifyContent:"center",gap:8,fontSize:14,padding:"14px"}}
                >
                  <span className="material-icons-round" style={{fontSize:20}}>
                    {isPrinting ? "hourglass_top" : "print"}
                  </span>
                  {isPrinting ? "Mencetak..." : "🖨️ Cetak Bluetooth (RPP02)"}
                </button>

                {/* Fallback: browser */}
                <button
                  className="btn btn-outlined"
                  onClick={()=>{
                    const ok = openPrintWindow(showReceipt, settings);
                    if(!ok){
                      const html = buildReceiptHTML(showReceipt, settings);
                      const blob = new Blob([html],{type:"text/html;charset=utf-8"});
                      const url = URL.createObjectURL(blob);
                      const a = document.createElement("a");
                      a.href=url; a.download=`struk-${showReceipt.id}.html`;
                      document.body.appendChild(a); a.click();
                      document.body.removeChild(a); URL.revokeObjectURL(url);
                      addToast("File struk diunduh","info");
                    } else {
                      addToast("Halaman cetak dibuka di tab baru");
                    }
                  }}
                  style={{display:"flex",alignItems:"center",justifyContent:"center",gap:6,fontSize:12}}
                >
                  <span className="material-icons-round" style={{fontSize:16}}>open_in_new</span>
                  Cetak via Browser (fallback)
                </button>
              </div>
            </div>
          </div>
        </div>
      )}

      {/* ── Product Form ── */}
      {showProductForm!==null && (
        <ProductFormModal
          product={showProductForm}
          onClose={()=>setShowProductForm(null)}
          onSave={p=>{
            if(p.id){setProducts(ps=>ps.map(x=>x.id===p.id?p:x));addToast("Produk diperbarui!");}
            else{setProducts(ps=>[...ps,{...p,id:Date.now()}]);addToast("Produk ditambahkan!");}
            setShowProductForm(null);
          }}
        />
      )}

      {/* ── Pembelian Form ── */}
      {showFormPembelian && (
        <div className="m3-overlay">
          <div className="m3-sheet">
            <div className="sheet-handle" />
            <div className="sheet-header">
              <span className="sheet-title">{editPembelianId?"✏️ Edit Pembelian":"➕ Catatan Pembelian"}</span>
              <button className="sheet-close" onClick={()=>{setShowFormPembelian(false);setEditPembelianId(null);}}>✕</button>
            </div>
            <div className="sheet-body">
              <div className="form-group">
                <label className="form-label">Barcode (Opsional)</label>
                <div style={{display:"flex",gap:6}}>
                  <input className="m3-form-input" value={formPembelian.barcode} onChange={e=>setFormPembelian(f=>({...f,barcode:e.target.value}))} placeholder="8991234..." style={{flex:1}} />
                </div>
              </div>
              
              <div className="form-group">
                <label className="form-label">Nama Produk *</label>
                <input className="m3-form-input" value={formPembelian.produk} onChange={e=>setFormPembelian(f=>({...f,produk:e.target.value}))} placeholder="Nama produk" />
                {formPembelian.produk && (
                  <div style={{fontSize:11,color:"var(--md-on-surface-variant)",marginTop:4,maxHeight:100,overflowY:"auto"}}>
                    {products.filter(p=>p.name.toLowerCase().includes(formPembelian.produk.toLowerCase())).map(p=>(
                      <div key={p.id} style={{padding:"4px 8px",background:"var(--md-surface-2)",borderRadius:6,marginTop:2,cursor:"pointer"}} onClick={()=>setFormPembelian(f=>({...f,produk:p.name}))}>
                        {p.name} (Stok: {p.stock})
                      </div>
                    ))}
                  </div>
                )}
              </div>

              <div className="form-row-2">
                <div className="form-group">
                  <label className="form-label">Satuan *</label>
                  <input className="m3-form-input" value={formPembelian.satuan} onChange={e=>setFormPembelian(f=>({...f,satuan:e.target.value}))} placeholder="pcs, box, dus" />
                </div>
                <div className="form-group">
                  <label className="form-label">Harga Satuan (Rp) *</label>
                  <input className="m3-form-input" type="number" inputMode="numeric" value={formPembelian.hargaSatuan} onChange={e=>setFormPembelian(f=>({...f,hargaSatuan:e.target.value}))} placeholder="0" />
                </div>
              </div>

              <div className="form-group">
                <label className="form-label">Jumlah Barang *</label>
                <input className="m3-form-input" type="number" inputMode="numeric" value={formPembelian.jumlahBarang} onChange={e=>setFormPembelian(f=>({...f,jumlahBarang:e.target.value}))} placeholder="0" />
              </div>

              <div className="form-group">
                <label className="form-label">Jumlah (Rp)</label>
                <div style={{padding:"12px",background:"var(--md-surface-2)",borderRadius:10,fontSize:14,fontWeight:700,color:"var(--md-primary)",fontFamily:"var(--font-mono)"}}>
                  {formPembelian.hargaSatuan && formPembelian.jumlahBarang ? formatRp(parseFloat(formPembelian.hargaSatuan) * parseInt(formPembelian.jumlahBarang)) : "Rp 0"}
                </div>
              </div>

              <div className="form-group">
                <label className="form-label">Supplier</label>
                <input className="m3-form-input" value={formPembelian.supplier} onChange={e=>setFormPembelian(f=>({...f,supplier:e.target.value}))} placeholder="Nama supplier (boleh kosong)" />
              </div>

              <div className="form-group">
                <label className="form-label">Tanggal</label>
                <input className="m3-form-input" type="date" value={formPembelian.tanggal} onChange={e=>setFormPembelian(f=>({...f,tanggal:e.target.value}))} />
              </div>

              <div className="form-group">
                <label className="form-label">Keterangan</label>
                <textarea className="m3-form-input" value={formPembelian.keterangan} onChange={e=>setFormPembelian(f=>({...f,keterangan:e.target.value}))} placeholder="Catatan tambahan..." style={{minHeight:60,resize:"vertical"}} />
              </div>
            </div>
            <div className="sheet-footer">
              <button className="btn btn-outlined" onClick={()=>{setShowFormPembelian(false);setEditPembelianId(null);}}>Batal</button>
              <button className="btn btn-filled" onClick={handleTambahPembelian}>💾 {editPembelianId?"Simpan":"Tambah"}</button>
            </div>
          </div>
        </div>
      )}

      {/* ── Penjualan Form ── */}
      {showFormPenjualan && (
        <div className="m3-overlay">
          <div className="m3-sheet">
            <div className="sheet-handle" />
            <div className="sheet-header">
              <span className="sheet-title">{editPenjualanId?"✏️ Edit Penjualan":"➕ Catatan Penjualan"}</span>
              <button className="sheet-close" onClick={()=>{setShowFormPenjualan(false);setEditPenjualanId(null);}}>✕</button>
            </div>
            <div className="sheet-body">
              <div className="form-group">
                <label className="form-label">Nama Produk *</label>
                <input className="m3-form-input" value={formPenjualan.produk} onChange={e=>setFormPenjualan(f=>({...f,produk:e.target.value}))} placeholder="Nama produk" />
                {formPenjualan.produk && (
                  <div style={{fontSize:11,color:"var(--md-on-surface-variant)",marginTop:4,maxHeight:100,overflowY:"auto"}}>
                    {products.filter(p=>p.name.toLowerCase().includes(formPenjualan.produk.toLowerCase())).map(p=>(
                      <div key={p.id} style={{padding:"4px 8px",background:"var(--md-surface-2)",borderRadius:6,marginTop:2,cursor:"pointer"}} onClick={()=>setFormPenjualan(f=>({...f,produk:p.name}))}>
                        {p.name} (Stok: {p.stock})
                      </div>
                    ))}
                  </div>
                )}
              </div>

              <div className="form-row-2">
                <div className="form-group">
                  <label className="form-label">Satuan *</label>
                  <input className="m3-form-input" value={formPenjualan.satuan} onChange={e=>setFormPenjualan(f=>({...f,satuan:e.target.value}))} placeholder="pcs, box, dus" />
                </div>
                <div className="form-group">
                  <label className="form-label">Harga Satuan (Rp) *</label>
                  <input className="m3-form-input" type="number" inputMode="numeric" value={formPenjualan.hargaSatuan} onChange={e=>setFormPenjualan(f=>({...f,hargaSatuan:e.target.value}))} placeholder="0" />
                </div>
              </div>

              <div className="form-group">
                <label className="form-label">Jumlah Barang *</label>
                <input className="m3-form-input" type="number" inputMode="numeric" value={formPenjualan.jumlahBarang} onChange={e=>setFormPenjualan(f=>({...f,jumlahBarang:e.target.value}))} placeholder="0" />
              </div>

              <div className="form-group">
                <label className="form-label">Jumlah (Rp)</label>
                <div style={{padding:"12px",background:"var(--md-surface-2)",borderRadius:10,fontSize:14,fontWeight:700,color:"var(--md-success)",fontFamily:"var(--font-mono)"}}>
                  {formPenjualan.hargaSatuan && formPenjualan.jumlahBarang ? formatRp(parseFloat(formPenjualan.hargaSatuan) * parseInt(formPenjualan.jumlahBarang)) : "Rp 0"}
                </div>
              </div>

              <div className="form-group">
                <label className="form-label">Tanggal</label>
                <input className="m3-form-input" type="date" value={formPenjualan.tanggal} onChange={e=>setFormPenjualan(f=>({...f,tanggal:e.target.value}))} />
              </div>

              <div className="form-group">
                <label className="form-label">Keterangan</label>
                <textarea className="m3-form-input" value={formPenjualan.keterangan} onChange={e=>setFormPenjualan(f=>({...f,keterangan:e.target.value}))} placeholder="Catatan tambahan..." style={{minHeight:60,resize:"vertical"}} />
              </div>
            </div>
            <div className="sheet-footer">
              <button className="btn btn-outlined" onClick={()=>{setShowFormPenjualan(false);setEditPenjualanId(null);}}>Batal</button>
              <button className="btn btn-filled" onClick={handleTambahPenjualan}>💾 {editPenjualanId?"Simpan":"Tambah"}</button>
            </div>
          </div>
        </div>
      )}

      {/* ── Toasts ── */}
      <div className="toast-wrap">
        {toasts.map(t=>(
          <div key={t.id} className={`m3-toast toast-${t.type}`}>
            {t.type==="success"?"✅":t.type==="error"?"❌":"ℹ️"} {t.msg}
          </div>
        ))}
      </div>
    </>
  );
}
