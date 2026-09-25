import React, { useState, useEffect, useRef, useMemo, useDeferredValue } from 'react';
import * as XLSX from 'xlsx';
import jsPDF from 'jspdf';
import 'jspdf-autotable';
import BarcodeModal from './BarcodeModal.jsx';

const ipcRenderer = typeof window !== 'undefined' && window.require ? window.require('electron').ipcRenderer : null;

const safeStorageGet = (key, fallback = null) => {
  try {
    if (typeof window === 'undefined' || !window.localStorage) return fallback;
    const value = window.localStorage.getItem(key);
    return value === null ? fallback : value;
  } catch (_) {
    return fallback;
  }
};

const safeStorageJson = (key, fallback) => {
  try {
    const raw = safeStorageGet(key, null);
    if (raw === null || raw === '') return fallback;
    const parsed = JSON.parse(raw);
    return parsed ?? fallback;
  } catch (_) {
    return fallback;
  }
};

const safeStorageSet = (key, value) => {
  try {
    if (typeof window !== 'undefined' && window.localStorage) {
      window.localStorage.setItem(key, value);
    }
  } catch (_) {}
};

class AppErrorBoundary extends React.Component {
  constructor(props) {
    super(props);
    this.state = { hasError: false, message: '' };
  }

  static getDerivedStateFromError(error) {
    return {
      hasError: true,
      message: error?.message || 'Unexpected application error'
    };
  }

  componentDidCatch(error) {
    // Keep the UI from becoming a completely blank page.
    try { console.error('POS application error:', error); } catch (_) {}
  }

  handleReload = () => {
    try { window.location.reload(); } catch (_) {}
  };

  handleResetStorage = () => {
    try {
      const keys = Object.keys(window.localStorage || {});
      keys.filter(k => k.startsWith('pos_')).forEach(k => window.localStorage.removeItem(k));
    } catch (_) {}
    try { window.location.reload(); } catch (_) {}
  };

  render() {
    if (!this.state.hasError) return this.props.children;

    return (
      <div style={{
        minHeight: '100vh',
        display: 'flex',
        alignItems: 'center',
        justifyContent: 'center',
        padding: '24px',
        background: '#f8fafc',
        fontFamily: 'Arial, sans-serif',
        boxSizing: 'border-box'
      }}>
        <div style={{
          width: '100%',
          maxWidth: '560px',
          background: '#fff',
          border: '1px solid #e2e8f0',
          borderRadius: '14px',
          padding: '28px',
          boxShadow: '0 12px 35px rgba(15,23,42,.10)',
          textAlign: 'center'
        }}>
          <div style={{ fontSize: '42px', marginBottom: '10px' }}>⚠️</div>
          <h2 style={{ margin: '0 0 10px', color: '#0f172a' }}>POS App එක load වෙන්න බැරි වුණා</h2>
          <p style={{ color: '#475569', lineHeight: 1.6, margin: '0 0 8px' }}>
            App එකේ runtime error එකක් ඇති. මුලින් Reload කරලා බලන්න.
          </p>
          <p style={{ color: '#94a3b8', fontSize: '12px', wordBreak: 'break-word', margin: '0 0 20px' }}>
            {this.state.message}
          </p>
          <div style={{ display: 'flex', gap: '10px', justifyContent: 'center', flexWrap: 'wrap' }}>
            <button onClick={this.handleReload} style={{
              padding: '10px 18px', border: 0, borderRadius: '8px',
              background: '#2563eb', color: '#fff', fontWeight: 700, cursor: 'pointer'
            }}>Reload</button>
            <button onClick={this.handleResetStorage} style={{
              padding: '10px 18px', border: '1px solid #cbd5e1', borderRadius: '8px',
              background: '#fff', color: '#334155', fontWeight: 700, cursor: 'pointer'
            }}>Reset POS Saved Data</button>
          </div>
        </div>
      </div>
    );
  }
}

// Translations Dictionary (Sinhala, English, Tamil)
const translations = {
  en: {
    loginTitle: "Poshitha Pos System",
    shopRegistration: "Shop Registration",
    username: "Username",
    password: "Password",
    loginBtn: "Login",
    registerBtn: "Register Shop",
    shopName: "Shop Name",
    shopMail: "Shop Email",
    address: "Address",
    phone: "Phone Number",
    secretKey: "Secret Key",
    alreadyHaveAccount: "Already registered? Login here",
    needToRegister: "Don't have an account? Register Shop",
    errorInvalidLogin: "Invalid Username or Password!",
    errorSecretKey: "Invalid Secret Key! Registration failed.",
    successReg: "Shop Registration Successful! Please login now.",
    billing: "Billing",
    inventory: "Inventory",
    suppliers: "Suppliers",
    customers: "Customer Dashboard",
    creditBook: "Credit Book",
    reports: "Reports & Dashboard",
    settings: "Shop Profile & Settings",
    logout: "Logout",
  },
  si: {
    loginTitle: "පොස් සිස්ටම් පිවිසුම",
    shopRegistration: "කඩය ලියාපදිංචි කිරීම",
    username: "පරිශීලක නාමය",
    password: "මුරපදය",
    loginBtn: "ඇතුළු වන්න",
    registerBtn: "කඩය ලියාපදිංචි කරන්න",
    shopName: "කඩයේ නම",
    shopMail: "විද්‍යුත් තැපෑල",
    address: "ලිපිනය",
    phone: "දුරකථන අංකය",
    secretKey: "රහස් යතුර",
    alreadyHaveAccount: " දැනටමත් ලියාපදිංචි වී ඇද්ද? පිවිසෙන්න",
    needToRegister: "ගිණුමක් නැද්ද? කඩය ලියාපදිංචි කරන්න",
    errorInvalidLogin: "වැරදි Username එකක් හෝ Password එකක්!",
    errorSecretKey: "වැරදි Secret Key එකකි! ලියාපදිංචිය අසාර්ථකයි.",
    successReg: "කඩය ලියාපදිංචිය සාර්ථකයි! දැන් Login වන්න.",
    billing: "බිල්පත්",
    inventory: "භාණ්ඩ තොග",
    suppliers: "සැපයුම්කරුවන්",
    customers: "පාරිභෝගික විස්තර",
    creditBook: "ක්‍රෙඩිට් බුක්",
    reports: "වාර්තා විශ්ලේෂණ",
    settings: "සැකසුම්",
    logout: "පිටවීම",
  },
  ta: {
    loginTitle: "POS அமைப்பு உள்நுழைவு",
    shopRegistration: "கடை பதிவு",
    username: "பயனர் பெயர்",
    password: "கடவுச்சொல்",
    loginBtn: "உள்நுழை (Login)",
    registerBtn: "கடையைப் பதிவு செய்",
    shopName: "கடை பெயர்",
    shopMail: "மின்னஞ்சல்",
    address: "முகவரி",
    phone: "தொலைபேசி எண்",
    secretKey: "ரகசிய விசை",
    alreadyHaveAccount: "ஏற்கனவே பதிவு செய்துள்ளீர்களா? உள்நுழைக",
    needToRegister: "கணக்கு இல்லையா? கடையைப் பதிவு செய்",
    errorInvalidLogin: "தவறான Username அல்லது Password!",
    errorSecretKey: "தவறான Secret Key! பதிவு தோல்வியடைந்தது.",
    successReg: "கடை பதிவு வெற்றிகரமாக முடிந்தது! இப்போது உள்நுழையவும்.",
    billing: "பில்லிங்",
    inventory: "சரக்கு (Inventory)",
    suppliers: "வழங்குநர்கள்",
    customers: " வாடிக்கையாளர் பக்கம்",
    creditBook: "கிரெடிட் புக்",
    reports: "அறிக்கைகள்",
    settings: "அமைப்புகள்",
    logout: "வெளியேறு",
  }
};


// CODE 128-B barcode helpers used by the inventory label printer.
const CODE128_PATTERNS = [
  '212222','222122','222221','121223','121322','131222','122213','122312','132212','221213',
  '221312','231212','112232','122132','122231','113222','123122','123221','223211','221132',
  '221231','213212','223112','312131','311222','321122','321221','312212','322112','322211',
  '212123','212321','232121','111323','131123','131321','112313','132113','132311','211313',
  '231113','231311','112133','112331','132131','113123','113321','133121','313121','211331',
  '231131','213113','213311','213131','311123','311321','331121','312113','312311','332111',
  '314111','221411','431111','111224','111422','121124','121421','141122','141221','112214',
  '112412','122114','122411','142112','142211','241211','221114','413111','241112','134111',
  '111242','121142','121241','114212','124112','124211','411212','421112','421211','212141',
  '214121','412121','111143','111341','131141','114113','114311','411113','411311','113141',
  '114131','311141','411131','211412','211214','211232','2331112'
];
const code128BValue = ch => {
  const n = ch.charCodeAt(0);
  return n >= 32 && n <= 126 ? n - 32 : null;
};
const buildCode128BPattern = value => {
  const text = String(value ?? '');
  if (!text) return '';
  const values = [];
  for (const ch of text) {
    const v = code128BValue(ch);
    if (v === null) return '';
    values.push(v);
  }
  let checksum = 104;
  values.forEach((v, i) => { checksum += v * (i + 1); });
  checksum %= 103;
  return [104, ...values, checksum, 106].map(v => CODE128_PATTERNS[v]).join('');
};
const barcodeSvgMarkup = (value, widthPx = 250, heightPx = 54) => {
  const pattern = buildCode128BPattern(value);
  if (!pattern) return '';
  const moduleWidth = widthPx / pattern.split('').reduce((sum, n) => sum + Number(n), 0);
  let x = 0;
  let bars = '';
  for (let i = 0; i < pattern.length; i++) {
    const w = Number(pattern[i]) * moduleWidth;
    if (i % 2 === 0) bars += `<rect x="${x.toFixed(3)}" y="0" width="${w.toFixed(3)}" height="${heightPx}" fill="#000"/>`;
    x += w;
  }
  return `<svg xmlns="http://www.w3.org/2000/svg" width="${widthPx}" height="${heightPx}" viewBox="0 0 ${widthPx} ${heightPx}" preserveAspectRatio="none">${bars}</svg>`;
};

function App() {
  // HIDDEN SECRET KEY (Obfuscated securely to prevent direct source code visibility)
  const MASTER_SECRET_KEY = [84, 108, 115, 49, 52, 53, 52, 55].map(code => String.fromCharCode(code)).join('');

  // LANGUAGE STATE
  const [lang, setLang] = useState('si');
  const t = translations[lang];
  const ui = {
    en: {
      returns:'Returns', searchBill:'Search Bill', clear:'Clear', processReturn:'Process Selected Return', recentReturns:'Recent Returns', deleteAll:'Delete All',
      shopOrderHistory:'Shop Order History', customerDashboard:'Customer Dashboard & Purchase History', creditBook:'Credit Book',
      reportsDashboard:'Reports & Sales Analytics Dashboard', dashboard:'Dashboard', monthlyReport:'Monthly Report', monthlySalesReport:'Monthly Sales Report',
      shopProfile:'Shop Profile & Settings', shopProfileDetails:'Shop Profile Details', receiptSettings:'Receipt Settings', receiptLogoSize:'Receipt Logo Size:',
      billing:'Billing', inventory:'Inventory', suppliers:'Suppliers', barcode:'Barcode', offers:'Offers',
      scanBarcode:'Scan Barcode or Select Product', itemsInBill:'Items in Bill', itemName:'Item Name', price:'Price', qty:'Qty', offer:'Offer', total:'Total', action:'Action',
      paymentCustomer:'Payment & Customer Details', customerInfo:'Customer Info (Optional)', paymentMethod:'Payment Method', cash:'Cash', card:'Card', online:'Online', credit:'Credit',
      remove:'Remove', noItems:'No items scanned yet. Scan a barcode or select above!', heldBills:'Held Bills',
      searchInventory:'Search Inventory', inventoryTable:'Stock Inventory Table', importExcel:'Import Excel', exportExcel:'Export Excel',
      addProduct:'Add New Product', editProduct:'Edit Product', saveProduct:'Save Product', updateProduct:'Update Product',
      supplierDashboard:'Supplier Dashboard & Order Management', addSupplier:'Add New Supplier', editSupplier:'Edit Supplier', saveSupplier:'Save Supplier', updateSupplier:'Update Supplier',
      orderQty:'Order Qty', supplier:'Supplier', generateOrder:'Generate Supplier Order', saveOrder:'Save Order', pending:'Pending', ongoing:'Ongoing', complete:'Complete',
      registerCustomer:'Register Customer', receivePayment:'Receive Credit Payment', customersExcel:'Customers Excel', registeredCreditCustomers:'Registered Credit Customers',
      payment:'Payment', outstanding:'Outstanding', creditLimit:'Credit Limit', actions:'Actions',
      exportSalesExcel:'Export Sales Excel', exportSalesPdf:'Export Sales PDF', clearAllSales:'Clear All Sales',
      todaySalesProfit:'Today Sales & Profit', monthSalesProfit:'This Month Sales & Profit', yearSalesProfit:'This Year Sales & Profit',
      monthlySalesProfitAnalysis:'Monthly Sales & Profit Analysis', fillRefresh:'Fill / Refresh', downloadExcel:'Download Excel',
      fullYearChart:'Full-Year Monthly Sales & Profit Bar Chart', month:'Month', sales:'Sales (Rs.)', profit:'Profit (Rs.)', profitMargin:'Profit Margin', yearTotal:'YEAR TOTAL',
      productPerformance:'Product Performance Charts', bestSellingChart:'Top 20 Best Selling - Bar Chart', slowMovingChart:'Top 20 Slow Moving - Bar Chart',
      bestSelling:'Top 20 Best Selling Products', slowMoving:'Top 20 Slow Moving Products', product:'Product', sold:'Sold', neverSold:'Never Sold Products',
      lowStock:'Low Stock Alert Report', downloadLowStock:'Download Low Stock Excel', productName:'Product Name', remainingStock:'Remaining Stock', alertLimit:'Alert Limit', statusAction:'Status / Action',
      databaseBackup:'Database Backup & Restore', terminal:'Terminal / Register Name:', footer:'Receipt Footer Message:', paperWidth:'Paper Width:', fontSize:'Font Size:', receiptFont:'Receipt Font:', receiptDesign:'Receipt Design:',
      cashierAccounts:'Cashier Accounts Management', selectCashier:'Select Cashier to Edit:', username:'Username:', password:'Password:',
      barcodeSize:'Barcode Size', barcodeQuantity:'Barcode Quantity', previewLabels:'Preview Labels', barcodePreview:'Barcode Label Preview', close:'Close', back:'Back', printLabels:'Print Labels',
      saveCustomer:'Save Customer', savePrintReceipt:'Save & Print Receipt', cancel:'Cancel', search:'Search',
      invoiceNo:'Invoice #', dateTime:'Date & Time', customerName:'Customer Name', phoneNumber:'Phone Number', totalAmount:'Total Amount',
    },
    si: {
      returns:'ආපසු ගැනීම්', searchBill:'බිල්පත සොයන්න', clear:'හිස් කරන්න', processReturn:'තෝරාගත් ආපසු ගැනීම ක්‍රියාත්මක කරන්න', recentReturns:'මෑත ආපසු ගැනීම්', deleteAll:'සියල්ල මකන්න',
      shopOrderHistory:'කඩ ඇණවුම් ඉතිහාසය', customerDashboard:'පාරිභෝගික Dashboard සහ මිලදී ගැනීම් ඉතිහාසය', creditBook:'ක්‍රෙඩිට් බුක්',
      reportsDashboard:'වාර්තා සහ විකුණුම් විශ්ලේෂණ Dashboard', dashboard:'Dashboard', monthlyReport:'මාසික වාර්තාව', monthlySalesReport:'මාසික විකුණුම් වාර්තාව',
      shopProfile:'කඩ තොරතුරු සහ සැකසුම්', shopProfileDetails:'කඩ තොරතුරු', receiptSettings:'බිල්පත් සැකසුම්', receiptLogoSize:'බිල්පත් Logo ප්‍රමාණය:',
      billing:'බිල්පත්', inventory:'භාණ්ඩ තොගය', suppliers:'සැපයුම්කරුවන්', barcode:'බාර්කෝඩ්', offers:'වට්ටම් / Offers',
      scanBarcode:'බාර්කෝඩ් Scan කරන්න හෝ භාණ්ඩය තෝරන්න', itemsInBill:'බිලේ භාණ්ඩ', itemName:'භාණ්ඩ නම', price:'මිල', qty:'ප්‍රමාණය', offer:'වට්ටම', total:'මුළු එකතුව', action:'ක්‍රියා',
      paymentCustomer:'ගෙවීම් සහ පාරිභෝගික විස්තර', customerInfo:'පාරිභෝගික විස්තර (විකල්ප)', paymentMethod:'ගෙවීම් ක්‍රමය', cash:'මුදල්', card:'කාඩ්', online:'ඔන්ලයින්', credit:'ණයට',
      remove:'ඉවත් කරන්න', noItems:'තවම භාණ්ඩ Scan කර නැත. බාර්කෝඩ් එකක් Scan කරන්න හෝ ඉහළින් තෝරන්න!', heldBills:'රඳවා ඇති බිල්',
      searchInventory:'භාණ්ඩ තොගය සොයන්න', inventoryTable:'භාණ්ඩ තොග වගුව', importExcel:'Excel Import', exportExcel:'Excel Export',
      addProduct:'නව භාණ්ඩයක් එක් කරන්න', editProduct:'භාණ්ඩය සංස්කරණය කරන්න', saveProduct:'භාණ්ඩය සුරකින්න', updateProduct:'භාණ්ඩය යාවත්කාලීන කරන්න',
      supplierDashboard:'සැපයුම්කරු Dashboard සහ ඇණවුම් කළමනාකරණය', addSupplier:'නව සැපයුම්කරුවෙකු එක් කරන්න', editSupplier:'සැපයුම්කරු සංස්කරණය', saveSupplier:'සැපයුම්කරු සුරකින්න', updateSupplier:'සැපයුම්කරු යාවත්කාලීන කරන්න',
      orderQty:'ඇණවුම් ප්‍රමාණය', supplier:'සැපයුම්කරු', generateOrder:'සැපයුම්කරු ඇණවුම සාදන්න', saveOrder:'ඇණවුම සුරකින්න', pending:'බලාපොරොත්තුවෙන්', ongoing:'ක්‍රියාත්මකයි', complete:'සම්පූර්ණයි',
      registerCustomer:'පාරිභෝගිකයා ලියාපදිංචි කරන්න', receivePayment:'ක්‍රෙඩිට් ගෙවීම ලබාගන්න', customersExcel:'පාරිභෝගික Excel', registeredCreditCustomers:'ලියාපදිංචි ක්‍රෙඩිට් පාරිභෝගිකයින්',
      payment:'ගෙවීම', outstanding:'ගෙවිය යුතු මුදල', creditLimit:'ක්‍රෙඩිට් සීමාව', actions:'ක්‍රියා',
      exportSalesExcel:'විකුණුම් Excel Export', exportSalesPdf:'විකුණුම් PDF Export', clearAllSales:'සියලු විකුණුම් මකන්න',
      todaySalesProfit:'අද විකුණුම් සහ ලාභය', monthSalesProfit:'මෙම මාසයේ විකුණුම් සහ ලාභය', yearSalesProfit:'මෙම වසරේ විකුණුම් සහ ලාභය',
      monthlySalesProfitAnalysis:'මාසික විකුණුම් සහ ලාභ විශ්ලේෂණය', fillRefresh:'පුරවන්න / යාවත්කාලීන කරන්න', downloadExcel:'Excel බාගන්න',
      fullYearChart:'වසරේ මාසික විකුණුම් සහ ලාභ Bar Chart', month:'මාසය', sales:'විකුණුම් (රු.)', profit:'ලාභය (රු.)', profitMargin:'ලාභ ප්‍රතිශතය', yearTotal:'වසරේ එකතුව',
      productPerformance:'භාණ්ඩ කාර්යසාධන Charts', bestSellingChart:'හොඳම විකිණෙන Top 20 - Bar Chart', slowMovingChart:'මන්දගාමී Top 20 - Bar Chart',
      bestSelling:'හොඳම විකිණෙන Top 20 භාණ්ඩ', slowMoving:'මන්දගාමී Top 20 භාණ්ඩ', product:'භාණ්ඩය', sold:'විකුණූ ප්‍රමාණය', neverSold:'තවම විකිණී නැති භාණ්ඩ',
      lowStock:'අඩු තොග අනතුරු ඇඟවීමේ වාර්තාව', downloadLowStock:'අඩු තොග Excel බාගන්න', productName:'භාණ්ඩ නම', remainingStock:'ඉතිරි තොගය', alertLimit:'අනතුරු සීමාව', statusAction:'තත්ත්වය / ක්‍රියා',
      databaseBackup:'දත්ත Backup සහ Restore', terminal:'Terminal / Register නම:', footer:'බිල්පත් පහළ පණිවිඩය:', paperWidth:'කඩදාසි පළල:', fontSize:'අකුරු ප්‍රමාණය:', receiptFont:'බිල්පත් අකුරු:', receiptDesign:'බිල්පත් Design:',
      cashierAccounts:'Cashier ගිණුම් කළමනාකරණය', selectCashier:'සංස්කරණය කිරීමට Cashier තෝරන්න:', username:'පරිශීලක නාමය:', password:'මුරපදය:',
      barcodeSize:'බාර්කෝඩ් ප්‍රමාණය', barcodeQuantity:'බාර්කෝඩ් ප්‍රමාණය / ගණන', previewLabels:'Labels Preview', barcodePreview:'බාර්කෝඩ් Label Preview', close:'වසන්න', back:'ආපසු', printLabels:'Labels මුද්‍රණය',
      saveCustomer:'පාරිභෝගිකයා සුරකින්න', savePrintReceipt:'සුරකින්න සහ බිල්පත මුද්‍රණය කරන්න', cancel:'අවලංගු කරන්න', search:'සොයන්න',
      invoiceNo:'බිල් අංකය', dateTime:'දිනය සහ වේලාව', customerName:'පාරිභෝගික නම', phoneNumber:'දුරකථන අංකය', totalAmount:'මුළු මුදල',
    },
    ta: {
      returns:'திருப்பிகள்', searchBill:'பில் தேடுக', clear:'அழி', processReturn:'தேர்ந்தெடுத்த திருப்பியை செயல்படுத்து', recentReturns:'சமீபத்திய திருப்பிகள்', deleteAll:'அனைத்தையும் நீக்கு',
      shopOrderHistory:'கடை ஆர்டர் வரலாறு', customerDashboard:'வாடிக்கையாளர் Dashboard மற்றும் கொள்முதல் வரலாறு', creditBook:'கிரெடிட் புக்',
      reportsDashboard:'அறிக்கைகள் மற்றும் விற்பனை பகுப்பாய்வு Dashboard', dashboard:'Dashboard', monthlyReport:'மாதாந்திர அறிக்கை', monthlySalesReport:'மாதாந்திர விற்பனை அறிக்கை',
      shopProfile:'கடை தகவல் மற்றும் அமைப்புகள்', shopProfileDetails:'கடை தகவல்கள்', receiptSettings:'ரசீது அமைப்புகள்', receiptLogoSize:'ரசீது Logo அளவு:',
      billing:'பில்லிங்', inventory:'சரக்கு', suppliers:'வழங்குநர்கள்', barcode:'பார்கோடு', offers:'சலுகைகள் / Offers',
      scanBarcode:'பார்கோடு Scan செய்யவும் அல்லது பொருளைத் தேர்ந்தெடுக்கவும்', itemsInBill:'பில்லில் உள்ள பொருட்கள்', itemName:'பொருள் பெயர்', price:'விலை', qty:'அளவு', offer:'சலுகை', total:'மொத்தம்', action:'செயல்',
      paymentCustomer:'கட்டணம் மற்றும் வாடிக்கையாளர் விவரங்கள்', customerInfo:'வாடிக்கையாளர் விவரங்கள் (விருப்பம்)', paymentMethod:'கட்டண முறை', cash:'பணம்', card:'அட்டை', online:'ஆன்லைன்', credit:'கடன்',
      remove:'நீக்கு', noItems:'இன்னும் பொருட்கள் Scan செய்யப்படவில்லை. பார்கோடு Scan செய்யவும் அல்லது மேலே தேர்ந்தெடுக்கவும்!', heldBills:'நிறுத்தி வைக்கப்பட்ட பில்கள்',
      searchInventory:'சரக்கைத் தேடுக', inventoryTable:'சரக்கு அட்டவணை', importExcel:'Excel Import', exportExcel:'Excel Export',
      addProduct:'புதிய பொருளைச் சேர்க்கவும்', editProduct:'பொருளைத் திருத்தவும்', saveProduct:'பொருளைச் சேமிக்கவும்', updateProduct:'பொருளை புதுப்பிக்கவும்',
      supplierDashboard:'வழங்குநர் Dashboard மற்றும் ஆர்டர் மேலாண்மை', addSupplier:'புதிய வழங்குநரைச் சேர்க்கவும்', editSupplier:'வழங்குநரைத் திருத்தவும்', saveSupplier:'வழங்குநரைச் சேமிக்கவும்', updateSupplier:'வழங்குநரை புதுப்பிக்கவும்',
      orderQty:'ஆர்டர் அளவு', supplier:'வழங்குநர்', generateOrder:'வழங்குநர் ஆர்டரை உருவாக்கு', saveOrder:'ஆர்டரைச் சேமிக்கவும்', pending:'நிலுவையில்', ongoing:'நடைபெறுகிறது', complete:'முடிந்தது',
      registerCustomer:'வாடிக்கையாளரை பதிவு செய்', receivePayment:'கிரெடிட் கட்டணம் பெறுக', customersExcel:'வாடிக்கையாளர் Excel', registeredCreditCustomers:'பதிவு செய்யப்பட்ட கிரெடிட் வாடிக்கையாளர்கள்',
      payment:'கட்டணம்', outstanding:'செலுத்த வேண்டியது', creditLimit:'கிரெடிட் வரம்பு', actions:'செயல்கள்',
      exportSalesExcel:'விற்பனை Excel Export', exportSalesPdf:'விற்பனை PDF Export', clearAllSales:'அனைத்து விற்பனையையும் அழி',
      todaySalesProfit:'இன்றைய விற்பனை மற்றும் லாபம்', monthSalesProfit:'இந்த மாத விற்பனை மற்றும் லாபம்', yearSalesProfit:'இந்த ஆண்டு விற்பனை மற்றும் லாபம்',
      monthlySalesProfitAnalysis:'மாதாந்திர விற்பனை மற்றும் லாப பகுப்பாய்வு', fillRefresh:'நிரப்புக / புதுப்பிக்க', downloadExcel:'Excel பதிவிறக்கம்',
      fullYearChart:'முழு ஆண்டு மாதாந்திர விற்பனை மற்றும் லாப Bar Chart', month:'மாதம்', sales:'விற்பனை (ரூ.)', profit:'லாபம் (ரூ.)', profitMargin:'லாப விகிதம்', yearTotal:'ஆண்டு மொத்தம்',
      productPerformance:'பொருள் செயல்திறன் Charts', bestSellingChart:'சிறந்த விற்பனை Top 20 - Bar Chart', slowMovingChart:'மெதுவாக நகரும் Top 20 - Bar Chart',
      bestSelling:'சிறந்த விற்பனை Top 20 பொருட்கள்', slowMoving:'மெதுவாக நகரும் Top 20 பொருட்கள்', product:'பொருள்', sold:'விற்ற அளவு', neverSold:'இதுவரை விற்கப்படாத பொருட்கள்',
      lowStock:'குறைந்த சரக்கு எச்சரிக்கை அறிக்கை', downloadLowStock:'குறைந்த சரக்கு Excel பதிவிறக்கம்', productName:'பொருள் பெயர்', remainingStock:'மீதமுள்ள சரக்கு', alertLimit:'எச்சரிக்கை வரம்பு', statusAction:'நிலை / செயல்',
      databaseBackup:'தரவுத்தள Backup மற்றும் Restore', terminal:'Terminal / Register பெயர்:', footer:'ரசீது Footer செய்தி:', paperWidth:'காகித அகலம்:', fontSize:'எழுத்து அளவு:', receiptFont:'ரசீது எழுத்து:', receiptDesign:'ரசீது Design:',
      cashierAccounts:'Cashier கணக்கு மேலாண்மை', selectCashier:'திருத்த Cashier-ஐ தேர்ந்தெடுக்கவும்:', username:'பயனர் பெயர்:', password:'கடவுச்சொல்:',
      barcodeSize:'பார்கோடு அளவு', barcodeQuantity:'பார்கோடு எண்ணிக்கை', previewLabels:'Labels Preview', barcodePreview:'பார்கோடு Label Preview', close:'மூடு', back:'பின்', printLabels:'Labels அச்சிடு',
      saveCustomer:'வாடிக்கையாளரைச் சேமிக்கவும்', savePrintReceipt:'சேமித்து ரசீதை அச்சிடு', cancel:'ரத்து செய்', search:'தேடுக',
      invoiceNo:'பில் #', dateTime:'தேதி மற்றும் நேரம்', customerName:'வாடிக்கையாளர் பெயர்', phoneNumber:'தொலைபேசி எண்', totalAmount:'மொத்த தொகை',
    }
  }[lang];


  const uiMore = {
    en: { billNo:'Bill No', date:'Date', customer:'Customer', originalTotal:'Original Total', soldQty:'Sold Qty', available:'Available', returnQty:'Return Qty', returnReason:'Return Reason', customerRequest:'Customer Request', damagedItem:'Damaged Item', wrongItem:'Wrong Item', expiredItem:'Expired Item', other:'Other', buyPrice:'Buy Price', sellPrice:'Sell Price', stock:'Stock', unit:'Unit', expiry:'Expiry', edit:'Edit', selectSupplier:'Select Supplier', orderNo:'Order No:', company:'Company:', supplierCompany:'Supplier Company:', supplierPhone:'Supplier Phone:', currentStock:'Current Stock', cost:'Cost', totalCost:'TOTAL COST', noCustomerHistory:'No customer sales history recorded yet.', noCreditCustomers:'No registered credit customers yet.', noShopOrders:'No Shop Orders saved yet.', invoice:'Invoice #', dateTime:'Date & Time', customerName:'Customer Name', phoneNumber:'Phone Number', totalAmount:'Total Amount' },
    si: { billNo:'බිල් අංකය', date:'දිනය', customer:'පාරිභෝගිකයා', originalTotal:'මුල් මුළු එකතුව', soldQty:'විකුණූ ප්‍රමාණය', available:'ලබාගත හැකි', returnQty:'ආපසු ප්‍රමාණය', returnReason:'ආපසු ගැනීමේ හේතුව', customerRequest:'පාරිභෝගික ඉල්ලීම', damagedItem:'හානි වූ භාණ්ඩය', wrongItem:'වැරදි භාණ්ඩය', expiredItem:'කල් ඉකුත් වූ භාණ්ඩය', other:'වෙනත්', buyPrice:'මිලදී ගැනීමේ මිල', sellPrice:'විකුණුම් මිල', stock:'තොගය', unit:'ඒකකය', expiry:'කල් ඉකුත්වීම', edit:'සංස්කරණය', selectSupplier:'සැපයුම්කරු තෝරන්න', orderNo:'ඇණවුම් අංකය:', company:'සමාගම:', supplierCompany:'සැපයුම්කරුගේ සමාගම:', supplierPhone:'සැපයුම්කරුගේ දුරකථනය:', currentStock:'දැනට ඇති තොගය', cost:'පිරිවැය', totalCost:'මුළු පිරිවැය', noCustomerHistory:'තවම පාරිභෝගික විකුණුම් ඉතිහාසයක් නැත.', noCreditCustomers:'තවම ක්‍රෙඩිට් පාරිභෝගිකයින් ලියාපදිංචි කර නැත.', noShopOrders:'තවම Shop Orders save කර නැත.', invoice:'බිල් අංකය', dateTime:'දිනය සහ වේලාව', customerName:'පාරිභෝගික නම', phoneNumber:'දුරකථන අංකය', totalAmount:'මුළු මුදල' },
    ta: { billNo:'பில் எண்', date:'தேதி', customer:'வாடிக்கையாளர்', originalTotal:'அசல் மொத்தம்', soldQty:'விற்ற அளவு', available:'கிடைக்கும் அளவு', returnQty:'திருப்பும் அளவு', returnReason:'திருப்பும் காரணம்', customerRequest:'வாடிக்கையாளர் கோரிக்கை', damagedItem:'சேதமடைந்த பொருள்', wrongItem:'தவறான பொருள்', expiredItem:'காலாவதியான பொருள்', other:'மற்றவை', buyPrice:'வாங்கிய விலை', sellPrice:'விற்பனை விலை', stock:'சரக்கு', unit:'அலகு', expiry:'காலாவதி', edit:'திருத்து', selectSupplier:'வழங்குநரைத் தேர்ந்தெடுக்கவும்', orderNo:'ஆர்டர் எண்:', company:'நிறுவனம்:', supplierCompany:'வழங்குநர் நிறுவனம்:', supplierPhone:'வழங்குநர் தொலைபேசி:', currentStock:'தற்போதைய சரக்கு', cost:'செலவு', totalCost:'மொத்த செலவு', noCustomerHistory:'வாடிக்கையாளர் விற்பனை வரலாறு எதுவும் இல்லை.', noCreditCustomers:'பதிவு செய்யப்பட்ட கிரெடிட் வாடிக்கையாளர்கள் இல்லை.', noShopOrders:'சேமிக்கப்பட்ட Shop Orders இல்லை.', invoice:'பில் #', dateTime:'தேதி மற்றும் நேரம்', customerName:'வாடிக்கையாளர் பெயர்', phoneNumber:'தொலைபேசி எண்', totalAmount:'மொத்த தொகை' }
  }[lang];

  // PRODUCT UNITS - loaded from Electron when available; fallback keeps the UI working.
  const DEFAULT_PRODUCT_UNITS = [
    'Pcs','Piece','Unit','Dozen','Pair','Set','Box','Carton','Pack','Packet','Bag','Sack','Bottle','Can','Tin','Jar','Tub','Tube','Pouch','Sachet','Roll','Bundle','Bunch','Tray','Crate','Case','Bucket','Drum','Barrel','Pallet','Container','Cup','Glass','Plate','Bowl','Loaf','Slice','Stick','Block','Head','Bunch','Kg','g','mg','ton','lb','oz','L','ml','cl','dl','m³','m²','m','cm','mm','km','ft','in','yd','sq ft','sq m','cu ft','cu m','Dozen (12)','Half Dozen (6)','Gross (144)'
  ];
  const [productUnits, setProductUnits] = useState(DEFAULT_PRODUCT_UNITS);

  // AUTH & REGISTRATION STATE
  const [isRegistered, setIsRegistered] = useState(false);
  const [isLoggedIn, setIsLoggedIn] = useState(false);
  const [showWelcomeAnimation, setShowWelcomeAnimation] = useState(false);
  const [authMode, setAuthMode] = useState('login'); // 'login', 'register', or 'forgot'

  // REGISTER FORM STATE
  const [regShopName, setRegShopName] = useState('');
  const [regEmail, setRegEmail] = useState('');
  const [regAddress, setRegAddress] = useState('');
  const [regPhone, setRegPhone] = useState('');
  const [regUsername, setRegUsername] = useState('');
  const [regPassword, setRegPassword] = useState('');
  const [regSecretKey, setRegSecretKey] = useState('');
  const [regError, setRegError] = useState('');

  // LOGIN FORM STATE
  const [loginUsername, setLoginUsername] = useState('');
  const [loginPassword, setLoginPassword] = useState('');
  const [loginError, setLoginError] = useState('');
// RETURNS STATE
const [activeTab, setActiveTab] = useState('pos'); // 'pos' = POS page, 'returns' = Returns page
// Collapsible dashboard sidebar (does not change any page options or controls).
const [sidebarCollapsed, setSidebarCollapsed] = useState(false);
// Dashboard navigation style: home cards, sidebar, or both.
const [dashboardStyle, setDashboardStyle] = useState(() => safeStorageGet('pos_dashboard_style', 'home') || 'home');
// Tracks dashboard shortcuts that open a specific Inventory section.
const [inventoryShortcutMode, setInventoryShortcutMode] = useState(null);
// Shop registration logo (kept separate until registration is completed).
const [regLogoUrl, setRegLogoUrl] = useState('');
const [searchBillNo, setSearchBillNo] = useState('');
const [returnBill, setReturnBill] = useState(null);
const [returnItems, setReturnItems] = useState({});
const [returnReason, setReturnReason] = useState('Customer Request');
const [returnRecords, setReturnRecords] = useState(() => safeStorageJson('pos_return_records', []));
  // FORGOT / RESET PASSWORD STATE
  const [resetUsername, setResetUsername] = useState('');
  const [resetSecretKey, setResetSecretKey] = useState('');
  const [newAdminUsername, setNewAdminUsername] = useState('');
  const [newAdminPassword, setNewAdminPassword] = useState('');
  const [resetError, setResetError] = useState('');
  const [resetSuccess, setResetSuccess] = useState('');

  // CURRENT USER ROLE & PROFILE ('admin', 'cashier1', 'cashier2', 'cashier3')
  const [currentUserRole, setCurrentUserRole] = useState('admin');

  // CASHIER ACCOUNTS MANAGEMENT (Admin can set up 3 cashiers)
  const [cashiers, setCashiers] = useState([
    { id: 1, username: 'cashier1', password: '123' },
    { id: 2, username: 'cashier2', password: '123' },
    { id: 3, username: 'cashier3', password: '123' },
  ]);

  // SETTINGS FORM STATE FOR CASHIER MANAGEMENT & PASSWORD CHANGE INSIDE SETTINGS
  const [settingsActiveCashierId, setSettingsActiveCashierId] = useState(1);
  const [settingsCashierUsername, setSettingsCashierUsername] = useState('cashier1');
  const [settingsCashierPassword, setSettingsCashierPassword] = useState('123');

    const [products, setProducts] = useState([]);
  const [salesHistory, setSalesHistory] = useState(() => safeStorageJson('pos_sales_history', []));
  const [reportsData, setReportsData] = useState(() => ({
    sales: safeStorageJson('pos_sales_items', []),
    lowStock: []
  }));
  
  // SUPPLIERS STATE
  const [suppliers, setSuppliers] = useState([
    { id: 1, name: 'Sample Supplier Ltd', phone: '0711234567', email: 'supplier@gmail.com', address: 'Colombo 01', company: 'Global Traders' }
  ]);
  const [newSupplier, setNewSupplier] = useState({ name: '', phone: '', email: '', address: '', company: '' });
  const [editingSupplierId, setEditingSupplierId] = useState(null);

  // ORDER REQUESTS STATE (Saved locally & managed via Supplier Dashboard)
  const [orderRequests, setOrderRequests] = useState([]);
  const [shopOrderHistory, setShopOrderHistory] = useState(() => safeStorageJson('pos_shop_order_history', []));
  const [selectedLowStockItems, setSelectedLowStockItems] = useState({});
  const [reportView, setReportView] = useState('overview');
  const [reportDeletedFlags, setReportDeletedFlags] = useState({});
  const [customOrderInputs, setCustomOrderInputs] = useState({}); // To edit order qty dynamically
  const [orderSupplierInputs, setOrderSupplierInputs] = useState({});
  const [showOrderGenerateModal, setShowOrderGenerateModal] = useState(false);
  const [orderDraft, setOrderDraft] = useState(null);
  const [inventorySearch, setInventorySearch] = useState('');
  // Deferred filtering keeps the input responsive with 10,000+ products.
  const deferredInventorySearch = useDeferredValue(inventorySearch);
  const [supplierLowStockSearch, setSupplierLowStockSearch] = useState('');
  const deferredSupplierLowStockSearch = useDeferredValue(supplierLowStockSearch);
  const [monthlyAnalyticsData, setMonthlyAnalyticsData] = useState(() => safeStorageJson('pos_monthly_analytics', []));

  // CUSTOMER DASHBOARD SEARCH STATE
  const [customerSearchPhone, setCustomerSearchPhone] = useState('');
  // Historical report selectors: default to the current local date/year, while keeping all billing records intact.
  const [selectedReportDate, setSelectedReportDate] = useState(() => {
    const d = new Date();
    return `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, '0')}-${String(d.getDate()).padStart(2, '0')}`;
  });
  const [selectedReportYear, setSelectedReportYear] = useState(() => new Date().getFullYear());

  // CREDIT / CUSTOMER LEDGER
  const [registeredCustomers, setRegisteredCustomers] = useState(() => safeStorageJson('pos_registered_customers', []));
  const [creditLedger, setCreditLedger] = useState(() => safeStorageJson('pos_credit_ledger', []));
  const [selectedCreditCustomerId, setSelectedCreditCustomerId] = useState('');
  const [showCustomerRegisterModal, setShowCustomerRegisterModal] = useState(false);
  const [showCreditPaymentModal, setShowCreditPaymentModal] = useState(false);
  const [creditPaymentSearch, setCreditPaymentSearch] = useState('');
  const [creditPaymentCustomer, setCreditPaymentCustomer] = useState(null);
  const [creditPaymentAmount, setCreditPaymentAmount] = useState('');
  const [customerForm, setCustomerForm] = useState({ name: '', phone: '', address: '', creditLimit: '', priceLevel: 'normal' });
  const [creditPaymentReceipt, setCreditPaymentReceipt] = useState(null);
  const [creditBookPrintSale, setCreditBookPrintSale] = useState(null);
  const [printMode, setPrintMode] = useState('sale');

  useEffect(() => {
    try { safeStorageSet('pos_registered_customers', JSON.stringify(registeredCustomers)); } catch {}
  }, [registeredCustomers]);
  useEffect(() => {
    try { safeStorageSet('pos_credit_ledger', JSON.stringify(creditLedger)); } catch {}
  }, [creditLedger]);
  useEffect(() => {
    if (!creditPaymentReceipt) return;
    const clear = () => setCreditPaymentReceipt(null);
    window.addEventListener('afterprint', clear);
    return () => window.removeEventListener('afterprint', clear);
  }, [creditPaymentReceipt]);
  useEffect(() => {
    if (!creditBookPrintSale) return;
    const clear = () => {
      setCreditBookPrintSale(null);
      setPrintMode('sale');
    };
    window.addEventListener('afterprint', clear);
    return () => window.removeEventListener('afterprint', clear);
  }, [creditBookPrintSale]);

  const [cart, setCart] = useState([]);
  const [barcodeInput, setBarcodeInput] = useState('');
  const [errorMessage, setErrorMessage] = useState('');
  const [notification, setNotification] = useState(null);
  const [pendingUnderpayment, setPendingUnderpayment] = useState(false);
  const [pendingClearAllProducts, setPendingClearAllProducts] = useState(false);
  // Excel import progress (Windows-style percentage indicator)
  const [excelImportProgress, setExcelImportProgress] = useState(null);
  const notificationTimerRef = useRef(null);
  const [paymentMethod, setPaymentMethod] = useState('Cash');
  const [priceLevel, setPriceLevel] = useState('normal');
  
  const [customerName, setCustomerName] = useState('');
  const [customerPhone, setCustomerPhone] = useState('');
  const [paidAmount, setPaidAmount] = useState('');
  // Whole-bill discount applied after existing item-level offers.
  // Percentage and direct LKR are separate inputs; using one clears the other.
  const [billDiscountPercent, setBillDiscountPercent] = useState('0');
  const [billDiscountLkr, setBillDiscountLkr] = useState('0');
  // Inventory scanner helper for quickly starting a new product with its barcode.
  const [inventoryScanCode, setInventoryScanCode] = useState('');
  const inventoryScanRef = useRef(null);
  const inventoryNameRef = useRef(null);
  // Optional live customer-facing bill display.
  const [showCustomerDisplay, setShowCustomerDisplay] = useState(false);
  // Customer-side loyalty / discount configuration and simple report controls.
  const [loyaltyRedeemCustomerId, setLoyaltyRedeemCustomerId] = useState('');
  const [loyaltyRedeemPoints, setLoyaltyRedeemPoints] = useState('');

  // LIVE BILLING DATE & TIME DISPLAY
  const [currentDateTime, setCurrentDateTime] = useState(() => new Date());
  // Keep the selected report date independent from the live clock so historical dates remain selectable.
  useEffect(() => {
    const timer = setInterval(() => setCurrentDateTime(new Date()), 1000);
    return () => clearInterval(timer);
  }, []);

  const [showPreviewModal, setShowPreviewModal] = useState(false);
  const [completedSale, setCompletedSale] = useState(null);
  const [shortcutPrintAfterConfirm, setShortcutPrintAfterConfirm] = useState(false);
  const saleSaveLockRef = useRef(false);

  const [heldBills, setHeldBills] = useState([]);

  const [editedQuantities, setEditedQuantities] = useState({});
  const [isEditing, setIsEditing] = useState({});

  // PERFORMANCE: never render thousands of DOM rows at once.
  // The data can still contain 10,000+ products/sales; only a small page is rendered.
  const TABLE_PAGE_SIZE = 50;
  const [inventoryPage, setInventoryPage] = useState(1);
  const [customerPage, setCustomerPage] = useState(1);
  const [lowStockPage, setLowStockPage] = useState(1);
  const [supplierPage, setSupplierPage] = useState(1);
  const [supplierLowStockPage, setSupplierLowStockPage] = useState(1);
  const [orderRequestPage, setOrderRequestPage] = useState(1);

  const [newProd, setNewProd] = useState({
    name: '', barcode: '', lot_number: '', grn_rate: '', grn_date: '', buying_price: '', price: '', wholesale_price: '', special_price: '', stock: '', min_stock_alert: '5', unit: 'Pcs', offer_type: 'none', offer_value: '0', offer_buy_qty: '1', offer_free_qty: '1', bulk_min_qty: '2', bulk_price: '', expiry_date: '', supplier: ''
  });
  // LOT SELECTION: same barcode can have multiple stock lots with different buy/sell prices.
  const [showLotSelectModal, setShowLotSelectModal] = useState(false);
  const [lotSelectCandidates, setLotSelectCandidates] = useState([]);
  const [lotSelectCode, setLotSelectCode] = useState('');
  const [editingId, setEditingId] = useState(null);
  const [showProductForm, setShowProductForm] = useState(false);

  // Inventory barcode label printing
  const [barcodePrintProduct, setBarcodePrintProduct] = useState(null);
  const [barcodePrintSize, setBarcodePrintSize] = useState('small');
  const [barcodePrintQuantity, setBarcodePrintQuantity] = useState(1);
  const [showBarcodePrintModal, setShowBarcodePrintModal] = useState(false);
  const [showBarcodePreview, setShowBarcodePreview] = useState(false);
  const [barcodePrintItems, setBarcodePrintItems] = useState([]);

  // Professional label-printer mode. The selected sticker size becomes the print page size.
  const [showBarcodePrintChoiceModal, setShowBarcodePrintChoiceModal] = useState(false);
  const [showBarcodeCustomModal, setShowBarcodeCustomModal] = useState(false);
  const [showBarcodeCustomPreview, setShowBarcodeCustomPreview] = useState(false);
  const [barcodeCustomPreset, setBarcodeCustomPreset] = useState('40x30');
  const [barcodeCustomWidth, setBarcodeCustomWidth] = useState('40');
  const [barcodeCustomHeight, setBarcodeCustomHeight] = useState('30');
  const [barcodeCustomGap, setBarcodeCustomGap] = useState('2');
  const [barcodeCustomQuantity, setBarcodeCustomQuantity] = useState(1);

  useEffect(() => {
    if (!barcodePrintItems.length) return;
    const clearAfterPrint = () => setBarcodePrintItems([]);
    window.addEventListener('afterprint', clearAfterPrint);
    return () => window.removeEventListener('afterprint', clearAfterPrint);
  }, [barcodePrintItems.length]);

  const barcodeInputRef = useRef(null);
  // POS-only keyboard navigation targets. These refs do not change any existing page/options.
  const paymentMethodRef = useRef(null);
  const paidAmountRef = useRef(null);
  const payPrintRef = useRef(null);
  const productSelectRef = useRef(null);

  // Dedicated keyboard targets for Billing's two result screens.
  // These are intentionally isolated from the rest of the app so no other
  // module/page controls are changed.
  const previewCancelRef = useRef(null);
  const previewConfirmRef = useRef(null);
  const completedPrintRef = useRef(null);
  const newBillRef = useRef(null);

  // Move focus between Billing controls using the physical arrow-key direction.
  // This is intentionally scoped to the Billing/POS page only.
  const moveBillingFocus = (direction) => {
    if (!isLoggedIn || activeTab !== 'pos') return;
    const active = document.activeElement;
    const nodes = Array.from(document.querySelectorAll('[data-pos-nav]'))
      .filter(el => {
        const r = el.getBoundingClientRect();
        return r.width > 0 && r.height > 0 && !el.disabled;
      });
    if (!nodes.length) return;

    const current = nodes.includes(active) ? active : null;
    if (!current) {
      (direction === 'up' || direction === 'left' ? nodes[nodes.length - 1] : nodes[0])?.focus();
      return;
    }

    const a = current.getBoundingClientRect();
    const ax = a.left + a.width / 2;
    const ay = a.top + a.height / 2;
    const candidates = nodes.filter(el => el !== current).map(el => {
      const r = el.getBoundingClientRect();
      const cx = r.left + r.width / 2;
      const cy = r.top + r.height / 2;
      const dx = cx - ax;
      const dy = cy - ay;
      let primary = 0;
      let secondary = 0;
      if (direction === 'up') { if (dy >= -2) return null; primary = -dy; secondary = Math.abs(dx); }
      if (direction === 'down') { if (dy <= 2) return null; primary = dy; secondary = Math.abs(dx); }
      if (direction === 'left') { if (dx >= -2) return null; primary = -dx; secondary = Math.abs(dy); }
      if (direction === 'right') { if (dx <= 2) return null; primary = dx; secondary = Math.abs(dy); }
      return { el, score: primary * 10 + secondary };
    }).filter(Boolean).sort((x, y) => x.score - y.score);

    if (candidates[0]) {
      candidates[0].el.focus();
      candidates[0].el.scrollIntoView({ block: 'nearest', inline: 'nearest' });
    }
  };

  const [receiptSettings, setReceiptSettings] = useState(() => {
    const saved = safeStorageJson('pos_receipt_settings', null);
    return saved || {
      shopName: 'MY SHOP NAME',
      address: 'No. 123, Main Street, Colombo',
      phone: '0712345678',
      email: 'myshop@gmail.com',
      cashier: 'ADMIN',
      terminal: 'MAIN / L-1',
      footerMsg: 'Thank you, Come Again!',
      footerComment: '',
      poweredByEnabled: false,
      poweredByText: 'Powered by Positha POS System',
      paperWidth: '80mm',
      fontFamily: "'Courier New', Courier, monospace",
      fontSize: '14px',
      receiptDesign: 'classic',
      logoUrl: '',
      logoSize: '40mm',
      elementStyles: {
        shopName: { fontSize: '24px', fontWeight: 900, color: '#000000', textAlign: 'center' },
        address: { fontSize: '16px', fontWeight: 900, color: '#000000', textAlign: 'center' },
        phone: { fontSize: '16px', fontWeight: 900, color: '#000000', textAlign: 'center' },
        items: { fontSize: '14px', fontWeight: 900, color: '#000000', textAlign: 'left' },
        grandTotal: { fontSize: '21px', fontWeight: 900, color: '#000000', textAlign: 'right' },
        footer: { fontSize: '19px', fontWeight: 900, color: '#000000', textAlign: 'center' },
        email: { fontSize: '12px', fontWeight: 700, color: '#000000', textAlign: 'center' },
        meta: { fontSize: '12px', fontWeight: 700, color: '#000000', textAlign: 'left' },
        footerComment: { fontSize: '11px', fontWeight: 700, color: '#000000', textAlign: 'center' },
        poweredBy: { fontSize: '8px', fontWeight: 400, color: '#666666', textAlign: 'center' },
        invoice: { fontSize: '14px', fontWeight: 700, color: '#000000', textAlign: 'left' }
      }
    };
  });

  // Receipt settings compatibility: old saved settings automatically receive
  // the new design/font options without changing existing shop data.
  useEffect(() => {
    setReceiptSettings(prev => ({
      ...prev,
      paperWidth: prev.paperWidth === '58mm' ? '58mm' : '80mm',
      fontFamily: prev.fontFamily || "'Courier New', Courier, monospace",
      fontSize: prev.fontSize || '14px',
      receiptDesign: prev.receiptDesign || 'classic',
      logoSize: prev.logoSize || '40mm',
      footerComment: typeof prev.footerComment === 'string' ? prev.footerComment : '',
      poweredByEnabled: prev.poweredByEnabled === true,
      poweredByText: typeof prev.poweredByText === 'string' && prev.poweredByText.trim() ? prev.poweredByText : 'Powered by Positha POS System',
      elementStyles: {
        shopName: { fontSize: '24px', fontWeight: 900, color: '#000000', textAlign: 'center', ...(prev.elementStyles?.shopName || {}) },
        address: { fontSize: '16px', fontWeight: 900, color: '#000000', textAlign: 'center', ...(prev.elementStyles?.address || {}) },
        phone: { fontSize: '16px', fontWeight: 900, color: '#000000', textAlign: 'center', ...(prev.elementStyles?.phone || {}) },
        items: { fontSize: '14px', fontWeight: 900, color: '#000000', textAlign: 'left', ...(prev.elementStyles?.items || {}) },
        grandTotal: { fontSize: '21px', fontWeight: 900, color: '#000000', textAlign: 'right', ...(prev.elementStyles?.grandTotal || {}) },
        footer: { fontSize: '19px', fontWeight: 900, color: '#000000', textAlign: 'center', ...(prev.elementStyles?.footer || {}) },
        email: { fontSize: '12px', fontWeight: 700, color: '#000000', textAlign: 'center', ...(prev.elementStyles?.email || {}) },
        meta: { fontSize: '12px', fontWeight: 700, color: '#000000', textAlign: 'left', ...(prev.elementStyles?.meta || {}) },
        footerComment: { fontSize: '11px', fontWeight: 700, color: '#000000', textAlign: 'center', ...(prev.elementStyles?.footerComment || {}) },
        poweredBy: { fontSize: '8px', fontWeight: 400, color: '#666666', textAlign: 'center', ...(prev.elementStyles?.poweredBy || {}) },
        invoice: { fontSize: '14px', fontWeight: 700, color: '#000000', textAlign: 'left', ...(prev.elementStyles?.invoice || {}) }
      }
    }));
  }, []);

  // Returns State

  // Check saved Shop Registration on load
  useEffect(() => {
    const savedLogin = localStorage.getItem('pos_last_login');
    if (savedLogin) {
      try {
        const remembered = JSON.parse(savedLogin);
        if (remembered.username) setLoginUsername(remembered.username);
        if (remembered.password) setLoginPassword(remembered.password);
      } catch (e) {}
    }
    const savedShop = localStorage.getItem('pos_shop_account');
    if (savedShop) {
      try {
        const parsed = JSON.parse(savedShop);
        setIsRegistered(true);
        if (parsed.logoUrl) setRegLogoUrl(parsed.logoUrl);
      } catch (e) {}
    }
    const savedCashiers = localStorage.getItem('pos_cashiers');
    if (savedCashiers) {
      try { 
        const parsedCashiers = JSON.parse(savedCashiers);
        setCashiers(parsedCashiers);
        if (parsedCashiers.length > 0) {
          setSettingsCashierUsername(parsedCashiers[0].username);
          setSettingsCashierPassword(parsedCashiers[0].password);
        }
      } catch (e) {}
    }
  }, []);

  useEffect(() => {
    safeStorageSet('pos_receipt_settings', JSON.stringify(receiptSettings));
  }, [receiptSettings]);

  useEffect(() => {
    safeStorageSet('pos_dashboard_style', dashboardStyle);
  }, [dashboardStyle]);

  useEffect(() => {
    safeStorageSet('pos_cashiers', JSON.stringify(cashiers));
  }, [cashiers]);

  const fetchProducts = async () => {
    if (ipcRenderer) {
      const data = await ipcRenderer.invoke('get-products');
      let expiryOverrides = {};
      let supplierOverrides = {};
      let lotOverrides = {};
      let priceOverrides = {};
      try {
        expiryOverrides = JSON.parse(localStorage.getItem('pos_expiry_dates') || '{}') || {};
      } catch (_) {
        expiryOverrides = {};
      }
      try {
        supplierOverrides = JSON.parse(localStorage.getItem('pos_supplier_overrides') || '{}') || {};
      } catch (_) {
        supplierOverrides = {};
      }
      try {
        lotOverrides = JSON.parse(localStorage.getItem('pos_lot_overrides') || '{}') || {};
      } catch (_) {
        lotOverrides = {};
      }
      try {
        priceOverrides = JSON.parse(localStorage.getItem('pos_price_overrides') || '{}') || {};
      } catch (_) {
        priceOverrides = {};
      }
      const mergedProducts = (data || []).map((product) => {
        const override = expiryOverrides[String(product.id)] || expiryOverrides[String(product.barcode)] || '';
        const supplierOverride = supplierOverrides[String(product.id)] ?? supplierOverrides[String(product.barcode)] ?? '';
        // Lot numbers MUST be keyed by product row ID because multiple lots intentionally share one barcode.
        const lotOverride = lotOverrides[String(product.id)] || '';
        const priceOverride = priceOverrides[String(product.id)] || priceOverrides[String(product.barcode)] || {};
        const isBulkOffer = String(product?.offer_type || '').toLowerCase() === 'bulk_price';
        const normalizedBulkMinQty = Math.max(2, Number(product?.bulk_min_qty) || (isBulkOffer ? Number(product?.offer_buy_qty) || 2 : 2));
        const normalizedBulkPrice = Math.max(0, Number(product?.bulk_price) || (isBulkOffer ? Number(product?.offer_value) || 0 : 0));
        return {
          ...product,
          expiry_date: product.expiry_date || override || '',
          supplier: product.supplier || supplierOverride || '',
          lot_number: product.lot_number || lotOverride || '',
          wholesale_price: product.wholesale_price ?? product.wholesalePrice ?? priceOverride.wholesale_price ?? '',
          special_price: product.special_price ?? product.specialPrice ?? priceOverride.special_price ?? '',
          // Bulk offer values are also encoded in the existing offer fields for DB compatibility.
          bulk_min_qty: normalizedBulkMinQty,
          bulk_price: normalizedBulkPrice
        };
      });
      setProducts(mergedProducts);
    }
  };

  const saveExpiryOverride = (product, expiryDate) => {
    try {
      const map = JSON.parse(localStorage.getItem('pos_expiry_dates') || '{}') || {};
      const value = expiryDate || '';
      if (product?.id !== undefined && product?.id !== null) map[String(product.id)] = value;
      if (product?.barcode) map[String(product.barcode)] = value;
      localStorage.setItem('pos_expiry_dates', JSON.stringify(map));
    } catch (_) {}
  };

  const saveSupplierOverride = (product, supplierName) => {
    try {
      const map = JSON.parse(localStorage.getItem('pos_supplier_overrides') || '{}') || {};
      const value = String(supplierName || '').trim();
      if (product?.id !== undefined && product?.id !== null) map[String(product.id)] = value;
      if (product?.barcode) map[String(product.barcode)] = value;
      localStorage.setItem('pos_supplier_overrides', JSON.stringify(map));
    } catch (_) {}
  };

  const saveLotOverride = (product, lotNumber) => {
    try {
      const map = JSON.parse(localStorage.getItem('pos_lot_overrides') || '{}') || {};
      const value = String(lotNumber || '').trim();
      // Never key lots by barcode: same barcode can have many lots.
      if (product?.id !== undefined && product?.id !== null) map[String(product.id)] = value;
      localStorage.setItem('pos_lot_overrides', JSON.stringify(map));
    } catch (_) {}
  };

  const savePriceOverride = (product, prices) => {
    try {
      const map = JSON.parse(localStorage.getItem('pos_price_overrides') || '{}') || {};
      const value = { wholesale_price: prices?.wholesale_price ?? '', special_price: prices?.special_price ?? '' };
      if (product?.id !== undefined && product?.id !== null) map[String(product.id)] = value;
      if (product?.barcode) map[String(product.barcode)] = value;
      localStorage.setItem('pos_price_overrides', JSON.stringify(map));
    } catch (_) {}
  };

  const removeLotOverride = (product) => {
    try {
      const map = JSON.parse(localStorage.getItem('pos_lot_overrides') || '{}') || {};
      if (product?.id !== undefined && product?.id !== null) delete map[String(product.id)];
      localStorage.setItem('pos_lot_overrides', JSON.stringify(map));
    } catch (_) {}
  };

  const removeSupplierOverride = (product) => {
    try {
      const map = JSON.parse(localStorage.getItem('pos_supplier_overrides') || '{}') || {};
      if (product?.id !== undefined && product?.id !== null) delete map[String(product.id)];
      if (product?.barcode) delete map[String(product.barcode)];
      localStorage.setItem('pos_supplier_overrides', JSON.stringify(map));
    } catch (_) {}
  };

  const removeExpiryOverride = (product) => {
    try {
      const map = JSON.parse(localStorage.getItem('pos_expiry_dates') || '{}') || {};
      if (product?.id !== undefined && product?.id !== null) delete map[String(product.id)];
      if (product?.barcode) delete map[String(product.barcode)];
      localStorage.setItem('pos_expiry_dates', JSON.stringify(map));
    } catch (_) {}
  };

  const getSaleInvoiceNo = (sale) => sale?.invoiceNo ?? sale?.invoice_no ?? sale?.id ?? '';
  const getSaleIdentity = (sale) => String(getSaleInvoiceNo(sale));
  const getSaleItemIdentity = (row) => String(row?.invoiceNo ?? row?.invoice_no ?? row?.sale_id ?? row?.saleId ?? '');

  const mergeSalesWithLocal = (remoteSales = []) => {
    const localSales = safeStorageJson('pos_sales_history', []);
    const merged = [];
    const seen = new Set();
    const allSales = [...(Array.isArray(localSales) ? localSales : []), ...(Array.isArray(remoteSales) ? remoteSales : [])];
    allSales.forEach(sale => {
      const key = getSaleIdentity(sale);
      const fallback = `${sale?.date || sale?.created_at || ''}|${sale?.customerPhone || sale?.customer_phone || ''}|${sale?.total || 0}`;
      const identity = key || fallback;
      if (seen.has(identity)) return;

      // Repair the old double-record pattern: renderer invoice + DB auto-id.
      // Matching is deliberately strict (same customer/amount and within 2 seconds).
      const saleTime = parseSaleDate(sale?.date ?? sale?.created_at)?.getTime() || 0;
      const saleTotal = Number(sale?.total ?? 0) || 0;
      const salePhone = String(sale?.customerPhone ?? sale?.customer_phone ?? '-').trim();
      const saleName = String(sale?.customerName ?? sale?.customer_name ?? 'General Customer').trim().toLowerCase();
      const twin = merged.find(x => {
        const xTime = parseSaleDate(x?.date ?? x?.created_at)?.getTime() || 0;
        return saleTime && xTime && Math.abs(saleTime - xTime) <= 2000
          && Math.abs((Number(x?.total ?? 0) || 0) - saleTotal) < 0.005
          && String(x?.customerPhone ?? x?.customer_phone ?? '-').trim() === salePhone
          && String(x?.customerName ?? x?.customer_name ?? 'General Customer').trim().toLowerCase() === saleName;
      });
      if (twin) return;

      seen.add(identity);
      merged.push(sale);
    });
    return merged.sort((a, b) => {
      const da = parseSaleDate(a?.date)?.getTime() || 0;
      const db = parseSaleDate(b?.date)?.getTime() || 0;
      return db - da;
    });
  };

  const mergeSaleItemsWithLocal = (remoteItems = []) => {
    const localItems = safeStorageJson('pos_sales_items', []);
    const localSales = safeStorageJson('pos_sales_history', []);
    const merged = [];
    const seen = new Set();

    // Keep the local item snapshot first. It contains the original renderer invoice
    // and final offer price. Remote rows are added only when they are not the old
    // DB-auto-id twin of an existing local bill.
    const addRow = (row) => {
      const saleKey = getSaleItemIdentity(row);
      const itemKey = String(row?.product_id ?? row?.productId ?? row?.barcode ?? row?.product_name ?? row?.name ?? '');
      const qty = Number(row?.qty ?? row?.quantity ?? 0) || 0;
      const identity = `${saleKey}|${itemKey}|${qty}|${row?.price ?? row?.unit_price ?? 0}`;
      if (seen.has(identity)) return;
      seen.add(identity);
      merged.push(row);
    };

    (Array.isArray(localItems) ? localItems : []).forEach(addRow);

    (Array.isArray(remoteItems) ? remoteItems : []).forEach(row => {
      const rowTime = parseSaleDate(row?.date)?.getTime() || 0;
      const rowTotal = Number(row?.sale_total ?? 0) || 0;
      const rowCustomer = String(row?.customer_name ?? 'General Customer').trim().toLowerCase();
      const legacyTwin = (Array.isArray(localSales) ? localSales : []).find(sale => {
        if (String(getSaleInvoiceNo(sale)) === String(row?.sale_id ?? '')) return false;
        const saleTime = parseSaleDate(sale?.date ?? sale?.created_at)?.getTime() || 0;
        return rowTime && saleTime && Math.abs(rowTime - saleTime) <= 2000
          && Math.abs((Number(sale?.total ?? 0) || 0) - rowTotal) < 0.005
          && String(sale?.customerName ?? sale?.customer_name ?? 'General Customer').trim().toLowerCase() === rowCustomer;
      });
      if (legacyTwin) return;
      addRow(row);
    });

    return merged;
  };

  const persistSalesSnapshot = (sales, items) => {
    safeStorageSet('pos_sales_history', JSON.stringify(Array.isArray(sales) ? sales : []));
    safeStorageSet('pos_sales_items', JSON.stringify(Array.isArray(items) ? items : []));
  };

  const fetchSales = async () => {
    const localSales = safeStorageJson('pos_sales_history', []);
    const localItems = safeStorageJson('pos_sales_items', []);
    if (!ipcRenderer) {
      setSalesHistory(Array.isArray(localSales) ? localSales : []);
      setReportsData(prev => ({ ...prev, sales: Array.isArray(localItems) ? localItems : [] }));
      return;
    }
    try {
      const [data, rep] = await Promise.all([
        ipcRenderer.invoke('get-sales'),
        ipcRenderer.invoke('get-reports-data')
      ]);
      const mergedSales = mergeSalesWithLocal(data || []);
      const remoteItems = Array.isArray(rep?.sales) ? rep.sales : [];
      const mergedItems = mergeSaleItemsWithLocal(remoteItems);
      setSalesHistory(mergedSales);
      setReportsData({ ...(rep || {}), sales: mergedItems, lowStock: rep?.lowStock || [] });
      persistSalesSnapshot(mergedSales, mergedItems);
    } catch (error) {
      console.error('Loading saved sales failed; using local sales snapshot:', error);
      setSalesHistory(Array.isArray(localSales) ? localSales : []);
      setReportsData(prev => ({ ...prev, sales: Array.isArray(localItems) ? localItems : [] }));
    }
  };
// Returns සඳහා State

// Returns සඳහා Functions
const getReturnItemKey = (billId, item) => `${billId}_${item.product_id ?? item.id ?? item.barcode ?? item.name}_${item.lot_number ?? item.lotNumber ?? ''}`;

const normalizeReturnItem = (item, index = 0) => ({
  ...item,
  id: item.id ?? item.product_id ?? item.productId ?? item.product_id ?? item.barcode ?? `item-${index}`,
  product_id: item.product_id ?? item.productId ?? item.id ?? null,
  barcode: item.barcode ?? item.product_barcode ?? '',
  lot_number: item.lot_number ?? item.lotNumber ?? '',
  name: item.name ?? item.product_name ?? item.productName ?? `Item ${index + 1}`,
  qty: Number(item.qty ?? item.quantity ?? item.sold_qty ?? 0),
  price: Number(item.price ?? item.unit_price ?? 0),
  unit: item.unit ?? 'Pcs'
});

const getBillItemsForReturn = (bill) => {
  const directItems = Array.isArray(bill?.items) ? bill.items : [];
  if (directItems.length > 0) return directItems.map(normalizeReturnItem);

  const saleId = bill?.invoiceNo ?? bill?.invoice_no ?? bill?.id ?? bill?.sale_id;
  const reportItems = (reportsData.sales || [])
    .filter(row => String(row.sale_id ?? row.saleId ?? row.invoice_no ?? row.invoiceNo) === String(saleId));

  return reportItems.map((row, index) => normalizeReturnItem({
    id: row.product_id ?? row.productId ?? row.id ?? row.barcode,
    product_id: row.product_id ?? row.productId,
    barcode: row.barcode ?? row.product_barcode,
    lot_number: row.lot_number ?? row.lotNumber ?? '',
    product_name: row.product_name ?? row.name,
    qty: row.qty ?? row.quantity,
    price: row.price ?? row.unit_price,
    unit: row.unit
  }, index));
};

const handleSearchBill = () => {
  const billNo = String(searchBillNo || '').trim().replace(/^#/, '');
  if (!billNo) {
    showNotification('Bill No එකක් ඇතුළත් කරන්න');
    return;
  }

  const bill = salesHistory.find((s) =>
    String(s.id) === billNo ||
    String(s.invoiceNo) === billNo ||
    String(s.invoice_no) === billNo
  );

  if (!bill) {
    showNotification('Bill No එක හම්බුනේ නෑ');
    setReturnBill(null);
    setReturnItems({});
    return;
  }

  const billId = bill.invoiceNo ?? bill.invoice_no ?? bill.id;
  const billItems = getBillItemsForReturn(bill);

  if (billItems.length === 0) {
    showNotification('Bill එක හම්බුනා. නමුත් ඒ Bill එකේ Items සොයාගත නොහැකි වුණා.');
    setReturnBill({ ...bill, items: [] });
    setReturnItems({});
    return;
  }

  const selected = {};
  billItems.forEach((item, index) => {
    const itemKey = getReturnItemKey(billId, { ...item, id: item.id ?? `${item.barcode || item.name}-${index}` });
    const soldQty = Number(item.qty ?? item.quantity ?? 0);
    const alreadyReturned = returnRecords
      .filter(r => String(r.originalBillNo) === String(billId))
      .flatMap(r => r.items || [])
      .filter(r => String(r.itemKey) === String(itemKey))
      .reduce((sum, r) => sum + Number(r.qty || 0), 0);

    const availableQty = Math.max(0, soldQty - alreadyReturned);
    selected[itemKey] = {
      selected: false,
      qty: availableQty,
      maxQty: availableQty
    };
  });

  setReturnBill({ ...bill, items: billItems });
  setReturnItems(selected);
};

const updateReturnItem = (key, changes) => {
  setReturnItems(prev => ({ ...prev, [key]: { ...prev[key], ...changes } }));
};

const handleDeleteReturnRecord = (recordId) => {
  const next = returnRecords.filter(r => r.id !== recordId);
  setReturnRecords(next);
  localStorage.setItem('pos_return_records', JSON.stringify(next));
  showNotification('Return record එක ඉවත් කරන ලදී.', 'success');
};

const handleDeleteAllReturnRecords = () => {
  setReturnRecords([]);
  localStorage.setItem('pos_return_records', JSON.stringify([]));
  showNotification('සියලුම Return records ඉවත් කරන ලදී.', 'success');
};

// Return financials use the exact amount the customer paid for the returned line.
// Existing item-offer / offer calculations are intentionally untouched. The only
// extra step here is allocating the already-calculated Whole Bill Discount across
// the returned lines, proportionally to each line's post-offer finalPrice.
const getReturnItemFinancials = (bill, item, billItems = []) => {
  const qty = Math.max(0, Number(item?.qty ?? item?.quantity ?? 0) || 0);
  const unitSell = Number(item?.price ?? item?.unit_price ?? item?.selling_price ?? 0) || 0;
  const subtotal = Number(item?.subtotal ?? (unitSell * qty)) || 0;
  const hasFinalPrice = item?.finalPrice !== undefined || item?.final_price !== undefined;
  const itemDiscount = Number(item?.discount ?? item?.item_discount ?? 0) || 0;
  const lineFinalPrice = hasFinalPrice
    ? Math.max(0, Number(item?.finalPrice ?? item?.final_price ?? 0) || 0)
    : Math.max(0, subtotal - itemDiscount);

  const lines = Array.isArray(billItems) && billItems.length ? billItems : (Array.isArray(bill?.items) ? bill.items : []);
  const postOfferTotal = lines.reduce((sum, row) => {
    const rowQty = Math.max(0, Number(row?.qty ?? row?.quantity ?? 0) || 0);
    const rowPrice = Number(row?.price ?? row?.unit_price ?? row?.selling_price ?? 0) || 0;
    const rowSubtotal = Number(row?.subtotal ?? (rowPrice * rowQty)) || 0;
    const rowHasFinal = row?.finalPrice !== undefined || row?.final_price !== undefined;
    const rowDiscount = Number(row?.discount ?? row?.item_discount ?? 0) || 0;
    const rowFinal = rowHasFinal
      ? Math.max(0, Number(row?.finalPrice ?? row?.final_price ?? 0) || 0)
      : Math.max(0, rowSubtotal - rowDiscount);
    return sum + rowFinal;
  }, 0);

  const billDiscount = Math.max(0, Number(bill?.billDiscountAmount ?? bill?.bill_discount_amount ?? 0) || 0);
  const allocatedBillDiscount = postOfferTotal > 0
    ? Math.min(lineFinalPrice, billDiscount * (lineFinalPrice / postOfferTotal))
    : 0;
  const netRevenue = Math.max(0, lineFinalPrice - allocatedBillDiscount);
  const buyPrice = Number(item?.buying_price ?? item?.buyPrice ?? item?.cost_price ?? item?.cost ?? 0) || 0;
  const profitImpact = netRevenue - (buyPrice * qty);
  const netUnitPrice = qty > 0 ? netRevenue / qty : 0;

  return { qty, lineFinalPrice, allocatedBillDiscount, netRevenue, netUnitPrice, buyPrice, profitImpact };
};

const handleProcessReturn = async () => {
  if (!returnBill) return;
  const billNo = returnBill.invoiceNo ?? returnBill.invoice_no ?? returnBill.id;
  const selectedItems = (returnBill.items || []).map(item => {
    const key = getReturnItemKey(billNo, item);
    const state = returnItems[key] || {};
    const qty = Math.min(Math.max(Number(state.qty || 0), 0), Number(state.maxQty || 0));
    return { item, key, qty, selected: !!state.selected };
  }).filter(x => x.selected && x.qty > 0);

  if (selectedItems.length === 0) { showNotification('Return කරන්න අවම වශයෙන් භාණ්ඩයක් තෝරන්න', 'error'); return; }

  try {
    const updatedProducts = [...products];
    const missingProducts = [];

    for (const entry of selectedItems) {
      const itemId = String(entry.item.product_id ?? entry.item.id ?? '').trim();
      const itemBarcode = String(entry.item.barcode ?? '').trim();
      const itemLot = String(entry.item.lot_number ?? entry.item.lotNumber ?? '').trim();
      const itemName = String(entry.item.name ?? '').trim().toLowerCase();

      // Return stock MUST go back to the exact sold stock lot.
      // 1) product_id / row id is the strongest match.
      // 2) Legacy records can use barcode + lot number.
      // 3) Barcode-only fallback is allowed only when exactly one inventory row matches.
      let productIndex = itemId
        ? updatedProducts.findIndex(p => String(p.id) === itemId)
        : -1;

      if (productIndex === -1 && itemBarcode && itemLot) {
        productIndex = updatedProducts.findIndex(p =>
          String(p.barcode ?? '').trim() === itemBarcode &&
          String(p.lot_number ?? '').trim() === itemLot
        );
      }

      if (productIndex === -1 && itemBarcode) {
        const barcodeMatches = updatedProducts
          .map((p, index) => ({ p, index }))
          .filter(x => String(x.p.barcode ?? '').trim() === itemBarcode);
        if (barcodeMatches.length === 1) productIndex = barcodeMatches[0].index;
        else if (barcodeMatches.length > 1) {
          throw new Error(`Return lot එක හඳුනාගත නොහැක: ${entry.item.name || itemBarcode}. Bill item එකේ Product ID / Lot Number අවශ්‍යයි.`);
        }
      }

      if (productIndex === -1 && itemName) {
        const nameMatches = updatedProducts
          .map((p, index) => ({ p, index }))
          .filter(x => String(x.p.name || '').trim().toLowerCase() === itemName);
        if (nameMatches.length === 1) productIndex = nameMatches[0].index;
      }

      if (productIndex === -1) {
        missingProducts.push(`${entry.item.name || entry.item.barcode || 'Unknown Item'}${itemLot ? ` (Lot: ${itemLot})` : ''}`);
        continue;
      }

      const current = updatedProducts[productIndex];
      const newStock = Number(current.stock || 0) + Number(entry.qty || 0);
      const updatedProduct = { ...current, stock: newStock };

      // Immediately update the Product Table / Low Stock screen.
      updatedProducts[productIndex] = updatedProduct;

      // Persist the new quantity to the Electron database.
      if (ipcRenderer) {
        await ipcRenderer.invoke('update-product', updatedProduct);
      }
    }

    if (missingProducts.length > 0) {
      throw new Error(`Product not found: ${missingProducts.join(', ')}`);
    }

    setProducts(updatedProducts);

    // Calculate the refund/profit impact from the same final amount the customer
    // actually paid. Item offers are already reflected in finalPrice; the existing
    // Whole Bill Discount is then allocated proportionally across the bill lines.
    const returnFinancials = selectedItems.map(x => ({
      ...x,
      financials: getReturnItemFinancials(returnBill, { ...x.item, qty: x.qty }, returnBill.items || [])
    }));
    const returnTotal = returnFinancials.reduce((sum, x) => sum + x.financials.netRevenue, 0);
    const returnProfitImpact = returnFinancials.reduce((sum, x) => sum + x.financials.profitImpact, 0);

    const record = {
      id: `RET-${Date.now()}`,
      originalBillNo: billNo,
      date: new Date().toLocaleString(),
      reason: returnReason || 'Customer Request',
      total: returnTotal,
      profitImpact: returnProfitImpact,
      processedBy: currentUserRole,
      items: returnFinancials.map(x => ({
        itemKey: x.key, id: x.item.id, product_id: x.item.product_id ?? x.item.id, barcode: x.item.barcode, lot_number: x.item.lot_number || '', name: x.item.name,
        qty: x.qty, unit: x.item.unit || 'Pcs', price: Number(x.item.price || 0),
        net_unit_price: x.financials.netUnitPrice,
        net_revenue: x.financials.netRevenue,
        allocated_bill_discount: x.financials.allocatedBillDiscount,
        profit_impact: x.financials.profitImpact,
        buying_price: x.financials.buyPrice
      }))
    };
    setReturnRecords(prev => { const next = [record, ...prev]; localStorage.setItem('pos_return_records', JSON.stringify(next)); return next; });
    await fetchProducts();
    await fetchSales();
    showNotification(`Return සාර්ථකයි! Refund Total: Rs. ${returnTotal.toFixed(2)} | Stock එකට නැවත එකතු කරන ලදී.`, 'success');
    setReturnBill(null); setReturnItems({}); setSearchBillNo(''); setReturnReason('Customer Request');
  } catch (error) {
    console.error('Return processing error:', error);
    showNotification(`Return process කිරීමේදී දෝෂයක් සිදු විය: ${error?.message || 'Unknown error'}`, 'error');
  }
};
  useEffect(() => {
    fetchProducts();
    fetchSales();
    if (ipcRenderer) {
      ipcRenderer.invoke('get-product-units').then((units) => {
        if (Array.isArray(units) && units.length > 0) {
          setProductUnits(Array.from(new Set(units.map(u => String(u).trim()).filter(Boolean))));
        }
      }).catch(() => {
        // Keep the local unit list if the Electron handler is unavailable.
      });
    }
    const savedSuppliers = localStorage.getItem('pos_suppliers');
    if (savedSuppliers) {
      try { setSuppliers(JSON.parse(savedSuppliers)); } catch (e) {}
    }
    const savedOrders = localStorage.getItem('pos_order_requests');
    if (savedOrders) {
      try { setOrderRequests(JSON.parse(savedOrders)); } catch (e) {}
    }
    const savedShopOrderHistory = localStorage.getItem('pos_shop_order_history');
    if (savedShopOrderHistory) {
      try { setShopOrderHistory(JSON.parse(savedShopOrderHistory)); } catch (e) {}
    }
  }, []);

  useEffect(() => {
    safeStorageSet('pos_suppliers', JSON.stringify(suppliers));
  }, [suppliers]);

  useEffect(() => {
    safeStorageSet('pos_order_requests', JSON.stringify(orderRequests));
  }, [orderRequests]);

  useEffect(() => {
    safeStorageSet('pos_shop_order_history', JSON.stringify(shopOrderHistory));
  }, [shopOrderHistory]);

  useEffect(() => {
    safeStorageSet('pos_monthly_analytics', JSON.stringify(monthlyAnalyticsData));
  }, [monthlyAnalyticsData]);
  useEffect(() => {
    if (Array.isArray(salesHistory)) safeStorageSet('pos_sales_history', JSON.stringify(salesHistory));
  }, [salesHistory]);
  useEffect(() => {
    if (Array.isArray(reportsData?.sales)) safeStorageSet('pos_sales_items', JSON.stringify(reportsData.sales));
  }, [reportsData?.sales]);


  useEffect(() => {
    setInventoryPage(1);
    setCustomerPage(1);
    setLowStockPage(1);
    setSupplierPage(1);
    setSupplierLowStockPage(1);
    setOrderRequestPage(1);
    if (barcodeInputRef.current) {
      barcodeInputRef.current.focus();
    }
  }, [activeTab, completedSale, isLoggedIn, inventorySearch, supplierLowStockSearch]);

  // ==================== POS KEYBOARD SHORTCUTS + ARROW NAVIGATION ====================
  // These controls are intentionally active only while the POS/Billing page is open.
  // Inventory, Returns, Suppliers, Shop Order History, Customer Dashboard, Reports,
  // Monthly Sales Report, Shop Profile, Offers, Barcode and all existing page options
  // remain unchanged.
  useEffect(() => {
    const handlePosShortcuts = (e) => {
      if (!isLoggedIn || activeTab !== 'pos') return;

      const ctrlOrCmd = e.ctrlKey || e.metaKey;
      if (ctrlOrCmd && !e.altKey && !e.shiftKey) {
        const key = String(e.key || '').toLowerCase();

        if (key === 'b') {
          e.preventDefault();
          e.stopPropagation();
          barcodeInputRef.current?.focus();
          barcodeInputRef.current?.select?.();
          return;
        }

        if (key === 'p') {
          e.preventDefault();
          e.stopPropagation();
          if (completedSale) {
            handlePrintReceipt();
          } else {
            payPrintRef.current?.focus();
          }
          return;
        }

        if (key === 'g') {
          e.preventDefault();
          e.stopPropagation();
          paidAmountRef.current?.focus();
          paidAmountRef.current?.select?.();
          return;
        }
      }

      // Result-screen navigation has its own dedicated handler above.
      // Do not let the normal Billing control navigator steal those arrows.
      if (showPreviewModal || completedSale) return;

      // Arrow keys navigate between visible Billing controls only.
      if (!ctrlOrCmd && !e.altKey && !e.shiftKey && ['ArrowUp','ArrowDown','ArrowLeft','ArrowRight'].includes(e.key)) {
        const active = document.activeElement;
        if (active?.closest?.('[data-pos-nav]') || active?.matches?.('[data-pos-nav]')) {
          e.preventDefault();
          e.stopPropagation();
          moveBillingFocus(e.key === 'ArrowUp' ? 'up' : e.key === 'ArrowDown' ? 'down' : e.key === 'ArrowLeft' ? 'left' : 'right');
          return;
        }
      }

      // Enter acts like OK/Select for Billing controls.
      if (!ctrlOrCmd && !e.altKey && !e.shiftKey && e.key === 'Enter') {
        const active = document.activeElement;
        const navEl = active?.closest?.('[data-pos-nav]') || (active?.matches?.('[data-pos-nav]') ? active : null);
        if (!navEl) return;

        e.preventDefault();
        e.stopPropagation();

        if (navEl.tagName === 'BUTTON') {
          if (navEl === payPrintRef.current) setShortcutPrintAfterConfirm(true);
          navEl.click();
          return;
        }

        if (navEl === barcodeInputRef.current) {
          handleBarcodeSubmit(e);
          return;
        }

        if (navEl === productSelectRef.current) {
          const value = navEl.value;
          const prod = products.find(p => String(p.id) === String(value));
          if (prod) addToCart(prod);
          navEl.value = '';
          return;
        }

        if (navEl === paidAmountRef.current) {
          if (payPrintRef.current && !payPrintRef.current.disabled) {
            payPrintRef.current.focus();
          }
          return;
        }

        // For editable Billing fields, Enter advances to the next Billing option.
        moveBillingFocus('down');
      }
    };

    window.addEventListener('keydown', handlePosShortcuts, true);
    return () => window.removeEventListener('keydown', handlePosShortcuts, true);
  }, [isLoggedIn, activeTab, products, cart.length, showPreviewModal, completedSale]);

  // ==================== BILL RESULT SCREEN KEYBOARD CONTROL ====================
  // Preview screen:
  //   Left / Up    -> Cancel / Edit
  //   Right / Down -> Confirm & Open Bill
  //   Enter       -> activate the focused button
  //
  // Completed screen:
  //   Left / Up    -> Receipt Print
  //   Right / Down -> New Bill
  //   Enter       -> activate the focused button
  //
  // This is scoped only to these two Billing result screens. No Inventory,
  // Returns, Suppliers, Shop Order History, Customer Dashboard, Reports,
  // Monthly Sales Report, Shop Profile, Offers, Barcode or Dashboard options
  // are touched.
  useEffect(() => {
    const handleBillResultKeys = (e) => {
      if (!isLoggedIn || activeTab !== 'pos') return;

      const previewOpen = !!showPreviewModal;
      const completedOpen = !!completedSale;
      if (!previewOpen && !completedOpen) return;

      const key = String(e.key || '');
      const isArrow = ['ArrowUp', 'ArrowDown', 'ArrowLeft', 'ArrowRight'].includes(key);
      const isEnter = key === 'Enter';

      if (!isArrow && !isEnter) return;
      if (e.ctrlKey || e.metaKey || e.altKey || e.shiftKey) return;

      e.preventDefault();
      e.stopPropagation();

      if (previewOpen) {
        const cancelBtn = previewCancelRef.current;
        const confirmBtn = previewConfirmRef.current;
        const active = document.activeElement;

        if (isArrow) {
          const goCancel = key === 'ArrowLeft' || key === 'ArrowUp';
          const target = goCancel ? cancelBtn : confirmBtn;
          target?.focus();
          target?.scrollIntoView?.({ block: 'nearest', inline: 'nearest' });
          return;
        }

        const target = active === confirmBtn ? confirmBtn : cancelBtn;
        target?.click();
        return;
      }

      if (completedOpen) {
        const printBtn = completedPrintRef.current;
        const newBtn = newBillRef.current;
        const active = document.activeElement;

        if (isArrow) {
          const goPrint = key === 'ArrowLeft' || key === 'ArrowUp';
          const target = goPrint ? printBtn : newBtn;
          target?.focus();
          target?.scrollIntoView?.({ block: 'nearest', inline: 'nearest' });
          return;
        }

        const target = active === newBtn ? newBtn : printBtn;
        target?.click();
      }
    };

    window.addEventListener('keydown', handleBillResultKeys, true);
    return () => window.removeEventListener('keydown', handleBillResultKeys, true);
  }, [isLoggedIn, activeTab, showPreviewModal, completedSale]);

  // Focus the first action immediately when either Billing result screen opens.
  // This removes the need to click with the mouse before using Arrow/Enter.
  useEffect(() => {
    if (!isLoggedIn || activeTab !== 'pos') return;

    if (showPreviewModal) {
      requestAnimationFrame(() => previewCancelRef.current?.focus());
      return;
    }

    if (completedSale) {
      requestAnimationFrame(() => completedPrintRef.current?.focus());
    }
  }, [isLoggedIn, activeTab, showPreviewModal, completedSale]);

  // ==================== APP-WIDE ARROW + ENTER KEYBOARD NAVIGATION ====================
  // Keeps normal caret movement inside text/number inputs while allowing buttons,
  // selects, checkboxes and other controls across the app to be operated without a mouse.
  useEffect(() => {
    const handleAppKeyboard = (e) => {
      if (!isLoggedIn) return;
      if (e.ctrlKey || e.metaKey || e.altKey) return;
      if (activeTab === 'pos' && (showPreviewModal || completedSale)) return;

      const key = e.key;
      const isArrow = ['ArrowUp','ArrowDown','ArrowLeft','ArrowRight'].includes(key);
      const isEnter = key === 'Enter';
      if (!isArrow && !isEnter) return;

      const active = document.activeElement;
      const tag = active?.tagName;
      const editable = tag === 'INPUT' || tag === 'TEXTAREA' || active?.isContentEditable;

      // Enter on ordinary form fields follows native form behavior. Barcode/search fields
      // with their own onKeyDown handlers therefore keep working unchanged.
      if (isEnter && editable) return;
      // Arrow keys inside editable fields retain native cursor/number behavior.
      if (isArrow && editable) return;

      const controls = Array.from(document.querySelectorAll(
        'button:not([disabled]), select:not([disabled]), input[type="checkbox"]:not([disabled]), input[type="radio"]:not([disabled]), [tabindex]:not([tabindex="-1"])'
      )).filter(el => {
        const r = el.getBoundingClientRect();
        return r.width > 0 && r.height > 0 && getComputedStyle(el).visibility !== 'hidden';
      });
      if (!controls.length) return;

      if (isEnter) {
        if (active?.tagName === 'BUTTON') { e.preventDefault(); active.click(); return; }
        if (active?.tagName === 'SELECT') { return; }
        if (active?.type === 'checkbox' || active?.type === 'radio') { e.preventDefault(); active.click(); return; }
      }

      if (!isArrow) return;
      e.preventDefault();
      const idx = controls.indexOf(active);
      const step = (key === 'ArrowLeft' || key === 'ArrowUp') ? -1 : 1;
      const next = idx < 0 ? controls[0] : controls[(idx + step + controls.length) % controls.length];
      next?.focus();
      next?.scrollIntoView?.({ block: 'nearest', inline: 'nearest' });
    };
    window.addEventListener('keydown', handleAppKeyboard, true);
    return () => window.removeEventListener('keydown', handleAppKeyboard, true);
  }, [isLoggedIn, activeTab, showPreviewModal, completedSale]);

  useEffect(() => {
    setCustomerPage(1);
  }, [customerSearchPhone]);

  const generateBarcode = () => {
    const randomCode = '200' + Math.floor(1000000009 + Math.random() * 9000000000).toString();
    setNewProd((prev) => ({ ...prev, barcode: randomCode }));
  };

  const getProductSellPrice = (product, level = priceLevel) => {
    const retail = Number(product?.retailPrice ?? product?.price) || 0;
    if (level === 'wholesale') {
      const wholesale = Number(product?.wholesale_price ?? product?.wholesalePrice);
      return Number.isFinite(wholesale) && wholesale > 0 ? wholesale : retail;
    }
    if (level === 'special') {
      const special = Number(product?.special_price ?? product?.specialPrice);
      return Number.isFinite(special) && special > 0 ? special : retail;
    }
    return retail;
  };

  const getPriceLevelLabel = (level) => level === 'wholesale' ? 'Wholesale' : level === 'special' ? 'Special' : 'Retail';

  const getLinePricing = (product, qty, level = product?.priceLevel || priceLevel) => {
    const quantity = Number(qty) || 0;
    const retailUnitPrice = Number(product?.retailPrice ?? product?.price) || 0;
    const unitPrice = getProductSellPrice(product, level);
    const subtotal = unitPrice * quantity;
    // Customer benefit from Wholesale/Special pricing is a discount against Retail Price.
    // Keep offer discount separate so existing offer calculations remain unchanged.
    const priceLevelDiscount = (level !== 'normal' && retailUnitPrice > unitPrice)
      ? (retailUnitPrice - unitPrice) * quantity
      : 0;
    let finalPrice = subtotal;
    let discount = 0;
    let offerLabel = '';

    if (product?.offer_type === 'percent' && Number(product.offer_value) > 0) {
      const percent = Math.min(100, Math.max(0, Number(product.offer_value)));
      discount = subtotal * (percent / 100);
      finalPrice = subtotal - discount;
      offerLabel = `${percent}% OFF`;
    } else if (product?.offer_type === 'b1g1') {
      const encodedOffer = Number(product.offer_value) || 0;
      const buyQty = Math.max(1, Number(product.offer_buy_qty) || (encodedOffer >= 1001 ? Math.floor(encodedOffer / 1000) : 1));
      const freeQty = Math.max(1, Number(product.offer_free_qty) || (encodedOffer >= 1001 ? (encodedOffer % 1000) : 1));
      const cycle = buyQty + freeQty;
      const payableQty = quantity - (Math.floor(quantity / cycle) * freeQty);
      finalPrice = unitPrice * Math.max(0, payableQty);
      discount = Math.max(0, subtotal - finalPrice);
      offerLabel = `BUY ${buyQty} GET ${freeQty} FREE`;
    } else if (product?.offer_type === 'bulk_price') {
      // Read from dedicated fields when available, with existing offer fields as DB-safe fallback.
      const minQty = Math.max(2, Number(product.bulk_min_qty) || Number(product.offer_buy_qty) || 2);
      const bulkUnitPrice = Math.max(0, Number(product.bulk_price) || Number(product.offer_value) || 0);
      if (quantity >= minQty && bulkUnitPrice > 0) {
        finalPrice = bulkUnitPrice * quantity;
        discount = Math.max(0, subtotal - finalPrice);
        offerLabel = `Rs. ${bulkUnitPrice} EACH FROM ${minQty}`;
      }
    }

    const effectiveUnitPrice = quantity > 0 ? (finalPrice / quantity) : unitPrice;
    return { unitPrice, retailUnitPrice, effectiveUnitPrice, subtotal, discount, priceLevelDiscount, finalPrice, offerLabel };
  };

  const calculatePrice = (product, qty) => getLinePricing(product, qty).finalPrice;

  const cartSubtotal = cart.reduce((sum, item) => sum + (Number(item.subtotal) || getLinePricing(item, item.qty).subtotal), 0);
  const cartRetailSubtotal = cart.reduce((sum, item) => {
    const retail = Number(item?.retailPrice ?? item?.price) || 0;
    return sum + retail * (Number(item?.qty) || 0);
  }, 0);
  const cartDiscount = cart.reduce((sum, item) => sum + (Number(item.discount) || getLinePricing(item, item.qty).discount), 0);
  const cartPriceLevelDiscount = cart.reduce((sum, item) => sum + (Number(item.priceLevelDiscount) || getLinePricing(item, item.qty, item?.priceLevel || priceLevel).priceLevelDiscount || 0), 0);
  const totalCustomerDiscount = cartDiscount + cartPriceLevelDiscount;

  // Fast lookup for BOTH barcode and product/item codes.
  // IMPORTANT: each key stores ALL matching rows so one barcode can represent multiple lots.
  const barcodeIndex = useMemo(() => {
    const index = new Map();
    const addKey = (value, product) => {
      const raw = String(value ?? '').trim();
      if (!raw) return;
      const keys = new Set([raw, raw.toLowerCase(), raw.toUpperCase()]);
      keys.forEach((key) => {
        const list = index.get(key) || [];
        if (!list.some(p => String(p.id) === String(product.id))) list.push(product);
        index.set(key, list);
      });
    };

    for (const product of products) {
      addKey(product?.barcode, product);
      addKey(product?.item_code, product);
      addKey(product?.itemCode, product);
      addKey(product?.product_code, product);
      addKey(product?.productCode, product);
      addKey(product?.code, product);
      addKey(product?.sku, product);
      addKey(product?.id, product);
    }
    return index;
  }, [products]);

  const closeLotSelectModal = () => {
    setShowLotSelectModal(false);
    setLotSelectCandidates([]);
    setLotSelectCode('');
    requestAnimationFrame(() => barcodeInputRef.current?.focus());
  };

  const selectLotForSale = (product) => {
    closeLotSelectModal();
    addToCart(product);
  };

  const handleBarcodeSubmit = (e) => {
    if (e) {
      e.preventDefault();
      e.stopPropagation?.();
    }

    const inputValue = e?.currentTarget?.value ?? barcodeInputRef.current?.value ?? barcodeInput;
    const code = String(inputValue ?? '').trim();
    if (!code) {
      barcodeInputRef.current?.focus();
      return;
    }

    const matches = barcodeIndex.get(code) || barcodeIndex.get(code.toLowerCase()) || barcodeIndex.get(code.toUpperCase()) || [];
    // Only lots with stock can be selected. Existing expiry validation still runs inside addToCart.
    const availableMatches = matches.filter(item => Number(item.stock || 0) > 0);

    if (availableMatches.length === 1) {
      setErrorMessage('');
      addToCart(availableMatches[0]);
      setBarcodeInput('');
      requestAnimationFrame(() => barcodeInputRef.current?.focus());
      return;
    }

    if (availableMatches.length > 1) {
      setErrorMessage('');
      setBarcodeInput('');
      setLotSelectCode(code);
      setLotSelectCandidates(
        [...availableMatches].sort((a, b) => {
          const da = String(a.grn_date || '');
          const db = String(b.grn_date || '');
          if (da !== db) return da.localeCompare(db); // oldest lot first in the list
          return Number(a.id || 0) - Number(b.id || 0);
        })
      );
      setShowLotSelectModal(true);
      return;
    }

    if (matches.length > 0) {
      setErrorMessage(`Barcode / Item Code (${code}) සඳහා Stock ඉවරයි!`);
    } else {
      setErrorMessage(`Barcode / Item Code (${code}) නොමැත!`);
    }
    setBarcodeInput('');
    requestAnimationFrame(() => barcodeInputRef.current?.focus());
  };

  // ==================== EXPIRY MANAGEMENT ====================
  // Expired = past expiry date.
  // Near Expiry = expires within the next 30 days.
  // Products without an expiry date are not treated as expired.
  const EXPIRY_WARNING_DAYS = 30;

  const getExpiryInfo = (product, referenceDate = new Date()) => {
    const raw = String(product?.expiry_date || '').trim();
    if (!raw) {
      return { hasExpiry: false, isExpired: false, isNearExpiry: false, daysRemaining: null, status: 'No Expiry Date' };
    }

    const parts = raw.split('-').map(Number);
    const expiry = parts.length === 3 && parts.every(Number.isFinite)
      ? new Date(parts[0], parts[1] - 1, parts[2])
      : new Date(raw);

    if (Number.isNaN(expiry.getTime())) {
      return { hasExpiry: false, isExpired: false, isNearExpiry: false, daysRemaining: null, status: 'Invalid Expiry Date' };
    }

    const today = new Date(referenceDate);
    today.setHours(0, 0, 0, 0);
    expiry.setHours(0, 0, 0, 0);

    const daysRemaining = Math.ceil((expiry.getTime() - today.getTime()) / 86400000);
    const isExpired = daysRemaining < 0;
    const isNearExpiry = !isExpired && daysRemaining <= EXPIRY_WARNING_DAYS;

    return {
      hasExpiry: true,
      isExpired,
      isNearExpiry,
      daysRemaining,
      status: isExpired ? 'Expired' : isNearExpiry ? 'Near Expiry' : 'Valid'
    };
  };

  const expirySummary = useMemo(() => {
    let expired = 0;
    let nearExpiry = 0;
    let validWithExpiry = 0;

    products.forEach(product => {
      const info = getExpiryInfo(product);
      if (info.isExpired) expired += 1;
      else if (info.isNearExpiry) nearExpiry += 1;
      else if (info.hasExpiry) validWithExpiry += 1;
    });

    return { expired, nearExpiry, validWithExpiry };
  }, [products]);

  const expiryReportProducts = useMemo(() => {
    return products
      .map(product => ({ product, info: getExpiryInfo(product) }))
      .filter(({ info }) => info.isExpired || info.isNearExpiry)
      .sort((a, b) => (a.info.daysRemaining ?? Infinity) - (b.info.daysRemaining ?? Infinity));
  }, [products]);

  const handleDownloadExpiryExcel = () => {
    if (!expiryReportProducts.length) {
      showNotification('Download කිරීමට Expired / Near Expiry items නැහැ.', 'info');
      return;
    }

    const rows = expiryReportProducts.map(({ product, info }, index) => ({
      No: index + 1,
      Barcode: product.barcode || '',
      'Product Name': product.name || '',
      Stock: Number(product.stock) || 0,
      Unit: product.unit || 'Pcs',
      'Expiry Date': product.expiry_date || '',
      'Days Remaining': info.daysRemaining,
      Status: info.status,
      'Buying Price': Number(product.buying_price) || 0,
      'Selling Price': Number(product.price) || 0,
      Supplier: product.supplier || ''
    }));

    const ws = XLSX.utils.json_to_sheet(rows);
    ws['!cols'] = [
      { wch: 7 }, { wch: 18 }, { wch: 32 }, { wch: 12 }, { wch: 12 },
      { wch: 16 }, { wch: 16 }, { wch: 18 }, { wch: 16 }, { wch: 16 }, { wch: 28 }
    ];

    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, 'Expiry Report');
    XLSX.writeFile(wb, `Expiry_Report_${new Date().toISOString().slice(0, 10)}.xlsx`);
    showNotification('Expired / Near Expiry Excel Report බාගත කිරීම ආරම්භ කළා.', 'success');
  };

  const addToCart = (product) => {
    if (Number(product.stock) <= 0) {
      setErrorMessage('Stock ඉවරයි!');
      return;
    }

    const expiryInfo = getExpiryInfo(product);

    if (expiryInfo.isExpired) {
      setErrorMessage(`❌ ${product.name} කල් ඉකුත් වී ඇත (${product.expiry_date}). Sale කරන්න බැහැ.`);
      showNotification(`❌ ${product.name} කල් ඉකුත් වී ඇත. Sale එකට එකතු නොකරන ලදී.`, 'error');
      return;
    }

    if (expiryInfo.isNearExpiry) {
      setErrorMessage(`⚠️ ${product.name} ${expiryInfo.daysRemaining} දිනකින් කල් ඉකුත් වේ.`);
      showNotification(
        `⚠️ ${product.name} — ${expiryInfo.daysRemaining} දිනකින් expiry. Expiry date එක පරීක්ෂා කරන්න.`,
        'info',
        4500
      );
    } else {
      setErrorMessage('');
    }

    setCart((prev) => {
      const existing = prev.find((item) => String(item.id) === String(product.id));
      const newQty = existing ? Number(existing.qty) + 1 : 1;
      if (newQty > Number(product.stock || 0)) {
        setErrorMessage(`මෙම Lot එකේ Stock ${product.stock} ${product.unit || 'Pcs'} පමණයි!`);
        return prev;
      }
      const appliedPrice = getProductSellPrice(product, priceLevel);
      const appliedProduct = { ...product, priceLevel, retailPrice: Number(product?.retailPrice ?? product?.price) || 0, appliedPrice, price: appliedPrice };
      const pricing = getLinePricing(appliedProduct, newQty, priceLevel);
      const bulkMinQty = Math.max(2, Number(product?.bulk_min_qty) || Number(product?.offer_buy_qty) || 2);
      const isBulkOffer = String(product?.offer_type || '').toLowerCase() === 'bulk_price';
      const previousQty = existing ? Number(existing.qty) || 0 : 0;
      if (isBulkOffer && newQty >= bulkMinQty && previousQty < bulkMinQty) {
        const bulkPrice = Number(product?.bulk_price) || Number(product?.offer_value) || 0;
        if (bulkPrice > 0) {
          showNotification(`🎉 ${product.name}: Bulk Price active! ${bulkMinQty}+ qty → Rs. ${bulkPrice.toFixed(2)} each.`, 'success', 5000);
        }
      }

      if (existing) {
        // Keep the most recently scanned/added item at the top of the cart.
        // Older cart items automatically move downward.
        const updatedItem = { ...existing, ...appliedProduct, ...pricing, qty: newQty };
        return [updatedItem, ...prev.filter((item) => item.id !== product.id)];
      }

      // New scans are inserted at the top; previous items move down.
      return [{ ...appliedProduct, ...pricing, qty: 1 }, ...prev];
    });
  };

  const removeFromCart = (id) => {
    setCart((prev) => prev.filter((item) => item.id !== id));
  };

  const handleQtyInputChange = (id, val) => {
    const newQty = parseFloat(val);
    setCart((prev) =>
      prev.map((item) => {
        if (item.id === id) {
          if (isNaN(newQty) || newQty <= 0) {
            return { ...item, qty: val, finalPrice: 0 };
          }
          const pricing = getLinePricing(item, newQty, item?.priceLevel || priceLevel);
          const bulkMinQty = Math.max(2, Number(item?.bulk_min_qty) || Number(item?.offer_buy_qty) || 2);
          const isBulkOffer = String(item?.offer_type || '').toLowerCase() === 'bulk_price';
          const previousQty = Number(item.qty) || 0;
          if (isBulkOffer && newQty >= bulkMinQty && previousQty < bulkMinQty) {
            const bulkPrice = Number(item?.bulk_price) || Number(item?.offer_value) || 0;
            if (bulkPrice > 0) {
              showNotification(`🎉 ${item.name}: Bulk Price active! ${bulkMinQty}+ qty → Rs. ${bulkPrice.toFixed(2)} each.`, 'success', 5000);
            }
          }
          return { ...item, qty: newQty, ...pricing };
        }
        return item;
      })
    );
  };

  const applyPriceLevelToCart = (nextLevel) => {
    setPriceLevel(nextLevel);
    setCart(prev => prev.map(item => ({
      ...item,
      priceLevel: nextLevel,
      retailPrice: Number(item?.retailPrice ?? item?.price) || 0,
      appliedPrice: getProductSellPrice(item, nextLevel),
      price: getProductSellPrice(item, nextLevel),
      ...getLinePricing(item, item.qty, nextLevel)
    })));
  };

  const itemOfferTotal = cart.reduce((sum, item) => sum + (parseFloat(item.finalPrice) || 0), 0);
  const normalizedBillDiscountPercent = Math.min(100, Math.max(0, Number(billDiscountPercent) || 0));
  const normalizedBillDiscountLkr = Math.min(itemOfferTotal, Math.max(0, Number(billDiscountLkr) || 0));
  const billDiscountType = normalizedBillDiscountLkr > 0 ? 'lkr' : (normalizedBillDiscountPercent > 0 ? 'percent' : 'none');
  const billDiscountAmount = billDiscountType === 'lkr'
    ? normalizedBillDiscountLkr
    : itemOfferTotal * (normalizedBillDiscountPercent / 100);
  const total = Math.max(0, itemOfferTotal - billDiscountAmount);
  const numericPaid = parseFloat(paidAmount) || 0;
  const changeAmount = numericPaid > total ? numericPaid - total : 0;

  const getCustomerBalance = (customerId) => creditLedger
    .filter(e => String(e.customerId) === String(customerId))
    .reduce((sum, e) => sum + (Number(e.type === 'credit' ? e.amount : -e.amount) || 0), 0);

  const getSelectedCreditCustomer = () => registeredCustomers.find(c => String(c.id) === String(selectedCreditCustomerId)) || null;

  const handleSaveCustomer = () => {
    const name = String(customerForm.name || '').trim();
    const phone = String(customerForm.phone || '').trim();
    if (!name || !phone) { showNotification('Customer Name සහ Phone Number දෙකම ඇතුළත් කරන්න.', 'error'); return; }
    const duplicate = registeredCustomers.find(c => String(c.phone).trim() === phone);
    if (duplicate) {
      setSelectedCreditCustomerId(String(duplicate.id));
      setCustomerName(duplicate.name); setCustomerPhone(duplicate.phone);
      setShowCustomerRegisterModal(false);
      showNotification('මෙම Phone Number එකට Customer කෙනෙක් දැනටමත් ඇත. එම Customer තෝරා ගත්තා.', 'info');
      return;
    }
    const customer = { id: `C-${Date.now()}`, name, phone, address: String(customerForm.address || '').trim(), creditLimit: Number(customerForm.creditLimit) || 0, priceLevel: customerForm.priceLevel || 'normal', createdAt: new Date().toISOString() };
    setRegisteredCustomers(prev => [customer, ...prev]);
    setSelectedCreditCustomerId(customer.id);
    setCustomerName(customer.name); setCustomerPhone(customer.phone); applyPriceLevelToCart(customer.priceLevel || 'normal');
    setCustomerForm({ name: '', phone: '', address: '', creditLimit: '', priceLevel: 'normal' });
    setShowCustomerRegisterModal(false);
    showNotification(`Customer ${customer.name} ලියාපදිංචි කළා.`, 'success');
  };

  const handleSelectCreditCustomer = (id) => {
    const c = registeredCustomers.find(x => String(x.id) === String(id));
    setSelectedCreditCustomerId(id);
    if (c) { setCustomerName(c.name); setCustomerPhone(c.phone); applyPriceLevelToCart(c.priceLevel || 'normal'); }
  };

  const handleDeleteRegisteredCustomer = (id) => {
    setRegisteredCustomers(prev => prev.filter(c => String(c.id) !== String(id)));
    setCreditLedger(prev => prev.filter(e => String(e.customerId) !== String(id)));
    if (String(selectedCreditCustomerId) === String(id)) { setSelectedCreditCustomerId(''); setCustomerName(''); setCustomerPhone(''); }
    showNotification('Registered Customer ඉවත් කරන ලදී.', 'success');
  };

  const exportRegisteredCustomersToExcel = () => {
    if (!registeredCustomers.length) { showNotification('Download කිරීමට Registered Customers නැහැ.', 'info'); return; }
    const rows = registeredCustomers.map((c, i) => ({ No: i + 1, 'Customer ID': c.id, 'Customer Name': c.name, 'Phone Number': c.phone, Address: c.address || '-', 'Credit Limit': Number(c.creditLimit || 0), 'Outstanding Balance': getCustomerBalance(c.id), 'Registered Date': c.createdAt ? new Date(c.createdAt).toLocaleString() : '-' }));
    const ws = XLSX.utils.json_to_sheet(rows); const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, 'Registered Customers');
    XLSX.writeFile(wb, `Registered_Customers_${new Date().toISOString().slice(0,10)}.xlsx`);
    showNotification('Registered Customers Excel Download ආරම්භ කළා.', 'success');
  };

  const handleOpenCreditPayment = () => {
    setCreditPaymentSearch(''); setCreditPaymentCustomer(null); setCreditPaymentAmount(''); setShowCreditPaymentModal(true);
  };

  const handleSearchCreditCustomer = () => {
    const q = String(creditPaymentSearch || '').trim().toLowerCase();
    if (!q) { showNotification('Customer ID හෝ Phone Number එකක් ඇතුළත් කරන්න.', 'error'); return; }
    const c = registeredCustomers.find(x => String(x.id).toLowerCase() === q || String(x.phone).toLowerCase() === q);
    if (!c) { setCreditPaymentCustomer(null); showNotification('Customer හමු වුණේ නැහැ.', 'error'); return; }
    setCreditPaymentCustomer(c); setCreditPaymentAmount('');
  };

  const handleSaveCreditPayment = () => {
    if (!creditPaymentCustomer) { showNotification('පළමුව Customer Search කරන්න.', 'error'); return; }
    const amount = Number(creditPaymentAmount); const balance = getCustomerBalance(creditPaymentCustomer.id);
    if (!Number.isFinite(amount) || amount <= 0) { showNotification('ගෙවන මුදල නිවැරදිව ඇතුළත් කරන්න.', 'error'); return; }
    if (amount > balance) { showNotification(`ගෙවන මුදල Outstanding Balance එකට වඩා වැඩියි. Balance: Rs. ${balance.toFixed(2)}`, 'error'); return; }
    const receipt = { id: `CP-${Date.now()}`, customerId: creditPaymentCustomer.id, customerName: creditPaymentCustomer.name, phone: creditPaymentCustomer.phone, previousBalance: balance, paidAmount: amount, remainingBalance: balance - amount, dateFormatted: currentFormattedDate() };
    setCreditLedger(prev => [...prev, { id: receipt.id, customerId: creditPaymentCustomer.id, type: 'payment', amount, date: new Date().toISOString(), reference: receipt.id }]);
    setCreditPaymentReceipt(receipt); setPrintMode('creditPayment'); setShowCreditPaymentModal(false);
    showNotification(`Rs. ${amount.toFixed(2)} payment එක save කළා. ඉතිරි ණය Rs. ${(balance - amount).toFixed(2)}.`, 'success');
    setTimeout(() => window.print(), 250);
  };

  const updateCustomerSettings = (customerId, patch) => {
    setRegisteredCustomers(prev => prev.map(c => String(c.id) === String(customerId) ? { ...c, ...patch } : c));
  };

  const handleRedeemLoyaltyPoints = () => {
    const c = registeredCustomers.find(x => String(x.id) === String(loyaltyRedeemCustomerId));
    const points = Math.floor(Number(loyaltyRedeemPoints) || 0);
    if (!c || points <= 0) { showNotification('Customer සහ redeem points නිවැරදිව තෝරන්න.', 'error'); return; }
    const balance = Number(c.loyaltyPoints || 0);
    if (points > balance) { showNotification(`Points balance ${balance} පමණයි.`, 'error'); return; }
    updateCustomerSettings(c.id, { loyaltyPoints: balance - points, redeemedPoints: Number(c.redeemedPoints || 0) + points });
    setLoyaltyRedeemPoints('');
    showNotification(`${c.name}: ${points} loyalty points redeem කළා.`, 'success');
  };

  const customerPaymentHistory = (customerId) => creditLedger
    .filter(e => String(e.customerId) === String(customerId))
    .sort((a,b) => new Date(b.date || 0) - new Date(a.date || 0));

  const isCustomerOverdue = (customerId) => customerPaymentHistory(customerId).some(e =>
    e.type === 'credit' && Number(e.amount || 0) > 0 && e.dueDate && new Date(`${e.dueDate}T23:59:59`) < new Date() && getCustomerBalance(customerId) > 0
  );

  const sendCustomerReminder = (customer) => {
    const due = getCustomerBalance(customer.id);
    const text = encodeURIComponent(`Hello ${customer.name}, your outstanding balance is Rs. ${due.toFixed(2)}. Please arrange payment. Thank you.`);
    const phone = String(customer.phone || '').replace(/[^0-9]/g, '');
    if (!phone) { showNotification('Customer phone number එක නැහැ.', 'error'); return; }
    window.open(`https://wa.me/${phone.startsWith('0') ? `94${phone.slice(1)}` : phone}?text=${text}`, '_blank', 'noopener,noreferrer');
  };

  const handleHoldBill = () => {
    if (cart.length === 0) {
      setErrorMessage('Bill එකේ Items කිසිවක් නැත!');
      return;
    }
    const newHold = {
      id: Date.now(),
      cart,
      customerName,
      customerPhone,
      priceLevel,
      time: new Date().toLocaleTimeString([], { hour: '2-digit', minute: '2-digit' })
    };
    setHeldBills([...heldBills, newHold]);
    setCart([]);
    setCustomerName('');
    setCustomerPhone('');
    setPaidAmount('');
    setErrorMessage('Bill එක Hold කරන ලදී! ⏸️');
  };

  const handleResumeBill = (heldBill) => {
    setCart(heldBill.cart);
    setCustomerName(heldBill.customerName);
    setCustomerPhone(heldBill.customerPhone);
    setPriceLevel(heldBill.priceLevel || 'normal');
    setHeldBills(heldBills.filter((b) => b.id !== heldBill.id));
    setErrorMessage('');
  };

  const handleOpenPreview = () => {
    if (cart.length === 0) return;
    if (paymentMethod === 'Credit' && !getSelectedCreditCustomer()) {
      showNotification('Credit Bill එකකට Registered Customer කෙනෙක් තෝරන්න.', 'error');
      setShowCustomerRegisterModal(true);
      return;
    }
    if (paymentMethod === 'Credit' && numericPaid > total) {
      showNotification('Credit Bill එකේ Paid Amount එක Total එකට වඩා වැඩි විය නොහැක.', 'error');
      return;
    }

    if (numericPaid < total && paidAmount !== '') {
      setPendingUnderpayment(true);
      showNotification(
        `ගෙවූ මුදල Rs. ${numericPaid.toFixed(2)}යි. Bill Total එක Rs. ${total.toFixed(2)}යි. ඉතිරි Rs. ${(total - numericPaid).toFixed(2)}යි. ඉදිරියට යන්නද?`,
        'info',
        0
      );
      return;
    }
    setShowPreviewModal(true);
  };

  const handleConfirmAndSave = async (forcePrint = false) => {
    if (saleSaveLockRef.current) return;
    saleSaveLockRef.current = true;
    try {
      setShowPreviewModal(false);
      const shouldPrintAfterConfirm = forcePrint || shortcutPrintAfterConfirm;
      setShortcutPrintAfterConfirm(false);

      const localNextInvoice = Math.max(140, ...salesHistory.map(s => Number(getSaleInvoiceNo(s))).filter(Number.isFinite)) + 1;
      let invoiceNo = localNextInvoice;
      if (ipcRenderer) {
        try {
          const dbNextInvoice = Number(await ipcRenderer.invoke('get-next-invoice-no'));
          if (Number.isFinite(dbNextInvoice) && dbNextInvoice > 0) invoiceNo = Math.max(localNextInvoice, dbNextInvoice);
        } catch (_) {}
      }

      const creditCustomer = paymentMethod === 'Credit' ? getSelectedCreditCustomer() : null;
      const salePaid = paymentMethod === 'Credit' ? Math.min(numericPaid, total) : (numericPaid || total);
      const saleDue = paymentMethod === 'Credit' ? Math.max(0, total - salePaid) : 0;
      // Save profit from the actual amount charged after item offers and bill discount.
      const actualSaleProfit = cart.reduce((sum, item) => {
        const qty = Number(item?.qty) || 0;
        const finalRevenue = Number(item?.finalPrice) || 0;
        const buyPrice = Number(item?.buying_price ?? item?.buyPrice ?? item?.cost_price ?? item?.cost) || 0;
        return sum + finalRevenue - (buyPrice * qty);
      }, 0) - billDiscountAmount;
      const salePayload = {
        invoiceNo,
        invoice_no: invoiceNo,
        subtotal: cartSubtotal,
        retailSubtotal: cartRetailSubtotal,
        discount: totalCustomerDiscount + billDiscountAmount,
        itemOfferDiscount: cartDiscount,
        priceLevelDiscount: cartPriceLevelDiscount,
        billDiscountPercent: billDiscountType === 'percent' ? normalizedBillDiscountPercent : 0,
        billDiscountLkr: billDiscountType === 'lkr' ? normalizedBillDiscountLkr : 0,
        billDiscountType,
        billDiscountAmount,
        profit: actualSaleProfit,
        date: new Date().toISOString(),
        priceLevel,
        priceLevelLabel: getPriceLevelLabel(priceLevel),
        total, items: cart, paymentMethod,
        paidAmount: salePaid, changeAmount: paymentMethod === 'Credit' ? 0 : changeAmount,
        customerName: creditCustomer?.name || customerName || 'General Customer',
        customerPhone: creditCustomer?.phone || customerPhone || '-',
        customer_id: creditCustomer?.id || null, creditAmount: saleDue, outstandingBalance: creditCustomer ? getCustomerBalance(creditCustomer.id) + saleDue : 0
      };

      // Save a complete local snapshot immediately. This keeps Customer Dashboard,
      // Daily Sales/Profit and historical reports available even when the Electron
      // database call is slow, unavailable, or returns an empty result after reload.
      const localSale = {
        ...salePayload,
        id: invoiceNo,
        invoiceNo,
        invoice_no: invoiceNo,
        customer_name: salePayload.customerName,
        customer_phone: salePayload.customerPhone,
        created_at: salePayload.date,
        date: salePayload.date,
        items: cart.map(item => ({ ...item, sale_id: invoiceNo, invoice_no: invoiceNo }))
      };
      const localSalesBefore = safeStorageJson('pos_sales_history', []);
      const localItemsBefore = safeStorageJson('pos_sales_items', []);
      const localSalesAfter = [localSale, ...(Array.isArray(localSalesBefore) ? localSalesBefore : [])]
        .filter((sale, index, arr) => index === arr.findIndex(x => String(getSaleIdentity(x)) === String(getSaleIdentity(sale))));
      const localItemsAfter = [
        ...cart.map(item => ({
          ...item,
          sale_id: invoiceNo,
          saleId: invoiceNo,
          invoice_no: invoiceNo,
          invoiceNo,
          product_id: item.product_id ?? item.id ?? null,
          product_name: item.product_name ?? item.name ?? '',
          qty: Number(item.qty) || 0,
          // Keep the same meaning in localStorage and SQLite:
          // price/unit_price = one-unit selling price, final_price = final line total.
          price: Number(item.price ?? item.unitPrice ?? item.unit_price ?? 0) || 0,
          unit_price: Number(item.unitPrice ?? item.unit_price ?? item.price ?? 0) || 0,
          item_discount: Number(item.discount ?? item.item_discount ?? item.itemDiscount ?? 0) || 0,
          final_price: Number(item.finalPrice ?? item.final_price ?? 0) || 0,
          buying_price: Number(item.buying_price ?? item.buyPrice ?? item.cost_price ?? item.cost ?? 0) || 0
        })),
        ...(Array.isArray(localItemsBefore) ? localItemsBefore : [])
      ];

      // SQLite is authoritative in the desktop app. Do not create a local
      // "ghost sale" when stock validation or the database write fails.
      if (ipcRenderer) {
        try {
          const saveResult = await ipcRenderer.invoke('save-sale', salePayload);
          if (!saveResult || saveResult.success === false) {
            const message = saveResult?.error || saveResult?.message || 'Unknown database error';
            console.warn('Database save-sale returned failure:', message);
            showNotification(`Database save අසාර්ථකයි: ${message}`, 'error', 5000);
            return;
          }
          // The backend may return an existing row for an accidental repeated invoke.
          // That is success, but it must never deduct stock a second time locally.
          if (saveResult.alreadySaved) {
            await fetchSales();
            showNotification(`Invoice #${invoiceNo} දැනටමත් save වී ඇත. Duplicate bill එකක් සෑදුවේ නැහැ.`, 'info', 3500);
            return;
          }
        } catch (saveError) {
          console.error('Database save-sale failed:', saveError);
          showNotification('Database save එක අසාර්ථකයි. Bill එක save කළේ නැහැ; නැවත උත්සාහ කරන්න.', 'error', 5000);
          return;
        }
      }

      // Only mirror a confirmed desktop save (or a browser-only session) to local storage.
      persistSalesSnapshot(localSalesAfter, localItemsAfter);
      setSalesHistory(localSalesAfter);
      setReportsData(prev => ({ ...prev, sales: localItemsAfter }));

      if (cart.length) {
        const soldById = new Map();
        const soldByBarcode = new Map();
        cart.forEach(item => {
          const qty = Number(item.qty) || 0;
          if (item.id != null) {
            const key = String(item.id);
            soldById.set(key, (soldById.get(key) || 0) + qty);
          } else {
            // Legacy rows without an ID can still fall back to barcode.
            const barcode = String(item.barcode ?? '').trim();
            if (barcode) soldByBarcode.set(barcode, (soldByBarcode.get(barcode) || 0) + qty);
          }
        });
        setProducts(prev => prev.map(p => {
          const idKey = String(p.id ?? '');
          const qty = soldById.has(idKey)
            ? soldById.get(idKey)
            : soldByBarcode.get(String(p.barcode ?? '').trim()) || 0;
          return qty ? { ...p, stock: Math.max(0, Number(p.stock || 0) - qty) } : p;
        }));
      }

      setPrintMode('sale');
      setCompletedSale({
        invoiceNo,
        items: [...cart],
        subtotal: cartSubtotal,
        retailSubtotal: cartRetailSubtotal,
        discount: totalCustomerDiscount + billDiscountAmount,
        itemOfferDiscount: cartDiscount,
        priceLevelDiscount: cartPriceLevelDiscount,
        billDiscountPercent: billDiscountType === 'percent' ? normalizedBillDiscountPercent : 0,
        billDiscountLkr: billDiscountType === 'lkr' ? normalizedBillDiscountLkr : 0,
        billDiscountType,
        billDiscountAmount,
        profit: actualSaleProfit,
        total,
        paymentMethod,
        paidAmount: salePaid,
        changeAmount: paymentMethod === 'Credit' ? 0 : changeAmount,
        creditAmount: saleDue,
        customerId: creditCustomer?.id || null,
        outstandingBalance: creditCustomer ? getCustomerBalance(creditCustomer.id) + saleDue : 0,
        customerName: creditCustomer?.name || customerName || 'General Customer',
        customerPhone: creditCustomer?.phone || customerPhone || '-',
        dateFormatted: currentFormattedDate()
      });

      // Refresh products/reports only after the completed bill is already rendered.
      if (ipcRenderer) {
        Promise.allSettled([fetchProducts(), fetchSales()]).catch(() => {});
      }

      if (creditCustomer && saleDue > 0) {
        const dueDate = new Date();
        dueDate.setDate(dueDate.getDate() + 30);
        setCreditLedger(prev => [...prev, { id: `CR-${invoiceNo}-${Date.now()}`, customerId: creditCustomer.id, type: 'credit', amount: saleDue, date: new Date().toISOString(), dueDate: dueDate.toISOString().slice(0,10), reference: invoiceNo, billNo: invoiceNo }]);
      }

      // Loyalty points: 1 point for every completed Rs.100 spent by a registered customer.
      const loyaltyCustomer = creditCustomer || registeredCustomers.find(c => String(c.phone || '').trim() && String(c.phone || '').trim() === String(customerPhone || '').trim());
      if (loyaltyCustomer) {
        const earnedPoints = Math.floor(total / 100);
        if (earnedPoints > 0) {
          setRegisteredCustomers(prev => prev.map(c => String(c.id) === String(loyaltyCustomer.id)
            ? { ...c, loyaltyPoints: Number(c.loyaltyPoints || 0) + earnedPoints, lifetimePoints: Number(c.lifetimePoints || 0) + earnedPoints }
            : c));
        }
      }

      if (shouldPrintAfterConfirm) {
        // Wait one paint cycle so the completed receipt is rendered before native printing.
        setTimeout(() => window.print(), 250);
      }    } finally {
      saleSaveLockRef.current = false;
    }
  };

  // Billing-only Ctrl+P: save the current bill and print the configured thermal receipt directly.
  // It intentionally skips the on-screen Bill Preview and leaves every other page shortcut unchanged.
  useEffect(() => {
    const handleBillingCtrlP = (e) => {
      if (!isLoggedIn || activeTab !== 'pos') return;
      if (!(e.ctrlKey || e.metaKey) || String(e.key).toLowerCase() !== 'p') return;
      if (showPreviewModal || completedSale) return;
      if (!cart.length) return;

      e.preventDefault();
      e.stopPropagation();

      if (paymentMethod === 'Credit' && !getSelectedCreditCustomer()) {
        showNotification('Credit Bill එකකට Registered Customer කෙනෙක් තෝරන්න.', 'error');
        setShowCustomerRegisterModal(true);
        return;
      }
      if (paymentMethod === 'Credit' && numericPaid > total) {
        showNotification('Credit Bill එකේ Paid Amount එක Total එකට වඩා වැඩි විය නොහැක.', 'error');
        return;
      }
      if (numericPaid < total && paidAmount !== '') {
        setPendingUnderpayment(true);
        showNotification(
          `ගෙවූ මුදල Rs. ${numericPaid.toFixed(2)}යි. Bill Total එක Rs. ${total.toFixed(2)}යි. ඉතිරි Rs. ${(total - numericPaid).toFixed(2)}යි.`,
          'info',
          0
        );
        return;
      }
      handleConfirmAndSave(true);
    };

    window.addEventListener('keydown', handleBillingCtrlP, true);
    return () => window.removeEventListener('keydown', handleBillingCtrlP, true);
  }, [isLoggedIn, activeTab, cart.length, showPreviewModal, completedSale, paymentMethod, numericPaid, total, paidAmount, selectedCreditCustomerId, registeredCustomers]);

  const handleStartNewBill = () => {
    setCart([]);
    setPaidAmount('');
    setBillDiscountPercent('0');
    setBillDiscountLkr('0');
    setCustomerName('');
    setCustomerPhone('');
    setCompletedSale(null);
    setSelectedCreditCustomerId('');
    setPaymentMethod('Cash');
    setPriceLevel('normal');
    setPrintMode('sale');
    setPendingUnderpayment(false);
    setNotification(null);
    setErrorMessage('');
  };

  const handlePrintReceipt = () => {
    window.print();
  };

  const openBarcodePrintModal = (product) => {
    setBarcodePrintProduct(product);
    setBarcodePrintSize('small');
    setBarcodePrintQuantity(1);
    setBarcodeCustomPreset('40x30');
    setBarcodeCustomWidth('40');
    setBarcodeCustomHeight('30');
    setBarcodeCustomGap('2');
    setBarcodeCustomQuantity(1);
    setShowBarcodePrintChoiceModal(true);
  };

  const openExistingA4BarcodePrinter = () => {
    setShowBarcodePrintChoiceModal(false);
    setShowBarcodeCustomModal(false);
    setShowBarcodeCustomPreview(false);
    setShowBarcodePrintModal(true);
  };

  const openCustomBarcodePrinter = () => {
    setShowBarcodePrintChoiceModal(false);
    setShowBarcodeCustomModal(true);
  };

  const barcodeLabelPresets = [
    { key: '30x20', label: '30 × 20 mm', note: 'Small', width: 30, height: 20 },
    { key: '40x20', label: '40 × 20 mm', note: 'Small barcode', width: 40, height: 20 },
    { key: '40x30', label: '40 × 30 mm', note: 'Standard POS', width: 40, height: 30 },
    { key: '50x25', label: '50 × 25 mm', note: 'Medium', width: 50, height: 25 },
    { key: '50x30', label: '50 × 30 mm', note: 'Medium POS', width: 50, height: 30 },
    { key: '60x30', label: '60 × 30 mm', note: 'Wide barcode', width: 60, height: 30 },
    { key: '60x40', label: '60 × 40 mm', note: 'Large', width: 60, height: 40 },
    { key: '70x40', label: '70 × 40 mm', note: 'Large POS', width: 70, height: 40 },
    { key: '80x50', label: '80 × 50 mm', note: 'Extra large', width: 80, height: 50 },
    { key: 'custom', label: 'Custom size', note: 'Enter your own', width: null, height: null }
  ];

  const applyBarcodeCustomPreset = (key) => {
    setBarcodeCustomPreset(key);
    const preset = barcodeLabelPresets.find(item => item.key === key);
    if (preset && preset.width && preset.height) {
      setBarcodeCustomWidth(String(preset.width));
      setBarcodeCustomHeight(String(preset.height));
    }
  };

  const prepareCustomBarcodePreview = () => {
    const width = Math.max(20, Math.min(100, parseFloat(barcodeCustomWidth) || 40));
    const height = Math.max(15, Math.min(100, parseFloat(barcodeCustomHeight) || 30));
    const gap = Math.max(0, Math.min(10, parseFloat(barcodeCustomGap) || 0));
    const qty = Math.max(1, Math.min(1000, parseInt(barcodeCustomQuantity, 10) || 1));
    if (!barcodePrintProduct || !String(barcodePrintProduct.barcode || '').trim()) {
      showNotification('මෙම Product එකට Barcode එකක් නැහැ.', 'error');
      return;
    }
    setBarcodeCustomWidth(String(width));
    setBarcodeCustomHeight(String(height));
    setBarcodeCustomGap(String(gap));
    setBarcodeCustomQuantity(qty);
    const items = Array.from({ length: qty }, (_, i) => ({ ...barcodePrintProduct, _printIndex: i }));
    setBarcodePrintItems(items);
    setShowBarcodeCustomModal(false);
    setShowBarcodeCustomPreview(true);
  };

  const handlePrintCustomBarcodeLabels = () => {
    if (!barcodePrintItems.length) {
      prepareCustomBarcodePreview();
      return;
    }
    setShowBarcodeCustomPreview(false);
    setTimeout(() => window.print(), 180);
  };

  const prepareBarcodePreview = () => {
    const qty = Math.max(1, Math.min(1000, parseInt(barcodePrintQuantity, 10) || 1));
    if (!barcodePrintProduct || !String(barcodePrintProduct.barcode || '').trim()) {
      showNotification('මෙම Product එකට Barcode එකක් නැහැ.', 'error');
      return;
    }
    const items = Array.from({ length: qty }, (_, i) => ({ ...barcodePrintProduct, _printIndex: i }));
    setBarcodePrintItems(items);
    setShowBarcodePrintModal(false);
    setShowBarcodePreview(true);
  };

  const handlePrintProductBarcodes = () => {
    if (!barcodePrintItems.length) {
      prepareBarcodePreview();
      return;
    }
    setShowBarcodePreview(false);
    // Give React one paint cycle to render the A4 sheet before opening the native print dialog.
    setTimeout(() => window.print(), 150);
  };

  const handleDeleteSaleItem = async (saleId) => {
    try {
      if (ipcRenderer) {
        await ipcRenderer.invoke('delete-sale', saleId);
        const nextSales = safeStorageJson('pos_sales_history', []).filter(s => String(getSaleIdentity(s)) !== String(saleId));
        const nextItems = safeStorageJson('pos_sales_items', []).filter(r => String(getSaleItemIdentity(r)) !== String(saleId));
        persistSalesSnapshot(nextSales, nextItems);
        await fetchSales();
        showNotification(`Sale Record #${saleId} ඉවත් කරන ලදී`, 'success');
      } else {
        setSalesHistory(prev => prev.filter(s => s.id !== saleId));
        showNotification(`Sale Record #${saleId} ඉවත් කරන ලදී`, 'success');
      }
    } catch (error) {
      console.error(error);
      showNotification('Sale Record ඉවත් කිරීම අසාර්ථකයි', 'error');
    }
  };

  const handleClearAllSales = async () => {
    try {
      if (ipcRenderer) {
        const res = await ipcRenderer.invoke('clear-all-sales');
        if (res && res.success === false) throw new Error(res.message || 'Clear failed');
      }
      setSalesHistory([]);
      setReportsData(prev => ({ ...prev, sales: [] }));
      persistSalesSnapshot([], []);
      await fetchSales();
      showNotification('🗑️ සියලුම Sales සහ Reports Records ඉවත් කරන ලදී.', 'success');
    } catch (error) {
      console.error(error);
      showNotification('Sales Records ඉවත් කිරීම අසාර්ථකයි', 'error');
    }
  };

  const handleDeleteAllCustomerHistory = async () => {
    if (salesHistory.length === 0) {
      showNotification('මකා දැමීමට Customer Purchase History එකක් නොමැත', 'info');
      return;
    }
    await handleClearAllSales();
  };

  const exportCustomerHistoryToExcel = () => {
    // Excel export intentionally uses ALL saved invoice dates. Date selection only controls the on-screen view.
    const rows = customerHistoryAllDays
      .map((sale, index) => ({
        No: index + 1,
        'Invoice #': getSaleInvoiceNo(sale),
        'Date & Time': sale.date ? new Date(sale.date).toLocaleString() : '-',
        'Customer Name': sale.customer_name || 'General Customer',
        'Phone Number': sale.customer_phone || '-',
        'Payment Method': sale.payment_method || '-',
        'Paid Amount': Number(sale.paid_amount || 0),
        'Credit/Due': Number(sale.credit_amount || sale.creditAmount || 0),
        'Outstanding After Sale': Number(sale.outstanding_balance || sale.outstandingBalance || 0),
        'Total Amount': Number(sale.total || 0)
      }));
    if (rows.length === 0) {
      showNotification('Download කිරීමට Customer History data නොමැත', 'info');
      return;
    }
    const worksheet = XLSX.utils.json_to_sheet(rows);
    worksheet['!cols'] = [
      { wch: 6 }, { wch: 14 }, { wch: 24 }, { wch: 26 }, { wch: 18 }, { wch: 18 }, { wch: 16 }
    ];
    const workbook = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(workbook, worksheet, 'Customer Purchase History');
    XLSX.writeFile(workbook, `Customer_Purchase_History_${new Date().toISOString().slice(0, 10)}.xlsx`);
    showNotification('Customer Purchase History Excel Download ආරම්භ කළා', 'success');
  };

  const showNotification = (message, type = 'info', duration = null) => {
    const text = String(message || 'Notification');
    if (notificationTimerRef.current) clearTimeout(notificationTimerRef.current);
    setNotification({ id: Date.now(), message: text, type });
    const ms = duration ?? (type === 'error' ? 4500 : 3200);
    if (ms > 0) {
      notificationTimerRef.current = setTimeout(() => setNotification(null), ms);
    }
  };

  const continueAfterUnderpayment = () => {
    setPendingUnderpayment(false);
    setNotification(null);
    if (notificationTimerRef.current) clearTimeout(notificationTimerRef.current);
    setShowPreviewModal(true);
  };

  const cancelUnderpayment = () => {
    setPendingUnderpayment(false);
    setNotification(null);
    if (notificationTimerRef.current) clearTimeout(notificationTimerRef.current);
  };

  // Safely ask Electron to close developer tools if the desktop wrapper supports it.
  useEffect(() => {
    try {
      if (ipcRenderer) ipcRenderer.invoke('close-devtools').catch(() => {});
      if (window.electronAPI && typeof window.electronAPI.closeDevTools === 'function') window.electronAPI.closeDevTools();
    } catch (_) {}
  }, []);

  const handleInventoryBarcodeScan = (e) => {
    if (e) e.preventDefault();
    const code = String(inventoryScanCode || '').trim();
    if (!code) {
      inventoryScanRef.current?.focus();
      return;
    }
    setNewProd(prev => ({ ...prev, barcode: code }));
    setInventoryScanCode('');
    showNotification(`Barcode ${code} Add New Product form එකට ඇතුළත් කළා. Details පුරවා Enter කරන්න.`, 'success');
    requestAnimationFrame(() => inventoryNameRef.current?.focus());
  };

  const handleSaveProduct = async (e) => {
    e.preventDefault();
    if (!newProd.name || newProd.price === '' || newProd.price === null) {
       showNotification('කරුණාකර Item Name සහ Selling Price ඇතුලත් කරන්න!', 'error');
      return;
    }

    const finalBarcode = newProd.barcode && newProd.barcode.trim() !== '' 
      ? newProd.barcode.trim() 
      : '200' + Math.floor(1000000009 + Math.random() * 9000000000).toString();

    const productData = {
      barcode: finalBarcode,
      name: newProd.name.trim(),
      lot_number: String(newProd.lot_number || '').trim() || `LOT-${Date.now()}`,
      grn_rate: parseFloat(newProd.grn_rate) || 0,
      grn_date: newProd.grn_date || '',
      buying_price: parseFloat(newProd.buying_price) || 0,
      price: parseFloat(newProd.price) || 0,
      wholesale_price: newProd.wholesale_price === '' ? '' : (parseFloat(newProd.wholesale_price) || 0),
      special_price: newProd.special_price === '' ? '' : (parseFloat(newProd.special_price) || 0),
      stock: parseFloat(newProd.stock) || 0,
      min_stock_alert: parseFloat(newProd.min_stock_alert) || 5,
      unit: newProd.unit || 'Pcs',
      offer_type: newProd.offer_type || 'none',
      offer_value: newProd.offer_type === 'b1g1'
        ? ((Math.max(1, parseFloat(newProd.offer_buy_qty) || 1) * 1000) + Math.max(1, parseFloat(newProd.offer_free_qty) || 1))
        : (parseFloat(newProd.offer_value) || 0),
      offer_buy_qty: Math.max(1, parseFloat(newProd.offer_buy_qty) || 1),
      offer_free_qty: Math.max(1, parseFloat(newProd.offer_free_qty) || 1),
      // Bulk pricing is stored in the existing offer fields so it works with the existing DB schema.
      bulk_min_qty: Math.max(1, parseFloat(newProd.bulk_min_qty) || 2),
      bulk_price: Math.max(0, parseFloat(newProd.bulk_price) || 0),
      ...(newProd.offer_type === 'bulk_price' ? {
        offer_value: Math.max(0, parseFloat(newProd.bulk_price) || 0),
        offer_buy_qty: Math.max(2, parseFloat(newProd.bulk_min_qty) || 2)
      } : {}),
      expiry_date: newProd.expiry_date || '',
      supplier: newProd.supplier || ''
    };

    if (ipcRenderer) {
      if (editingId) {
        await ipcRenderer.invoke('update-product', { ...productData, id: editingId });
        saveExpiryOverride({ id: editingId, barcode: finalBarcode }, productData.expiry_date);
        saveSupplierOverride({ id: editingId, barcode: finalBarcode }, productData.supplier);
        saveLotOverride({ id: editingId }, productData.lot_number);
        savePriceOverride({ id: editingId, barcode: finalBarcode }, productData);
        showNotification('Update සාර්ථකයි!', 'success');
      } else {
        const addedProduct = await ipcRenderer.invoke('add-product', productData);
        const savedProductRef = {
          id: addedProduct?.id ?? addedProduct?.lastInsertRowid ?? addedProduct?.insertId,
          barcode: finalBarcode
        };
        saveExpiryOverride(savedProductRef, productData.expiry_date);
        saveSupplierOverride(savedProductRef, productData.supplier);
        saveLotOverride(savedProductRef, productData.lot_number);
        savePriceOverride(savedProductRef, productData);
        // Keep the supplier attached even when the Electron product handler does not persist that optional field.
        if (savedProductRef.id !== undefined && savedProductRef.id !== null) {
          try {
            await ipcRenderer.invoke('update-product', { ...productData, id: savedProductRef.id });
          } catch (_) {}
        }
        showNotification('අලුතින් එකතු කරන ලදි!', 'success');
      }
      await fetchProducts();
      await fetchSales();
    }

    resetForm();
  };

  const handleSaveSupplier = (e) => {
    e.preventDefault();
    if (!newSupplier.name) {
      showNotification('කරුණාකර Supplier Name ඇතුළත් කරන්න!');
      return;
    }
    if (editingSupplierId) {
      setSuppliers(suppliers.map(s => s.id === editingSupplierId ? { ...newSupplier, id: editingSupplierId } : s));
      setEditingSupplierId(null);
      showNotification('Supplier යාවත්කාලීන කරන ලදී!');
    } else {
      const newSupObj = { ...newSupplier, id: Date.now() };
      setSuppliers([...suppliers, newSupObj]);
      showNotification('නව Supplier එකතු කරන ලදී!');
    }
    setNewSupplier({ name: '', phone: '', email: '', address: '', company: '' });
  };

  const handleEditSupplier = (sup) => {
    setEditingSupplierId(sup.id);
    setNewSupplier({
      name: sup.name || '',
      phone: sup.phone || '',
      email: sup.email || '',
      address: sup.address || '',
      company: sup.company || ''
    });
  };

  const handleDeleteSupplier = (id) => {
    setSuppliers(prev => prev.filter(s => s.id !== id));
    showNotification('Supplier ඉවත් කරන ලදී', 'success');
  };

  // ORDER REQUEST HANDLERS FOR SUPPLIER DASHBOARD (Supports items without supplier & custom editable order quantities)
  const handleCustomOrderInput = (itemId, val) => {
    setCustomOrderInputs(prev => ({ ...prev, [itemId]: val }));
  };

  const handleCustomSupplierInput = (itemId, val) => {
    setOrderSupplierInputs(prev => ({ ...prev, [itemId]: val }));
  };

  const getSuggestedOrderQty = (item) => {
    const defaultQty = (parseFloat(item.min_stock_alert) || 5) * 2 - (parseFloat(item.stock) || 0);
    return defaultQty > 0 ? defaultQty : 10;
  };

  const handleToggleLowStockSelection = (itemId) => {
    setSelectedLowStockItems(prev => ({ ...prev, [itemId]: !prev[itemId] }));
  };

  const handleSelectAllLowStock = () => {
    const next = { ...selectedLowStockItems };
    filteredSupplierLowStockList.forEach(item => { next[item.id] = true; });
    setSelectedLowStockItems(next);
    showNotification(`${filteredSupplierLowStockList.length} low-stock items selected.`, 'success');
  };

  const handleClearLowStockSelection = () => {
    setSelectedLowStockItems({});
    showNotification('Low Stock selection cleared.', 'info');
  };

  const handleStartOrderGeneration = () => {
    const selectedItems = supplierLowStockList
      .filter(item => selectedLowStockItems[item.id])
      .map(item => ({
        id: item.id, name: item.name, barcode: item.barcode,
        currentStock: Number(item.stock) || 0,
        orderQty: Number(customOrderInputs[item.id] ?? getSuggestedOrderQty(item)) || 1,
        unit: item.unit || 'Pcs',
        supplier: String(orderSupplierInputs[item.id] ?? item.supplier ?? '').trim(),
        buyingPrice: Number(item.buying_price) || 0
      }));

    if (!selectedItems.length) {
      showNotification('කරුණාකර Order කිරීමට අවශ්‍ය Low Stock items select කරන්න.', 'error');
      return;
    }
    const missingSupplier = selectedItems.find(item => !item.supplier || item.supplier === 'General');
    if (missingSupplier) {
      showNotification(`${missingSupplier.name} සඳහා Supplier එකක් තෝරන්න.`, 'error');
      return;
    }
    const orderNumber = `ORD-${new Date().getFullYear()}-${String(Date.now()).slice(-8)}`;
    const totalCost = selectedItems.reduce((sum, item) => sum + (item.buyingPrice * item.orderQty), 0);
    setOrderDraft({
      orderNumber, supplierName: 'Multiple Suppliers', supplierCompany: '', supplierPhone: '',
      supplierAddress: '', supplierEmail: '', date: currentFormattedDate(), status: 'Pending',
      items: selectedItems, totalCost
    });
    setShowOrderGenerateModal(true);
  };

  const handleDraftOrderItemChange = (index, field, value) => {
    setOrderDraft(prev => {
      if (!prev) return prev;
      const items = prev.items.map((item, i) => i === index ? { ...item, [field]: field === 'orderQty' ? Math.max(1, Number(value) || 1) : value } : item);
      return { ...prev, items, totalCost: items.reduce((sum, item) => sum + ((Number(item.buyingPrice) || 0) * (Number(item.orderQty) || 0)), 0) };
    });
  };

  const handleConfirmGeneratedOrder = () => {
    if (!orderDraft || !orderDraft.items?.length) return;
    const finalItems = orderDraft.items.filter(i => Number(i.orderQty) > 0);
    if (!finalItems.length) {
      showNotification('Order Qty අවම වශයෙන් item එකකටවත් ඇතුළත් කරන්න.', 'error');
      return;
    }
    const totalCost = finalItems.reduce((sum, item) => sum + ((Number(item.buyingPrice) || 0) * (Number(item.orderQty) || 0)), 0);
    const savedOrder = { ...orderDraft, items: finalItems, totalCost, id: Date.now(), savedAt: new Date().toISOString() };
    setShopOrderHistory(prev => [savedOrder, ...prev]);
    setSelectedLowStockItems({});
    setShowOrderGenerateModal(false);
    setOrderDraft(null);
    showNotification(`✅ ${savedOrder.orderNumber} Shop Order History වෙත සුරැකිණි.`, 'success');
  };

  const handleDeleteShopOrder = (orderId) => {
    if (window.confirm('මෙම Shop Order එක ඉවත් කිරීමට අවශ්‍යද?')) {
      setShopOrderHistory(prev => prev.filter(o => o.id !== orderId));
    }
  };

  const handleDeleteAllShopOrders = () => {
    if (!shopOrderHistory.length) { showNotification('Delete කිරීමට Shop Orders නැහැ.', 'info'); return; }
    if (window.confirm('සියලුම Shop Orders ඉවත් කරන්නද?')) {
      setShopOrderHistory([]);
      showNotification('සියලුම Shop Orders ඉවත් කරන ලදී.', 'success');
    }
  };

  const handleDownloadShopOrderHistoryExcel = () => {
    const rows = [];
    shopOrderHistory.forEach(ord => (ord.items || []).forEach(it => rows.push({
      'Order Number': ord.orderNumber, 'Order Date': ord.date, Status: ord.status, Supplier: it.supplier,
      Barcode: it.barcode, 'Item Name': it.name, 'Buying Price': Number(it.buyingPrice) || 0,
      'Order Qty': Number(it.orderQty) || 0, Unit: it.unit || 'Pcs',
      'Item Cost': (Number(it.buyingPrice) || 0) * (Number(it.orderQty) || 0), 'Order Total Cost': Number(ord.totalCost) || 0
    })));
    if (!rows.length) { showNotification('Download කිරීමට Shop Order History data නැහැ.', 'info'); return; }
    const ws = XLSX.utils.json_to_sheet(rows); const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, 'Shop Order History');
    XLSX.writeFile(wb, `Shop_Order_History_${new Date().toISOString().slice(0,10)}.xlsx`);
  };

  const handleUpdateOrderStatus = (orderId, newStatus) => {
    setOrderRequests(prev => prev.map(o => o.id === orderId ? { ...o, status: newStatus } : o));
  };

  const handleDeleteOrderRequest = (orderId) => {
    if (window.confirm('මෙම Order Request එක ඉවත් කිරීමට අවශ්‍යද?')) {
      setOrderRequests(prev => prev.filter(o => o.id !== orderId));
    }
  };

  const handleDeleteAllOrderRequests = () => {
    if (!orderRequests.length) {
      showNotification('Delete කිරීමට Saved Orders නැහැ.', 'info');
      return;
    }
    if (window.confirm('සියලුම Saved Order Requests ඉවත් කරන්නද?')) {
      setOrderRequests([]);
      showNotification('සියලුම Saved Order Requests ඉවත් කරන ලදී.', 'success');
    }
  };

  // EXCEL & PDF EXPORT FOR ORDER REQUESTS WITH SHOP DETAILS
  const handleDownloadOrderExcel = (ord) => {
    const supplier = suppliers.find(s => String(s.name).trim() === String(ord.supplierName).trim()) || {};
    const headerRows = [
      { 'Field': 'SHOP / COMPANY', 'Value': receiptSettings.shopName },
      { 'Field': 'Shop Address', 'Value': receiptSettings.address },
      { 'Field': 'Shop Phone', 'Value': receiptSettings.phone },
      { 'Field': 'Supplier', 'Value': ord.supplierName },
      { 'Field': 'Supplier Company', 'Value': ord.supplierCompany || supplier.company || '' },
      { 'Field': 'Supplier Address', 'Value': ord.supplierAddress || supplier.address || '' },
      { 'Field': 'Supplier Phone', 'Value': ord.supplierPhone || supplier.phone || '' },
      { 'Field': 'Supplier Email', 'Value': ord.supplierEmail || supplier.email || '' },
      { 'Field': 'Order Number', 'Value': ord.orderNumber || ord.id },
      { 'Field': 'Order Date', 'Value': ord.date },
      { 'Field': 'Status', 'Value': ord.status }
    ];
    const itemRows = ord.items.map((it, idx) => ({
      '#': idx + 1, 'Item Name': it.name, Barcode: it.barcode,
      'Current Stock': it.currentStock, 'Buy Price': Number(it.buyingPrice) || 0, Unit: it.unit || 'Pcs',
      'Order Qty': it.orderQty, 'Item Cost': (Number(it.buyingPrice) || 0) * (Number(it.orderQty) || 0), Supplier: it.supplier || ord.supplierName
    }));
    const ws = XLSX.utils.json_to_sheet(headerRows);
    XLSX.utils.sheet_add_json(ws, itemRows, { origin: `A${headerRows.length + 3}` });
    ws['!cols'] = [{wch:22},{wch:34},{wch:18},{wch:14},{wch:12},{wch:14},{wch:22}];
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, 'Supplier Order');
    XLSX.writeFile(wb, `${ord.orderNumber || 'Order_Request'}_${String(ord.supplierName).replace(/[^a-z0-9_-]/gi,'_')}.xlsx`);
    showNotification('Supplier Order Excel Report බාගත කිරීම ආරම්භ කළා.', 'success');
  };

  const handleDownloadOrderPDF = (ord) => {
    const doc = new jsPDF();
    if (receiptSettings.logoUrl) {
      try {
        doc.addImage(receiptSettings.logoUrl, 'JPEG', 14, 10, 15, 15);
      } catch (e) {}
    }
    doc.setFontSize(16);
    doc.text(receiptSettings.shopName, receiptSettings.logoUrl ? 32 : 14, 18);
    doc.setFontSize(10);
    doc.text(`Address: ${receiptSettings.address} | Tel: ${receiptSettings.phone}`, receiptSettings.logoUrl ? 32 : 14, 24);
    doc.text(`Email: ${receiptSettings.email || '-'}`, receiptSettings.logoUrl ? 32 : 14, 29);

    doc.setFontSize(12);
    doc.text(`Order Request Report - Supplier: ${ord.supplierName}`, 14, 38);
    doc.setFontSize(10);
    doc.text(`Date: ${ord.date} | Status: ${ord.status}`, 14, 44);

    const tableColumn = ['#', 'Item Name', 'Barcode', 'Current Stock', 'Order Qty'];
    const tableRows = ord.items.map((it, idx) => [
      idx + 1,
      it.name,
      it.barcode,
      `${it.currentStock} ${it.unit}`,
      `${it.orderQty} ${it.unit}`
    ]);

    doc.autoTable({ head: [tableColumn], body: tableRows, startY: 48 });
    doc.save(`Order_Request_${ord.supplierName}_${ord.id}.pdf`);
  };

  const handleExcelImport = (e) => {
    const file = e.target.files && e.target.files[0];
    if (!file) return;

    const reader = new FileReader();
    reader.onload = async (evt) => {
      setExcelImportProgress({ percent: 1, stage: 'Reading Excel file...', total: 0, processed: 0, added: 0, updated: 0 });
      try {
        // Read as ArrayBuffer to preserve Sinhala / Tamil / English Unicode.
        const data = new Uint8Array(evt.target.result);
        const wb = XLSX.read(data, {
          type: 'array',
          cellDates: true,
          raw: true,
          codepage: 65001
        });

        if (!wb.SheetNames || wb.SheetNames.length === 0) {
          showNotification('Excel File එකේ Sheet එකක් හමු නොවීය!', 'error');
          setExcelImportProgress(null);
          return;
        }

        const ws = wb.Sheets[wb.SheetNames[0]];
        if (!ws || !ws['!ref']) {
          showNotification('Excel Sheet එකේ ඩේටා කිසිවක් නැත!', 'error');
          setExcelImportProgress(null);
          return;
        }

        const rawData = XLSX.utils.sheet_to_json(ws, {
          defval: '',
          raw: true,
          blankrows: false
        });

        if (!rawData || rawData.length === 0) {
          showNotification('Excel Sheet එකේ ඩේටා කිසිවක් නැත!', 'error');
          setExcelImportProgress(null);
          return;
        }

        const textValue = (value) => {
          if (value === null || value === undefined) return '';
          return String(value).replace(/\r\n/g, '\n').replace(/\r/g, '\n').trim();
        };

        const getValue = (row, keys) => {
          for (const key of keys) {
            if (Object.prototype.hasOwnProperty.call(row, key)) {
              const value = row[key];
              if (value !== null && value !== undefined && String(value).trim() !== '') return value;
            }
          }
          const normalizedKeys = keys.map(k => String(k).toLowerCase().replace(/[\s_-]+/g, ''));
          const actualKey = Object.keys(row).find(k =>
            normalizedKeys.includes(String(k).toLowerCase().replace(/[\s_-]+/g, ''))
          );
          return actualKey ? row[actualKey] : '';
        };

        const normalizeBarcode = (value) => {
          if (value === null || value === undefined || String(value).trim() === '') return '';
          let barcode = String(value).trim();
          if (/^\d+\.0+$/.test(barcode)) barcode = barcode.replace(/\.0+$/, '');
          return barcode;
        };

        const formattedProducts = rawData.map((row, index) => {
          const name = textValue(getValue(row, [
            'Name', 'Product Name', 'Item Name', 'name', 'product_name', 'item_name'
          ])) || `Item ${index + 1}`;

          const barcode = normalizeBarcode(getValue(row, [
            'Barcode', 'barcode', 'BARCODE'
          ]));

          const buying_price = parseFloat(getValue(row, [
            'Buying Price', 'buying_price', 'Buy Price', 'BuyingPrice'
          ])) || 0;

          const price = parseFloat(getValue(row, [
            'Selling Price', 'price', 'Price', 'Sell Price', 'SellingPrice'
          ])) || 0;

          const stock = parseFloat(getValue(row, [
            'Stock', 'stock', 'Qty', 'Stock Qty', 'Quantity'
          ])) || 0;

          const min_stock_alert = parseFloat(getValue(row, [
            'Low Stock Alert', 'min_stock_alert', 'Low Stock Alert Qty', 'Alert Limit'
          ])) || 5;

          const unit = textValue(getValue(row, ['Unit', 'unit'])) || 'Pcs';
          const offer_type = textValue(getValue(row, ['Offer Type', 'offer_type'])) || 'none';
          const offer_value = parseFloat(getValue(row, ['Offer Value', 'offer_value'])) || 0;
          const offer_buy_qty = Math.max(1, parseFloat(getValue(row, ['Buy Qty', 'Offer Buy Qty', 'offer_buy_qty'])) || 1);
          const offer_free_qty = Math.max(1, parseFloat(getValue(row, ['Free Qty', 'Offer Free Qty', 'offer_free_qty'])) || 1);
          const bulk_min_qty = Math.max(2, parseFloat(getValue(row, ['Bulk Min Qty', 'Minimum Qty', 'bulk_min_qty'])) || 2);
          const bulk_price = Math.max(0, parseFloat(getValue(row, ['Bulk Price', 'Bulk Unit Price', 'bulk_price'])) || 0);
          // Keep bulk settings compatible with the existing offer columns used by the DB handler.
          const saved_offer_value = offer_type === 'bulk_price' ? bulk_price : offer_value;
          const saved_offer_buy_qty = offer_type === 'bulk_price' ? bulk_min_qty : offer_buy_qty;
          const expiry_date = textValue(getValue(row, ['Expiry Date', 'expiry_date'])) || '';
          const supplier = textValue(getValue(row, ['Supplier', 'supplier'])) || '';
          // LOT NUMBER: keep the exact lot value through Excel export -> import.
          // Support the current display name as well as legacy/database-style keys.
          const lot_number = textValue(getValue(row, [
            'Lot', 'Lot Number', 'Lot No', 'lot_number', 'lotNumber', 'LOT', 'LOT NUMBER'
          ])) || '';

          return {
            barcode, name, lot_number, buying_price, price, stock, min_stock_alert, unit,
            offer_type, offer_value: saved_offer_value, offer_buy_qty: saved_offer_buy_qty, offer_free_qty, bulk_min_qty, bulk_price, expiry_date, supplier
          };
        }).filter(item => item.name && item.name.trim() !== '');

        if (formattedProducts.length === 0) {
          showNotification('Excel Sheet එකේ valid product data කිසිවක් හමු නොවීය!', 'error');
          setExcelImportProgress(null);
          return;
        }

        setExcelImportProgress({
          percent: 5,
          stage: 'Checking existing inventory...',
          total: formattedProducts.length,
          processed: 0,
          added: 0,
          updated: 0
        });

        // IMPORTANT: Import now behaves like an inventory update.
        // Existing products are matched by barcode first, then by name when barcode is blank.
        // New products are added through the existing bulk database handler.
        const existingProducts = ipcRenderer
          ? (await ipcRenderer.invoke('get-products') || [])
          : products;

        const barcodeMap = new Map();
        const nameMap = new Map();
        existingProducts.forEach(p => {
          const b = normalizeBarcode(p.barcode);
          if (b) barcodeMap.set(b, p);
          const n = textValue(p.name).toLowerCase();
          if (n) nameMap.set(n, p);
        });

        const toAdd = [];
        const toUpdate = [];
        const seenKeys = new Set();

        formattedProducts.forEach(item => {
          const key = item.barcode ? `b:${item.barcode}` : `n:${item.name.toLowerCase()}`;
          if (seenKeys.has(key)) return;
          seenKeys.add(key);

          const existing = item.barcode
            ? barcodeMap.get(item.barcode)
            : nameMap.get(item.name.toLowerCase());

          if (existing) {
            toUpdate.push({ ...existing, ...item, id: existing.id, barcode: item.barcode || existing.barcode });
          } else {
            // Generate a barcode only for genuinely new rows without a barcode.
            toAdd.push({
              ...item,
              barcode: item.barcode || ('200' + Math.floor(1000000009 + Math.random() * 9000000000).toString())
            });
          }
        });

        let updatedCount = 0;
        let addedCount = 0;

        // Update in small batches so thousands of rows do not freeze the renderer.
        if (ipcRenderer && toUpdate.length) {
          const chunkSize = 25;
          for (let i = 0; i < toUpdate.length; i += chunkSize) {
            const chunk = toUpdate.slice(i, i + chunkSize);
            await Promise.all(chunk.map(async item => {
              await ipcRenderer.invoke('update-product', item);
              // Keep Lot in the same ID-keyed override store used by manual editing.
              // This also preserves lots when the underlying Electron DB handler has
              // an older schema that does not persist the optional lot_number field.
              saveLotOverride(item, item.lot_number);
            }));
            updatedCount += chunk.length;
            const processed = Math.min(formattedProducts.length, updatedCount);
            const percent = Math.min(85, 5 + Math.round((processed / formattedProducts.length) * 55));
            setExcelImportProgress(prev => prev ? {
              ...prev,
              percent,
              stage: `Updating existing stock... ${processed}/${formattedProducts.length}`,
              processed,
              updated: updatedCount
            } : prev);
          }
        }

        // Add new products in one database bulk operation for performance.
        if (ipcRenderer && toAdd.length) {
          setExcelImportProgress(prev => prev ? {
            ...prev,
            percent: 88,
            stage: `Adding ${toAdd.length} new products...`,
            processed: updatedCount
          } : prev);

          const res = await ipcRenderer.invoke('bulk-add-products', toAdd);
          if (res && res.success === false) {
            throw new Error(res.message || 'Bulk product import failed');
          }
          // Some older bulk-add handlers return only success/count. Refresh first,
          // then persist each imported Lot by the newly-created product ID.
          if (ipcRenderer && toAdd.length) {
            const refreshedAfterAdd = await ipcRenderer.invoke('get-products') || [];
            const addedByBarcode = new Map(
              refreshedAfterAdd
                .filter(p => p && p.barcode)
                .map(p => [normalizeBarcode(p.barcode), p])
            );
            toAdd.forEach(item => {
              const saved = addedByBarcode.get(normalizeBarcode(item.barcode));
              if (saved && item.lot_number) saveLotOverride(saved, item.lot_number);
            });
          }
          addedCount = toAdd.length;
        }

        // Refresh once at the end rather than after every row.
        if (ipcRenderer) {
          await fetchProducts();
          await fetchSales();
        } else {
          setProducts(prev => {
            const map = new Map(prev.map(p => [String(p.id), p]));
            toUpdate.forEach(p => map.set(String(p.id), p));
            return [...map.values(), ...toAdd];
          });
        }

        setExcelImportProgress({
          percent: 100,
          stage: 'Import completed successfully',
          total: formattedProducts.length,
          processed: formattedProducts.length,
          added: addedCount,
          updated: updatedCount
        });

        showNotification(
          `🎉 Excel Inventory Update complete! ${updatedCount} updated, ${addedCount} new items added.`,
          'success',
          5000
        );

        window.setTimeout(() => setExcelImportProgress(null), 1400);
      } catch (error) {
        console.error('Excel Import Error:', error);
        setExcelImportProgress(null);
        showNotification(
          `Excel Import කිරීමේදී දෝෂයක් සිදු විය${error?.message ? `: ${error.message}` : '.'}`,
          'error'
        );
      } finally {
        e.target.value = '';
      }
    };

    reader.onerror = () => {
      setExcelImportProgress(null);
      showNotification('Excel File එක කියවීමට නොහැකි විය!', 'error');
      e.target.value = '';
    };

    reader.readAsArrayBuffer(file);
  };

  const resetForm = () => {
    setNewProd({ name: '', barcode: '', lot_number: '', grn_rate: '', grn_date: '', buying_price: '', price: '', wholesale_price: '', special_price: '', stock: '', min_stock_alert: '5', unit: 'Pcs', offer_type: 'none', offer_value: '0', offer_buy_qty: '1', offer_free_qty: '1', bulk_min_qty: '2', bulk_price: '', expiry_date: '', supplier: '' });
    setInventoryScanCode('');
    setEditingId(null);
    setShowProductForm(false);
  };

  const handleEditClick = (p) => {
    setEditingId(p.id);
    setShowProductForm(true);
    window.setTimeout(() => document.getElementById('inventory-product-form')?.scrollIntoView({ behavior:'smooth', block:'start' }), 50);
    setNewProd({
      name: p.name || '',
      barcode: p.barcode || '',
      lot_number: p.lot_number || '',
      grn_rate: p.grn_rate !== undefined && p.grn_rate !== null ? p.grn_rate.toString() : '',
      grn_date: p.grn_date || '',
      buying_price: p.buying_price !== undefined && p.buying_price !== null ? p.buying_price.toString() : '',
      price: p.price !== undefined && p.price !== null ? p.price.toString() : '',
      wholesale_price: p.wholesale_price !== undefined && p.wholesale_price !== null && p.wholesale_price !== '' ? p.wholesale_price.toString() : '',
      special_price: p.special_price !== undefined && p.special_price !== null && p.special_price !== '' ? p.special_price.toString() : '',
      stock: p.stock !== undefined && p.stock !== null ? p.stock.toString() : '',
      min_stock_alert: p.min_stock_alert !== undefined && p.min_stock_alert !== null ? p.min_stock_alert.toString() : '5',
      unit: p.unit || 'Pcs',
      offer_type: p.offer_type || 'none',
      offer_value: p.offer_value !== undefined && p.offer_value !== null ? p.offer_value.toString() : '0',
      offer_buy_qty: p.offer_buy_qty !== undefined && p.offer_buy_qty !== null ? p.offer_buy_qty.toString() : '1',
      offer_free_qty: p.offer_free_qty !== undefined && p.offer_free_qty !== null ? p.offer_free_qty.toString() : '1',
      bulk_min_qty: p.bulk_min_qty !== undefined && p.bulk_min_qty !== null ? p.bulk_min_qty.toString() : (p.offer_type === 'bulk_price' ? String(Math.max(2, Number(p.offer_buy_qty) || 2)) : '2'),
      bulk_price: p.bulk_price !== undefined && p.bulk_price !== null && Number(p.bulk_price) > 0 ? p.bulk_price.toString() : (p.offer_type === 'bulk_price' && Number(p.offer_value) > 0 ? String(p.offer_value) : ''),
      expiry_date: p.expiry_date || '',
      supplier: p.supplier || ''
    });
  };

  const handleDeleteProduct = async (id) => {
    try {
      const product = products.find(p => String(p.id) === String(id));
      if (ipcRenderer) {
        await ipcRenderer.invoke('delete-product', id);
        removeExpiryOverride(product);
        removeLotOverride(product);
        await fetchProducts();
        await fetchSales();
        showNotification(`🗑️ ${product?.name || 'Product'} ඉවත් කරන ලදී!`, 'success');
      } else {
        setProducts(prev => prev.filter(p => String(p.id) !== String(id)));
        showNotification('🗑️ Product ඉවත් කරන ලදී!', 'success');
      }
    } catch (error) {
      console.error('Delete product error:', error);
      showNotification('❌ Product ඉවත් කිරීම අසාර්ථකයි!', 'error');
    }
  };

  const handleClearAllProducts = async () => {
    if (!products || products.length === 0) {
      showNotification('ℹ️ ඉවත් කිරීමට Products නැහැ!', 'info');
      return;
    }
    if (notificationTimerRef.current) clearTimeout(notificationTimerRef.current);
    setPendingClearAllProducts(true);
    setNotification({ id: Date.now(), message: `⚠️ Inventory එකේ තියෙන Products ${products.length} ම සියල්ලම Delete කරන්නද? OK කළොත් සියල්ලම මැකේ.`, type: 'info' });
  };

  const confirmClearAllProducts = async () => {
    setPendingClearAllProducts(false);
    setNotification(null);
    if (notificationTimerRef.current) clearTimeout(notificationTimerRef.current);
    try {
      if (ipcRenderer) {
        await ipcRenderer.invoke('clear-all-products');
        try { localStorage.removeItem('pos_expiry_dates'); } catch (_) {}
        try { localStorage.removeItem('pos_price_overrides'); } catch (_) {}
        await fetchProducts();
        await fetchSales();
      } else {
        setProducts([]);
      }
      showNotification('🗑️ සියලුම Products ඉවත් කරන ලදී!', 'success');
    } catch (error) {
      console.error('Clear all products error:', error);
      showNotification('❌ සියලුම Products ඉවත් කිරීම අසාර්ථකයි!', 'error');
    }
  };

  const cancelClearAllProducts = () => {
    setPendingClearAllProducts(false);
    setNotification(null);
    if (notificationTimerRef.current) clearTimeout(notificationTimerRef.current);
  };

  const handleLogoUpload = (e) => {
    const file = e.target.files[0];
    if (!file) return;
    const reader = new FileReader();
    reader.onloadend = () => {
      setReceiptSettings({ ...receiptSettings, logoUrl: reader.result });
    };
    reader.readAsDataURL(file);
  };

  const handleQuantityChange = (barcode, value) => {
    setEditedQuantities(prev => ({ ...prev, [barcode]: value }));
  };

  const toggleEdit = (barcode, currentStock) => {
    setIsEditing(prev => ({ ...prev, [barcode]: !prev[barcode] }));
  };

  const handleDeleteLowStockItem = async (item) => {
    try {
      if (ipcRenderer) {
        await ipcRenderer.invoke('delete-product', item.id);
      }
      setProducts(prev => prev.filter(p => String(p.id) !== String(item.id)));
      setReportsData(prev => ({ ...prev, lowStock: (prev.lowStock || []).filter(p => String(p.id) !== String(item.id)) }));
      showNotification(`${item.name} product එක ඉවත් කරන ලදී`, 'success');
    } catch (error) {
      console.error(error);
      showNotification('Low Stock product ඉවත් කිරීම අසාර්ථකයි', 'error');
    }
  };

  const handleDeleteAllLowStockItems = async () => {
    const items = [...(reportsData.lowStock || [])];
    if (items.length === 0) {
      showNotification('Delete කිරීමට Low Stock products නැහැ', 'info');
      return;
    }
    try {
      if (ipcRenderer) {
        for (const item of items) {
          await ipcRenderer.invoke('delete-product', item.id);
        }
      }
      const ids = new Set(items.map(i => String(i.id)));
      setProducts(prev => prev.filter(p => !ids.has(String(p.id))));
      setReportsData(prev => ({ ...prev, lowStock: [] }));
      showNotification('සියලුම Low Stock products ඉවත් කරන ලදී', 'success');
    } catch (error) {
      console.error(error);
      showNotification('Low Stock products ඉවත් කිරීම අසාර්ථකයි', 'error');
    }
  };

  const handleDownloadLowStockExcel = () => {
    if (window.XLSX && reportsData.lowStock) {
      const dataToExport = reportsData.lowStock.map(item => ({
        Barcode: item.barcode,
        "Product Name": item.name,
        "Remaining Stock": editedQuantities[item.barcode] !== undefined ? editedQuantities[item.barcode] : item.stock,
        "Alert Limit": item.min_stock_alert || 5,
        "Buy Price": Number(item.buying_price) || 0,
        "Unit": item.unit || "Pcs"
      }));
      const worksheet = XLSX.utils.json_to_sheet(dataToExport);
      const workbook = XLSX.utils.book_new();
      XLSX.utils.book_append_sheet(workbook, worksheet, "Low Stock Report");
      XLSX.writeFile(workbook, "Low_Stock_Report.xlsx");
    }
  };

  const getSaleDateValue = (sale) => sale?.date ?? sale?.created_at ?? sale?.createdAt ?? sale?.dateFormatted ?? sale?.timestamp ?? null;
  const parseSaleDate = (value) => {
    if (value instanceof Date) return Number.isNaN(value.getTime()) ? null : value;
    if (!value) return null;
    const raw = String(value).trim();
    if (!raw) return null;
    let d = new Date(raw);
    if (!Number.isNaN(d.getTime())) return d;
    const datePart = raw.split(',')[0].trim();
    const parts = datePart.split(/[\/\-.]/).map(x => x.trim());
    if (parts.length === 3 && parts.every(Boolean)) {
      const a=Number(parts[0]), b=Number(parts[1]), c=Number(parts[2]);
      if ([a,b,c].every(Number.isFinite) && c >= 1000) {
        const year=c, month=a > 12 && b <= 12 ? b : a, day=a > 12 && b <= 12 ? a : b;
        d=new Date(year, month-1, day);
        if (!Number.isNaN(d.getTime()) && d.getFullYear()===year && d.getMonth()===month-1 && d.getDate()===day) return d;
      }
    }
    const m=raw.match(/^(\d{4})-(\d{1,2})-(\d{1,2})/);
    if(m){ d=new Date(Number(m[1]),Number(m[2])-1,Number(m[3])); if(!Number.isNaN(d.getTime())) return d; }
    return null;
  };
  const localDateKey = (value) => {
    const d = value instanceof Date ? value : parseSaleDate(value);
    if (!d) return '';
    return `${d.getFullYear()}-${String(d.getMonth()+1).padStart(2,'0')}-${String(d.getDate()).padStart(2,'0')}`;
  };

  const filteredCustomerHistory = useMemo(() => {
    const q = customerSearchPhone.trim().toLowerCase();
    const dateFiltered = selectedReportDate ? salesHistory.filter(s => localDateKey(getSaleDateValue(s)) === selectedReportDate) : salesHistory;
    if (!q) return dateFiltered;
    return dateFiltered.filter(s => String(s.customer_phone ?? s.customerPhone ?? '').toLowerCase().includes(q) || String(s.customer_name ?? s.customerName ?? '').toLowerCase().includes(q));
  }, [salesHistory, customerSearchPhone, selectedReportDate]);

  const customerHistoryAllDays = useMemo(() => {
    const q = customerSearchPhone.trim().toLowerCase();
    if (!q) return salesHistory;
    return salesHistory.filter(s => String(s.customer_phone ?? s.customerPhone ?? '').toLowerCase().includes(q) || String(s.customer_name ?? s.customerName ?? '').toLowerCase().includes(q));
  }, [salesHistory, customerSearchPhone]);

  const inventoryPageCount = Math.max(1, Math.ceil(products.length / TABLE_PAGE_SIZE));
  const customerPageCount = Math.max(1, Math.ceil(filteredCustomerHistory.length / TABLE_PAGE_SIZE));
  const lowStockList = useMemo(
    () => products.filter(p => (parseFloat(p.stock) || 0) <= (parseFloat(p.min_stock_alert) || 5)),
    [products]
  );
  const lowStockPageCount = Math.max(1, Math.ceil(lowStockList.length / TABLE_PAGE_SIZE));
  const supplierPageCount = Math.max(1, Math.ceil(suppliers.length / TABLE_PAGE_SIZE));
  const supplierLowStockList = lowStockList;
  const filteredSupplierLowStockList = useMemo(() => {
    const q = deferredSupplierLowStockSearch.trim().toLowerCase();
    if (!q) return supplierLowStockList;
    return supplierLowStockList.filter(p => {
      const name = String(p.name || '').toLowerCase();
      const barcode = String(p.barcode || '').toLowerCase();
      const supplier = String(p.supplier || '').toLowerCase();
      return name.includes(q) || barcode.includes(q) || supplier.includes(q);
    });
  }, [supplierLowStockList, deferredSupplierLowStockSearch]);
  const supplierLowStockPageCount = Math.max(1, Math.ceil(filteredSupplierLowStockList.length / TABLE_PAGE_SIZE));
  const visibleOrderRequests = useMemo(
    () => orderRequests.filter(o => String(o.supplierName || '').trim() && String(o.supplierName || '').trim() !== 'General'),
    [orderRequests]
  );
  const orderRequestPageCount = Math.max(1, Math.ceil(visibleOrderRequests.length / TABLE_PAGE_SIZE));

  // GLOBAL INVENTORY SEARCH: this searches every product, not just the current page.
  const filteredInventory = useMemo(() => {
    const q = deferredInventorySearch.trim().toLowerCase();
    if (!q) return products;
    const exact = [];
    const partial = [];
    for (const p of products) {
      const name = String(p.name || '').trim();
      const barcode = String(p.barcode || '').trim();
      const supplier = String(p.supplier || '').trim();
      const nameLower = name.toLowerCase();
      const barcodeLower = barcode.toLowerCase();
      const supplierLower = supplier.toLowerCase();
      if (barcodeLower === q || nameLower === q) exact.push(p);
      else if (nameLower.includes(q) || barcodeLower.includes(q) || supplierLower.includes(q)) partial.push(p);
    }
    return exact.length ? [...exact, ...partial] : partial;
  }, [products, deferredInventorySearch]);
  const filteredInventoryPageCount = Math.max(1, Math.ceil(filteredInventory.length / TABLE_PAGE_SIZE));
  const pagedInventory = filteredInventory.slice((inventoryPage - 1) * TABLE_PAGE_SIZE, inventoryPage * TABLE_PAGE_SIZE);
  const pagedCustomerHistory = filteredCustomerHistory.slice((customerPage - 1) * TABLE_PAGE_SIZE, customerPage * TABLE_PAGE_SIZE);
  const pagedLowStock = lowStockList.slice((lowStockPage - 1) * TABLE_PAGE_SIZE, lowStockPage * TABLE_PAGE_SIZE);
  const pagedSuppliers = suppliers.slice((supplierPage - 1) * TABLE_PAGE_SIZE, supplierPage * TABLE_PAGE_SIZE);
  const pagedSupplierLowStock = filteredSupplierLowStockList.slice((supplierLowStockPage - 1) * TABLE_PAGE_SIZE, supplierLowStockPage * TABLE_PAGE_SIZE);
  const pagedOrderRequests = visibleOrderRequests.slice((orderRequestPage - 1) * TABLE_PAGE_SIZE, orderRequestPage * TABLE_PAGE_SIZE);

  const PaginationControls = ({ page, pageCount, setPage, totalRows, label }) => (
    <div className="table-pagination no-print">
      <span>{label}: <b>{totalRows}</b> | Page <b>{page}</b> / <b>{pageCount}</b></span>
      <div style={{ display: 'flex', gap: '6px' }}>
        <button disabled={page <= 1} onClick={() => setPage(1)}>⏮ First</button>
        <button disabled={page <= 1} onClick={() => setPage(p => Math.max(1, p - 1))}>◀ Prev</button>
        <button disabled={page >= pageCount} onClick={() => setPage(p => Math.min(pageCount, p + 1))}>Next ▶</button>
        <button disabled={page >= pageCount} onClick={() => setPage(pageCount)}>Last ⏭</button>
      </div>
    </div>
  );

  const stockSummary = useMemo(() => {
    let cost = 0;
    let selling = 0;
    const keys = new Set();
    for (const p of products) {
      const stock = parseFloat(p.stock) || 0;
      cost += (parseFloat(p.buying_price) || 0) * stock;
      selling += (parseFloat(p.price) || 0) * stock;
      const key = String(p.barcode || p.id || '').trim();
      if (key) keys.add(key);
    }
    return { cost, selling, count: keys.size };
  }, [products]);
  const totalStockCostValue = stockSummary.cost;
  const totalStockSellingValue = stockSummary.selling;
  const totalStockItemsCount = stockSummary.count;

  const now = new Date();
  const currentYear = now.getFullYear();
  const currentMonth = now.getMonth();
  const currentDate = now.getDate();

  // RETURN IMPACT: returns reduce sales, profit and product performance totals.
  const returnAmountByOriginalBill = useMemo(() => {
    const map = new Map();
    (returnRecords || []).forEach((ret) => {
      const billKey = String(ret.originalBillNo ?? '');
      map.set(billKey, (map.get(billKey) || 0) + (Number(ret.total) || 0));
    });
    return map;
  }, [returnRecords]);

  // Exact return-profit impact for bills processed after this fix. Older return
  // records do not have this field, so reports keep the previous proportional
  // fallback for backward compatibility.
  const returnProfitImpactByOriginalBill = useMemo(() => {
    const map = new Map();
    (returnRecords || []).forEach((ret) => {
      const billKey = String(ret.originalBillNo ?? '');
      const itemImpact = Array.isArray(ret.items)
        ? ret.items.reduce((sum, item) => sum + (Number(item?.profit_impact) || 0), 0)
        : 0;
      const impact = ret.profitImpact !== undefined ? Number(ret.profitImpact) || 0 : itemImpact;
      map.set(billKey, (map.get(billKey) || 0) + impact);
    });
    return map;
  }, [returnRecords]);

  // Recalculate Dashboard sales/profit after customer returns.
  // Build this index once per render instead of filtering the entire sales-detail table for every sale.
  const saleItemsBySaleId = useMemo(() => {
    const map = new Map();
    (reportsData.sales || []).forEach((item) => {
      const key = String(item.sale_id ?? item.saleId ?? '');
      if (!map.has(key)) map.set(key, []);
      map.get(key).push(item);
    });
    return map;
  }, [reportsData.sales]);

  const getSaleItemsForProfit = (sale) => {
    const keys = [
      sale?.id,
      sale?.invoiceNo,
      sale?.invoice_no,
      sale?.sale_id
    ].filter(v => v !== undefined && v !== null && String(v) !== '');
    for (const key of keys) {
      const rows = saleItemsBySaleId.get(String(key));
      if (rows && rows.length) return rows;
    }
    return Array.isArray(sale?.items) ? sale.items : [];
  };

  // PROFIT DISCOUNT IMPACT: calculate profit from the ACTUAL amount charged.
  // The report rows can contain only the original selling price, while the
  // saved bill contains the real item offer/final price. Merge both sources so
  // Dashboard/Reports never fall back to original Sell Price when a discount
  // was applied.
  const getSaleDiscountBreakdown = (sale, itemsForSale = []) => {
    const saleItems = itemsForSale.length ? itemsForSale : (Array.isArray(sale?.items) ? sale.items : []);
    const detailItemDiscount = saleItems.reduce((sum, item) => sum + (Number(item.discount ?? item.item_discount ?? 0) || 0), 0);
    const explicitItemDiscount = Number(sale?.itemOfferDiscount ?? sale?.item_offer_discount ?? 0) || 0;
    const explicitBillDiscount = Number(sale?.billDiscountAmount ?? sale?.bill_discount_amount ?? 0) || 0;

    // IMPORTANT: older saved bills may not have itemOfferDiscount /
    // billDiscountAmount fields, but they still have subtotal + total.
    // Recover the real discount from the amount actually charged.
    const savedSubtotal = Number(sale?.subtotal ?? sale?.sub_total ?? 0) || 0;
    const savedTotal = Number(sale?.total ?? sale?.grandTotal ?? sale?.grand_total ?? 0) || 0;
    const savedDiscountField = Number(sale?.discount ?? sale?.total_discount ?? 0) || 0;
    const recoveredTotalDiscount = savedDiscountField > 0
      ? savedDiscountField
      : (savedSubtotal > savedTotal ? savedSubtotal - savedTotal : 0);

    // If the report item itself carries the product offer, recover that too.
    const offerDerivedDiscount = saleItems.reduce((sum, item) => {
      const qty = Number(item?.qty ?? item?.quantity ?? 0) || 0;
      const price = Number(item?.price ?? item?.selling_price ?? item?.sellPrice ?? 0) || 0;
      if (qty <= 0 || price <= 0) return sum;
      if (String(item?.offer_type || '').toLowerCase() === 'percent') {
        const pct = Math.min(100, Math.max(0, Number(item?.offer_value) || 0));
        return sum + (price * qty * pct / 100);
      }
      if (String(item?.offer_type || '').toLowerCase() === 'b1g1') {
        const buyQty = Math.max(1, Number(item?.offer_buy_qty) || 1);
        const freeQty = Math.max(1, Number(item?.offer_free_qty) || 1);
        const free = Math.floor(qty / (buyQty + freeQty)) * freeQty;
        return sum + price * free;
      }
      return sum;
    }, 0);

    const itemOfferDiscount = Math.max(0, explicitItemDiscount, detailItemDiscount, offerDerivedDiscount);
    const billDiscount = Math.max(
      0,
      explicitBillDiscount,
      recoveredTotalDiscount - itemOfferDiscount
    );

    return {
      itemOfferDiscount,
      billDiscount,
      totalProfitLost: Math.max(0, itemOfferDiscount + billDiscount)
    };
  };

  const getAdjustedSaleProfit = (sale, itemsForSale = []) => {
    const reportItems = Array.isArray(itemsForSale) ? itemsForSale : [];
    const savedItems = Array.isArray(sale?.items) ? sale.items : [];
    const saleItems = savedItems.length ? savedItems : reportItems;
    const { itemOfferDiscount, billDiscount } = getSaleDiscountBreakdown(sale, saleItems);

    const saleTotal = Number(sale?.total ?? sale?.grandTotal ?? sale?.grand_total ?? 0) || 0;
    const explicitWholeBillDiscount = Number(sale?.billDiscountAmount ?? sale?.bill_discount_amount ?? 0) || 0;

    // Build the amount the customer would pay AFTER product/item offers,
    // but BEFORE the separate Whole Bill Discount box. This is important
    // because some report rows keep the original selling price even though
    // the saved bill total is already discounted.
    const itemFinalRevenue = saleItems.reduce((sum, item) => {
      const qty = Number(item?.qty ?? item?.quantity ?? item?.sold_qty ?? 0) || 0;
      if (qty <= 0) return sum;
      const sellPrice = Number(item?.price ?? item?.selling_price ?? item?.sellPrice ?? item?.unitPrice ?? 0) || 0;
      const subtotal = Number(item?.subtotal ?? (sellPrice * qty)) || 0;
      const hasFinalPrice = item?.finalPrice !== undefined || item?.final_price !== undefined;
      if (hasFinalPrice) return sum + Math.max(0, Number(item?.finalPrice ?? item?.final_price ?? 0) || 0);
      const itemDiscount = Number(item?.discount ?? item?.item_discount ?? 0) || 0;
      return sum + Math.max(0, subtotal - itemDiscount);
    }, 0);

    const totalBuyCost = saleItems.reduce((sum, item) => {
      const qty = Number(item?.qty ?? item?.quantity ?? item?.sold_qty ?? 0) || 0;
      let buyPrice = Number(item?.buying_price ?? item?.buyPrice ?? item?.cost_price ?? item?.cost ?? 0) || 0;
      // Some report rows do not carry buying_price. Fall back to the current
      // inventory record using product id/barcode/name so Whole Bill Discount
      // can still calculate: Final Bill Total - actual buying cost.
      if (buyPrice <= 0 && Array.isArray(products)) {
        const itemId = item?.product_id ?? item?.productId ?? item?.id;
        const itemBarcode = String(item?.barcode ?? item?.product_barcode ?? '').trim();
        const itemName = String(item?.product_name ?? item?.name ?? '').trim().toLowerCase();
        const product = products.find(p =>
          (itemId != null && String(p?.id) === String(itemId)) ||
          (itemBarcode && String(p?.barcode ?? '').trim() === itemBarcode) ||
          (itemName && String(p?.name ?? '').trim().toLowerCase() === itemName)
        );
        buyPrice = Number(product?.buying_price ?? product?.buyPrice ?? product?.cost_price ?? product?.cost ?? 0) || 0;
      }
      return sum + (buyPrice * qty);
    }, 0);

    // WHOLE BILL DISCOUNT RULE:
    // Once the separate Whole Bill Discount has reduced the bill total,
    // profit must be calculated from the FINAL BILL TOTAL, not the original
    // selling-price total. This works even when the backend/report row does
    // not preserve billDiscountAmount/subtotal fields.
    // Example: Sell 1350 - Whole Bill Discount 350 = Final Total 1000;
    // Buy cost 600 => Profit 400.
    const hasWholeBillDiscount = saleItems.length > 0 && saleTotal >= 0 && (
      explicitWholeBillDiscount > 0 ||
      billDiscount > 0 ||
      (itemFinalRevenue > 0 && saleTotal < itemFinalRevenue - 0.005)
    );
    if (hasWholeBillDiscount && totalBuyCost > 0) {
      return saleTotal - totalBuyCost;
    }

    if (!saleItems.length) {
      // New bills store the already-discounted profit. Do not subtract the
      // Whole Bill Discount a second time when no item rows are available.
      const storedProfit = Number(sale?.profit ?? sale?.grossProfit ?? sale?.gross_profit);
      const hasExplicitSavedBillDiscount = sale?.billDiscountAmount !== undefined || sale?.bill_discount_amount !== undefined;
      if (hasExplicitSavedBillDiscount && Number.isFinite(storedProfit)) return storedProfit;
      const grossProfit = Number.isFinite(storedProfit)
        ? storedProfit
        : ((saleTotal || 0) * 0.2);
      return grossProfit - itemOfferDiscount - billDiscount;
    }

    // No Whole Bill Discount: preserve the existing item-offer calculation.
    // Match report rows to Billing cart rows so saved finalPrice/discount is
    // used instead of an original selling price shown by the report table.
    const usedSaved = new Set();
    const mergedItems = saleItems.map((row) => {
      const rowKey = String(row?.id ?? row?.product_id ?? row?.barcode ?? row?.product_name ?? '');
      let saved = savedItems.find((x, idx) => {
        if (usedSaved.has(idx)) return false;
        const savedKey = String(x?.id ?? x?.product_id ?? x?.barcode ?? x?.name ?? x?.product_name ?? '');
        const sameKey = rowKey && savedKey && rowKey === savedKey;
        const rowName = String(row?.product_name ?? row?.name ?? '').trim().toLowerCase();
        const savedName = String(x?.product_name ?? x?.name ?? '').trim().toLowerCase();
        const sameName = rowName && rowName === savedName;
        return sameKey || sameName;
      });
      if (saved) usedSaved.add(savedItems.indexOf(saved));
      return saved ? {
        ...row,
        ...saved,
        __billingFinalPrice: saved.finalPrice !== undefined || saved.final_price !== undefined,
        sale_id: row.sale_id ?? saved.sale_id,
        saleId: row.saleId ?? saved.saleId,
        invoice_no: row.invoice_no ?? saved.invoice_no
      } : row;
    });

    const rawSubtotals = mergedItems.map((item) => {
      const qty = Number(item.qty ?? item.quantity ?? 0) || 0;
      const sell = Number(item.price ?? item.selling_price ?? item.sellPrice ?? 0) || 0;
      return Number(item.subtotal ?? (sell * qty)) || 0;
    });
    const subtotalTotal = rawSubtotals.reduce((a, b) => a + b, 0);
    const knownItemDiscount = mergedItems.reduce((sum, item) => sum + (Number(item.discount ?? item.item_discount ?? 0) || 0), 0);
    const missingItemDiscount = Math.max(0, itemOfferDiscount - knownItemDiscount);

    let itemProfit = 0;
    mergedItems.forEach((item, index) => {
      const qty = Number(item.qty ?? item.quantity ?? 0) || 0;
      if (qty <= 0) return;
      const buyPrice = Number(item.buying_price ?? item.buyPrice ?? item.cost_price ?? item.cost ?? 0) || 0;
      const sellPrice = Number(item.price ?? item.selling_price ?? item.sellPrice ?? 0) || 0;
      const lineSubtotal = rawSubtotals[index] || 0;
      const lineDiscount = Number(item.discount ?? item.item_discount ?? 0) || 0;
      const hasStoredFinal = item.finalPrice !== undefined || item.final_price !== undefined;
      let lineFinalPrice = Number(item.finalPrice ?? item.final_price ?? 0) || 0;
      const allocatedMissing = subtotalTotal > 0 ? missingItemDiscount * (lineSubtotal / subtotalTotal) : 0;
      if (!hasStoredFinal || !item.__billingFinalPrice) {
        lineFinalPrice = Math.max(0, lineSubtotal - lineDiscount - allocatedMissing);
      }
      itemProfit += lineFinalPrice - (buyPrice * qty);
    });

    return itemProfit - billDiscount;
  };

  const selectedDateSales = useMemo(() => {
    const rows = selectedReportDate ? salesHistory.filter(s => localDateKey(getSaleDateValue(s)) === selectedReportDate) : salesHistory;
    let sales = 0, profit = 0;
    rows.forEach(sale => {
      const originalTotal = Number(sale.total) || 0;
      const returnAmount = returnAmountByOriginalBill.get(String(sale.invoiceNo ?? sale.invoice_no ?? sale.id ?? '')) || 0;
      const netTotal = Math.max(0, originalTotal - returnAmount);
      const items = getSaleItemsForProfit(sale);
      const originalProfit = getAdjustedSaleProfit(sale, items);
      const exactReturnProfit = returnProfitImpactByOriginalBill.get(String(sale.invoiceNo ?? sale.invoice_no ?? sale.id ?? ''));
      const netProfit = exactReturnProfit !== undefined
        ? originalProfit - exactReturnProfit
        : originalProfit * (originalTotal > 0 ? netTotal / originalTotal : 1);
      sales += netTotal;
      profit += netProfit;
    });
    return { rows, sales, profit };
  }, [salesHistory, selectedReportDate, returnAmountByOriginalBill, returnProfitImpactByOriginalBill, saleItemsBySaleId]);

  const getAdjustedItemProfit = (item, sale, saleItems = []) => {
    if (!sale) return Number(item?.profit) || 0;

    const qty = Number(item?.qty ?? item?.quantity ?? 0) || 0;
    const buyPrice = Number(item?.buying_price ?? item?.buyPrice ?? item?.cost_price ?? item?.cost ?? 0) || 0;
    const sellPrice = Number(item?.price ?? item?.selling_price ?? item?.sellPrice ?? 0) || 0;
    const subtotal = Number(item?.subtotal ?? (sellPrice * qty)) || 0;
    const { itemOfferDiscount, billDiscount } = getSaleDiscountBreakdown(sale, saleItems);

    const explicitItemDiscount = Number(item?.discount ?? item?.item_discount);
    const hasItemDiscount = Number.isFinite(explicitItemDiscount) && explicitItemDiscount > 0;
    const hasFinalPrice = item?.finalPrice !== undefined || item?.final_price !== undefined;
    const knownItemDiscount = (saleItems || []).reduce((sum, row) => sum + (Number(row?.discount ?? row?.item_discount ?? 0) || 0), 0);
    const missingItemDiscount = Math.max(0, itemOfferDiscount - knownItemDiscount);
    const saleSubtotal = (saleItems || []).reduce((sum, row) => {
      const rq = Number(row?.qty ?? row?.quantity ?? 0) || 0;
      const rp = Number(row?.price ?? row?.selling_price ?? 0) || 0;
      return sum + (Number(row?.subtotal ?? (rp * rq)) || 0);
    }, 0);
    const allocatedMissing = !hasItemDiscount && !hasFinalPrice && saleSubtotal > 0
      ? missingItemDiscount * (subtotal / saleSubtotal)
      : 0;
    const finalRevenue = hasFinalPrice
      ? Math.max(0, Number(item?.finalPrice ?? item?.final_price ?? 0) || 0)
      : Math.max(0, subtotal - (hasItemDiscount ? explicitItemDiscount : 0) - allocatedMissing);

    const saleFinalRevenue = (saleItems || []).reduce((sum, row) => {
      const rq = Number(row?.qty ?? row?.quantity ?? 0) || 0;
      const rp = Number(row?.price ?? row?.selling_price ?? 0) || 0;
      const rs = Number(row?.subtotal ?? (rp * rq)) || 0;
      const rd = Number(row?.discount ?? row?.item_discount ?? 0) || 0;
      const rf = row?.finalPrice !== undefined || row?.final_price !== undefined
        ? Number(row?.finalPrice ?? row?.final_price ?? 0) || 0
        : Math.max(0, rs - rd - (saleSubtotal > 0 ? missingItemDiscount * (rs / saleSubtotal) : 0));
      return sum + rf;
    }, 0);
    const billDiscountShare = saleFinalRevenue > 0 ? billDiscount * (finalRevenue / saleFinalRevenue) : 0;
    return finalRevenue - (buyPrice * qty) - billDiscountShare;
  };

  const salesSummary = useMemo(() => {
    let todaySales = 0, todayProfit = 0;
    let monthSales = 0, monthProfit = 0;
    let yearSales = 0, yearProfit = 0;
    salesHistory.forEach((sale) => {
      const saleDateValue = getSaleDateValue(sale);
      if (!saleDateValue) return;
      const sDate = parseSaleDate(saleDateValue);
      if (!sDate) return;
      const saleKey = String(sale.invoiceNo ?? sale.invoice_no ?? sale.id ?? '');
      const returnAmount = returnAmountByOriginalBill.get(saleKey) || 0;
      const originalTotal = Number(sale.total) || 0;
      const netTotal = Math.max(0, originalTotal - returnAmount);
      const profitRatio = originalTotal > 0 ? netTotal / originalTotal : 1;
      const itemsForSale = getSaleItemsForProfit(sale);
      const originalProfit = getAdjustedSaleProfit(sale, itemsForSale);
      const exactReturnProfit = returnProfitImpactByOriginalBill.get(saleKey);
      const netProfit = exactReturnProfit !== undefined
        ? originalProfit - exactReturnProfit
        : originalProfit * profitRatio;
      if (sDate.getFullYear() === currentYear) {
        yearSales += netTotal; yearProfit += netProfit;
        if (sDate.getMonth() === currentMonth) {
          monthSales += netTotal; monthProfit += netProfit;
          if (sDate.getDate() === currentDate) { todaySales += netTotal; todayProfit += netProfit; }
        }
      }
    });
    return { todaySales, todayProfit, monthSales, monthProfit, yearSales, yearProfit };
  }, [salesHistory, saleItemsBySaleId, returnAmountByOriginalBill, returnProfitImpactByOriginalBill, currentYear, currentMonth, currentDate]);

  const { todaySales, todayProfit, monthSales, monthProfit, yearSales, yearProfit } = salesSummary;

  // PRODUCT PERFORMANCE ANALYTICS - Best Selling, Slow Moving & Never Sold
  // Indexed lookups avoid O(products × sales) scans when the inventory grows to 10,000+ items.
  const productPerformanceData = useMemo(() => {
    // Indexed lookups avoid O(products × sales) scans when the inventory grows.
    const productLookup = new Map();
    products.forEach((product) => {
      [product.id, product.product_id, product.barcode].filter(v => v !== undefined && v !== null && String(v) !== '').forEach(v => productLookup.set(String(v), product));
      const nameKey = String(product.name || '').trim().toLowerCase();
      if (nameKey) productLookup.set(`name:${nameKey}`, product);
    });
    const productPerformanceMap = new Map();
    products.forEach((product) => {
      const key = String(product.id ?? product.product_id ?? product.barcode ?? product.name);
      productPerformanceMap.set(key, {
        key, name: product.name || 'Unknown Product', barcode: product.barcode || '-', unit: product.unit || 'Pcs', totalSold: 0, revenue: 0
      });
    });

    (reportsData.sales || []).forEach((saleItem) => {
      const possibleKeys = [saleItem.product_id, saleItem.id, saleItem.barcode, saleItem.product_name]
        .filter(v => v !== undefined && v !== null).map(v => String(v));
      let matchedKey = possibleKeys.find(key => productPerformanceMap.has(key));
      if (!matchedKey) {
        const byNameProduct = productLookup.get(`name:${String(saleItem.product_name || '').trim().toLowerCase()}`);
        if (byNameProduct) matchedKey = String(byNameProduct.id ?? byNameProduct.product_id ?? byNameProduct.barcode ?? byNameProduct.name);
      }
      const key = matchedKey || String(saleItem.product_id ?? saleItem.id ?? saleItem.barcode ?? saleItem.product_name ?? 'unknown');
      if (!productPerformanceMap.has(key)) {
        productPerformanceMap.set(key, { key, name: saleItem.product_name || 'Unknown Product', barcode: saleItem.barcode || '-', unit: saleItem.unit || 'Pcs', totalSold: 0, revenue: 0 });
      }
      const record = productPerformanceMap.get(key);
      const qty = Number(saleItem.qty ?? saleItem.quantity ?? 0) || 0;
      const price = Number(saleItem.price ?? saleItem.selling_price ?? 0) || 0;
      record.totalSold += qty;
      record.revenue += qty * price;
    });

    (returnRecords || []).forEach((ret) => {
      (ret.items || []).forEach((returnedItem) => {
        const possibleKeys = [returnedItem.id, returnedItem.product_id, returnedItem.barcode, returnedItem.name]
          .filter(v => v !== undefined && v !== null).map(v => String(v));
        let matchedKey = possibleKeys.find(key => productPerformanceMap.has(key));
        if (!matchedKey) {
          const targetName = String(returnedItem.name || returnedItem.product_name || '').trim().toLowerCase();
          for (const [k, item] of productPerformanceMap.entries()) {
            if (String(item.name).trim().toLowerCase() === targetName) { matchedKey = k; break; }
          }
        }
        if (!matchedKey) return;
        const record = productPerformanceMap.get(matchedKey);
        const returnedQty = Number(returnedItem.qty ?? returnedItem.quantity ?? 0) || 0;
        const returnedPrice = Number(returnedItem.price ?? returnedItem.selling_price ?? 0) || 0;
        record.totalSold = Math.max(0, record.totalSold - returnedQty);
        record.revenue = Math.max(0, record.revenue - (returnedQty * returnedPrice));
      });
    });

    const productPerformance = [...productPerformanceMap.values()];
    const totalUnitsSold = productPerformance.reduce((sum, item) => sum + item.totalSold, 0);
    const performanceWithPercent = productPerformance.map(item => ({ ...item, salesPercentage: totalUnitsSold > 0 ? (item.totalSold / totalUnitsSold) * 100 : 0 }));
    const bestSellingProducts = performanceWithPercent.filter(item => item.totalSold > 0).sort((a, b) => b.totalSold - a.totalSold || b.revenue - a.revenue).slice(0, 20);
    const slowMovingProducts = performanceWithPercent.filter(item => item.totalSold > 0).sort((a, b) => a.totalSold - b.totalSold || a.revenue - b.revenue).slice(0, 20);
    const neverSoldProducts = performanceWithPercent.filter(item => item.totalSold === 0).sort((a, b) => String(a.name).localeCompare(String(b.name)));
    return { productPerformance, bestSellingProducts, slowMovingProducts, neverSoldProducts };
  }, [products, reportsData.sales, returnRecords]);

  const { productPerformance, bestSellingProducts, slowMovingProducts, neverSoldProducts } = productPerformanceData;

  const exportAnalyticsToExcel = () => {
    const rawSales = reportsData.sales || [];
    const saleLookup = new Map((salesHistory || []).map(sale => [String(sale.invoiceNo ?? sale.invoice_no ?? sale.id ?? ''), sale]));
    const adjustedItemProfit = (item) => {
      const saleKey = String(item.sale_id ?? item.saleId ?? '');
      const sale = saleLookup.get(saleKey);
      const saleItems = getSaleItemsForProfit(sale);
      return getAdjustedItemProfit(item, sale, saleItems);
    };
    const excelData = rawSales.map((s) => ({
      'Sale ID': s.sale_id,
      'Date': s.date,
      'Customer': s.customer_name,
      'Payment Method': s.payment_method,
      'Product Name': s.product_name,
      'Buying Price (Cost)': s.buying_price,
      'Selling Price': s.price,
      'Qty': s.qty,
      'Total Sales': (parseFloat(s.price) || 0) * (parseFloat(s.qty) || 0),
      'Profit': Number(adjustedItemProfit(s).toFixed(2))
    }));
    const worksheet = XLSX.utils.json_to_sheet(excelData);
    const bestSellingSheet = XLSX.utils.json_to_sheet(bestSellingProducts.map((item, index) => ({
      Rank: index + 1, Barcode: item.barcode, 'Product Name': item.name,
      'Total Sold Qty': item.totalSold, 'Sales Percentage': Number(item.salesPercentage.toFixed(2)),
      Revenue: Number(item.revenue.toFixed(2)), Unit: item.unit
    })));
    const slowMovingSheet = XLSX.utils.json_to_sheet(slowMovingProducts.map((item, index) => ({
      Rank: index + 1, Barcode: item.barcode, 'Product Name': item.name,
      'Total Sold Qty': item.totalSold, 'Sales Percentage': Number(item.salesPercentage.toFixed(2)),
      Revenue: Number(item.revenue.toFixed(2)), Unit: item.unit
    })));
    const neverSoldSheet = XLSX.utils.json_to_sheet(neverSoldProducts.map((item, index) => ({
      No: index + 1, Barcode: item.barcode, 'Product Name': item.name, Unit: item.unit,
      'Total Sold Qty': 0, 'Sales Percentage': 0
    })));
    const workbook = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(workbook, worksheet, 'Sales Analytics Report');
    XLSX.utils.book_append_sheet(workbook, bestSellingSheet, 'Top 20 Best Selling');
    XLSX.utils.book_append_sheet(workbook, slowMovingSheet, 'Top 20 Slow Moving');
    XLSX.utils.book_append_sheet(workbook, neverSoldSheet, 'Never Sold');
    XLSX.writeFile(workbook, `SalesReport_${new Date().toISOString().slice(0, 10)}.xlsx`);
  };

  const refreshMonthlyAnalytics = () => {
    const year = Number(selectedReportYear) || new Date().getFullYear();
    const rows = Array.from({ length: 12 }, (_, i) => ({
      monthNo: i + 1,
      month: new Date(year, i, 1).toLocaleString(undefined, { month: 'long' }),
      year,
      sales: 0,
      profit: 0
    }));
    const itemMap = new Map();
    (reportsData.sales || []).forEach(item => {
      const key = String(item.sale_id ?? item.saleId ?? '');
      if (!itemMap.has(key)) itemMap.set(key, []);
      itemMap.get(key).push(item);
    });
    (salesHistory || []).forEach(sale => {
      const d = parseSaleDate(getSaleDateValue(sale));
      if (!d || d.getFullYear() !== year) return;
      const idx = d.getMonth();
      const originalTotal = Number(sale.total) || 0;
      const returnAmt = returnAmountByOriginalBill.get(String(sale.invoiceNo ?? sale.invoice_no ?? sale.id ?? '')) || 0;
      const netSales = Math.max(0, originalTotal - returnAmt);
      const saleItems = itemMap.get(String(sale.invoiceNo ?? sale.invoice_no ?? sale.id ?? '')) || (Array.isArray(sale.items) ? sale.items : []);
      const profit = getAdjustedSaleProfit(sale, saleItems);
      const exactReturnProfit = returnProfitImpactByOriginalBill.get(String(sale.invoiceNo ?? sale.invoice_no ?? sale.id ?? ''));
      const netProfit = exactReturnProfit !== undefined
        ? profit - exactReturnProfit
        : profit * (originalTotal > 0 ? netSales / originalTotal : 1);
      rows[idx].sales += netSales;
      rows[idx].profit += netProfit;
    });
    setMonthlyAnalyticsData(rows.map(r => ({ ...r, sales: Number(r.sales.toFixed(2)), profit: Number(r.profit.toFixed(2)) })));
    showNotification('Monthly Sales & Profit Analysis යාවත්කාලීන කළා.', 'success');
  };

  const exportMonthlyAnalyticsToExcel = () => {
    const rows = monthlyAnalyticsData.length ? monthlyAnalyticsData : (() => {
      const year = Number(selectedReportYear) || new Date().getFullYear();
      const rows = Array.from({ length: 12 }, (_, i) => ({ monthNo: i + 1, month: new Date(year, i, 1).toLocaleString(undefined, { month: 'long' }), year, sales: 0, profit: 0 }));
      const itemMap = new Map();
      (reportsData.sales || []).forEach(item => { const key = String(item.sale_id ?? item.saleId ?? ''); if (!itemMap.has(key)) itemMap.set(key, []); itemMap.get(key).push(item); });
      (salesHistory || []).forEach(sale => {
        const d = new Date(sale.date); if (isNaN(d.getTime()) || d.getFullYear() !== year) return;
        const originalTotal = Number(sale.total) || 0; const returnAmt = returnAmountByOriginalBill.get(String(sale.invoiceNo ?? sale.invoice_no ?? sale.id ?? '')) || 0;
        const netSales = Math.max(0, originalTotal - returnAmt); const saleItems = itemMap.get(String(sale.invoiceNo ?? sale.invoice_no ?? sale.id ?? '')) || (Array.isArray(sale.items) ? sale.items : []);
        const profit = getAdjustedSaleProfit(sale, saleItems);
        const exactReturnProfit = returnProfitImpactByOriginalBill.get(String(sale.invoiceNo ?? sale.invoice_no ?? sale.id ?? ''));
        const netProfit = exactReturnProfit !== undefined ? profit - exactReturnProfit : profit * (originalTotal > 0 ? netSales / originalTotal : 1);
        rows[d.getMonth()].sales += netSales; rows[d.getMonth()].profit += netProfit;
      });
      return rows.map(r => ({ ...r, sales: Number(r.sales.toFixed(2)), profit: Number(r.profit.toFixed(2)) }));
    })();
    const data = rows.map(r => ({ Year: r.year, Month: r.month, 'Sales (Rs.)': r.sales, 'Profit (Rs.)': r.profit }));
    if (!data.length) return;
    const ws = XLSX.utils.json_to_sheet(data);
    ws['!cols'] = [{wch:10},{wch:18},{wch:18},{wch:18}];
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, 'Monthly Sales Profit');
    XLSX.writeFile(wb, `Monthly_Sales_Profit_${selectedReportYear}.xlsx`);
    showNotification('Monthly Sales & Profit Excel Report බාගත කළා.', 'success');
  };

  const handleDeleteMonthlyAnalytics = () => {
    if (!monthlyAnalyticsData.length) {
      showNotification('Delete කිරීමට Monthly Sales & Profit Analysis data නැහැ.', 'info');
      return;
    }
    setMonthlyAnalyticsData([]);
    localStorage.removeItem('pos_monthly_analytics');
    showNotification('🗑️ Monthly Sales & Profit Analysis table එක මකා දමන ලදී.', 'success');
  };

  useEffect(() => {
    if (salesHistory.length > 0 && monthlyAnalyticsData.length === 0) {
      refreshMonthlyAnalytics();
    }
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [salesHistory, reportsData.sales, returnRecords]);

  // ==================== ADVANCED REPORT PAGE DATA ====================
  // Each advanced page is calculated locally from the same live Billing/Return/
  // Inventory/Supplier data already used by the existing reports.  Defaults are
  // intentionally kept as arrays so an empty database never causes a render error.
  const advancedReportData = useMemo(() => {
    const sales = Array.isArray(salesHistory) ? salesHistory : [];
    const items = Array.isArray(reportsData?.sales) ? reportsData.sales : [];
    const returns = Array.isArray(returnRecords) ? returnRecords : [];
    const orders = Array.isArray(shopOrderHistory) ? shopOrderHistory : [];
    const requests = Array.isArray(orderRequests) ? orderRequests : [];

    const amountOf = (sale) => Number(sale?.total ?? sale?.grandTotal ?? sale?.amount ?? 0) || 0;
    const dateOf = (sale) => {
      const d = new Date(sale?.date ?? sale?.created_at ?? sale?.createdAt ?? '');
      return Number.isNaN(d.getTime()) ? null : d;
    };
    const labelDate = (d) => d ? d.toISOString().slice(0, 10) : 'Unknown';
    const addToMap = (map, key, value) => map.set(key, (map.get(key) || 0) + (Number(value) || 0));

    const dailyMap = new Map();
    const paymentMap = new Map();
    const customerMap = new Map();
    const cashierMap = new Map();
    const discountMap = new Map();
    sales.forEach((sale) => {
      const total = amountOf(sale);
      const saleKey = String(sale?.invoiceNo ?? sale?.invoice_no ?? sale?.id ?? '');
      const returnAmt = returnAmountByOriginalBill.get(saleKey) || 0;
      const netSaleTotal = Math.max(0, total - returnAmt);
      const d = dateOf(sale);
      addToMap(dailyMap, labelDate(d), netSaleTotal);
      addToMap(paymentMap, String(sale?.paymentMethod ?? sale?.payment_method ?? 'Cash'), netSaleTotal);
      addToMap(customerMap, String(sale?.customerName ?? sale?.customer_name ?? 'General Customer'), netSaleTotal);
      addToMap(cashierMap, String(sale?.cashierName ?? sale?.cashier ?? sale?.cashier_name ?? sale?.username ?? 'Cashier'), netSaleTotal);
      const saleItems = getSaleItemsForProfit(sale);
      const { itemOfferDiscount: itemDiscount, billDiscount, totalProfitLost } = getSaleDiscountBreakdown(sale, saleItems);
      if (itemDiscount) addToMap(discountMap, 'Item Offers', itemDiscount);
      if (billDiscount) addToMap(discountMap, 'Whole Bill Discount', billDiscount);
      if (!itemDiscount && !billDiscount && totalProfitLost) addToMap(discountMap, 'Discount', totalProfitLost);
    });

    const toRows = (map, firstKey, secondKey) => [...map.entries()]
      .sort((a, b) => b[1] - a[1])
      .map(([name, value]) => ({ [firstKey]: name, [secondKey]: Number(value.toFixed(2)) }));
    const toChart = (map) => [...map.entries()]
      .sort((a, b) => b[1] - a[1])
      .map(([name, value]) => ({ name, value: Number(value.toFixed(2)) }));

    const dailySalesReport = toRows(dailyMap, 'Date', 'Sales');
    const paymentMethodReport = toRows(paymentMap, 'Payment Method', 'Sales');
    const customerSalesReport = toRows(customerMap, 'Customer', 'Sales');
    const discountReport = toRows(discountMap, 'Discount Type', 'Discount');
    const cashierSalesReport = toRows(cashierMap, 'Cashier', 'Sales');
    const profitLostReport = sales.map((sale) => {
      const saleKey = String(sale?.invoiceNo ?? sale?.invoice_no ?? sale?.id ?? '');
      const saleItems = getSaleItemsForProfit(sale);
      const { itemOfferDiscount, billDiscount, totalProfitLost } = getSaleDiscountBreakdown(sale, saleItems);
      return {
        'Invoice No': sale?.invoiceNo ?? sale?.invoice_no ?? sale?.id ?? '-',
        'Date': sale?.date ?? sale?.created_at ?? sale?.createdAt ?? '-',
        'Item Offers Lost': Number(itemOfferDiscount.toFixed(2)),
        'Whole Bill Discount Lost': Number(billDiscount.toFixed(2)),
        'Total Profit Lost': Number(totalProfitLost.toFixed(2))
      };
    }).filter(row => row['Total Profit Lost'] > 0).reverse();
    const profitLostChart = [
      { name: 'Item Offers', value: Number(profitLostReport.reduce((sum, row) => sum + Number(row['Item Offers Lost'] || 0), 0).toFixed(2)) },
      { name: 'Whole Bill Discount', value: Number(profitLostReport.reduce((sum, row) => sum + Number(row['Whole Bill Discount Lost'] || 0), 0).toFixed(2)) }
    ].filter(row => row.value > 0);

    const returnReasonMap = new Map();
    returns.forEach((r) => addToMap(returnReasonMap, String(r?.reason ?? 'Other'), r?.total));
    const returnReport = toRows(returnReasonMap, 'Reason', 'Return Total');

    const creditMap = new Map();
    sales.forEach((sale) => {
      const outstanding = Number(sale?.creditAmount ?? sale?.outstandingBalance ?? 0) || 0;
      if (outstanding > 0) addToMap(creditMap, String(sale?.customerName ?? sale?.customer_name ?? 'General Customer'), outstanding);
    });
    const creditReport = [...creditMap.entries()].sort((a, b) => b[1] - a[1]).map(([name, outstanding]) => ({ name, outstanding: Number(outstanding.toFixed(2)) }));

    const supplierMap = new Map();
    [...orders, ...requests].forEach((order) => {
      const supplier = String(order?.supplierName ?? order?.supplier ?? order?.supplier_name ?? 'Supplier');
      const total = Number(order?.total ?? order?.amount ?? order?.orderTotal ?? 0) || 0;
      const qty = Number(order?.qty ?? order?.quantity ?? order?.orderQty ?? 0) || 0;
      const old = supplierMap.get(supplier) || { orders: 0, amount: 0, qty: 0 };
      old.orders += 1; old.amount += total; old.qty += qty;
      supplierMap.set(supplier, old);
    });
    const supplierReport = [...supplierMap.entries()].map(([supplier, v]) => ({ Supplier: supplier, 'Orders': v.orders, 'Amount': Number(v.amount.toFixed(2)), 'Qty': v.qty })).sort((a, b) => b.Amount - a.Amount || b.Orders - a.Orders);

    const stockUnitMap = new Map();
    (Array.isArray(products) ? products : []).forEach((product) => {
      const unit = String(product?.unit ?? 'Pcs');
      addToMap(stockUnitMap, unit, Number(product?.stock ?? 0));
    });
    const inventoryCategoryReport = toChart(stockUnitMap);

    const reportSalesRows = sales.map((sale) => {
      const gross = amountOf(sale);
      return { _id: sale?.invoiceNo ?? sale?.invoice_no ?? sale?.id ?? '-', date: sale?.date ?? sale?.created_at ?? '-', customerName: sale?.customerName ?? sale?.customer_name ?? 'Walk-in', paymentMethod: sale?.paymentMethod ?? sale?.payment_method ?? 'Cash', _gross: gross, _returns: 0, _net: gross };
    });

    const dailyChart = toChart(dailyMap);
    const paymentChart = toChart(paymentMap);
    const customerChart = toChart(customerMap);
    const discountChart = toChart(discountMap);
    const cashierChart = toChart(cashierMap);
    const returnChart = toChart(returnReasonMap);
    const supplierChart = supplierReport.map((r) => ({ name: r.Supplier, value: r.Amount || r.Orders }));

    return { dailySalesReport, paymentMethodReport, customerSalesReport, discountReport, profitLostReport, returnReport, creditReport, cashierSalesReport, supplierReport, inventoryCategoryReport, reportSalesRows, dailyChart, paymentChart, customerChart, discountChart, profitLostChart, cashierChart, returnChart, supplierChart, items };
  }, [salesHistory, reportsData?.sales, returnRecords, returnAmountByOriginalBill, returnProfitImpactByOriginalBill, shopOrderHistory, orderRequests, products, saleItemsBySaleId]);

  const { dailySalesReport, paymentMethodReport, customerSalesReport, discountReport, profitLostReport, returnReport, creditReport, cashierSalesReport, supplierReport, inventoryCategoryReport, reportSalesRows, dailyChart, paymentChart, customerChart, discountChart, profitLostChart, cashierChart, returnChart, supplierChart } = advancedReportData;

  // ==================== ADVANCED REPORT DASHBOARD HELPERS ====================
  // These reports are live views of Billing/Inventory/Customer/Return data.
  const reportRowsByPage = (page) => {
    if (page === 'sales') return reportSalesRows.map(s => ({'Invoice No':s._id,Date:s.date,Customer:s.customerName ?? s.customer_name ?? 'Walk-in',Payment:s.paymentMethod ?? s.payment_method ?? 'Cash','Gross Sales':s._gross,Returns:s._returns,'Net Sales':s._net}));
    if (page === 'daily') return dailySalesReport;
    if (page === 'payments') return paymentMethodReport;
    if (page === 'customers') return customerSalesReport;
    if (page === 'discounts') return discountReport;
    if (page === 'profitlost') return profitLostReport;
    if (page === 'returns') return (returnRecords||[]).map(r=>({'Return ID':r.id,'Bill No':r.originalBillNo,Date:r.date,Reason:r.reason,Total:r.total}));
    if (page === 'credit') return creditReport;
    if (page === 'cashier') return cashierSalesReport;
    if (page === 'suppliers') return supplierReport;
    if (page === 'stock') return (products||[]).map(p=>({'Product Name':p.name,Barcode:p.barcode,Stock:Number(p.stock||0),Unit:p.unit}));
    if (page === 'products') return productPerformance.map(p=>({'Product Name':p.name,Barcode:p.barcode,'Sold Qty':p.totalSold,Revenue:p.revenue,Unit:p.unit}));
    if (page === 'lowstock') return (reportsData.lowStock||[]).map(p=>({'Product Name':p.name??p.product_name,Barcode:p.barcode,Stock:p.stock,'Alert Limit':p.min_stock_alert??5}));
    return [{TodaySales:todaySales,TodayProfit:todayProfit,MonthSales:monthSales,MonthProfit:monthProfit,YearSales:yearSales,YearProfit:yearProfit}];
  };
  const reportChartData = (page) => {
    if (page==='payments') return paymentChart;
    if (page==='customers') return customerChart;
    if (page==='discounts') return discountChart;
    if (page==='profitlost') return profitLostChart;
    if (page==='returns') return returnChart;
    if (page==='credit') return creditReport.map(x=>({name:x.name,value:x.outstanding}));
    if (page==='cashier') return cashierChart;
    if (page==='suppliers') return supplierChart;
    if (page==='stock') return inventoryCategoryReport;
    if (page==='products') return bestSellingProducts.map(x=>({name:x.name,value:x.totalSold}));
    if (page==='lowstock') return (reportsData.lowStock||[]).slice(0,20).map(x=>({name:x.name??x.product_name??'Product',value:Number(x.stock||0)}));
    if (page==='daily') return dailySalesReport;
    if (page==='monthly') return monthlyAnalyticsData.map(x=>({name:x.month,value:x.sales}));
    return [{name:'Today',value:todaySales},{name:'Month',value:monthSales},{name:'Year',value:yearSales}];
  };
  const reportChartTitle = {overview:'Sales Summary',sales:'Sales by Date',daily:'Daily Sales',monthly:'Monthly Sales',products:'Best Selling Products',lowstock:'Remaining Stock',payments:'Payment Methods',customers:'Customer Sales',discounts:'Discount Types',profitlost:'GIVE PROFIT LOST',returns:'Returns by Reason',credit:'Outstanding Credit',cashier:'Cashier Sales',suppliers:'Supplier Orders',stock:'Stock Distribution'};
  // Compatibility aliases used by the report dashboard UI.
  const reportPageTitle = reportChartTitle;
  const reportPageRows = reportRowsByPage;
  const chartPalette=['#2563eb','#7c3aed','#059669','#f59e0b','#dc2626','#0891b2','#db2777','#65a30d','#9333ea','#ea580c'];
  const BarChart=({data=[],title='Bar Chart',valuePrefix='Rs. '})=>{const rows=data.slice(0,12),max=Math.max(...rows.map(x=>Number(x.value)||0),1);return <div style={{background:'#fff',border:'1px solid #e5e7eb',borderRadius:'10px',padding:'14px',overflowX:'auto'}}><h3 style={{margin:'0 0 10px'}}>📊 {title}</h3>{rows.length?<div style={{minWidth:'560px',height:'280px',display:'flex',alignItems:'flex-end',gap:'10px',padding:'10px 10px 42px',borderBottom:'2px solid #94a3b8'}}>{rows.map((r,i)=><div key={`${r.name}-${i}`} style={{flex:1,minWidth:'34px',height:'100%',display:'flex',alignItems:'flex-end',justifyContent:'center',position:'relative'}}><div title={`${r.name}: ${valuePrefix}${Number(r.value||0).toFixed(2)}`} style={{width:'70%',maxWidth:'42px',height:`${Math.max(4,(Number(r.value)||0)/max*220)}px`,background:chartPalette[i%chartPalette.length],borderRadius:'5px 5px 0 0'}}/><span style={{position:'absolute',bottom:'-38px',fontSize:'10px',maxWidth:'58px',overflow:'hidden',textOverflow:'ellipsis',whiteSpace:'nowrap'}}>{r.name}</span></div>)}</div>:<div style={{padding:'70px 10px',textAlign:'center',color:'#64748b'}}>No data yet.</div>}</div>};
  const LineChart=({data=[],title='Line Chart',valuePrefix='Rs. '})=>{const rows=data.slice(-20),w=760,h=250,pad=35,max=Math.max(...rows.map(x=>Number(x.value)||0),1),points=rows.map((r,i)=>`${pad+(rows.length===1?0:i*(w-2*pad)/(rows.length-1))},${h-pad-(Number(r.value)||0)/max*(h-2*pad)}`).join(' ');return <div style={{background:'#fff',border:'1px solid #e5e7eb',borderRadius:'10px',padding:'14px',overflowX:'auto'}}><h3 style={{margin:'0 0 10px'}}>📈 {title}</h3>{rows.length?<svg viewBox={`0 0 ${w} ${h}`} style={{width:'100%',minWidth:'600px',height:'250px'}}><polyline points={points} fill="none" stroke="#2563eb" strokeWidth="4"/>{rows.map((r,i)=>{const x=pad+(rows.length===1?0:i*(w-2*pad)/(rows.length-1)),y=h-pad-(Number(r.value)||0)/max*(h-2*pad);return <g key={`${r.name}-${i}`}><circle cx={x} cy={y} r="5" fill="#7c3aed"/><text x={x} y={h-8} textAnchor="middle" fontSize="9">{String(r.name).slice(0,9)}</text><title>{r.name}: {valuePrefix}{Number(r.value||0).toFixed(2)}</title></g>})}</svg>:<div style={{padding:'70px 10px',textAlign:'center',color:'#64748b'}}>No data yet.</div>}</div>};
  const PieChart=({data=[],title='Pie Chart'})=>{const rows=data.filter(x=>Number(x.value)>0).slice(0,10),total=rows.reduce((a,x)=>a+Number(x.value||0),0);let angle=-90;const polar=(cx,cy,r,a)=>[cx+r*Math.cos(a*Math.PI/180),cy+r*Math.sin(a*Math.PI/180)];const arcs=rows.map((r,i)=>{const start=angle,sweep=Number(r.value)/total*360;angle+=sweep;const end=angle,[x1,y1]=polar(120,120,90,start),[x2,y2]=polar(120,120,90,end);return {r,i,pct:Number(r.value)/total*100,path:`M 120 120 L ${x1} ${y1} A 90 90 0 ${sweep>180?1:0} 1 ${x2} ${y2} Z`};});return <div style={{background:'#fff',border:'1px solid #e5e7eb',borderRadius:'10px',padding:'14px'}}><h3 style={{margin:'0 0 10px'}}>🥧 {title}</h3>{rows.length?<div style={{display:'flex',alignItems:'center',gap:'18px',flexWrap:'wrap'}}><svg viewBox="0 0 240 240" style={{width:'230px',height:'230px'}}>{arcs.map(a=><path key={a.i} d={a.path} fill={chartPalette[a.i%chartPalette.length]} stroke="#fff" strokeWidth="2"><title>{a.r.name}: {a.pct.toFixed(1)}%</title></path>)}</svg><div>{arcs.map(a=><div key={a.i} style={{fontSize:'12px',marginBottom:'5px'}}><span style={{display:'inline-block',width:'10px',height:'10px',background:chartPalette[a.i%chartPalette.length],marginRight:'6px'}}/>{a.r.name}: {a.pct.toFixed(1)}%</div>)}</div></div>:<div style={{padding:'70px 10px',textAlign:'center',color:'#64748b'}}>No data yet.</div>}</div>};
  const reportTableStyle={width:'100%',borderCollapse:'collapse'},reportThStyle={padding:'9px',background:'#f1f5f9',textAlign:'left'},reportTdStyle={padding:'9px',borderBottom:'1px solid #e5e7eb'};
  const downloadReportExcel=(name,sheets)=>{const wb=XLSX.utils.book_new();if(Array.isArray(sheets)){XLSX.utils.book_append_sheet(wb,XLSX.utils.json_to_sheet(sheets),'Report');}else{Object.entries(sheets||{}).forEach(([sheet,rows])=>XLSX.utils.book_append_sheet(wb,XLSX.utils.json_to_sheet(Array.isArray(rows)?rows:[]),String(sheet).slice(0,31)));}XLSX.writeFile(wb,`${name}_${new Date().toISOString().slice(0,10)}.xlsx`);showNotification(`${name} Excel report බාගත කළා.`,'success');};
  const markReportDeleted=(key,name=reportPageTitle[key]||'Report')=>{setReportDeletedFlags(prev=>({...prev,[key]:true}));showNotification(`🗑️ ${name} report data view clear කළා. Billing data එකට බලපෑමක් නැහැ.`,'success');};
  const refreshReportPage=(key)=>{setReportDeletedFlags(prev=>({...prev,[key]:false}));fetchSales();fetchProducts();showNotification('Report data refresh කළා.','success');};


  const exportAnalyticsToPDF = () => {
    const rawSales = reportsData.sales || [];
    const saleLookup = new Map((salesHistory || []).map(sale => [String(sale.invoiceNo ?? sale.invoice_no ?? sale.id ?? ''), sale]));
    const adjustedItemProfit = (item) => {
      const saleKey = String(item.sale_id ?? item.saleId ?? '');
      const sale = saleLookup.get(saleKey);
      const saleItems = getSaleItemsForProfit(sale);
      return getAdjustedItemProfit(item, sale, saleItems);
    };
    const doc = new jsPDF();
    doc.text('Sales & Profit Analytics Summary Report', 14, 15);
    doc.setFontSize(10);
    doc.text(`Today Sales: Rs. ${todaySales.toFixed(2)} | Profit: Rs. ${todayProfit.toFixed(2)}`, 14, 22);
    doc.text(`This Month Sales: Rs. ${monthSales.toFixed(2)} | Profit: Rs. ${monthProfit.toFixed(2)}`, 14, 28);
    doc.text(`This Year Sales: Rs. ${yearSales.toFixed(2)} | Profit: Rs. ${yearProfit.toFixed(2)}`, 14, 34);

    const tableColumn = ['Sale ID', 'Date', 'Product', 'Buy Price', 'Sell Price', 'Qty', 'Total', 'Profit'];
    const tableRows = rawSales.map((s) => [
      `#${s.sale_id}`,
      new Date(s.date).toLocaleDateString(),
      s.product_name,
      `Rs. ${s.buying_price}`,
      `Rs. ${s.price}`,
      s.qty,
      `Rs. ${(s.price * s.qty).toFixed(2)}`,
      `Rs. ${adjustedItemProfit(s).toFixed(2)}`
    ]);
    doc.autoTable({ head: [tableColumn], body: tableRows, startY: 40 });
    doc.save(`SalesReport_${new Date().toISOString().slice(0, 10)}.pdf`);
  };

  const exportToExcel = () => {
    const excelData = products.map((p) => ({
      Barcode: p.barcode,
      'Product Name': p.name,
      Lot: p.lot_number || '',
      'Buying Price': p.buying_price || 0,
      'Selling Price': p.price,
      Stock: p.stock,
      'Low Stock Alert Threshold': p.min_stock_alert || 5,
      Unit: p.unit,
      'Offer Type': p.offer_type,
      'Offer Value': p.offer_value,
      'Offer Buy Qty': p.offer_buy_qty || 1,
      'Offer Free Qty': p.offer_free_qty || 1,
      'Bulk Min Qty': p.bulk_min_qty || 2,
      'Bulk Price': p.bulk_price || 0,
      'Expiry Date': p.expiry_date,
      'Expiry Status': getExpiryInfo(p).status,
      'Days Remaining': getExpiryInfo(p).daysRemaining,
      Supplier: p.supplier || ''
    }));
    const worksheet = XLSX.utils.json_to_sheet(excelData);
    const workbook = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(workbook, worksheet, 'Products');
    XLSX.writeFile(workbook, 'Product_Inventory_Report.xlsx');
  };

  const exportToPDF = () => {
    const doc = new jsPDF();
    doc.text('Product Inventory Stock Report', 14, 15);
    const tableColumn = ['Barcode', 'Name', 'Buy Price', 'Retail Price', 'Wholesale Price', 'Special Price', 'Stock', 'Supplier'];
    const tableRows = products.map((p) => [
      p.barcode, p.name, `Rs. ${p.buying_price || 0}.00`, `Rs. ${p.price}.00`,
      `${p.stock} ${p.unit}`, p.supplier || ''
    ]);
    doc.autoTable({ head: [tableColumn], body: tableRows, startY: 20 });
    doc.save('Product_Inventory_Report.pdf');
  };

  const downloadCustomerBillPDF = (sale) => {
    const doc = new jsPDF();
    doc.text(receiptSettings.shopName, 14, 15);
    doc.setFontSize(10);
    doc.text(receiptSettings.address, 14, 21);
    doc.text(`Tel: ${receiptSettings.phone}`, 14, 27);
    doc.text(`Invoice No: #${getSaleInvoiceNo(sale)}`, 14, 37);
    doc.text(`Customer Name: ${sale.customer_name}`, 14, 43);
    doc.text(`Phone: ${sale.customer_phone}`, 14, 49);
    doc.text(`Date & Time: ${(parseSaleDate(getSaleDateValue(sale)) || new Date()).toLocaleString()}`, 14, 55);
    doc.text(`Payment Method: ${sale.payment_method || sale.paymentMethod || '-'}`, 14, 61);

    const saleItems = (reportsData.sales || []).filter(r => String(r.sale_id) === String(getSaleInvoiceNo(sale)));
    const tableColumn = ['Item Name', 'Unit Price', 'Qty', 'Offer', 'Discount', 'Total'];
    const tableRows = saleItems.map(item => {
      const unitPrice = Number(item.unitPrice ?? item.price ?? item.selling_price ?? 0) || 0;
      const qty = Number(item.qty ?? item.quantity ?? 0) || 0;
      const lineSubtotal = Number(item.subtotal ?? (unitPrice * qty)) || 0;
      const lineDiscount = Number(item.discount ?? 0) || 0;
      const lineTotal = Number(item.finalPrice ?? item.final_price ?? (lineSubtotal - lineDiscount)) || 0;
      const offer = item.offerLabel || (item.offer_type === 'percent' ? `${item.offer_value || 0}% OFF` : item.offer_type === 'b1g1' ? `BUY ${item.offer_buy_qty || (Number(item.offer_value) >= 1001 ? Math.floor(Number(item.offer_value) / 1000) : 1)} GET ${item.offer_free_qty || (Number(item.offer_value) >= 1001 ? Number(item.offer_value) % 1000 : 1)} FREE` : item.offer_type === 'bulk_price' ? `Rs. ${item.bulk_price || 0} EACH FROM ${item.bulk_min_qty || 2}` : '');
      return [item.product_name || item.name || '-', unitPrice.toFixed(2), qty, offer, lineDiscount.toFixed(2), lineTotal.toFixed(2)];
    });
    doc.autoTable({ head: [tableColumn], body: tableRows, startY: 67, styles: { fontSize: 7 } });
    const finalY = doc.lastAutoTable.finalY || 80;
    const saleDiscount = Number(sale.discount ?? sale.total_discount ?? 0) || 0;
    const saleSubtotal = Number(sale.subtotal ?? ((Number(sale.total) || 0) + saleDiscount)) || 0;
    doc.text(`Subtotal: Rs. ${saleSubtotal.toFixed(2)}`, 14, finalY + 10);
    doc.text(`Discount: -Rs. ${saleDiscount.toFixed(2)}`, 14, finalY + 16);
    doc.text(`Grand Total: Rs. ${(parseFloat(sale.total) || 0).toFixed(2)}`, 14, finalY + 22);
    doc.text(receiptSettings.footerMsg, 14, finalY + 18);
    if (receiptSettings.footerComment) {
      const commentLines = doc.splitTextToSize(String(receiptSettings.footerComment), 75);
      doc.text(commentLines, 14, finalY + 24);
    }
    doc.save(`Customer_Invoice_${getSaleInvoiceNo(sale)}.pdf`);
  };

  const currentFormattedDate = () => {
    const d = new Date();
    const day = String(d.getDate()).padStart(2, '0');
    const month = String(d.getMonth() + 1).padStart(2, '0');
    const year = d.getFullYear();
    const hours = String(d.getHours()).padStart(2, '0');
    const mins = String(d.getMinutes()).padStart(2, '0');
    return `${day}/${month}/${year}, ${hours}:${mins}`;
  };

  const activeSaleObj = completedSale || {
    invoiceNo: salesHistory.length > 0 ? Number(getSaleInvoiceNo(salesHistory[0])) + 1 : 141,
    items: cart,
    subtotal: cartSubtotal,
    discount: cartDiscount,
    total,
    paymentMethod,
    paidAmount: numericPaid || total,
    changeAmount,
    customerName: customerName || 'General Customer',
    customerPhone: customerPhone || '-',
    dateFormatted: currentFormattedDate()
  };

  // SHOP REGISTRATION HANDLER
  const handleRegisterSubmit = (e) => {
    e.preventDefault();
    if (regSecretKey.trim() !== MASTER_SECRET_KEY) {
      setRegError(t.errorSecretKey);
      return;
    }
    const shopData = {
      shopName: regShopName,
      email: regEmail,
      address: regAddress,
      phone: regPhone,
      adminUsername: regUsername,
      adminPassword: regPassword,
      logoUrl: regLogoUrl || ''
    };
    localStorage.setItem('pos_shop_account', JSON.stringify(shopData));
    
    setReceiptSettings(prev => ({
      ...prev,
      shopName: regShopName,
      address: regAddress,
      phone: regPhone,
      email: regEmail,
      logoUrl: regLogoUrl || prev.logoUrl || ''
    }));

    setIsRegistered(true);
    setAuthMode('login');
    showNotification(t.successReg);
  };

  // PASSWORD RESET HANDLER
  const handleResetSubmit = (e) => {
    e.preventDefault();
    if (resetSecretKey.trim() !== MASTER_SECRET_KEY) {
      setResetError('වැරදි Secret Key එකකි! (Invalid Secret Key)');
      return;
    }

    const savedShopStr = localStorage.getItem('pos_shop_account');
    let adminU = 'admin';
    let adminP = '123';
    if (savedShopStr) {
      try {
        const s = JSON.parse(savedShopStr);
        adminU = s.adminUsername || 'admin';
        adminP = s.adminPassword || '123';

        if (resetUsername.trim() === adminU || resetUsername.trim() === 'admin') {
          s.adminUsername = newAdminUsername.trim() || adminU;
          s.adminPassword = newAdminPassword.trim() || adminP;
          localStorage.setItem('pos_shop_account', JSON.stringify(s));
          setResetSuccess('මුරපදය සාර්ථකව වෙනස් කරන ලදී! (Password Reset Successful)');
          setResetError('');
          setTimeout(() => {
            setAuthMode('login');
            setResetSuccess('');
          }, 2000);
          return;
        }
      } catch (err) {}
    }

    const cashierIndex = cashiers.findIndex(c => c.username === resetUsername.trim());
    if (cashierIndex !== -1) {
      const updatedCashiers = [...cashiers];
      updatedCashiers[cashierIndex].username = newAdminUsername.trim() || updatedCashiers[cashierIndex].username;
      updatedCashiers[cashierIndex].password = newAdminPassword.trim() || updatedCashiers[cashierIndex].password;
      setCashiers(updatedCashiers);
      setResetSuccess('අයකැමි ගිණුමේ මුරපදය සාර්ථකව වෙනස් කරන ලදී!');
      setResetError('');
      setTimeout(() => {
        setAuthMode('login');
        setResetSuccess('');
      }, 2000);
      return;
    }

    setResetError('ඇතුළත් කළ Username එක හමු නොවීය! (Username not found)');
  };

  // HOME DASHBOARD NAVIGATION
  // Existing module/page code is intentionally left unchanged.
  const goToInventorySection = (section) => {
    setInventoryShortcutMode(section);
    setActiveTab('inventory');
    window.setTimeout(() => {
      const targetId = section === 'offers' ? 'inventory-product-form' : 'inventory-table-section';
      document.getElementById(targetId)?.scrollIntoView({ behavior: 'smooth', block: 'start' });
    }, 50);
  };

  // LOGIN HANDLER + branded transition into Home Dashboard
  const enterPoshithaSystem = (role) => {
    setCurrentUserRole(role);
    setInventoryShortcutMode(null);
    setActiveTab((dashboardStyle === 'home' || dashboardStyle === 'both') ? 'home' : 'pos');
    setLoginError('');
    setIsLoggedIn(true);
    setShowWelcomeAnimation(true);
    window.setTimeout(() => setShowWelcomeAnimation(false), 7000);
  };

  // LOGIN HANDLER
  const handleLoginSubmit = (e) => {
    e.preventDefault();
    const savedShopStr = localStorage.getItem('pos_shop_account');
    let adminU = 'admin';
    let adminP = '123';
    if (savedShopStr) {
      try {
        const s = JSON.parse(savedShopStr);
        adminU = s.adminUsername || 'admin';
        adminP = s.adminPassword || '123';
      } catch (err) {}
    }

    if (loginUsername.trim() === adminU && loginPassword === adminP) {
      localStorage.setItem('pos_last_login', JSON.stringify({
        username: loginUsername.trim(),
        password: loginPassword,
        role: 'admin',
        savedAt: new Date().toISOString()
      }));
      enterPoshithaSystem('admin');
      return;
    }

    const foundCashier = cashiers.find(c => c.username === loginUsername.trim() && c.password === loginPassword);
    if (foundCashier) {
      localStorage.setItem('pos_last_login', JSON.stringify({
        username: loginUsername.trim(),
        password: loginPassword,
        role: `cashier${foundCashier.id}`,
        savedAt: new Date().toISOString()
      }));
      enterPoshithaSystem(`cashier${foundCashier.id}`);
      return;
    }

    setLoginError(t.errorInvalidLogin);
  };

  // SHOP TERMS & CONDITIONS — displayed/exported in the currently selected language.
  const shopConditions = {
    en: [
      '1. Daily backup must be completed immediately when the shop is closed.',
      '2. Your software Username and Password must be kept safely and must not be shared with unauthorized persons.',
      '3. The Username is especially important and must be kept safely for future login and recovery purposes.',
      '4. The software password cannot be changed casually by the user; keep the registered credentials secure.',
      '5. After installing the software, up to 3 free software updates may be obtained through our service.',
      '6. After the 3 free updates are used, each additional update service is charged at Rs. 1,000.',
      '7. After one year, if you need to clear/reset application data or cache to improve software speed, this service can be performed on request.',
      '8. We recommend a data/cache cleanup once every 1 to 1.5 years when required for performance maintenance.',
      '9. If a cleanup is requested, always keep a recent backup before the service is performed.',
      '10. Optional annual / 1.5-year cleanup service is charged at Rs. 5,000 when requested.',
      '11. Keep backup files in a safe location. The shop is responsible for protecting its login credentials and backup files.',
      '12. Billing, Inventory, Returns, Suppliers, Shop Order History, Customer Dashboard, Reports & Dashboard, Monthly Sales Report, Shop Profile, Offers and Barcode features remain subject to the installed software version and service terms.'
    ],
    si: [
      '1. සෑම දිනකම කඩය වසා දැමූ මොහොතේම Daily Backup එක අනිවාර්යයෙන් සිදු කළ යුතුය.',
      '2. ඔබගේ Software Username සහ Password ආරක්ෂිතව තබා ගත යුතු අතර අනවසර පුද්ගලයන් සමඟ බෙදා නොගත යුතුය.',
      '3. විශේෂයෙන් Username එක අනිවාර්යයෙන් ආරක්ෂිතව මතක තබා ගත යුතුය. එය Login සහ Recovery කටයුතු සඳහා වැදගත් වේ.',
      '4. Software එක තුළ Password එක අහඹුවෙන් වෙනස් නොකළ යුතු අතර ලියාපදිංචි Login තොරතුරු ආරක්ෂිතව තබා ගත යුතුය.',
      '5. Software එක Install කරගත් පසු අප හරහා Free Updates 3ක් ලබාගත හැක.',
      '6. Free Updates 3 අවසන් වූ පසු සිදු කරන සෑම අමතර Update Service එකක් සඳහා Rs. 1,000ක මුදලක් අය කෙරේ.',
      '7. වසරකට පසු Software එකේ වේගය වැඩි කිරීම සඳහා Application Data / Cache Clear කිරීම අවශ්‍ය නම්, ඔබගේ ඉල්ලීම මත එම Service එක සිදු කළ හැක.',
      '8. Software Performance Maintenance සඳහා අවශ්‍යතාවය අනුව වසරකට හෝ වසර 1.5කට වරක් Cleanup / Clear කිරීම අපි නිර්දේශ කරමු.',
      '9. Cleanup / Clear Service එකක් සිදු කිරීමට පෙර අලුත්ම Backup එකක් අනිවාර්යයෙන් තබා ගත යුතුය.',
      '10. ඉල්ලීම මත සිදු කරන වාර්ෂික / වසර 1.5ක Cleanup Service සඳහා Rs. 5,000ක මුදලක් අය කෙරේ.',
      '11. Backup Files ආරක්ෂිත ස්ථානයක තබා ගැනීම සහ Login Credentials ආරක්ෂා කිරීම කඩයේ වගකීම වේ.',
      '12. Billing, Inventory, Returns, Suppliers, Shop Order History, Customer Dashboard, Reports & Dashboard, Monthly Sales Report, Shop Profile, Offers සහ Barcode යන විශේෂාංග ස්ථාපිත Software Version එක සහ Service Terms අනුව පවතී.'
    ],
    ta: [
      '1. கடை மூடப்படும் உடனேயே தினசரி Backup கட்டாயமாக செய்ய வேண்டும்.',
      '2. உங்கள் Software Username மற்றும் Password பாதுகாப்பாக வைத்திருக்க வேண்டும்; அனுமதியற்ற நபர்களுடன் பகிரக்கூடாது.',
      '3. குறிப்பாக Username-ஐ பாதுகாப்பாக நினைவில் வைத்திருப்பது அவசியம்; Login மற்றும் Recovery க்கு அது முக்கியமானது.',
      '4. Software-இல் Password-ஐ தேவையில்லாமல் மாற்றக்கூடாது; பதிவு செய்யப்பட்ட Login தகவல்களை பாதுகாப்பாக வைத்திருக்க வேண்டும்.',
      '5. Software Install செய்த பிறகு எங்கள் சேவையின் மூலம் 3 Free Updates பெறலாம்.',
      '6. 3 Free Updates முடிந்த பிறகு ஒவ்வொரு கூடுதல் Update Service க்கும் Rs. 1,000 கட்டணம் வசூலிக்கப்படும்.',
      '7. ஒரு வருடத்திற்குப் பிறகு Software வேகத்தை அதிகரிக்க Application Data / Cache Clear செய்ய வேண்டுமெனில் கோரிக்கையின் பேரில் அந்த Service செய்யலாம்.',
      '8. Performance Maintenance க்காக தேவைக்கேற்ப வருடத்திற்கு ஒருமுறை அல்லது 1.5 ஆண்டுகளுக்கு ஒருமுறை Cleanup / Clear செய்ய பரிந்துரைக்கப்படுகிறது.',
      '9. Cleanup / Clear Service செய்வதற்கு முன் சமீபத்திய Backup ஒன்றை கட்டாயமாக வைத்திருக்க வேண்டும்.',
      '10. கோரிக்கையின் பேரில் செய்யப்படும் வருடாந்திர / 1.5 ஆண்டு Cleanup Service க்கு Rs. 5,000 கட்டணம் வசூலிக்கப்படும்.',
      '11. Backup Files மற்றும் Login Credentials-ஐ பாதுகாப்பாக வைத்திருப்பது கடையின் பொறுப்பாகும்.',
      '12. Billing, Inventory, Returns, Suppliers, Shop Order History, Customer Dashboard, Reports & Dashboard, Monthly Sales Report, Shop Profile, Offers மற்றும் Barcode அம்சங்கள் நிறுவப்பட்ட Software Version மற்றும் Service Terms படி வழங்கப்படும்.'
    ]
  };

  const downloadShopConditionsExcel = () => {
    const rows = shopConditions[lang].map((condition, index) => ({
      No: index + 1,
      Condition: condition
    }));
    const ws = XLSX.utils.json_to_sheet(rows);
    ws['!cols'] = [{ wch: 8 }, { wch: 115 }];
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, 'Shop Conditions');
    XLSX.writeFile(wb, `POShitha_POS_Shop_Conditions_${lang}_${new Date().toISOString().slice(0,10)}.xlsx`);
    showNotification('Shop Conditions Excel එක බාගත කිරීම ආරම්භ කළා.', 'success');
  };

  // BACKUP & RESTORE HANDLERS FOR SETTINGS DASHBOARD
  const handleDownloadBackup = async () => {
    try {
      let dbProducts = products;
      let dbSales = salesHistory;
      if (ipcRenderer) {
        dbProducts = await ipcRenderer.invoke('get-products') || [];
        dbSales = await ipcRenderer.invoke('get-sales') || [];
      }

      const backupData = {
        shopAccount: localStorage.getItem('pos_shop_account') ? JSON.parse(localStorage.getItem('pos_shop_account')) : {},
        receiptSettings: receiptSettings,
        cashiers: cashiers,
        suppliers: suppliers,
        orderRequests: orderRequests,
        shopOrderHistory: shopOrderHistory,
        products: dbProducts,
        salesHistory: dbSales,
        backupDate: new Date().toISOString()
      };

      const dataStr = "data:text/json;charset=utf-8," + encodeURIComponent(JSON.stringify(backupData, null, 2));
      const downloadAnchor = document.createElement('a');
      downloadAnchor.setAttribute("href", dataStr);
      downloadAnchor.setAttribute("download", `POS_System_Backup_${new Date().toISOString().slice(0, 10)}.json`);
      document.body.appendChild(downloadAnchor);
      downloadAnchor.click();
      downloadAnchor.remove();
      showNotification('💾 Backup ගොනුව සාර්ථකව බාගත කරගන්න ලදී!');
    } catch (err) {
      showNotification('Backup නිර්මාණය කිරීමේදී දෝෂයක් සිදු විය!');
    }
  };

  const handleRestoreBackup = (e) => {
    const file = e.target.files[0];
    if (!file) return;

    const reader = new FileReader();
    reader.onload = async (evt) => {
      try {
        const parsedData = JSON.parse(evt.target.result);
        if (!parsedData.receiptSettings && !parsedData.products) {
          showNotification('මෙම ගොනුව නිවැරදි POS Backup ගොනුවක් නොවේ!');
          return;
        }

        // Blocking confirmation dialog removed: restore proceeds and reports status through notifications.
        showNotification('🔄 Backup Restore ආරම්භ කරමින් පවතී...', 'info');
        {
          if (parsedData.shopAccount) {
            localStorage.setItem('pos_shop_account', JSON.stringify(parsedData.shopAccount));
          }
          if (parsedData.receiptSettings) {
            setReceiptSettings(parsedData.receiptSettings);
            localStorage.setItem('pos_receipt_settings', JSON.stringify(parsedData.receiptSettings));
          }
          if (parsedData.cashiers) {
            setCashiers(parsedData.cashiers);
            localStorage.setItem('pos_cashiers', JSON.stringify(parsedData.cashiers));
          }
          if (parsedData.suppliers) {
            setSuppliers(parsedData.suppliers);
            localStorage.setItem('pos_suppliers', JSON.stringify(parsedData.suppliers));
          }
          if (parsedData.orderRequests) {
            setOrderRequests(parsedData.orderRequests);
            localStorage.setItem('pos_order_requests', JSON.stringify(parsedData.orderRequests));
          }
          if (parsedData.shopOrderHistory) {
            setShopOrderHistory(parsedData.shopOrderHistory);
            localStorage.setItem('pos_shop_order_history', JSON.stringify(parsedData.shopOrderHistory));
          }

          if (ipcRenderer) {
            if (parsedData.products && Array.isArray(parsedData.products)) {
              await ipcRenderer.invoke('clear-all-products');
              await ipcRenderer.invoke('bulk-add-products', parsedData.products);
            }
            if (parsedData.salesHistory && Array.isArray(parsedData.salesHistory)) {
              await ipcRenderer.invoke('clear-all-sales');
              for (const sale of parsedData.salesHistory) {
                await ipcRenderer.invoke('save-sale', sale);
              }
            }
            await fetchProducts();
            await fetchSales();
          }

          showNotification('🎉 Backup ප්‍රතිස්ථාපනය (Restore) කිරීම සාර්ථකයි!');
        }
      } catch (err) {
        showNotification('Backup ගොනුව කියවීමේදී දෝෂයක් සිදු විය. කරුණාකර නිවැරදි JSON ගොනුවක් තෝරන්න.');
      }
    };
    reader.readAsText(file);
    e.target.value = '';
  };


  // Receipt-only item name formatter. Long names are shortened so Qty and Amount never overlap.
  const formatReceiptItemName = (value) => {
    const raw = String(value ?? '').replace(/\s+/g, ' ').trim();
    if (!raw) return 'Item';
    const maxChars = receiptSettings.paperWidth === '58mm' ? 27 : 39;
    if (raw.length <= maxChars) return raw;
    const cut = raw.slice(0, maxChars - 1).trim();
    const lastSpace = cut.lastIndexOf(' ');
    return `${(lastSpace > Math.floor(maxChars * 0.55) ? cut.slice(0, lastSpace) : cut).trim()}…`;
  };

  // Per-section receipt typography. These settings affect the printed receipt only.
  const receiptElementStyle = (key, fallback = {}) => ({
    fontSize: receiptSettings.elementStyles?.[key]?.fontSize || fallback.fontSize,
    fontWeight: receiptSettings.elementStyles?.[key]?.fontWeight ?? fallback.fontWeight,
    color: receiptSettings.elementStyles?.[key]?.color || fallback.color || '#000000',
    textAlign: receiptSettings.elementStyles?.[key]?.textAlign || fallback.textAlign || 'left',
    overflowWrap: 'anywhere',
    wordBreak: 'break-word',
    maxWidth: '100%',
    boxSizing: 'border-box'
  });

  // Shared thermal receipt renderer.
  // Billing Preview, Billing Print and Credit Book Print use this same receipt design.
  const renderReceiptContent = (saleOverride = activeSaleObj) => {
    const sale = saleOverride || activeSaleObj || {};
    const items = Array.isArray(sale.items) ? sale.items : [];
    const paymentType = String(sale.paymentMethod || sale.payment_method || 'Cash');
    const salePaid = Number(sale.paidAmount ?? sale.paid_amount ?? 0) || 0;
    const saleChange = Number(sale.changeAmount ?? sale.change_amount ?? 0) || 0;
    const saleCredit = Number(sale.creditAmount ?? sale.credit_amount ?? 0) || 0;
    const receiptText = {
      en: { tel:'Tel', email:'Email', invoice:'Invoice No', date:'Date', cashier:'Cashier', terminal:'Terminal', customer:'Customer', item:'Item & Description', qty:'Qty', amount:'Amount', unitPrice:'Unit Price', retailPrice:'Retail Price', discount:'Discount', subtotal:'Subtotal', offerDiscount:'Price / Offer Discount', billDiscount:'Bill Discount', billDiscountLkr:'Bill Discount (Direct LKR)', totalSavings:'TOTAL SAVINGS', grandTotal:'GRAND TOTAL', saleType:'Sale Type', tendered:'Tendered', balance:'Balance', creditDue:'Credit / Due', totalItems:'Total Items', noItem:'Item', retail:'Retail', wholesale:'Wholesale', special:'Special', cash:'Cash', card:'Card', online:'Online', credit:'Credit', offer:'Offer', off:'OFF', buy:'BUY', get:'GET', free:'FREE', eachFrom:'EACH FROM' },
      si: { tel:'දුරකථන', email:'විද්‍යුත් තැපෑල', invoice:'බිල් අංකය', date:'දිනය', cashier:'අයකැමි', terminal:'ටර්මිනල්', customer:'පාරිභෝගිකයා', item:'අයිතමය සහ විස්තරය', qty:'ප්‍රමාණය', amount:'මුදල', unitPrice:'ඒකක මිල', retailPrice:'සාමාන්‍ය මිල', discount:'වට්ටම', subtotal:'උප එකතුව', offerDiscount:'මිල / වට්ටම් වට්ටම', billDiscount:'බිල් වට්ටම', billDiscountLkr:'බිල් වට්ටම (සෘජු රු.)', totalSavings:'මුළු ඉතිරිය', grandTotal:'මුළු එකතුව', saleType:'විකුණුම් වර්ගය', tendered:'ගෙවූ මුදල', balance:'ඉතිරි මුදල', creditDue:'ණය / ගෙවිය යුතු', totalItems:'මුළු භාණ්ඩ', noItem:'භාණ්ඩය', retail:'සාමාන්‍ය', wholesale:'තොග', special:'විශේෂ', cash:'මුදල්', card:'කාඩ්', online:'ඔන්ලයින්', credit:'ණයට', offer:'වට්ටම', off:'වට්ටම', buy:'මිලදී ගන්න', get:'ලබාගන්න', free:'නොමිලේ', eachFrom:'එක් එකක් (සිට)' },
      ta: { tel:'தொலைபேசி', email:'மின்னஞ்சல்', invoice:'பில் எண்', date:'தேதி', cashier:'காசாளர்', terminal:'டெர்மினல்', customer:'வாடிக்கையாளர்', item:'பொருள் மற்றும் விவரம்', qty:'அளவு', amount:'தொகை', unitPrice:'அலகு விலை', retailPrice:'சில்லறை விலை', discount:'தள்ளுபடி', subtotal:'உப மொத்தம்', offerDiscount:'விலை / சலுகை தள்ளுபடி', billDiscount:'பில் தள்ளுபடி', billDiscountLkr:'பில் தள்ளுபடி (நேரடி ரூ.)', totalSavings:'மொத்த சேமிப்பு', grandTotal:'மொத்தம்', saleType:'விற்பனை வகை', tendered:'செலுத்தியது', balance:'மீதம்', creditDue:'கடன் / செலுத்த வேண்டியது', totalItems:'மொத்த பொருட்கள்', noItem:'பொருள்', retail:'சில்லறை', wholesale:'மொத்த விற்பனை', special:'சிறப்பு', cash:'பணம்', card:'அட்டை', online:'ஆன்லைன்', credit:'கடன்', offer:'சலுகை', off:'தள்ளுபடி', buy:'வாங்கு', get:'பெறு', free:'இலவசம்', eachFrom:'ஒவ்வொன்றும் முதல்' }
    }[lang];

    const receiptPaymentType = {
      Cash: receiptText.cash,
      Card: receiptText.card,
      Online: receiptText.online,
      Credit: receiptText.credit
    }[paymentType] || paymentType;

    const receiptPriceLevel = {
      normal: receiptText.retail,
      wholesale: receiptText.wholesale,
      special: receiptText.special
    };

    const receiptOfferLabel = (item, pricing) => {
      const type = String(item.offer_type || '').toLowerCase();
      if (type === 'percent') return `${Number(item.offer_value || 0)}% ${receiptText.off}`;
      if (type === 'b1g1') return `${receiptText.buy} ${item.offer_buy_qty || 1} ${receiptText.get} ${item.offer_free_qty || 1} ${receiptText.free}`;
      if (type === 'bulk_price') return `Rs. ${Number(item.bulk_price || item.offer_value || 0).toFixed(2)} ${receiptText.eachFrom} ${item.bulk_min_qty || item.offer_buy_qty || 2}`;
      return item.offerLabel || item.offer_label || pricing.offerLabel || '';
    };

    return (
      <div className="thermal-receipt-content">
        <div style={{ textAlign: 'center', marginBottom: '3mm' }}>
          {receiptSettings.logoUrl && (
            <div style={{ marginBottom: '1.5mm' }}>
              <img className="thermal-receipt-logo" src={receiptSettings.logoUrl} alt="Logo" style={{ width: receiptSettings.logoSize || '40mm', height: receiptSettings.logoSize || '40mm', minWidth: receiptSettings.logoSize || '40mm', minHeight: receiptSettings.logoSize || '40mm', maxWidth: receiptSettings.logoSize || '40mm', maxHeight: receiptSettings.logoSize || '40mm', objectFit: 'contain', borderRadius: '0', display: 'block', margin: '0 auto' }} />
            </div>
          )}
          <div className="receipt-shop-name" style={{ margin: 0, lineHeight: 1.05, letterSpacing: '.5px', ...receiptElementStyle('shopName', { fontSize: '24px', fontWeight: 900, textAlign: 'center' }) }}>
            {receiptSettings.shopName}
          </div>
          <div style={{ marginTop: '1mm', lineHeight: 1.35, whiteSpace: 'pre-line', ...receiptElementStyle('address', { fontSize: '16px', fontWeight: 900, textAlign: 'center' }) }}>
            {receiptSettings.address}
          </div>
          <div style={{ marginTop: '.7mm', lineHeight: 1.15, ...receiptElementStyle('phone', { fontSize: '16px', fontWeight: 900, textAlign: 'center' }) }}>{receiptText.tel}: {receiptSettings.phone}</div>
          {receiptSettings.email && <div style={{ lineHeight: 1.15, overflowWrap: 'anywhere', ...receiptElementStyle('email', { fontSize: '12px', fontWeight: 700, textAlign: 'center' }) }}>{receiptText.email}: {receiptSettings.email}</div>}
        </div>

        <div className="receipt-meta-box" style={{ marginBottom: '2.8mm', borderTop: '2px dashed #000', borderBottom: '2px dashed #000', padding: '2mm 0', ...receiptElementStyle('invoice', receiptSettings.elementStyles?.meta || { fontSize: '12px', fontWeight: 700, textAlign: 'left' }) }}>
          <div className="thermal-row"><span>{receiptText.invoice}:</span><span className="thermal-number">#{sale.invoiceNo ?? sale.invoice_no ?? sale.id ?? '-'}</span></div>
          <div className="thermal-row"><span>{receiptText.date}:</span><span className="thermal-number">{sale.dateFormatted || (sale.date ? new Date(sale.date).toLocaleString() : '-')}</span></div>
          <div className="thermal-row"><span>{receiptText.cashier}:</span><span className="thermal-number">{String(sale.cashier || currentUserRole || 'ADMIN').toUpperCase()}</span></div>
          <div className="thermal-row"><span>{receiptText.terminal}:</span><span className="thermal-number">{receiptSettings.terminal}</span></div>
          {(sale.customerName && sale.customerName !== 'General Customer') && (
            <div className="thermal-row"><span>{receiptText.customer}:</span><span className="thermal-number" style={{ maxWidth: '60%', overflowWrap: 'anywhere' }}>{sale.customerName}</span></div>
          )}
        </div>

        <div className="receipt-items-heading" style={{ display: 'grid', gridTemplateColumns: 'minmax(0, 1fr) 9mm 19mm', columnGap: '1mm', fontWeight: '900', borderBottom: '2px solid #000', paddingBottom: '1.2mm', marginBottom: '1.8mm', fontSize: '1.05em', minWidth: 0 }}>
          <span style={{ minWidth: 0 }}># {receiptText.item}</span>
          <span style={{ textAlign: 'center', whiteSpace: 'nowrap' }}>{receiptText.qty}</span>
          <span style={{ textAlign: 'right', whiteSpace: 'nowrap' }}>{receiptText.amount}</span>
        </div>

        {items.map((item, index) => {
          const pricing = getLinePricing(item, item.qty);
          const unitPrice = Number(item.unitPrice ?? item.price ?? item.selling_price ?? pricing.unitPrice) || 0;
          const retailUnitPrice = Number(item.retailPrice ?? pricing.retailUnitPrice ?? item.price ?? unitPrice) || 0;
          const qty = Number(item.qty ?? item.quantity ?? 0) || 0;
          const lineSubtotal = Number(item.subtotal ?? pricing.subtotal ?? (unitPrice * qty)) || 0;
          const offerDiscount = Number(item.discount ?? pricing.discount ?? 0) || 0;
          const priceLevelDiscount = Number(item.priceLevelDiscount ?? pricing.priceLevelDiscount ?? 0) || 0;
          const lineDiscount = offerDiscount + priceLevelDiscount;
          const lineTotal = Number(item.finalPrice ?? item.final_price ?? pricing.finalPrice ?? (lineSubtotal - offerDiscount)) || 0;
          const offerLabel = receiptOfferLabel(item, pricing);
          const itemPriceLevel = item.priceLevel || sale.priceLevel || 'normal';
          const hasCustomerPrice = itemPriceLevel !== 'normal' && retailUnitPrice > unitPrice;

          return (
            <div className="thermal-item-row" key={item.id || item.barcode || index} style={{ marginBottom: '2.6mm', paddingBottom: '1.2mm', borderBottom: '.6px dashed #9ca3af', minWidth: 0, width: '100%', boxSizing: 'border-box', breakInside: 'avoid', pageBreakInside: 'avoid', ...receiptElementStyle('items', { fontSize: '14px', fontWeight: 900, textAlign: 'left' }) }}>
              <div style={{ display: 'grid', gridTemplateColumns: 'minmax(0, 1fr) 9mm 19mm', columnGap: '1mm', alignItems: 'start', minWidth: 0, width: '100%' }}>
                <span className="thermal-item-name" title={String(item.name || item.product_name || 'Item')} style={{ minWidth: 0, maxWidth: '100%', fontSize: '1em', lineHeight: 1.18, fontWeight: 900, overflowWrap: 'anywhere', wordBreak: 'break-word' }}>{index + 1}. {formatReceiptItemName(item.name || item.product_name || 'Item')}</span>
                <span className="thermal-number" style={{ width: '9mm', fontSize: '.96em', lineHeight: 1.15, fontWeight: 900, whiteSpace: 'nowrap', textAlign: 'center' }}>{qty}</span>
                <span className="thermal-number" style={{ width: '19mm', fontSize: '.96em', lineHeight: 1.15, fontWeight: 900, whiteSpace: 'nowrap', textAlign: 'right' }}>{lineTotal.toFixed(2)}</span>
              </div>
              <div style={{ fontSize: '0.82em', lineHeight: 1.2, fontWeight: '700', paddingLeft: '4mm', marginTop: '.5mm', overflowWrap: 'anywhere' }}>
                {hasCustomerPrice ? (
                  <>
                    <span>{receiptText.retailPrice}: <span style={{ textDecoration: 'line-through', textDecorationThickness: '1.5px' }}>Rs. {retailUnitPrice.toFixed(2)}</span></span>
                    <span> &nbsp;{receiptPriceLevel[itemPriceLevel] || receiptText.retail} {receiptText.unitPrice}: Rs. {unitPrice.toFixed(2)}</span>
                  </>
                ) : (
                  <span>{receiptText.unitPrice}: Rs. {unitPrice.toFixed(2)}</span>
                )}
                {offerLabel ? ` | ${offerLabel}` : ''}
              </div>
              {lineDiscount > 0 && (
                <div style={{ fontSize: '0.82em', lineHeight: 1.2, fontWeight: '700', paddingLeft: '4mm', marginTop: '.3mm' }}>
                  {receiptText.discount}: -Rs. {lineDiscount.toFixed(2)} (Sub: Rs. {lineSubtotal.toFixed(2)})
                </div>
              )}
            </div>
          );
        })}

        <div style={{ borderTop: '2px dashed #000', margin: '2.8mm 0' }}></div>

        <div className="receipt-totals" style={{ fontSize: '1.1em', fontWeight: '900', lineHeight: 1.3 }}>
          <div className="thermal-row"><span>{receiptText.subtotal}:</span><span className="thermal-number">{Number(sale.retailSubtotal ?? sale.subtotal ?? sale.total ?? 0).toFixed(2)}</span></div>
          <div className="thermal-row"><span>{receiptText.offerDiscount}:</span><span className="thermal-number">-{Math.max(0, Number(sale.discount ?? 0) - Number(sale.billDiscountAmount ?? sale.bill_discount_amount ?? 0)).toFixed(2)}</span></div>
          {Number(sale.billDiscountAmount ?? sale.bill_discount_amount ?? 0) > 0 && <div className="thermal-row"><span>{String(sale.billDiscountType ?? sale.bill_discount_type ?? '').toLowerCase() === 'lkr' || Number(sale.billDiscountLkr ?? sale.bill_discount_lkr ?? 0) > 0 ? receiptText.billDiscountLkr : `${receiptText.billDiscount} (${Number(sale.billDiscountPercent ?? sale.bill_discount_percent ?? 0).toFixed(2)}%):`}</span><span className="thermal-number">-{Number(sale.billDiscountAmount ?? sale.bill_discount_amount ?? 0).toFixed(2)}</span></div>}
          <div className="thermal-row" style={{ fontWeight: 900 }}><span>{receiptText.totalSavings}:</span><span className="thermal-number">-{Number(sale.discount ?? 0).toFixed(2)}</span></div>

          <div className="thermal-grand-total receipt-grand-total" style={{ display: 'grid', gridTemplateColumns: 'minmax(0, 1fr) auto', columnGap: '2mm', alignItems: 'center', margin: '2.8mm 0', borderTop: '2px solid #000', borderBottom: '2px solid #000', padding: '2mm 0', lineHeight: 1.05, ...receiptElementStyle('grandTotal', { fontSize: '21px', fontWeight: 900, textAlign: 'right' }) }}>
            <span>{receiptText.grandTotal}:</span>
            <span className="thermal-number">{Number(sale.total || 0).toFixed(2)}</span>
          </div>

          <div style={{ margin: '2.5mm 0' }}>
            <div className="thermal-row"><span>{receiptText.saleType}:</span><span className="thermal-number">{receiptPaymentType}</span></div>
            <div className="thermal-row"><span>{receiptText.tendered}:</span><span className="thermal-number">{salePaid.toFixed(2)}</span></div>
            <div className="thermal-row"><span>{receiptText.balance}:</span><span className="thermal-number">{saleChange.toFixed(2)}</span></div>
            {paymentType.toLowerCase() === 'credit' && (
              <div className="thermal-row"><span>{receiptText.creditDue}:</span><span className="thermal-number">{saleCredit.toFixed(2)}</span></div>
            )}
          </div>

          <div style={{ textAlign: 'center', marginTop: '4mm', paddingTop: '2mm', borderTop: '1px dashed #000' }}>
            <div style={{ fontSize: '1.1em', fontWeight: '900' }}>{receiptText.totalItems}: {items.length}</div>
            <div className="receipt-footer-message" style={{ marginTop: '2mm', letterSpacing: '.5px', ...receiptElementStyle('footer', { fontSize: '19px', fontWeight: 900, textAlign: 'center' }) }}>
              {receiptSettings.footerMsg}
            </div>
            {receiptSettings.footerComment && (
              <div style={{ marginTop: '2mm', whiteSpace: 'pre-line', overflowWrap: 'anywhere', ...receiptElementStyle('footerComment', { fontSize: '11px', fontWeight: 700, textAlign: 'center' }) }}>
                {receiptSettings.footerComment}
              </div>
            )}
            {receiptSettings.poweredByEnabled && receiptSettings.poweredByText && (
              <div className="receipt-powered-by" style={{ marginTop: '2.5mm', paddingTop: '1.5mm', borderTop: '.5px solid #d1d5db', letterSpacing: '.1px', lineHeight: 1.1, ...receiptElementStyle('poweredBy', { fontSize: '8px', fontWeight: 400, textAlign: 'center', color: '#666666' }) }}>
                {receiptSettings.poweredByText}
              </div>
            )}
          </div>
        </div>
      </div>
    );
  };

  // Convert an existing Customer/Credit Book sale into the same receipt structure used by Billing.
  const buildHistoricalSaleReceipt = (sale) => {
    const saleItems = (reportsData.sales || [])
      .filter(row => String(row.sale_id ?? row.saleId ?? '') === String(sale.id))
      .map((item, index) => {
        const unitPrice = Number(item.unitPrice ?? item.price ?? item.selling_price ?? 0) || 0;
        const qty = Number(item.qty ?? item.quantity ?? 0) || 0;
        const subtotal = Number(item.subtotal ?? (unitPrice * qty)) || 0;
        const discount = Number(item.discount ?? 0) || 0;
        const finalPrice = Number(item.finalPrice ?? item.final_price ?? (subtotal - discount)) || 0;
        return {
          ...item,
          id: item.id ?? `${sale.id}-${index}`,
          name: item.name ?? item.product_name ?? 'Item',
          qty,
          price: unitPrice,
          unitPrice,
          subtotal,
          discount,
          finalPrice,
          offerLabel: item.offerLabel || item.offer_label || (
            item.offer_type === 'percent' ? `${item.offer_value || 0}% OFF` :
            item.offer_type === 'b1g1' ? `BUY ${item.offer_buy_qty || 1} GET ${item.offer_free_qty || 1} FREE` : item.offer_type === 'bulk_price' ? `Rs. ${item.bulk_price || 0} EACH FROM ${item.bulk_min_qty || 2}` : ''
          )
        };
      });

    const total = Number(sale.total || 0) || 0;
    const paid = Number(sale.paid_amount ?? sale.paidAmount ?? 0) || 0;
    const due = Number(sale.credit_amount ?? sale.creditAmount ?? (String(sale.payment_method || '').toLowerCase() === 'credit' ? Math.max(0, total - paid) : 0)) || 0;

    return {
      invoiceNo: getSaleInvoiceNo(sale),
      items: saleItems,
      subtotal: Number(sale.subtotal ?? total) || total,
      discount: Number(sale.discount ?? 0) || 0,
      total,
      paymentMethod: sale.payment_method || sale.paymentMethod || 'Cash',
      paidAmount: paid,
      changeAmount: Number(sale.change_amount ?? sale.changeAmount ?? 0) || 0,
      creditAmount: due,
      customerName: sale.customer_name || sale.customerName || 'General Customer',
      customerPhone: sale.customer_phone || sale.customerPhone || '-',
      cashier: sale.cashier || currentUserRole,
      terminal: receiptSettings.terminal,
      dateFormatted: sale.date ? new Date(sale.date).toLocaleString() : currentFormattedDate()
    };
  };

  const handlePrintCreditBookBill = (sale) => {
    const receipt = buildHistoricalSaleReceipt(sale);
    if (!receipt.items.length) {
      showNotification('මෙම Bill එකේ Item details හමු වුණේ නැහැ.', 'error');
      return;
    }
    setCreditBookPrintSale(receipt);
    setPrintMode('creditBook');
    setTimeout(() => window.print(), 250);
  };

  if (!isLoggedIn) {
    return (
      <div style={{
        display: 'flex', justifyContent: 'center', alignItems: 'center', height: '100vh',
        backgroundColor: '#0f172a',
        backgroundImage: 'repeating-linear-gradient(0deg, rgba(148,163,184,.055) 0, rgba(148,163,184,.055) 1px, transparent 1px, transparent 44px), repeating-linear-gradient(90deg, rgba(148,163,184,.055) 0, rgba(148,163,184,.055) 1px, transparent 1px, transparent 44px), radial-gradient(circle at 15% 20%, rgba(37,99,235,.45) 0, transparent 32%), radial-gradient(circle at 85% 15%, rgba(124,58,237,.38) 0, transparent 30%), radial-gradient(circle at 75% 85%, rgba(5,150,105,.30) 0, transparent 32%), linear-gradient(135deg, #020617 0%, #111827 48%, #1e293b 100%)',
        fontFamily: 'sans-serif', flexDirection: 'column', gap: '15px', position: 'relative', overflow: 'hidden'
      }}>
        <div style={{ position: 'absolute', inset: 0, background: 'linear-gradient(120deg, rgba(255,255,255,.05), transparent 45%, rgba(255,255,255,.03))', pointerEvents: 'none' }} />
        
        <div style={{ display: 'flex', gap: '10px', backgroundColor: '#374151', padding: '6px 12px', borderRadius: '8px' }}>
          <button onClick={() => setLang('si')} style={{ padding: '4px 10px', backgroundColor: lang === 'si' ? '#2563eb' : 'transparent', color: '#fff', border: 'none', borderRadius: '4px', cursor: 'pointer', fontWeight: 'bold' }}>සිංහල</button>
          <button onClick={() => setLang('en')} style={{ padding: '4px 10px', backgroundColor: lang === 'en' ? '#2563eb' : 'transparent', color: '#fff', border: 'none', borderRadius: '4px', cursor: 'pointer', fontWeight: 'bold' }}>English</button>
          <button onClick={() => setLang('ta')} style={{ padding: '4px 10px', backgroundColor: lang === 'ta' ? '#2563eb' : 'transparent', color: '#fff', border: 'none', borderRadius: '4px', cursor: 'pointer', fontWeight: 'bold' }}>தமிழ்</button>
        </div>

        <div style={{ backgroundColor: '#fff', padding: '30px', borderRadius: '16px', width: '400px', boxShadow: '0 18px 55px rgba(0,0,0,0.42)', maxHeight: '90vh', overflowY: 'auto' }}>
          
          {authMode === 'login' && (
            <div style={{ textAlign: 'center', marginBottom: '20px' }}>
              <div style={{ display: 'inline-flex', alignItems: 'center', justifyContent: 'center', width: '64px', height: '64px', borderRadius: '18px', background: 'linear-gradient(135deg, #2563eb, #7c3aed)', color: '#fff', fontSize: '30px', fontWeight: '900', boxShadow: '0 10px 25px rgba(37,99,235,.30)' }}>POS</div>
              <h1 style={{ margin: '12px 0 4px', fontSize: '25px', letterSpacing: '1.2px', fontWeight: '900', color: '#0f172a' }}>POSHITHA POS SYSTEM</h1>
              <div style={{ width: '100%', height: '3px', borderRadius: '999px', background: 'linear-gradient(90deg, transparent, #2563eb, #7c3aed, #059669, transparent)', margin: '8px 0 10px' }} />
              <p style={{ margin: 0, color: '#64748b', fontSize: '12px', fontWeight: '700', letterSpacing: '.4px' }}>Smart Billing • Inventory • Sales Analytics</p>
            </div>
          )}

          {authMode === 'register' ? (
            <div>
              <div style={{ textAlign: 'center', marginBottom: '15px' }}>
                <h2 style={{ color: '#1f2937', margin: '0 0 5px 0' }}>🏪 {t.shopRegistration}</h2>
              </div>

              {regError && (
                <div style={{ backgroundColor: '#fee2e2', color: '#dc2626', padding: '10px', borderRadius: '6px', fontSize: '13px', marginBottom: '15px', fontWeight: 'bold', border: '1px solid #fecaca' }}>
                  {regError}
                </div>
              )}

              <form onSubmit={handleRegisterSubmit} style={{ display: 'flex', flexDirection: 'column', gap: '10px' }}>
                <div>
                  <label style={{ fontSize: '12px', fontWeight: 'bold', color: '#374151' }}>{t.shopName}:</label>
                  <input type="text" placeholder="My Awesome Shop" value={regShopName} onChange={(e) => setRegShopName(e.target.value)} required style={{ width: '100%', padding: '8px', borderRadius: '6px', border: '1px solid #d1d5db' }} />
                </div>
                <div>
                  <label style={{ fontSize: '12px', fontWeight: 'bold', color: '#374151' }}>{t.shopMail}:</label>
                  <input type="email" placeholder="shop@gmail.com" value={regEmail} onChange={(e) => setRegEmail(e.target.value)} required style={{ width: '100%', padding: '8px', borderRadius: '6px', border: '1px solid #d1d5db' }} />
                </div>
                <div>
                  <label style={{ fontSize: '12px', fontWeight: 'bold', color: '#374151' }}>{t.address}:</label>
                  <input type="text" placeholder="No. 123, Main Street" value={regAddress} onChange={(e) => setRegAddress(e.target.value)} required style={{ width: '100%', padding: '8px', borderRadius: '6px', border: '1px solid #d1d5db' }} />
                </div>
                <div>
                  <label style={{ fontSize: '12px', fontWeight: 'bold', color: '#374151' }}>{t.phone}:</label>
                  <input type="text" placeholder="0712345678" value={regPhone} onChange={(e) => setRegPhone(e.target.value)} required style={{ width: '100%', padding: '8px', borderRadius: '6px', border: '1px solid #d1d5db' }} />
                </div>
                <div>
                  <label style={{ fontSize: '12px', fontWeight: 'bold', color: '#374151' }}>Shop Logo:</label>
                  <input type="file" accept="image/*" onChange={(e) => {
                    const file = e.target.files?.[0];
                    if (!file) return;
                    const reader = new FileReader();
                    reader.onload = () => setRegLogoUrl(reader.result);
                    reader.readAsDataURL(file);
                  }} style={{ width: '100%', padding: '6px' }} />
                  {regLogoUrl && <img src={regLogoUrl} alt="Shop Logo Preview" style={{ width: '65px', height: '65px', objectFit: 'contain', marginTop: '7px', border: '1px solid #d1d5db', borderRadius: '8px', background: '#fff' }} />}
                </div>
                <div>
                  <label style={{ fontSize: '12px', fontWeight: 'bold', color: '#374151' }}>Admin {t.username}:</label>
                  <input type="text" placeholder="admin" value={regUsername} onChange={(e) => setRegUsername(e.target.value)} required style={{ width: '100%', padding: '8px', borderRadius: '6px', border: '1px solid #d1d5db' }} />
                </div>
                <div>
                  <label style={{ fontSize: '12px', fontWeight: 'bold', color: '#374151' }}>Admin {t.password}:</label>
                  <input type="password" placeholder="********" value={regPassword} onChange={(e) => setRegPassword(e.target.value)} required style={{ width: '100%', padding: '8px', borderRadius: '6px', border: '1px solid #d1d5db' }} />
                </div>
                <div>
                  <label style={{ fontSize: '12px', fontWeight: 'bold', color: '#dc2626' }}>{t.secretKey}:</label>
                  <input type="password" placeholder="Enter approval key" value={regSecretKey} onChange={(e) => setRegSecretKey(e.target.value)} required style={{ width: '100%', padding: '8px', borderRadius: '6px', border: '2px solid #dc2626', backgroundColor: '#fef2f2', fontWeight: 'bold' }} />
                </div>

                <button type="submit" style={{ padding: '12px', backgroundColor: '#059669', color: '#fff', border: 'none', borderRadius: '6px', fontWeight: 'bold', fontSize: '15px', cursor: 'pointer', marginTop: '5px' }}>
                  ✅ {t.registerBtn}
                </button>
              </form>

              <div style={{ marginTop: '15px', textAlign: 'center' }}>
                <button onClick={() => setAuthMode('login')} style={{ background: 'none', border: 'none', color: '#2563eb', cursor: 'pointer', fontSize: '13px', textDecoration: 'underline' }}>
                  {t.alreadyHaveAccount}
                </button>
              </div>
            </div>
          ) : authMode === 'forgot' ? (
            <div>
              <div style={{ textAlign: 'center', marginBottom: '15px' }}>
                <h2 style={{ color: '#1f2937', margin: '0 0 5px 0' }}>🔑 මුරපදය නැවත සකස් කිරීම</h2>
                <p style={{ color: '#6b7280', fontSize: '12px', margin: 0 }}>Password Reset via Secret Key</p>
              </div>

              {resetError && (
                <div style={{ backgroundColor: '#fee2e2', color: '#dc2626', padding: '10px', borderRadius: '6px', fontSize: '13px', marginBottom: '15px', fontWeight: 'bold', border: '1px solid #fecaca' }}>
                  {resetError}
                </div>
              )}

              {resetSuccess && (
                <div style={{ backgroundColor: '#ecfdf5', color: '#059669', padding: '10px', borderRadius: '6px', fontSize: '13px', marginBottom: '15px', fontWeight: 'bold', border: '1px solid #a7f3d0' }}>
                  {resetSuccess}
                </div>
              )}

              <form onSubmit={handleResetSubmit} style={{ display: 'flex', flexDirection: 'column', gap: '12px' }}>
                <div>
                  <label style={{ fontSize: '12px', fontWeight: 'bold', color: '#374151' }}>Shop Username:</label>
                  <input type="text" placeholder="Enter your username" value={resetUsername} onChange={(e) => setResetUsername(e.target.value)} required style={{ width: '100%', padding: '10px', borderRadius: '6px', border: '1px solid #d1d5db', fontSize: '14px' }} />
                </div>
                <div>
                  <label style={{ fontSize: '12px', fontWeight: 'bold', color: '#dc2626' }}>Secret Key:</label>
                  <input type="password" placeholder="Enter secret key" value={resetSecretKey} onChange={(e) => setResetSecretKey(e.target.value)} required style={{ width: '100%', padding: '10px', borderRadius: '6px', border: '2px solid #dc2626', backgroundColor: '#fef2f2', fontWeight: 'bold', fontSize: '14px' }} />
                </div>
                <div>
                  <label style={{ fontSize: '12px', fontWeight: 'bold', color: '#374151' }}>New Username:</label>
                  <input type="text" placeholder="Enter new username" value={newAdminUsername} onChange={(e) => setNewAdminUsername(e.target.value)} required style={{ width: '100%', padding: '10px', borderRadius: '6px', border: '1px solid #d1d5db', fontSize: '14px' }} />
                </div>
                <div>
                  <label style={{ fontSize: '12px', fontWeight: 'bold', color: '#374151' }}>New Password:</label>
                  <input type="password" placeholder="Enter new password" value={newAdminPassword} onChange={(e) => setNewAdminPassword(e.target.value)} required style={{ width: '100%', padding: '10px', borderRadius: '6px', border: '1px solid #d1d5db', fontSize: '14px' }} />
                </div>

                <button type="submit" style={{ padding: '12px', backgroundColor: '#d97706', color: '#fff', border: 'none', borderRadius: '6px', fontWeight: 'bold', fontSize: '15px', cursor: 'pointer', marginTop: '5px' }}>
                  🔄 Reset Password
                </button>
              </form>

              <div style={{ marginTop: '15px', textAlign: 'center' }}>
                <button onClick={() => setAuthMode('login')} style={{ background: 'none', border: 'none', color: '#2563eb', cursor: 'pointer', fontSize: '13px', textDecoration: 'underline' }}>
                  ⬅️ Back to Login
                </button>
              </div>
            </div>
          ) : (
            <div>
              <div style={{ textAlign: 'center', marginBottom: '20px' }}>
                <h2 style={{ color: '#1f2937', margin: '0 0 5px 0' }}>🔐 {t.loginTitle}</h2>
                <p style={{ color: '#6b7280', fontSize: '13px', margin: 0 }}>Admin / Cashier Login</p>
              </div>

              {loginError && (
                <div style={{ backgroundColor: '#fee2e2', color: '#dc2626', padding: '10px', borderRadius: '6px', fontSize: '13px', marginBottom: '15px', fontWeight: 'bold', border: '1px solid #fecaca' }}>
                  {loginError}
                </div>
              )}

              <form onSubmit={handleLoginSubmit} style={{ display: 'flex', flexDirection: 'column', gap: '15px' }}>
                <div>
                  <label style={{ display: 'block', fontSize: '13px', fontWeight: 'bold', color: '#374151', marginBottom: '5px' }}>{t.username}:</label>
                  <input type="text" placeholder="admin / cashier1" value={loginUsername} onChange={(e) => setLoginUsername(e.target.value)} required style={{ width: '100%', padding: '10px', borderRadius: '6px', border: '1px solid #d1d5db', fontSize: '14px', outline: 'none' }} />
                </div>
                <div>
                  <label style={{ display: 'block', fontSize: '13px', fontWeight: 'bold', color: '#374151', marginBottom: '5px' }}>{t.password}:</label>
                  <input type="password" placeholder="Enter password" value={loginPassword} onChange={(e) => setLoginPassword(e.target.value)} required style={{ width: '100%', padding: '10px', borderRadius: '6px', border: '1px solid #d1d5db', fontSize: '14px', outline: 'none' }} />
                </div>

                <button type="submit" style={{ padding: '12px', backgroundColor: '#2563eb', color: '#fff', border: 'none', borderRadius: '6px', fontWeight: 'bold', fontSize: '15px', cursor: 'pointer', marginTop: '5px' }}>
                  🚀 {t.loginBtn}
                </button>
              </form>

              <div style={{ marginTop: '20px', textAlign: 'center', display: 'flex', flexDirection: 'column', gap: '10px' }}>
                <button onClick={() => setAuthMode('forgot')} style={{ background: 'none', border: 'none', color: '#d97706', cursor: 'pointer', fontSize: '13px', fontWeight: 'bold', textDecoration: 'underline' }}>
                  🔑 Username හෝ Password අමතකද? (Password Reset)
                </button>
                <button onClick={() => setAuthMode('register')} style={{ background: 'none', border: 'none', color: '#059669', cursor: 'pointer', fontSize: '13px', fontWeight: 'bold', textDecoration: 'underline' }}>
                  ➕ {t.needToRegister}
                </button>
                <div style={{ fontSize: '11px', color: '#9ca3af', marginTop: '5px' }}>
                  Default Admin: <b>admin</b> / <b>123</b> | Cashier: <b>cashier1</b> / <b>123</b>
                </div>
              </div>
            </div>
          )}

        </div>
      </div>
    );
  }

  return (
    <div style={{ display: 'flex', flexDirection: dashboardStyle === 'classic' ? 'row' : 'column', height: '100vh', minWidth: '1100px', fontFamily: 'sans-serif', backgroundColor: '#f3f4f6', overflow: 'hidden' }}>
      {notification && (
        <div className="no-print" style={{
          position: 'fixed', top: '18px', right: '18px', zIndex: 999999,
          minWidth: '300px', maxWidth: '460px', padding: '14px 16px',
          borderRadius: '10px', color: '#fff', boxShadow: '0 10px 30px rgba(0,0,0,0.25)',
          backgroundColor: notification.type === 'success' ? '#059669' : notification.type === 'error' ? '#dc2626' : '#2563eb',
          display: 'flex', alignItems: 'center', gap: '12px'
        }}>
          <span style={{ fontSize: '20px' }}>{notification.type === 'success' ? '✓' : notification.type === 'error' ? '⚠️' : 'ℹ️'}</span>
          <div style={{ flex: 1, minWidth: 0 }}>
            <div style={{ fontWeight: '600', lineHeight: 1.4 }}>{notification.message}</div>
            {pendingUnderpayment && (
              <div style={{ display: 'flex', gap: '8px', marginTop: '10px' }}>
                <button onClick={continueAfterUnderpayment} style={{ padding: '7px 13px', border: 'none', borderRadius: '6px', background: '#fff', color: '#2563eb', fontWeight: '800', cursor: 'pointer' }}>ඔව්, ඉදිරියට</button>
                <button onClick={cancelUnderpayment} style={{ padding: '7px 13px', border: '1px solid rgba(255,255,255,.65)', borderRadius: '6px', background: 'transparent', color: '#fff', fontWeight: '800', cursor: 'pointer' }}>නැහැ</button>
              </div>
            )}
            {pendingClearAllProducts && (
              <div style={{ display: 'flex', gap: '8px', marginTop: '10px' }}>
                <button onClick={confirmClearAllProducts} style={{ padding: '7px 13px', border: 'none', borderRadius: '6px', background: '#fff', color: '#dc2626', fontWeight: '900', cursor: 'pointer' }}>OK, සියල්ල Delete කරන්න</button>
                <button onClick={cancelClearAllProducts} style={{ padding: '7px 13px', border: '1px solid rgba(255,255,255,.65)', borderRadius: '6px', background: 'transparent', color: '#fff', fontWeight: '800', cursor: 'pointer' }}>NO</button>
              </div>
            )}
          </div>
          <button onClick={() => { setNotification(null); if (pendingUnderpayment) setPendingUnderpayment(false); if (pendingClearAllProducts) setPendingClearAllProducts(false); if (notificationTimerRef.current) clearTimeout(notificationTimerRef.current); }} style={{ alignSelf: 'flex-start', background: 'transparent', color: '#fff', border: 'none', fontSize: '20px', cursor: 'pointer', padding: 0 }}>×</button>
        </div>
      )}
      
      {showWelcomeAnimation && (
        <div className="no-print poshitha-welcome-overlay">
          <div className="poshitha-welcome-glow" />
          <div className="poshitha-welcome-card">
            {receiptSettings.logoUrl ? (
              <img
                src={receiptSettings.logoUrl}
                alt="Shop Logo"
                className="poshitha-welcome-logo"
                style={{ objectFit: 'contain', background: '#fff', padding: '8px', boxSizing: 'border-box' }}
              />
            ) : (
              <div className="poshitha-welcome-logo">POS</div>
            )}
            <h1 className="poshitha-welcome-title">
              {lang === 'si' ? 'POSHITHA POS SYSTEM වෙත සාදරයෙන් පිළිගනිමු' :
               lang === 'ta' ? 'POSHITHA POS SYSTEM-க்கு வரவேற்கிறோம்' :
               'WELCOME TO POSHITHA POS SYSTEM'}
            </h1>
            <div className="poshitha-welcome-subtitle">
              {lang === 'si' ? 'Billing Dashboard විවෘත වෙමින් පවතී...' :
               lang === 'ta' ? 'Billing Dashboard திறக்கப்படுகிறது...' :
               'Billing Dashboard is opening...'}
            </div>
          </div>
        </div>
      )}

      {excelImportProgress && (
        <div className="no-print" style={{
          position:'fixed', inset:0, zIndex:999998, background:'rgba(15,23,42,.62)',
          display:'flex', alignItems:'center', justifyContent:'center', padding:'20px'
        }}>
          <div style={{
            width:'min(520px, 95vw)', background:'#fff', borderRadius:'14px', padding:'24px',
            boxShadow:'0 24px 80px rgba(0,0,0,.35)'
          }}>
            <div style={{display:'flex',justifyContent:'space-between',gap:'12px',alignItems:'center'}}>
              <h3 style={{margin:0}}>📊 Excel Inventory Update</h3>
              <b style={{fontSize:'22px',color:'#2563eb'}}>{excelImportProgress.percent}%</b>
            </div>
            <div style={{marginTop:'8px',color:'#475569',fontSize:'13px',fontWeight:700}}>
              {excelImportProgress.stage}
            </div>
            <div style={{marginTop:'16px',height:'18px',background:'#e5e7eb',borderRadius:'999px',overflow:'hidden'}}>
              <div style={{
                width:`${excelImportProgress.percent}%`,height:'100%',
                background:'linear-gradient(90deg,#2563eb,#7c3aed,#059669)',
                transition:'width .25s ease'
              }} />
            </div>
            <div style={{display:'grid',gridTemplateColumns:'repeat(3,1fr)',gap:'8px',marginTop:'15px'}}>
              <div style={{padding:'10px',background:'#eff6ff',borderRadius:'8px',textAlign:'center'}}><b>{excelImportProgress.total || 0}</b><div style={{fontSize:'11px'}}>Total</div></div>
              <div style={{padding:'10px',background:'#ecfdf5',borderRadius:'8px',textAlign:'center'}}><b>{excelImportProgress.updated || 0}</b><div style={{fontSize:'11px'}}>Updated</div></div>
              <div style={{padding:'10px',background:'#fefce8',borderRadius:'8px',textAlign:'center'}}><b>{excelImportProgress.added || 0}</b><div style={{fontSize:'11px'}}>New</div></div>
            </div>
            <div style={{marginTop:'14px',fontSize:'11px',color:'#64748b'}}>
              Existing items are updated by Barcode; new items are added. The screen is kept responsive during large imports.
            </div>
          </div>
        </div>
      )}

      <style>{`
        .barcode-print-sheet { display: none; }
        /* BARCODE LABEL PREVIEW - order is Shop -> Item -> Price/Unit -> Bars -> Number */
        .barcode-preview-grid { display: grid; grid-template-columns: repeat(6, 30mm); grid-auto-rows: 20mm; gap: 5px; align-content: start; justify-content: center; }
        .barcode-preview-grid[data-size="large"] { grid-template-columns: repeat(5, 38mm); grid-auto-rows: 25mm; }

        .preview-label {
          box-sizing: border-box;
          width: 30mm;
          height: 20mm;
          border: 0.45mm dashed #222;
          background: #fff;
          padding: 1mm 1.15mm 0.8mm;
          overflow: hidden;
          display: flex;
          flex-direction: column;
          align-items: center;
          justify-content: flex-start;
          gap: 0.55mm;
          color: #000;
          font-family: Arial, sans-serif;
        }
        .barcode-preview-grid[data-size="large"] .preview-label {
          width: 38mm;
          height: 25mm;
          padding: 1.2mm 1.3mm 1mm;
          gap: 0.7mm;
        }

        .preview-label .barcode-shop {
          flex: 0 0 auto;
          font-size: 6.2px;
          line-height: 1;
          font-weight: 700;
          width: 100%;
          text-align: center;
          white-space: nowrap;
          overflow: hidden;
          text-overflow: ellipsis;
        }
        .preview-label .barcode-name {
          flex: 0 0 auto;
          font-size: 7px;
          line-height: 1.05;
          font-weight: 600;
          width: 100%;
          text-align: center;
          white-space: nowrap;
          overflow: hidden;
          text-overflow: ellipsis;
        }
        .preview-label .barcode-meta {
          flex: 0 0 auto;
          width: 100%;
          display: flex;
          align-items: baseline;
          justify-content: center;
          gap: 4px;
          line-height: 1;
          white-space: nowrap;
        }
        .preview-label .barcode-price {
          font-size: 7.6px;
          line-height: 1;
          font-weight: 800;
        }
        .preview-label .barcode-unit {
          font-size: 5.2px;
          line-height: 1;
          font-weight: 600;
        }
        .preview-label .barcode-svg {
          flex: 0 0 auto;
          display: block;
          width: 100%;
          height: 7.2mm;
          line-height: 0;
          margin: 0;
        }
        .preview-label .barcode-svg svg {
          display: block;
          width: 100%;
          height: 100%;
        }
        .preview-label .barcode-number {
          flex: 0 0 auto;
          font-family: Arial, sans-serif;
          font-size: 6.2px;
          line-height: 1;
          font-weight: 700;
          letter-spacing: 0;
          white-space: nowrap;
          text-align: center;
          width: 100%;
          margin-top: 0.1mm;
        }
        .barcode-preview-grid[data-size="large"] .preview-label .barcode-svg { height: 9.3mm; }
        .barcode-preview-grid[data-size="large"] .preview-label .barcode-shop { font-size: 7px; }
        .barcode-preview-grid[data-size="large"] .preview-label .barcode-name { font-size: 8px; }
        .barcode-preview-grid[data-size="large"] .preview-label .barcode-price { font-size: 8.5px; }
        .barcode-preview-grid[data-size="large"] .preview-label .barcode-unit { font-size: 5.8px; }
        .barcode-preview-grid[data-size="large"] .preview-label .barcode-number { font-size: 6.8px; }
        .table-pagination {
          display: flex; justify-content: space-between; align-items: center; gap: 12px;
          margin: 10px 0 20px; padding: 10px 12px; background: #f8fafc;
          border: 1px solid #e5e7eb; border-radius: 8px; font-size: 12px; color: #475569;
        }
        .table-pagination button {
          padding: 6px 9px; border: 1px solid #cbd5e1; border-radius: 5px;
          background: #fff; cursor: pointer; font-weight: 600;
        }
        .table-pagination button:disabled { opacity: .45; cursor: not-allowed; }
        /* ===== Receipt Design Presets (print + preview) ===== */
        .receipt-design-classic .receipt-meta-box { background: #fff; }
        .receipt-design-modern .receipt-shop-name { letter-spacing: .6px !important; text-transform: uppercase; }
        .receipt-design-modern .receipt-meta-box { border-style: solid !important; border-width: 1px !important; padding: 2mm !important; }
        .receipt-design-modern .receipt-items-heading { border-bottom-style: solid !important; }
        .receipt-design-bold .receipt-shop-name { font-size: 1.9em !important; }
        .receipt-design-bold .receipt-totals { font-weight: 900 !important; }
        .receipt-design-bold .receipt-grand-total { letter-spacing: .2px; }
        .receipt-design-compact { line-height: 1.05 !important; }
        .receipt-design-compact .receipt-shop-name { font-size: 1.45em !important; }
        .receipt-design-compact .receipt-meta-box { padding: 1mm 0 !important; margin-bottom: 1.2mm !important; }
        .receipt-design-compact .thermal-item-row { margin-bottom: 1mm !important; }
        .receipt-design-elegant .receipt-shop-name { letter-spacing: 1px !important; }
        .receipt-design-elegant .receipt-meta-box { border-top-style: double !important; border-bottom-style: double !important; }
        .receipt-design-elegant .receipt-footer-message { letter-spacing: .35px !important; }

        /* ===== Billing workspace visual polish only ===== */
        .billing-workspace {
          background: linear-gradient(180deg, #f8fafc 0%, #eef2ff 100%) !important;
        }
        .billing-panel {
          border: 1px solid #dbeafe !important;
          box-shadow: 0 8px 24px rgba(15,23,42,.07) !important;
        }
        .billing-panel-title {
          font-weight: 900 !important;
          letter-spacing: .2px !important;
        }
        .billing-total-highlight {
          box-shadow: inset 0 0 0 1px #bbf7d0, 0 4px 14px rgba(5,150,105,.08) !important;
        }
        .receipt-preview-card {
          box-shadow: 0 10px 30px rgba(15,23,42,.10) !important;
        }

                .receipt-preview-overlay .receipt-preview-card { flex: 0 0 auto !important; }
        .receipt-preview-overlay .thermal-item-row {
          margin-bottom: 2.6mm !important;
          padding-bottom: 1.2mm !important;
          border-bottom: .6px dashed #9ca3af !important;
        }
        .receipt-preview-overlay .thermal-item-name {
          font-size: 1.12em !important;
          line-height: 1.25 !important;
          font-weight: 900 !important;
        }

        @keyframes poshithaWelcomeFade {
          0%, 88% { opacity: 1; }
          100% { opacity: 0; pointer-events: none; }
        }
        @keyframes poshithaWelcomeZoom {
          0% { opacity: 0; transform: scale(.72) translateY(18px); filter: blur(6px); }
          35% { opacity: 1; transform: scale(1.04) translateY(0); filter: blur(0); }
          75% { opacity: 1; transform: scale(1); }
          100% { opacity: 0; transform: scale(1.08); }
        }
        @keyframes poshithaWelcomeGlow {
          0%,100% { transform: scale(.85); opacity: .25; }
          50% { transform: scale(1.15); opacity: .65; }
        }
        .poshitha-welcome-overlay {
          position: fixed !important; inset: 0 !important; z-index: 999999 !important;
          display: flex !important; align-items: center !important; justify-content: center !important;
          background: radial-gradient(circle at center, #172554 0%, #020617 72%) !important;
          animation: poshithaWelcomeFade 7s ease forwards !important;
        }
        .poshitha-welcome-glow {
          position: absolute; width: 280px; height: 280px; border-radius: 50%;
          background: radial-gradient(circle, rgba(59,130,246,.35), transparent 68%);
          animation: poshithaWelcomeGlow 1.4s ease-in-out infinite;
        }
        .poshitha-welcome-card { position: relative; text-align: center; color: #fff; animation: poshithaWelcomeZoom 7s cubic-bezier(.2,.8,.2,1) forwards; }
        .poshitha-welcome-logo {
          width: 92px; height: 92px; margin: 0 auto 18px; border-radius: 26px;
          display: flex; align-items: center; justify-content: center; font-size: 34px; font-weight: 1000;
          background: linear-gradient(135deg, #2563eb, #7c3aed, #059669);
          box-shadow: 0 0 45px rgba(59,130,246,.45);
        }
        .poshitha-welcome-title { margin: 0; font-size: clamp(24px, 4vw, 38px); font-weight: 1000; letter-spacing: 1.5px; }
        .poshitha-welcome-subtitle { margin-top: 9px; color: #bfdbfe; font-size: 14px; font-weight: 800; letter-spacing: .6px; }

@media (max-width: 800px) { .table-pagination { flex-direction: column; align-items: stretch; } }
        @media print {
          @page { size: ${receiptSettings.paperWidth === '58mm' ? '58mm' : '80mm'} auto; margin: 0; }
          body * { visibility: hidden !important; }
          .printable-receipt.active-print-receipt, .printable-receipt.active-print-receipt * { visibility: visible !important; }
          .credit-payment-receipt, .credit-payment-receipt * { visibility: visible !important; }
          .credit-book-print-receipt, .credit-book-print-receipt * { visibility: visible !important; }
          .barcode-print-sheet, .barcode-print-sheet * { visibility: visible !important; }
          .barcode-print-sheet {
            display: grid !important; position: absolute !important; left: 0 !important; top: 0 !important;
            width: 210mm !important; min-height: 297mm !important; padding: 8mm !important; margin: 0 !important;
            background: #fff !important; z-index: 1000000 !important; box-sizing: border-box !important;
          }
          .barcode-label {
            box-sizing: border-box !important;
            border: 0.45mm dashed #222 !important;
            width: 30mm !important;
            height: 20mm !important;
            padding: 1mm 1.15mm 0.8mm !important;
            background: #fff !important;
            overflow: hidden !important;
            display: flex !important;
            flex-direction: column !important;
            align-items: center !important;
            justify-content: flex-start !important;
            gap: 0.55mm !important;
            color: #000 !important;
            font-family: Arial, sans-serif !important;
            break-inside: avoid !important;
            page-break-inside: avoid !important;
          }
          .barcode-print-sheet[data-size="small"] {
            grid-template-columns: repeat(6, 30mm) !important;
            grid-auto-rows: 20mm !important;
            gap: 0 !important;
            align-content: start !important;
          }
          .barcode-print-sheet[data-size="large"] {
            grid-template-columns: repeat(5, 38mm) !important;
            grid-auto-rows: 25mm !important;
            gap: 0 !important;
            align-content: start !important;
          }
          .barcode-print-sheet[data-size="large"] .barcode-label {
            width: 38mm !important;
            height: 25mm !important;
            padding: 1.2mm 1.3mm 1mm !important;
            gap: 0.7mm !important;
          }
          .barcode-label .barcode-shop {
            flex: 0 0 auto !important;
            font-size: 5.2pt !important;
            line-height: 1 !important;
            font-weight: 700 !important;
            width: 100% !important;
            text-align: center !important;
            white-space: nowrap !important;
            overflow: hidden !important;
            text-overflow: ellipsis !important;
          }
          .barcode-label .barcode-name {
            flex: 0 0 auto !important;
            font-size: 5.8pt !important;
            line-height: 1.05 !important;
            font-weight: 600 !important;
            width: 100% !important;
            text-align: center !important;
            white-space: nowrap !important;
            overflow: hidden !important;
            text-overflow: ellipsis !important;
          }
          .barcode-label .barcode-meta {
            flex: 0 0 auto !important;
            width: 100% !important;
            display: flex !important;
            align-items: baseline !important;
            justify-content: center !important;
            gap: 1.2mm !important;
            line-height: 1 !important;
            white-space: nowrap !important;
          }
          .barcode-label .barcode-price {
            font-size: 6.6pt !important;
            line-height: 1 !important;
            font-weight: 800 !important;
          }
          .barcode-label .barcode-unit {
            font-size: 4.4pt !important;
            line-height: 1 !important;
            font-weight: 600 !important;
          }
          .barcode-label .barcode-svg {
            flex: 0 0 auto !important;
            display: block !important;
            width: 100% !important;
            height: 7.2mm !important;
            line-height: 0 !important;
            margin: 0 !important;
          }
          .barcode-label .barcode-svg svg {
            display: block !important;
            width: 100% !important;
            height: 100% !important;
          }
          .barcode-label .barcode-number {
            flex: 0 0 auto !important;
            font-family: Arial, sans-serif !important;
            font-size: 5.4pt !important;
            line-height: 1 !important;
            font-weight: 700 !important;
            letter-spacing: 0 !important;
            white-space: nowrap !important;
            text-align: center !important;
            width: 100% !important;
            margin-top: 0.1mm !important;
          }
          .barcode-print-sheet[data-size="large"] .barcode-label .barcode-svg { height: 9.3mm !important; }
          .barcode-print-sheet[data-size="large"] .barcode-label .barcode-shop { font-size: 5.8pt !important; }
          .barcode-print-sheet[data-size="large"] .barcode-label .barcode-name { font-size: 6.5pt !important; }
          .barcode-print-sheet[data-size="large"] .barcode-label .barcode-price { font-size: 7.2pt !important; }
          .barcode-print-sheet[data-size="large"] .barcode-label .barcode-unit { font-size: 4.8pt !important; }
          .barcode-print-sheet[data-size="large"] .barcode-label .barcode-number { font-size: 5.8pt !important; }
          /* ===== PROFESSIONAL LABEL-PRINTER MODE ===== */
          .barcode-custom-print-sheet { display: none; }
          .barcode-custom-preview-grid { display: grid; grid-template-columns: repeat(auto-fill, minmax(150px, 1fr)); gap: 14px; align-items: start; }
          .barcode-custom-preview-card { display: flex; align-items: center; justify-content: center; min-height: 125px; padding: 10px; background: #e2e8f0; border-radius: 10px; overflow: hidden; }
          /* Custom label mode deliberately reuses the existing .barcode-label design.
             Only the physical width/height changes; typography, order and proportions stay identical. */
          .barcode-custom-label {
            box-sizing: border-box !important;
            width: ${Math.max(20, Math.min(100, parseFloat(barcodeCustomWidth) || 40))}mm !important;
            height: ${Math.max(15, Math.min(100, parseFloat(barcodeCustomHeight) || 30))}mm !important;
          }
          .barcode-custom-label .barcode-shop,
          .barcode-custom-label .barcode-name,
          .barcode-custom-label .barcode-meta,
          .barcode-custom-label .barcode-price,
          .barcode-custom-label .barcode-unit,
          .barcode-custom-label .barcode-svg,
          .barcode-custom-label .barcode-number { box-sizing: border-box; }

          .credit-book-print-receipt {
            display: block !important;
            position: absolute !important;
            left: 0 !important;
            top: 0 !important;
            width: ${receiptSettings.paperWidth === '58mm' ? '58mm' : '80mm'} !important;
            max-width: ${receiptSettings.paperWidth === '58mm' ? '58mm' : '80mm'} !important;
            margin: 0 !important;
            padding: 2.2mm !important;
            box-sizing: border-box !important;
            background: #fff !important;
            z-index: 1000002 !important;
            font-family: ${receiptSettings.fontFamily} !important;
            font-size: ${receiptSettings.fontSize} !important;
            font-weight: 700 !important;
            overflow: visible !important;
            word-break: normal !important;
            overflow-wrap: anywhere !important;
          }
          .credit-payment-receipt {
            display: block !important;
            position: absolute !important;
            left: 0 !important;
            top: 0 !important;
            width: ${receiptSettings.paperWidth === '58mm' ? '58mm' : '80mm'} !important;
            max-width: ${receiptSettings.paperWidth === '58mm' ? '58mm' : '80mm'} !important;
            margin: 0 !important;
            padding: 2.2mm !important;
            box-sizing: border-box !important;
            background: #fff !important;
            z-index: 1000001 !important;
            font-family: ${receiptSettings.fontFamily} !important;
            font-size: ${receiptSettings.fontSize} !important;
            font-weight: 700 !important;
            overflow: visible !important;
            word-break: normal !important;
            overflow-wrap: anywhere !important;
          }
          .printable-receipt.active-print-receipt {
            display: block !important;
            position: absolute !important;
            left: 0 !important;
            top: 0 !important;
            width: ${receiptSettings.paperWidth === '58mm' ? '58mm' : '80mm'} !important;
            max-width: ${receiptSettings.paperWidth === '58mm' ? '58mm' : '80mm'} !important;
            margin: 0 !important;
            padding: 2mm !important;
            box-sizing: border-box !important;
            background: #fff !important;
            z-index: 999999 !important;
            font-family: ${receiptSettings.fontFamily} !important;
            font-size: ${receiptSettings.fontSize} !important;
            font-weight: 700 !important;
            overflow: visible !important;
            word-break: normal !important;
            overflow-wrap: anywhere !important;
          }
          /* Thermal-safe width: keep every receipt element inside the selected roll. */
          .printable-receipt.active-print-receipt,
          .credit-book-print-receipt,
          .credit-payment-receipt {
            width: ${receiptSettings.paperWidth === '58mm' ? '58mm' : '80mm'} !important;
            max-width: ${receiptSettings.paperWidth === '58mm' ? '58mm' : '80mm'} !important;
            box-sizing: border-box !important;
            overflow-x: hidden !important;
          }
          .thermal-receipt-content,
          .printable-receipt.active-print-receipt *,
          .credit-book-print-receipt *,
          .credit-payment-receipt * {
            max-width: 100% !important;
            box-sizing: border-box !important;
          }
          .printable-receipt.active-print-receipt img,
          .credit-book-print-receipt img,
          .credit-payment-receipt img {
            max-width: 100% !important;
            height: auto !important;
          }
          /* Prevent fixed-width children from pushing text outside 58/80mm. */
          .printable-receipt.active-print-receipt table,
          .credit-book-print-receipt table,
          .credit-payment-receipt table {
            width: 100% !important;
            max-width: 100% !important;
            table-layout: fixed !important;
          }
          .thermal-receipt-content {
            width: 100% !important;
            max-width: 100% !important;
            box-sizing: border-box !important;
            overflow: visible !important;
          }
          .thermal-receipt-logo {
            width: ${receiptSettings.logoSize || '40mm'} !important;
            height: ${receiptSettings.logoSize || '40mm'} !important;
            min-width: ${receiptSettings.logoSize || '40mm'} !important;
            min-height: ${receiptSettings.logoSize || '40mm'} !important;
            max-width: ${receiptSettings.logoSize || '40mm'} !important;
            max-height: ${receiptSettings.logoSize || '40mm'} !important;
            object-fit: contain !important;
            border-radius: 0 !important;
            display: block !important;
            margin: 0 auto !important;
            transform: none !important;
          }
          .printable-receipt.active-print-receipt, .credit-book-print-receipt {
            transform: none !important;
            zoom: 1 !important;
            -webkit-print-color-adjust: exact !important;
            print-color-adjust: exact !important;
          }
          .thermal-row {
            display: grid !important;
            grid-template-columns: minmax(0, 1fr) auto !important;
            column-gap: 2mm !important;
            width: 100% !important;
            box-sizing: border-box !important;
          }
          .thermal-item-row {
            break-inside: avoid !important;
            page-break-inside: avoid !important;
            margin-bottom: 1.8mm !important;
            width: 100% !important;
            min-width: 0 !important;
            overflow: visible !important;
          }
          .thermal-item-row > div:first-child, .receipt-items-heading {
            display: grid !important;
            /* Qty + amount are fixed; item name gets all remaining width. */
            grid-template-columns: minmax(0, 1fr) 8mm 18mm !important;
            column-gap: 0.8mm !important;
            width: 100% !important;
            max-width: 100% !important;
            min-width: 0 !important;
            box-sizing: border-box !important;
          }
          .thermal-item-row > div:first-child > *, .receipt-items-heading > * {
            min-width: 0 !important;
            max-width: 100% !important;
          }
          .thermal-item-name {
            min-width: 0 !important;
            max-width: 100% !important;
            overflow-wrap: anywhere !important;
            word-break: break-word !important;
            white-space: normal !important;
          }
          .thermal-item-name + .thermal-number, .thermal-item-name ~ .thermal-number {
            white-space: nowrap !important;
          }
          .thermal-number {
            white-space: nowrap !important;
            text-align: right !important;
          }
          .thermal-grand-total {
            break-inside: avoid !important;
            page-break-inside: avoid !important;
          }
          .no-print { display: none !important; }
          .printable-receipt.active-print-receipt, .credit-book-print-receipt {
            height: auto !important;
            min-height: 0 !important;
            max-height: none !important;
          }
          @page {
            size: ${receiptSettings.paperWidth === '58mm' ? '58mm' : '80mm'} auto;
            margin: 0 !important;
          }
        }
      `}</style>
      {/* ↩️ POS STANDARD RETURNS PAGE */}
      {activeTab === 'returns' && (
        <div style={{ flex: 1, order: 1, padding: '20px', maxWidth: '1200px', margin: '0 auto', overflowY: 'auto' }}>
          <div style={{ background: '#fff', border: '1px solid #e5e7eb', borderRadius: '10px', padding: '20px', boxShadow: '0 2px 8px rgba(0,0,0,.06)' }}>
            <div style={{ marginBottom: '20px' }}>
              <div style={{ display: 'flex', justifyContent: 'center', alignItems: 'center', gap: '10px', flexWrap: 'wrap', marginBottom: '14px' }}>
                <button onClick={() => setActiveTab('home')} style={{ padding: '9px 14px', border: 'none', background: '#2563eb', color:'#fff', borderRadius: '6px', cursor: 'pointer', fontWeight:'bold' }}>🏠 Home</button>
                <button onClick={handleDeleteAllReturnRecords} disabled={!returnRecords.length} style={{ padding: '9px 14px', border: 'none', background: returnRecords.length ? '#dc2626' : '#9ca3af', color:'#fff', borderRadius: '6px', cursor: returnRecords.length ? 'pointer':'not-allowed', fontWeight:'bold' }}>🗑️ {ui.deleteAll}</button>
              </div>
              <div style={{ textAlign: 'center' }}><h2 style={{ margin: 0 }}>↩️ Sales Return / භාණ්ඩ ආපසු ගැනීම</h2><p style={{ margin: '5px 0 0', color: '#6b7280' }}>Bill No එකෙන් sale එක සොයා partial හෝ full item return කරන්න.</p></div>
            </div>
            <div style={{ display: 'flex', gap: '10px', marginBottom: '20px' }}>
              <input type="text" placeholder={uiMore.billNo} value={searchBillNo} onChange={(e) => setSearchBillNo(e.target.value)} onKeyDown={(e) => e.key === 'Enter' && handleSearchBill()} style={{ flex: 1, padding: '12px', fontSize: '15px', border: '1px solid #cbd5e1', borderRadius: '6px' }} />
              <button onClick={handleSearchBill} style={{ padding: '12px 24px', background: '#2563eb', color: '#fff', border: 'none', borderRadius: '6px', fontWeight: 'bold', cursor: 'pointer' }}>🔍 {ui.searchBill}</button>
            </div>
            {returnBill && (
              <>
                <div style={{ display: 'grid', gridTemplateColumns: 'repeat(4, 1fr)', gap: '12px', marginBottom: '20px', background: '#f8fafc', padding: '14px', borderRadius: '8px' }}>
                  <div><small>{uiMore.billNo}</small><div style={{ fontWeight: 'bold' }}>{returnBill.invoiceNo ?? returnBill.invoice_no ?? returnBill.id}</div></div>
                  <div><small>{uiMore.date}</small><div style={{ fontWeight: 'bold' }}>{returnBill.date || returnBill.dateFormatted || '-'}</div></div>
                  <div><small>{uiMore.customer}</small><div style={{ fontWeight: 'bold' }}>{returnBill.customerName || 'General Customer'}</div></div>
                  <div><small>{uiMore.originalTotal}</small><div style={{ fontWeight: 'bold' }}>Rs. {Number(returnBill.total || 0).toFixed(2)}</div></div>
                </div>
                <div style={{ overflowX: 'auto' }}>
                  <table style={{ width: '100%', borderCollapse: 'collapse' }}>
                    <thead><tr style={{ background: '#f1f5f9', textAlign: 'left' }}><th style={{ padding: '10px' }}>Return</th><th style={{ padding: '10px' }}>Item</th><th style={{ padding: '10px' }}>{uiMore.soldQty}</th><th style={{ padding: '10px' }}>{uiMore.available}</th><th style={{ padding: '10px' }}>{uiMore.returnQty}</th><th style={{ padding: '10px' }}>{ui.price}</th></tr></thead>
                    <tbody>{(returnBill.items || []).map(item => {
                      const billNo = returnBill.invoiceNo ?? returnBill.invoice_no ?? returnBill.id;
                      const key = getReturnItemKey(billNo, item); const state = returnItems[key] || {}; const soldQty = Number(item.qty ?? item.quantity ?? 0);
                      return <tr key={key} style={{ borderBottom: '1px solid #e5e7eb' }}>
                        <td style={{ padding: '10px' }}><input type="checkbox" checked={!!state.selected} disabled={!state.maxQty} onChange={e => updateReturnItem(key, { selected: e.target.checked })} /></td>
                        <td style={{ padding: '10px', fontWeight: 'bold' }}>{item.name}<div style={{ fontSize: '12px', color: '#6b7280' }}>{item.barcode || ''}</div>{item.lot_number && <div style={{ marginTop: '3px', fontSize: '12px', color: '#7c3aed', fontWeight: '800' }}>Lot: {item.lot_number}</div>}</td>
                        <td style={{ padding: '10px' }}>{soldQty} {item.unit || ''}</td>
                        <td style={{ padding: '10px', color: state.maxQty ? '#059669' : '#dc2626' }}>{state.maxQty || 0}</td>
                        <td style={{ padding: '10px' }}><input type="number" min="0" max={state.maxQty || 0} step="any" value={state.qty ?? 0} disabled={!state.selected || !state.maxQty} onChange={e => updateReturnItem(key, { qty: e.target.value })} style={{ width: '90px', padding: '7px', border: '1px solid #cbd5e1', borderRadius: '5px' }} /></td>
                        <td style={{ padding: '10px' }}>Rs. {Number(item.price || 0).toFixed(2)}</td>
                      </tr>;
                    })}</tbody>
                  </table>
                </div>
                <div style={{ display: 'flex', gap: '15px', alignItems: 'center', marginTop: '20px', flexWrap: 'wrap' }}>
                  <select value={returnReason} onChange={e => setReturnReason(e.target.value)} style={{ padding: '11px', border: '1px solid #cbd5e1', borderRadius: '6px' }}><option>{uiMore.customerRequest}</option><option>{uiMore.damagedItem}</option><option>{uiMore.wrongItem}</option><option>{uiMore.expiredItem}</option><option>{uiMore.other}</option></select>
                  <button onClick={handleProcessReturn} style={{ padding: '12px 24px', background: '#dc2626', color: '#fff', border: 'none', borderRadius: '6px', fontWeight: 'bold', cursor: 'pointer' }}>↩️ Process Selected Return</button>
                </div>
              </>
            )}
          </div>
          {returnRecords.length > 0 && (
            <div style={{ marginTop: '20px', background: '#fff', border: '1px solid #e5e7eb', borderRadius: '10px', padding: '18px' }}>
              <div style={{ marginBottom: '10px' }}>
                <h3 style={{ margin: 0 }}>🧾 Recent Returns</h3>
              </div>
              {returnRecords.map(r => (
                <div key={r.id} style={{ padding: '10px 0', borderBottom: '1px solid #f1f5f9', display: 'flex', justifyContent: 'space-between', alignItems: 'center', gap: '10px' }}>
                  <span><b>{r.id}</b> • Bill #{r.originalBillNo} • {r.reason}<div style={{ fontSize: '12px', color: '#6b7280' }}>{r.date}</div></span>
                  <div style={{ display: 'flex', alignItems: 'center', gap: '12px' }}><span><b>Rs. {Number(r.total || 0).toFixed(2)}</b></span><button onClick={() => handleDeleteReturnRecord(r.id)} style={{ padding: '7px 11px', background: '#fff', color: '#dc2626', border: '1px solid #fecaca', borderRadius: '6px', cursor: 'pointer' }}>🗑️ Delete</button></div>
                </div>
              ))}
            </div>
          )}
        </div>
      )}
      {/* 🧾 PRINTABLE RECEIPT - optimized for 58mm / 80mm thermal rolls */}
      <div
        className={`printable-receipt ${printMode === 'sale' ? 'active-print-receipt' : ''} receipt-design-${receiptSettings.receiptDesign || 'classic'}`}
        style={{
          display: 'none',
          width: receiptSettings.paperWidth === '58mm' ? '58mm' : '80mm',
          maxWidth: receiptSettings.paperWidth === '58mm' ? '58mm' : '80mm',
          boxSizing: 'border-box',
          fontFamily: receiptSettings.fontFamily,
          fontSize: receiptSettings.fontSize,
          color: '#000',
          padding: '2.2mm',
          backgroundColor: '#fff',
          fontWeight: 700,
          overflow: 'visible',
          overflowWrap: 'anywhere'
        }}
      >
        {renderReceiptContent()}
      </div>

      {creditBookPrintSale && (
        <div
          className={`credit-book-print-receipt receipt-design-${receiptSettings.receiptDesign || 'classic'}`}
          style={{
            display: 'none',
            width: receiptSettings.paperWidth === '58mm' ? '58mm' : '80mm',
            maxWidth: receiptSettings.paperWidth === '58mm' ? '58mm' : '80mm',
            boxSizing: 'border-box',
            fontFamily: receiptSettings.fontFamily,
            fontSize: receiptSettings.fontSize,
            color: '#000',
            padding: '2.2mm',
            background: '#fff',
            fontWeight: 700,
            overflow: 'visible',
            overflowWrap: 'anywhere'
          }}
        >
          {renderReceiptContent(creditBookPrintSale)}
        </div>
      )}

      {creditPaymentReceipt && (
        <div className="credit-payment-receipt" style={{ display: 'none', width: receiptSettings.paperWidth === '58mm' ? '58mm' : '80mm', fontFamily: receiptSettings.fontFamily, fontSize: receiptSettings.fontSize, color: '#000', padding: '5px', background: '#fff', fontWeight: 'bold' }}>
          <div style={{ textAlign: 'center', marginBottom: '10px' }}>
            {receiptSettings.logoUrl && <img className="thermal-receipt-logo" src={receiptSettings.logoUrl} alt="Logo" style={{ width: receiptSettings.logoSize || '40mm', height: receiptSettings.logoSize || '40mm', minWidth: receiptSettings.logoSize || '40mm', minHeight: receiptSettings.logoSize || '40mm', maxWidth: receiptSettings.logoSize || '40mm', maxHeight: receiptSettings.logoSize || '40mm', objectFit: 'contain', borderRadius: '0', display: 'block', margin: '0 auto' }} />}
            <h2 style={{ margin: '4px 0' }}>{receiptSettings.shopName}</h2>
            <div style={{ whiteSpace: 'pre-line' }}>{receiptSettings.address}</div>
            <div>Tel: {receiptSettings.phone}</div>
          </div>
          <div style={{ borderTop: '2px dashed #000', borderBottom: '2px dashed #000', padding: '7px 0', marginBottom: '8px' }}>
            <div style={{ textAlign: 'center', fontSize: '1.2em' }}>CREDIT PAYMENT RECEIPT</div>
            <div style={{ display: 'flex', justifyContent: 'space-between' }}><span>Receipt No:</span><span>{creditPaymentReceipt.id}</span></div>
            <div style={{ display: 'flex', justifyContent: 'space-between' }}><span>Date:</span><span>{creditPaymentReceipt.dateFormatted}</span></div>
          </div>
          <div>Customer: {creditPaymentReceipt.customerName}</div>
          <div>Phone: {creditPaymentReceipt.phone}</div>
          <div style={{ borderTop: '1px dashed #000', marginTop: '8px', paddingTop: '8px' }}>
            <div style={{ display: 'flex', justifyContent: 'space-between' }}><span>Previous Due:</span><span>Rs. {creditPaymentReceipt.previousBalance.toFixed(2)}</span></div>
            <div style={{ display: 'flex', justifyContent: 'space-between' }}><span>Payment:</span><span>Rs. {creditPaymentReceipt.paidAmount.toFixed(2)}</span></div>
            <div style={{ display: 'flex', justifyContent: 'space-between', fontSize: '1.25em', borderTop: '2px solid #000', borderBottom: '2px solid #000', padding: '5px 0', marginTop: '5px' }}><span>Remaining Due:</span><span>Rs. {creditPaymentReceipt.remainingBalance.toFixed(2)}</span></div>
          </div>
          <div style={{ textAlign: 'center', marginTop: '14px' }}>
            <div>{receiptSettings.footerMsg}</div>
            {receiptSettings.footerComment && (
              <div style={{ marginTop: '6px', whiteSpace: 'pre-line', overflowWrap: 'anywhere' }}>{receiptSettings.footerComment}</div>
            )}
          </div>
        </div>
      )}

      {showCustomerRegisterModal && (
        <div className="no-print" style={{ position: 'fixed', inset: 0, background: 'rgba(0,0,0,.6)', zIndex: 6000, display: 'flex', alignItems: 'center', justifyContent: 'center', padding: '20px' }}>
          <div style={{ background: '#fff', width: '440px', maxWidth: '95vw', padding: '22px', borderRadius: '12px' }}>
            <h3 style={{ marginTop: 0 }}>👤 Register Credit Customer</h3>
            <div style={{ display: 'grid', gap: '9px' }}>
              <input placeholder="Customer Name *" value={customerForm.name} onChange={e => setCustomerForm({...customerForm, name:e.target.value})} style={{ padding:'10px', border:'1px solid #cbd5e1', borderRadius:'6px' }} />
              <input placeholder="Phone Number *" value={customerForm.phone} onChange={e => setCustomerForm({...customerForm, phone:e.target.value})} style={{ padding:'10px', border:'1px solid #cbd5e1', borderRadius:'6px' }} />
              <input placeholder="Address (Optional)" value={customerForm.address} onChange={e => setCustomerForm({...customerForm, address:e.target.value})} style={{ padding:'10px', border:'1px solid #cbd5e1', borderRadius:'6px' }} />
              <input type="number" placeholder="Credit Limit (Optional)" value={customerForm.creditLimit} onChange={e => setCustomerForm({...customerForm, creditLimit:e.target.value})} style={{ padding:'10px', border:'1px solid #cbd5e1', borderRadius:'6px' }} />
              <select value={customerForm.priceLevel} onChange={e => setCustomerForm({...customerForm, priceLevel:e.target.value})} style={{ padding:'10px', border:'1px solid #cbd5e1', borderRadius:'6px' }}>
                <option value="normal">Normal Customer — Retail</option>
                <option value="wholesale">Wholesale Customer</option>
                <option value="special">Special Customer</option>
              </select>
            </div>
            <div style={{ display:'flex', gap:'8px', marginTop:'15px' }}>
              <button onClick={() => setShowCustomerRegisterModal(false)} style={{ flex:1, padding:'10px', background:'#6b7280', color:'#fff', border:0, borderRadius:'6px' }}>Cancel</button>
              <button onClick={handleSaveCustomer} style={{ flex:1, padding:'10px', background:'#059669', color:'#fff', border:0, borderRadius:'6px', fontWeight:'bold' }}>Save Customer</button>
            </div>
          </div>
        </div>
      )}

      {showCreditPaymentModal && (
        <div className="no-print" style={{ position:'fixed', inset:0, background:'rgba(0,0,0,.6)', zIndex:6000, display:'flex', alignItems:'center', justifyContent:'center', padding:'20px' }}>
          <div style={{ background:'#fff', width:'520px', maxWidth:'95vw', padding:'22px', borderRadius:'12px' }}>
            <h3 style={{ marginTop:0 }}>💰 Receive Customer Credit Payment</h3>
            <div style={{ display:'flex', gap:'8px' }}>
              <input autoFocus placeholder="Customer ID or Phone Number" value={creditPaymentSearch} onChange={e=>setCreditPaymentSearch(e.target.value)} onKeyDown={e=>e.key==='Enter'&&handleSearchCreditCustomer()} style={{ flex:1, padding:'10px', border:'1px solid #cbd5e1', borderRadius:'6px' }} />
              <button onClick={handleSearchCreditCustomer} style={{ padding:'10px 14px', background:'#2563eb', color:'#fff', border:0, borderRadius:'6px', fontWeight:'bold' }}>🔍 Search</button>
            </div>
            {creditPaymentCustomer && <div style={{ marginTop:'15px', padding:'12px', background:'#f8fafc', border:'1px solid #e2e8f0', borderRadius:'8px' }}>
              <div style={{ fontWeight:'bold', fontSize:'16px' }}>{creditPaymentCustomer.name}</div><div>{creditPaymentCustomer.phone} • ID: {creditPaymentCustomer.id}</div>
              <div style={{ marginTop:'6px', color:'#b45309', fontWeight:'bold' }}>Current Outstanding: Rs. {getCustomerBalance(creditPaymentCustomer.id).toFixed(2)}</div>
              <input type="number" min="0.01" placeholder="Payment Amount" value={creditPaymentAmount} onChange={e=>setCreditPaymentAmount(e.target.value)} style={{ width:'100%', boxSizing:'border-box', marginTop:'10px', padding:'11px', border:'2px solid #10b981', borderRadius:'6px', fontSize:'16px', fontWeight:'bold' }} />
            </div>}
            <div style={{ display:'flex', gap:'8px', marginTop:'15px' }}><button onClick={()=>setShowCreditPaymentModal(false)} style={{ flex:1, padding:'10px', background:'#6b7280', color:'#fff', border:0, borderRadius:'6px' }}>Cancel</button><button onClick={handleSaveCreditPayment} disabled={!creditPaymentCustomer} style={{ flex:1, padding:'10px', background:creditPaymentCustomer?'#059669':'#9ca3af', color:'#fff', border:0, borderRadius:'6px', fontWeight:'bold' }}>Save & Print Receipt</button></div>
          </div>
        </div>
      )}

      {showPreviewModal && (
        <div className="no-print receipt-preview-overlay" style={{ position: 'fixed', inset: 0, background: 'rgba(15,23,42,.78)', display: 'flex', justifyContent: 'center', alignItems: 'center', zIndex: 9000, padding: '20px', overflowY: 'auto' }}>
          <div style={{ width: 'min(760px, 96vw)', maxHeight: '94vh', overflowY: 'auto', background: '#e2e8f0', borderRadius: '16px', padding: '18px', boxShadow: '0 24px 80px rgba(0,0,0,.40)' }}>
            <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', gap: '12px', marginBottom: '14px' }}>
              <div>
                <h3 style={{ margin: 0, color: '#0f172a', fontWeight: 900 }}>🧾 BILL PREVIEW — EXACT PRINT LAYOUT</h3>
                <div style={{ marginTop: '4px', fontSize: '12px', color: '#475569', fontWeight: 700 }}>{receiptSettings.paperWidth} • {receiptSettings.fontSize} • {receiptSettings.receiptDesign || 'classic'}</div>
              </div>
              <span style={{ padding: '6px 10px', background: '#dcfce7', color: '#166534', borderRadius: '999px', fontSize: '11px', fontWeight: 900 }}>PRINT MATCH</span>
            </div>
            <div style={{ display: 'flex', gap: '10px', marginBottom: '14px' }}>
              <button
                ref={previewCancelRef}
                data-pos-nav="true"
                type="button"
                onClick={() => setShowPreviewModal(false)}
                style={{ flex: 1, padding: '12px', backgroundColor: '#64748b', color: '#fff', border: 'none', borderRadius: '8px', fontWeight: 900, cursor: 'pointer' }}
              >❌ Cancel / Edit</button>
              <button
                ref={previewConfirmRef}
                data-pos-nav="true"
                type="button"
                onClick={handleConfirmAndSave}
                style={{ flex: 1, padding: '12px', backgroundColor: '#059669', color: '#fff', border: 'none', borderRadius: '8px', fontWeight: 900, cursor: 'pointer' }}
              >✅ Confirm & Open Bill</button>
            </div>
            <div className={`receipt-preview-card printable-receipt receipt-design-${receiptSettings.receiptDesign || 'classic'}`}
              style={{ display: 'block', width: receiptSettings.paperWidth === '58mm' ? '58mm' : '80mm', maxWidth: '100%', margin: '0 auto', boxSizing: 'border-box', fontFamily: receiptSettings.fontFamily, fontSize: receiptSettings.fontSize, color: '#000', padding: '2.2mm', backgroundColor: '#fff', fontWeight: 700, overflow: 'visible', overflowWrap: 'anywhere', boxShadow: '0 12px 30px rgba(15,23,42,.18)' }}>
              {renderReceiptContent()}
            </div>
          </div>
        </div>
      )}

      {barcodePrintItems.length > 0 && (
        <div className="barcode-print-sheet" data-size={barcodePrintSize}>
          {barcodePrintItems.map((item, index) => (
            <div className="barcode-label" key={`${item.barcode}-${index}`}>
              {/* Exact printed order: Shop Name -> Item Name -> Price + Unit -> Barcode Bars -> Barcode Number */}
              <div className="barcode-shop" title={receiptSettings.shopName}>{receiptSettings.shopName || 'POSHITHA POS SYSTEM'}</div>
              <div className="barcode-name" title={item.name}>{item.name}</div>
              <div className="barcode-meta">
                <span className="barcode-price">Rs. {Number(item.price || 0).toFixed(2)}</span>
                <span className="barcode-unit">{item.unit || 'Pcs'}</span>
              </div>
              <div
                className="barcode-svg"
                dangerouslySetInnerHTML={{ __html: barcodeSvgMarkup(item.barcode, barcodePrintSize === 'large' ? 250 : 205, barcodePrintSize === 'large' ? 58 : 43) }}
              />
              <div className="barcode-number" title={String(item.barcode || '')}>{item.barcode}</div>
            </div>
          ))}
        </div>
      )}

      {barcodePrintItems.length > 0 && (
        <div className="barcode-custom-print-sheet">
          <style>{`@media print { @page { size: ${Math.max(20, Math.min(100, parseFloat(barcodeCustomWidth) || 40))}mm ${Math.max(15, Math.min(100, parseFloat(barcodeCustomHeight) || 30))}mm; margin: 0; } .barcode-custom-print-sheet { display: block !important; position: fixed !important; left: 0 !important; top: 0 !important; width: ${Math.max(20, Math.min(100, parseFloat(barcodeCustomWidth) || 40))}mm !important; } .barcode-custom-print-page { width: ${Math.max(20, Math.min(100, parseFloat(barcodeCustomWidth) || 40))}mm !important; height: ${Math.max(15, Math.min(100, parseFloat(barcodeCustomHeight) || 30))}mm !important; margin: 0 0 ${Math.max(0, Math.min(10, parseFloat(barcodeCustomGap) || 0))}mm 0 !important; padding: 0 !important; box-sizing: border-box !important; break-after: page; page-break-after: always; display: flex !important; align-items: flex-start !important; justify-content: center !important; } .barcode-custom-print-page:last-child { break-after: auto; page-break-after: auto; } .barcode-custom-label.print-label { width: ${Math.max(20, Math.min(100, parseFloat(barcodeCustomWidth) || 40))}mm !important; height: ${Math.max(15, Math.min(100, parseFloat(barcodeCustomHeight) || 30))}mm !important; box-sizing: border-box !important; } .barcode-custom-print-sheet * { visibility: visible !important; } body > *:not(#root) { visibility: hidden !important; } #root > *:not(.barcode-custom-print-sheet) { visibility: hidden !important; } }`}</style>
          {barcodePrintItems.map((item, index) => (
            <div className="barcode-custom-print-page" key={`custom-print-${item.barcode}-${index}`}>
              <div className="barcode-label barcode-custom-label print-label">
                <div className="barcode-shop" title={receiptSettings.shopName}>{receiptSettings.shopName || 'POSHITHA POS SYSTEM'}</div>
                <div className="barcode-name" title={item.name}>{item.name}</div>
                <div className="barcode-meta"><span className="barcode-price">Rs. {Number(item.price || 0).toFixed(2)}</span><span className="barcode-unit">{item.unit || 'Pcs'}</span></div>
                <div className="barcode-svg" dangerouslySetInnerHTML={{ __html: barcodeSvgMarkup(item.barcode, 300, 70) }} />
                <div className="barcode-number" title={String(item.barcode || '')}>{item.barcode}</div>
              </div>
            </div>
          ))}
        </div>
      )}

      {showLotSelectModal && (
        <div style={{ position:'fixed', inset:0, background:'rgba(0,0,0,0.55)', zIndex:10050, display:'flex', alignItems:'center', justifyContent:'center', padding:'20px' }}>
          <div style={{ width:'min(720px, 96vw)', maxHeight:'85vh', overflowY:'auto', background:'#fff', borderRadius:'12px', boxShadow:'0 20px 60px rgba(0,0,0,0.35)', padding:'20px' }}>
            <div style={{ display:'flex', justifyContent:'space-between', alignItems:'center', gap:'12px', marginBottom:'14px' }}>
              <div>
                <h3 style={{ margin:'0 0 4px', color:'#1f2937' }}>📦 Select Stock Lot</h3>
                <div style={{ fontSize:'13px', color:'#6b7280' }}>Barcode / Item Code: <b>{lotSelectCode}</b> — මෙම customer ගත්ත Lot එක තෝරන්න.</div>
              </div>
              <button type="button" onClick={closeLotSelectModal} style={{ border:0, background:'#e5e7eb', borderRadius:'7px', padding:'8px 12px', cursor:'pointer', fontWeight:'bold' }}>✕ Close</button>
            </div>
            <div style={{ display:'grid', gap:'10px' }}>
              {lotSelectCandidates.map((lot, index) => {
                const expiryInfo = getExpiryInfo(lot);
                return (
                  <button
                    type="button"
                    key={`lot-choice-${lot.id}`}
                    onClick={() => selectLotForSale(lot)}
                    style={{ textAlign:'left', padding:'14px', border:'2px solid #ddd6fe', borderRadius:'10px', background:index === 0 ? '#faf5ff' : '#fff', cursor:'pointer' }}
                  >
                    <div style={{ display:'flex', justifyContent:'space-between', gap:'12px', flexWrap:'wrap' }}>
                      <div>
                        <div style={{ fontSize:'16px', fontWeight:'900', color:'#111827' }}>{lot.name}</div>
                        <div style={{ marginTop:'4px', color:'#7c3aed', fontWeight:'800' }}>Lot: {lot.lot_number || `ID-${lot.id}`}{index === 0 ? ' • Oldest available' : ''}</div>
                        <div style={{ marginTop:'4px', fontSize:'12px', color:'#6b7280' }}>GRN Date: {lot.grn_date || 'N/A'} | Expiry: {lot.expiry_date || 'N/A'} {expiryInfo.isExpired ? '🔴 EXPIRED' : expiryInfo.isNearExpiry ? '🟡 Near Expiry' : ''}</div>
                      </div>
                      <div style={{ minWidth:'220px', textAlign:'right' }}>
                        <div>Buy: <b>Rs. {Number(lot.buying_price || 0).toFixed(2)}</b></div>
                        <div style={{ color:'#059669', fontSize:'18px' }}>Sell: <b>Rs. {Number(lot.price || 0).toFixed(2)}</b></div>
                        <div>Stock: <b>{lot.stock} {lot.unit || 'Pcs'}</b></div>
                      </div>
                    </div>
                  </button>
                );
              })}
            </div>
          </div>
        </div>
      )}

      {showBarcodePrintChoiceModal && barcodePrintProduct && (
        <div className="no-print" style={{ position:'fixed', inset:0, zIndex:100002, background:'rgba(15,23,42,.72)', display:'flex', alignItems:'center', justifyContent:'center', padding:'20px' }}>
          <div style={{ width:'min(620px,100%)', background:'#fff', borderRadius:'16px', padding:'24px', boxShadow:'0 25px 80px rgba(0,0,0,.35)' }}>
            <div style={{ display:'flex', justifyContent:'space-between', alignItems:'flex-start', gap:'12px', marginBottom:'18px' }}>
              <div><h2 style={{ margin:'0 0 5px' }}>🖨️ Barcode Label Print</h2><div style={{ color:'#64748b', fontSize:'13px' }}>{barcodePrintProduct.name} — {barcodePrintProduct.barcode}</div></div>
              <button type="button" onClick={() => setShowBarcodePrintChoiceModal(false)} style={{ border:0, background:'#e2e8f0', borderRadius:'8px', padding:'8px 11px', cursor:'pointer', fontWeight:800 }}>✕</button>
            </div>
            <div style={{ display:'grid', gridTemplateColumns:'1fr 1fr', gap:'12px' }}>
              <button type="button" onClick={openExistingA4BarcodePrinter} style={{ padding:'22px 14px', border:'2px solid #cbd5e1', background:'#fff', borderRadius:'12px', cursor:'pointer', textAlign:'left' }}>
                <div style={{ fontSize:'27px' }}>📄</div><div style={{ fontSize:'16px', fontWeight:900, marginTop:'7px' }}>Existing A4 Label Print</div><div style={{ fontSize:'12px', color:'#64748b', marginTop:'5px' }}>Current A4 barcode label layout.</div>
              </button>
              <button type="button" onClick={openCustomBarcodePrinter} style={{ padding:'22px 14px', border:'2px solid #2563eb', background:'#eff6ff', borderRadius:'12px', cursor:'pointer', textAlign:'left' }}>
                <div style={{ fontSize:'27px' }}>🏷️</div><div style={{ fontSize:'16px', fontWeight:900, marginTop:'7px', color:'#1d4ed8' }}>Label Printer / Custom</div><div style={{ fontSize:'12px', color:'#475569', marginTop:'5px' }}>Use the selected sticker size as the actual printer page.</div>
              </button>
            </div>
          </div>
        </div>
      )}

      {showBarcodeCustomModal && barcodePrintProduct && (
        <div className="no-print" style={{ position:'fixed', inset:0, zIndex:100003, background:'rgba(15,23,42,.72)', display:'flex', alignItems:'center', justifyContent:'center', padding:'20px', overflowY:'auto' }}>
          <div style={{ width:'min(760px,100%)', background:'#fff', borderRadius:'16px', padding:'24px', boxShadow:'0 25px 80px rgba(0,0,0,.35)' }}>
            <h2 style={{ margin:'0 0 5px' }}>🏷️ Label Printer Settings</h2>
            <div style={{ color:'#64748b', fontSize:'13px' }}>{barcodePrintProduct.name} — {barcodePrintProduct.barcode}</div>
            <div style={{ marginTop:'14px', fontSize:'12px', color:'#475569', background:'#f8fafc', border:'1px solid #e2e8f0', borderRadius:'10px', padding:'11px 13px' }}>Selected sticker size will be used as the real print-page size. This is for roll / die-cut label printers, not A4.</div>
            <label style={{ display:'block', fontWeight:900, marginTop:'18px', marginBottom:'8px' }}>Choose Sticker Size</label>
            <div style={{ display:'grid', gridTemplateColumns:'repeat(3,minmax(0,1fr))', gap:'9px' }}>
              {barcodeLabelPresets.map(preset => (
                <button key={preset.key} type="button" onClick={() => applyBarcodeCustomPreset(preset.key)} style={{ padding:'11px 8px', borderRadius:'9px', border:barcodeCustomPreset===preset.key ? '2px solid #2563eb' : '1px solid #cbd5e1', background:barcodeCustomPreset===preset.key ? '#eff6ff' : '#fff', cursor:'pointer', textAlign:'center' }}>
                  <div style={{ fontWeight:900, fontSize:'13px' }}>{preset.label}</div><div style={{ fontSize:'10px', color:'#64748b', marginTop:'3px' }}>{preset.note}</div>
                </button>
              ))}
            </div>
            <div style={{ display:'grid', gridTemplateColumns:'repeat(3,1fr)', gap:'10px', marginTop:'15px' }}>
              <div><label style={{ display:'block', fontSize:'12px', fontWeight:800, marginBottom:'5px' }}>Width (mm)</label><input type="number" min="20" max="100" step="0.1" value={barcodeCustomWidth} onChange={e => { setBarcodeCustomWidth(e.target.value); setBarcodeCustomPreset('custom'); }} style={{ width:'100%', boxSizing:'border-box', padding:'10px', border:'1px solid #cbd5e1', borderRadius:'7px' }} /></div>
              <div><label style={{ display:'block', fontSize:'12px', fontWeight:800, marginBottom:'5px' }}>Height (mm)</label><input type="number" min="15" max="100" step="0.1" value={barcodeCustomHeight} onChange={e => { setBarcodeCustomHeight(e.target.value); setBarcodeCustomPreset('custom'); }} style={{ width:'100%', boxSizing:'border-box', padding:'10px', border:'1px solid #cbd5e1', borderRadius:'7px' }} /></div>
              <div><label style={{ display:'block', fontSize:'12px', fontWeight:800, marginBottom:'5px' }}>Gap (mm)</label><input type="number" min="0" max="10" step="0.1" value={barcodeCustomGap} onChange={e => setBarcodeCustomGap(e.target.value)} style={{ width:'100%', boxSizing:'border-box', padding:'10px', border:'1px solid #cbd5e1', borderRadius:'7px' }} /></div>
            </div>
            <div style={{ marginTop:'13px' }}><label style={{ display:'block', fontSize:'12px', fontWeight:800, marginBottom:'5px' }}>Number of Labels</label><input type="number" min="1" max="1000" step="1" value={barcodeCustomQuantity} onChange={e => setBarcodeCustomQuantity(e.target.value)} style={{ width:'220px', maxWidth:'100%', boxSizing:'border-box', padding:'10px', border:'2px solid #2563eb', borderRadius:'7px', fontSize:'16px' }} /></div>
            <div style={{ display:'flex', justifyContent:'space-between', gap:'10px', alignItems:'center', marginTop:'20px', flexWrap:'wrap' }}>
              <div style={{ fontSize:'11px', color:'#64748b' }}>Print size: <b>{barcodeCustomWidth || 0} × {barcodeCustomHeight || 0} mm</b> • gap: <b>{barcodeCustomGap || 0} mm</b></div>
              <div style={{ display:'flex', gap:'8px' }}><button type="button" onClick={() => { setShowBarcodeCustomModal(false); setShowBarcodePrintChoiceModal(true); }} style={{ padding:'10px 15px', border:'1px solid #cbd5e1', background:'#fff', borderRadius:'7px', cursor:'pointer' }}>← Back</button><button type="button" onClick={prepareCustomBarcodePreview} style={{ padding:'10px 18px', border:0, background:'#2563eb', color:'#fff', borderRadius:'7px', cursor:'pointer', fontWeight:900 }}>👁️ Preview Labels</button></div>
            </div>
          </div>
        </div>
      )}

      {showBarcodeCustomPreview && barcodePrintItems.length > 0 && (
        <div className="no-print" style={{ position:'fixed', inset:0, zIndex:100004, background:'rgba(15,23,42,.82)', display:'flex', flexDirection:'column', alignItems:'center', overflowY:'auto', padding:'24px' }}>
          <div style={{ width:'min(1050px,100%)', background:'#fff', borderRadius:'14px', padding:'18px', boxShadow:'0 20px 70px rgba(0,0,0,.35)' }}>
            <h2 style={{ margin:0 }}>👁️ Label Printer Preview</h2>
            <div style={{ color:'#64748b', fontSize:'12px', margin:'4px 0 14px' }}>{barcodePrintItems.length} label(s) • {barcodeCustomWidth} × {barcodeCustomHeight} mm • gap {barcodeCustomGap} mm • 1 label per print page</div>
            <div style={{ background:'#f1f5f9', border:'1px solid #e2e8f0', borderRadius:'10px', padding:'18px', maxHeight:'65vh', overflowY:'auto' }}>
              <div className="barcode-custom-preview-grid">
                {barcodePrintItems.slice(0,60).map((item,index) => (
                  <div className="barcode-custom-preview-card" key={`custom-preview-${item.barcode}-${index}`}>
                    <div className="barcode-label barcode-custom-label" style={{ width:`${barcodeCustomWidth}mm`, height:`${barcodeCustomHeight}mm` }}>
                      <div className="barcode-shop" title={receiptSettings.shopName}>{receiptSettings.shopName || 'POSHITHA POS SYSTEM'}</div>
                      <div className="barcode-name" title={item.name}>{item.name}</div>
                      <div className="barcode-meta"><span className="barcode-price">Rs. {Number(item.price || 0).toFixed(2)}</span><span className="barcode-unit">{item.unit || 'Pcs'}</span></div>
                      <div className="barcode-svg" dangerouslySetInnerHTML={{ __html: barcodeSvgMarkup(item.barcode, 300, 70) }} />
                      <div className="barcode-number" title={String(item.barcode || '')}>{item.barcode}</div>
                    </div>
                  </div>
                ))}
              </div>
              {barcodePrintItems.length > 60 && <div style={{ textAlign:'center', marginTop:'12px', fontSize:'12px', color:'#64748b' }}>Preview shows the first 60 labels. All {barcodePrintItems.length} labels will print.</div>}
            </div>
            <div style={{ display:'flex', justifyContent:'flex-end', gap:'8px', marginTop:'14px' }}>
              <button type="button" onClick={() => { setShowBarcodeCustomPreview(false); setShowBarcodeCustomModal(true); setShowBarcodePrintChoiceModal(false); }} style={{ padding:'10px 16px', border:'1px solid #cbd5e1', background:'#fff', borderRadius:'7px', cursor:'pointer' }}>← Back</button>
              <button type="button" onClick={handlePrintCustomBarcodeLabels} style={{ padding:'10px 20px', border:'none', background:'#16a34a', color:'#fff', borderRadius:'7px', cursor:'pointer', fontWeight:'900' }}>🖨️ Print on Label Printer</button>
            </div>
          </div>
        </div>
      )}

      {showBarcodePrintModal && barcodePrintProduct && (
        <div className="no-print" style={{ position: 'fixed', inset: 0, zIndex: 100000, background: 'rgba(0,0,0,.6)', display: 'flex', alignItems: 'center', justifyContent: 'center', padding: '20px' }}>
          <div style={{ width: '420px', maxWidth: '100%', background: '#fff', borderRadius: '12px', padding: '22px', boxShadow: '0 20px 60px rgba(0,0,0,.3)' }}>
            <h2 style={{ margin: '0 0 6px' }}>🖨️ Print Barcode</h2>
            <div style={{ color: '#6b7280', marginBottom: '18px' }}>{barcodePrintProduct.name} — {barcodePrintProduct.barcode}</div>
            <label style={{ display: 'block', fontWeight: '700', marginBottom: '7px' }}>Barcode Size</label>
            <div style={{ display: 'grid', gridTemplateColumns: '1fr 1fr', gap: '10px', marginBottom: '18px' }}>
              <button type="button" onClick={() => setBarcodePrintSize('large')} style={{ padding: '13px 8px', borderRadius: '8px', border: barcodePrintSize === 'large' ? '2px solid #2563eb' : '1px solid #d1d5db', background: barcodePrintSize === 'large' ? '#eff6ff' : '#fff', cursor: 'pointer', fontWeight: '700' }}>Large<br/><span style={{ fontSize: '12px', fontWeight: '500' }}>3.8cm × 2.5cm</span></button>
              <button type="button" onClick={() => setBarcodePrintSize('small')} style={{ padding: '13px 8px', borderRadius: '8px', border: barcodePrintSize === 'small' ? '2px solid #2563eb' : '1px solid #d1d5db', background: barcodePrintSize === 'small' ? '#eff6ff' : '#fff', cursor: 'pointer', fontWeight: '700' }}>Small<br/><span style={{ fontSize: '12px', fontWeight: '500' }}>3cm × 2cm</span></button>
            </div>
            <label style={{ display: 'block', fontWeight: '700', marginBottom: '7px' }}>Barcode Quantity</label>
            <input type="number" min="1" max="1000" step="1" value={barcodePrintQuantity} onChange={e => setBarcodePrintQuantity(e.target.value)} style={{ width: '100%', boxSizing: 'border-box', padding: '11px', border: '2px solid #3b82f6', borderRadius: '7px', fontSize: '16px', marginBottom: '18px' }} />
            <div style={{ background: '#f8fafc', border: '1px solid #e5e7eb', borderRadius: '8px', padding: '10px 12px', fontSize: '12px', color: '#475569', marginBottom: '18px' }}>
              A4 paper එකේ labels grid එකක් ලෙස print වෙනවා. හැම label එකක් වටේම cutting line එකක් තියෙනවා.
            </div>
            <div style={{ display: 'flex', justifyContent: 'flex-end', gap: '8px' }}>
              <button type="button" onClick={() => setShowBarcodePrintModal(false)} style={{ padding: '10px 16px', border: '1px solid #d1d5db', background: '#fff', borderRadius: '7px', cursor: 'pointer' }}>Cancel</button>
              <button type="button" onClick={prepareBarcodePreview} style={{ padding: '10px 18px', border: 'none', background: '#2563eb', color: '#fff', borderRadius: '7px', cursor: 'pointer', fontWeight: '800' }}>👁️ Preview Labels</button>
            </div>
          </div>
        </div>
      )}

      {showBarcodePreview && barcodePrintItems.length > 0 && (
        <div className="no-print" style={{ position: 'fixed', inset: 0, zIndex: 100001, background: 'rgba(15,23,42,.78)', display: 'flex', flexDirection: 'column', alignItems: 'center', overflowY: 'auto', padding: '24px' }}>
          <div style={{ width: 'min(1000px, 100%)', background: '#fff', borderRadius: '14px', padding: '18px', boxShadow: '0 20px 70px rgba(0,0,0,.35)' }}>
            <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', gap: '12px', marginBottom: '14px' }}>
              <div>
                <h2 style={{ margin: 0 }}>👁️ Barcode Label Preview</h2>
                <div style={{ color: '#64748b', fontSize: '12px', marginTop: '4px' }}>{barcodePrintItems.length} label(s) • {barcodePrintSize === 'large' ? '3.8cm × 2.5cm' : '3cm × 2cm'} • A4 cutting-line layout</div>
              </div>
              <button type="button" onClick={() => setShowBarcodePreview(false)} style={{ padding: '8px 12px', border: '1px solid #cbd5e1', background: '#fff', borderRadius: '7px', cursor: 'pointer' }}>✕ Close</button>
            </div>
            <div style={{ background: '#f1f5f9', border: '1px solid #e2e8f0', borderRadius: '10px', padding: '18px', maxHeight: '65vh', overflowY: 'auto' }}>
              <div className="barcode-preview-grid" data-size={barcodePrintSize}>
                {barcodePrintItems.slice(0, 60).map((item, index) => (
                  <div className="barcode-label preview-label" key={`preview-${item.barcode}-${index}`}>
                    <div className="barcode-shop" title={receiptSettings.shopName}>{receiptSettings.shopName || 'POSHITHA POS SYSTEM'}</div>
                    <div className="barcode-name" title={item.name}>{item.name}</div>
                    <div className="barcode-meta"><span className="barcode-price">Rs. {Number(item.price || 0).toFixed(2)}</span><span className="barcode-unit">{item.unit || 'Pcs'}</span></div>
                    <div className="barcode-svg" dangerouslySetInnerHTML={{ __html: barcodeSvgMarkup(item.barcode, barcodePrintSize === 'large' ? 250 : 205, barcodePrintSize === 'large' ? 58 : 43) }} />
                    <div className="barcode-number" title={String(item.barcode || '')}>{item.barcode}</div>
                  </div>
                ))}
              </div>
              {barcodePrintItems.length > 60 && <div style={{ textAlign: 'center', marginTop: '12px', fontSize: '12px', color: '#64748b' }}>Preview shows the first 60 labels. All {barcodePrintItems.length} labels will print.</div>}
            </div>
            <div style={{ display: 'flex', justifyContent: 'flex-end', gap: '8px', marginTop: '14px' }}>
              <button type="button" onClick={() => { setShowBarcodePreview(false); setShowBarcodePrintModal(true); }} style={{ padding: '10px 16px', border: '1px solid #cbd5e1', background: '#fff', borderRadius: '7px', cursor: 'pointer' }}>← Back</button>
              <button type="button" onClick={handlePrintProductBarcodes} style={{ padding: '10px 20px', border: 'none', background: '#2563eb', color: '#fff', borderRadius: '7px', cursor: 'pointer', fontWeight: '800' }}>🖨️ Print Labels</button>
            </div>
          </div>
        </div>
      )}

      {dashboardStyle === 'classic' && (
        <>
          {/* SIDEBAR NAVIGATION — collapsible; existing page options/actions are unchanged. */}
      <div className="no-print" style={{ width: sidebarCollapsed ? '76px' : '240px', minWidth: sidebarCollapsed ? '76px' : '240px', height: '100vh', boxSizing: 'border-box', flexShrink: 0, order: -1, backgroundColor: '#1f2937', color: '#fff', padding: sidebarCollapsed ? '12px 8px' : '20px', display: 'flex', flexDirection: 'column', gap: '10px', overflowY: 'auto' }}>
        <div style={{ display: 'flex', alignItems: 'center', gap: '10px', marginBottom: '10px' }}>
          {receiptSettings.logoUrl ? (
            <img src={receiptSettings.logoUrl} alt="Logo" style={{ width: '35px', height: '35px', objectFit: 'cover', borderRadius: '50%' }} />
          ) : (
            <div style={{ width: '35px', height: '35px', backgroundColor: '#3b82f6', borderRadius: '50%', display: 'flex', alignItems: 'center', justifyContent: 'center', fontWeight: 'bold' }}>POS</div>
          )}
          {!sidebarCollapsed && <div>
            <h2 style={{ fontSize: '15px', fontWeight: 'bold', color: '#3b82f6', margin: 0, overflow: 'hidden', textOverflow: 'ellipsis', whiteSpace: 'nowrap' }}>{receiptSettings.shopName}</h2>
            <span style={{ fontSize: '11px', color: '#9ca3af', textTransform: 'uppercase' }}>Role: {currentUserRole}</span>
          </div>}
        </div>

        <button onClick={() => setSidebarCollapsed(prev => !prev)} title={sidebarCollapsed ? 'Expand Dashboard' : 'Collapse Dashboard'} aria-label={sidebarCollapsed ? 'Expand Dashboard' : 'Collapse Dashboard'} style={{ width: '100%', padding: '8px', background: '#374151', color: '#fff', border: '1px solid #4b5563', borderRadius: '7px', cursor: 'pointer', fontSize: '16px', fontWeight: '900', lineHeight: 1 }}>
          {sidebarCollapsed ? '☰' : '◀'}
        </button>

        {!sidebarCollapsed && <div style={{ display: 'flex', gap: '5px', backgroundColor: '#374151', padding: '4px', borderRadius: '6px', marginBottom: '10px' }}>
          <button onClick={() => setLang('si')} style={{ flex: 1, padding: '4px', fontSize: '11px', backgroundColor: lang === 'si' ? '#2563eb' : 'transparent', color: '#fff', border: 'none', borderRadius: '4px', cursor: 'pointer' }}>සිංහල</button>
          <button onClick={() => setLang('en')} style={{ flex: 1, padding: '4px', fontSize: '11px', backgroundColor: lang === 'en' ? '#2563eb' : 'transparent', color: '#fff', border: 'none', borderRadius: '4px', cursor: 'pointer' }}>English</button>
          <button onClick={() => setLang('ta')} style={{ flex: 1, padding: '4px', fontSize: '11px', backgroundColor: lang === 'ta' ? '#2563eb' : 'transparent', color: '#fff', border: 'none', borderRadius: '4px', cursor: 'pointer' }}>தமிழ்</button>
        </div>}

        <button title={t.billing} onClick={() => setActiveTab('pos')} style={{ padding: '8px', textAlign: sidebarCollapsed ? 'center' : 'left', display: 'flex', alignItems: 'center', justifyContent: sidebarCollapsed ? 'center' : 'flex-start', backgroundColor: activeTab === 'pos' ? '#374151' : 'transparent', color: '#fff', border: 'none', borderRadius: '6px', cursor: 'pointer' }}> <span style={{ display: 'inline-flex', width: '32px', height: '32px', alignItems: 'center', justifyContent: 'center', fontSize: '24px', borderRadius: '8px', backgroundColor: '#2563eb', marginRight: sidebarCollapsed ? 0 : '8px' }}>🛒</span>{!sidebarCollapsed && t.billing}</button>
        <button onClick={() => setActiveTab('inventory')} style={{ padding: '8px', textAlign: sidebarCollapsed ? 'center' : 'left', display: 'flex', alignItems: 'center', justifyContent: sidebarCollapsed ? 'center' : 'flex-start', backgroundColor: activeTab === 'inventory' ? '#374151' : 'transparent', color: '#fff', border: 'none', borderRadius: '6px', cursor: 'pointer' }}> <span style={{ display: 'inline-flex', width: '32px', height: '32px', alignItems: 'center', justifyContent: 'center', fontSize: '24px', borderRadius: '8px', backgroundColor: '#059669', marginRight: sidebarCollapsed ? 0 : '8px' }}>📦</span>{!sidebarCollapsed && t.inventory}</button>
        <button onClick={() => setActiveTab('returns')} style={{ padding: '8px', textAlign: sidebarCollapsed ? 'center' : 'left', display: 'flex', alignItems: 'center', justifyContent: sidebarCollapsed ? 'center' : 'flex-start', backgroundColor: activeTab === 'returns' ? '#374151' : 'transparent', color: '#fff', border: 'none', borderRadius: '6px', cursor: 'pointer' }}> <span style={{ display: 'inline-flex', width: '32px', height: '32px', alignItems: 'center', justifyContent: 'center', fontSize: '24px', borderRadius: '8px', backgroundColor: '#f59e0b', marginRight: sidebarCollapsed ? 0 : '8px' }}>↩️</span>{!sidebarCollapsed && ui.returns}</button>
        
        {currentUserRole === 'admin' && (
          <>
            <button onClick={() => setActiveTab('suppliers')} style={{ padding: '8px', textAlign: sidebarCollapsed ? 'center' : 'left', display: 'flex', alignItems: 'center', justifyContent: sidebarCollapsed ? 'center' : 'flex-start', backgroundColor: activeTab === 'suppliers' ? '#374151' : 'transparent', color: '#fff', border: 'none', borderRadius: '6px', cursor: 'pointer' }}> <span style={{ display: 'inline-flex', width: '32px', height: '32px', alignItems: 'center', justifyContent: 'center', fontSize: '24px', borderRadius: '8px', backgroundColor: '#8b5cf6', marginRight: sidebarCollapsed ? 0 : '8px' }}>🚚</span>{!sidebarCollapsed && t.suppliers}</button>
            <button onClick={() => setActiveTab('shopOrderHistory')} style={{ padding: '8px', textAlign: sidebarCollapsed ? 'center' : 'left', display: 'flex', alignItems: 'center', justifyContent: sidebarCollapsed ? 'center' : 'flex-start', backgroundColor: activeTab === 'shopOrderHistory' ? '#374151' : 'transparent', color: '#fff', border: 'none', borderRadius: '6px', cursor: 'pointer' }}> <span style={{ display: 'inline-flex', width: '32px', height: '32px', alignItems: 'center', justifyContent: 'center', fontSize: '24px', borderRadius: '8px', backgroundColor: '#14b8a6', marginRight: sidebarCollapsed ? 0 : '8px' }}>🧾</span>{!sidebarCollapsed && ui.shopOrderHistory}</button>
            <button onClick={() => setActiveTab('customers')} style={{ padding: '8px', textAlign: sidebarCollapsed ? 'center' : 'left', display: 'flex', alignItems: 'center', justifyContent: sidebarCollapsed ? 'center' : 'flex-start', backgroundColor: activeTab === 'customers' ? '#374151' : 'transparent', color: '#fff', border: 'none', borderRadius: '6px', cursor: 'pointer' }}> <span style={{ display: 'inline-flex', width: '32px', height: '32px', alignItems: 'center', justifyContent: 'center', fontSize: '24px', borderRadius: '8px', backgroundColor: '#ec4899', marginRight: sidebarCollapsed ? 0 : '8px' }}>👥</span>{!sidebarCollapsed && t.customers}</button>
            <button onClick={() => setActiveTab('creditBook')} style={{ padding: '8px', textAlign: sidebarCollapsed ? 'center' : 'left', display: 'flex', alignItems: 'center', justifyContent: sidebarCollapsed ? 'center' : 'flex-start', backgroundColor: activeTab === 'creditBook' ? '#374151' : 'transparent', color: '#fff', border: 'none', borderRadius: '6px', cursor: 'pointer' }}> <span style={{ display: 'inline-flex', width: '32px', height: '32px', alignItems: 'center', justifyContent: 'center', fontSize: '24px', borderRadius: '8px', backgroundColor: '#f97316', marginRight: sidebarCollapsed ? 0 : '8px' }}>📘</span>{!sidebarCollapsed && t.creditBook}</button>
            <button onClick={() => setActiveTab('reports')} style={{ padding: '8px', textAlign: sidebarCollapsed ? 'center' : 'left', display: 'flex', alignItems: 'center', justifyContent: sidebarCollapsed ? 'center' : 'flex-start', backgroundColor: activeTab === 'reports' ? '#374151' : 'transparent', color: '#fff', border: 'none', borderRadius: '6px', cursor: 'pointer' }}> <span style={{ display: 'inline-flex', width: '32px', height: '32px', alignItems: 'center', justifyContent: 'center', fontSize: '24px', borderRadius: '8px', backgroundColor: '#06b6d4', marginRight: sidebarCollapsed ? 0 : '8px' }}>📊</span>{!sidebarCollapsed && t.reports}</button>
            <button onClick={() => { setActiveTab('reports'); setReportView('monthly'); }} style={{ padding: '8px', textAlign: sidebarCollapsed ? 'center' : 'left', display: 'flex', alignItems: 'center', justifyContent: sidebarCollapsed ? 'center' : 'flex-start', backgroundColor: activeTab === 'reports' && reportView === 'monthly' ? '#374151' : 'transparent', color: '#fff', border: 'none', borderRadius: '6px', cursor: 'pointer' }}> <span style={{ display: 'inline-flex', width: '32px', height: '32px', alignItems: 'center', justifyContent: 'center', fontSize: '24px', borderRadius: '8px', backgroundColor: '#7c3aed', marginRight: sidebarCollapsed ? 0 : '8px' }}>📈</span>{!sidebarCollapsed && ui.monthlySalesReport}</button>
            <button onClick={() => setActiveTab('settings')} style={{ padding: '8px', textAlign: sidebarCollapsed ? 'center' : 'left', display: 'flex', alignItems: 'center', justifyContent: sidebarCollapsed ? 'center' : 'flex-start', backgroundColor: activeTab === 'settings' ? '#374151' : 'transparent', color: '#fff', border: 'none', borderRadius: '6px', cursor: 'pointer' }}> <span style={{ display: 'inline-flex', width: '32px', height: '32px', alignItems: 'center', justifyContent: 'center', fontSize: '24px', borderRadius: '8px', backgroundColor: '#64748b', marginRight: sidebarCollapsed ? 0 : '8px' }}>⚙️</span>{!sidebarCollapsed && t.settings}</button>
          </>
        )}
        
        <div style={{ marginTop: 'auto', borderTop: '1px solid #374151', paddingTop: '10px' }}>
          {!sidebarCollapsed && <div style={{ fontSize: '12px', color: '#cbd5e1', fontWeight: 'bold', marginBottom: '8px', textAlign: 'center' }}>Poshitha Pos System 01</div>}
          <button onClick={() => setIsLoggedIn(false)} style={{ width: '100%', padding: '10px', textAlign: 'left', backgroundColor: '#dc2626', color: '#fff', border: 'none', borderRadius: '6px', cursor: 'pointer', fontWeight: 'bold' }}> <span style={{ display: 'inline-flex', width: '30px', height: '30px', alignItems: 'center', justifyContent: 'center', fontSize: '22px', marginRight: sidebarCollapsed ? 0 : '8px' }}>🚪</span>{!sidebarCollapsed && t.logout}</button>
        </div>
      </div>

        </>
      )}

      <div className="no-print" style={{ flex: 1, order: 1, padding: '20px', overflowY: 'auto' }}>

        {(dashboardStyle === 'sidebar' || dashboardStyle === 'both') && (
          <div style={{
            display: 'flex', alignItems: 'center', gap: '8px', flexWrap: 'wrap',
            background: '#0f172a', color: '#fff', padding: '9px 10px', borderRadius: '10px', marginBottom: '12px',
            boxShadow: '0 4px 12px rgba(15,23,42,.12)'
          }}>
            {[
              { key: 'home', icon: '🏠', label: lang === 'si' ? 'මුල් පිටුව' : lang === 'ta' ? 'முகப்பு' : 'Home' },
              { key: 'pos', icon: '🧾', label: t.billing },
              { key: 'inventory', icon: '📦', label: t.inventory },
              { key: 'returns', icon: '↩️', label: ui.returns },
              ...(currentUserRole === 'admin' ? [
                { key: 'suppliers', icon: '🚚', label: t.suppliers },
                { key: 'shopOrderHistory', icon: '🧾', label: ui.shopOrderHistory },
                { key: 'customers', icon: '👥', label: t.customers },
                { key: 'creditBook', icon: '📘', label: t.creditBook },
                { key: 'reports', icon: '📊', label: t.reports },
                { key: 'monthlySalesReport', icon: '📈', label: ui.monthlySalesReport },
                { key: 'settings', icon: '⚙️', label: t.settings },
                { key: 'offers', icon: '🏷️', label: ui.offers },
                { key: 'barcode', icon: '▥', label: ui.barcode }
              ] : [])
            ].map(nav => (
              <button key={nav.key} type="button" onClick={() => {
                if (nav.key === 'home') { setInventoryShortcutMode(null); setActiveTab('home'); }
                else if (nav.key === 'monthlySalesReport') { setReportView('monthly'); setActiveTab('reports'); }
                else if (nav.key === 'offers' || nav.key === 'barcode') goToInventorySection(nav.key);
                else { setInventoryShortcutMode(null); setActiveTab(nav.key); }
              }} style={{
                padding: '8px 11px', border: '1px solid rgba(255,255,255,.15)', borderRadius: '7px',
                background: activeTab === nav.key ? '#2563eb' : 'rgba(255,255,255,.07)', color: '#fff',
                cursor: 'pointer', fontWeight: '800', fontSize: '12px'
              }}>
                {nav.icon} {nav.label}
              </button>
            ))}
          </div>
        )}

        {dashboardStyle !== 'classic' && activeTab !== 'pos' && (
          <>
        <div style={{
          display: 'flex',
          alignItems: 'center',
          justifyContent: 'space-between',
          gap: '12px',
          flexWrap: 'wrap',
          background: '#ffffff',
          border: '1px solid #e5e7eb',
          borderRadius: '12px',
          padding: '10px 14px',
          marginBottom: '16px',
          boxShadow: '0 2px 8px rgba(0,0,0,0.05)'
        }}>
          <div style={{ display: 'flex', alignItems: 'center', gap: '10px', minWidth: 0 }}>
            {receiptSettings.logoUrl ? (
              <img src={receiptSettings.logoUrl} alt="Logo" style={{ width: '38px', height: '38px', objectFit: 'cover', borderRadius: '50%' }} />
            ) : (
              <div style={{ width: '38px', height: '38px', background: '#2563eb', color: '#fff', borderRadius: '50%', display: 'flex', alignItems: 'center', justifyContent: 'center', fontWeight: '900' }}>POS</div>
            )}
            <div style={{ minWidth: 0 }}>
              <div style={{ fontWeight: '900', color: '#111827', overflow: 'hidden', textOverflow: 'ellipsis', whiteSpace: 'nowrap' }}>{receiptSettings.shopName || 'POSHITHA POS SYSTEM'}</div>
              <div style={{ fontSize: '11px', color: '#6b7280' }}>Role: {currentUserRole}</div>
            </div>
          </div>

          <div style={{ display: 'flex', alignItems: 'center', gap: '7px', flexWrap: 'wrap' }}>
            {activeTab !== 'home' && (
              <button
                type="button"
                onClick={() => { setInventoryShortcutMode(null); setActiveTab('home'); }}
                style={{ padding: '9px 14px', background: '#2563eb', color: '#fff', border: 'none', borderRadius: '7px', cursor: 'pointer', fontWeight: '800' }}
              >🏠 Home</button>
            )}
            <div style={{ display: 'flex', gap: '3px', background: '#f1f5f9', padding: '3px', borderRadius: '7px' }}>
              <button type="button" onClick={() => setLang('si')} style={{ padding: '6px 8px', fontSize: '11px', background: lang === 'si' ? '#2563eb' : 'transparent', color: lang === 'si' ? '#fff' : '#334155', border: 'none', borderRadius: '5px', cursor: 'pointer' }}>සිංහල</button>
              <button type="button" onClick={() => setLang('en')} style={{ padding: '6px 8px', fontSize: '11px', background: lang === 'en' ? '#2563eb' : 'transparent', color: lang === 'en' ? '#fff' : '#334155', border: 'none', borderRadius: '5px', cursor: 'pointer' }}>English</button>
              <button type="button" onClick={() => setLang('ta')} style={{ padding: '6px 8px', fontSize: '11px', background: lang === 'ta' ? '#2563eb' : 'transparent', color: lang === 'ta' ? '#fff' : '#334155', border: 'none', borderRadius: '5px', cursor: 'pointer' }}>தமிழ்</button>
            </div>
            <button
              type="button"
              onClick={() => setIsLoggedIn(false)}
              style={{ padding: '9px 14px', background: '#dc2626', color: '#fff', border: 'none', borderRadius: '7px', cursor: 'pointer', fontWeight: '800' }}
            >🚪 {t.logout}</button>
          </div>
        </div>
          </>
        )}


        {inventoryShortcutMode && activeTab === 'inventory' && dashboardStyle !== 'classic' && (
          <button
            type="button"
            onClick={() => {
              setInventoryShortcutMode(null);
              setActiveTab((dashboardStyle === 'home' || dashboardStyle === 'both') ? 'home' : 'pos');
            }}
            title="Back to Home"
            style={{
              position: 'fixed', top: '18px', right: '22px', zIndex: 9999,
              padding: '10px 16px', border: 'none', borderRadius: '999px',
              background: '#2563eb', color: '#fff', cursor: 'pointer', fontWeight: '900',
              boxShadow: '0 8px 24px rgba(37,99,235,.30)'
            }}
          >🏠 Home</button>
        )}

        {activeTab === 'home' && (dashboardStyle === 'home' || dashboardStyle === 'both') && (
          <div style={{ maxWidth: '1250px', margin: '0 auto', padding: '8px 4px 30px' }}>
            <div style={{ textAlign: 'center', marginBottom: '24px' }}>
              {receiptSettings.logoUrl ? (
                <img src={receiptSettings.logoUrl} alt="Shop Logo" style={{ width: '78px', height: '78px', objectFit: 'contain', borderRadius: '18px', border: '1px solid #e5e7eb', background: '#fff', padding: '5px' }} />
              ) : (
                <div style={{ width: '78px', height: '78px', margin: '0 auto', background: '#2563eb', color: '#fff', borderRadius: '18px', display: 'flex', alignItems: 'center', justifyContent: 'center', fontSize: '22px', fontWeight: '900', boxShadow: '0 10px 25px rgba(37,99,235,.25)' }}>POS</div>
              )}
              <h1 style={{ margin: '12px 0 4px', color: '#111827', fontSize: '28px', fontWeight: '900' }}>{receiptSettings.shopName || 'POSHITHA POS SYSTEM'}</h1>
              <div style={{ color: '#334155', fontSize: '13px', fontWeight: '900', letterSpacing: '1.4px', marginTop: '2px' }}>
                POSHITHA POS SYSTEM
              </div>
              <div style={{ color: '#64748b', fontSize: '14px', marginTop: '3px' }}>
                {lang === 'si' ? 'ප්‍රධාන Home Dashboard' : lang === 'ta' ? 'முதன்மை Home Dashboard' : 'Main Home Dashboard'}
              </div>
            </div>

            <div style={{
              display: 'grid',
              gridTemplateColumns: 'repeat(auto-fit, minmax(180px, 1fr))',
              gap: '16px'
            }}>
              {[
                { key: 'pos', icon: '🧾', label: t.billing, bg: '#2563eb' },
                { key: 'inventory', icon: '📦', label: t.inventory, bg: '#059669' },
                { key: 'returns', icon: '↩️', label: ui.returns, bg: '#f59e0b' },
                { key: 'suppliers', icon: '🚚', label: t.suppliers, bg: '#8b5cf6', admin: true },
                { key: 'shopOrderHistory', icon: '🧾', label: ui.shopOrderHistory, bg: '#14b8a6', admin: true },
                { key: 'customers', icon: '👥', label: t.customers, bg: '#ec4899', admin: true },
                { key: 'creditBook', icon: '📘', label: t.creditBook, bg: '#f97316', admin: true },
                { key: 'reports', icon: '📊', label: t.reports, bg: '#06b6d4', admin: true },
                { key: 'monthlySalesReport', icon: '📈', label: ui.monthlySalesReport, bg: '#7c3aed', admin: true },
                { key: 'settings', icon: '⚙️', label: t.settings, bg: '#64748b', admin: true },
                { key: 'offers', icon: '🏷️', label: ui.offers, bg: '#db2777' },
                { key: 'barcode', icon: '▥', label: ui.barcode, bg: '#0f766e' }
              ].filter(item => !item.admin || currentUserRole === 'admin').map(item => (
                <button
                  key={item.key}
                  type="button"
                  onClick={() => {
                    if (item.key === 'monthlySalesReport') {
                      setReportView('monthly');
                      setActiveTab('reports');
                    } else if (item.key === 'offers' || item.key === 'barcode') {
                      goToInventorySection(item.key);
                    } else {
                      setInventoryShortcutMode(null);
                      setActiveTab(item.key);
                    }
                  }}
                  style={{
                    minHeight: '145px',
                    padding: '20px 14px',
                    border: 'none',
                    borderRadius: '16px',
                    background: '#fff',
                    boxShadow: '0 5px 18px rgba(15,23,42,.10)',
                    borderTop: `6px solid ${item.bg}`,
                    cursor: 'pointer',
                    display: 'flex',
                    flexDirection: 'column',
                    alignItems: 'center',
                    justifyContent: 'center',
                    gap: '10px',
                    transition: 'transform .15s ease, box-shadow .15s ease'
                  }}
                >
                  <span style={{
                    width: '68px',
                    height: '68px',
                    borderRadius: '18px',
                    background: item.bg,
                    color: '#fff',
                    display: 'flex',
                    alignItems: 'center',
                    justifyContent: 'center',
                    fontSize: item.key === 'barcode' ? '38px' : '34px',
                    fontWeight: '900',
                    boxShadow: '0 8px 18px rgba(15,23,42,.12)'
                  }}>{item.icon}</span>
                  <span style={{ color: '#111827', fontSize: '15px', fontWeight: '900', textAlign: 'center', lineHeight: 1.25 }}>{item.label}</span>
                </button>
              ))}
            </div>

            <div style={{ marginTop: '22px' }}>
              <div style={{ fontSize: '12px', fontWeight: '900', color: '#64748b', textTransform: 'uppercase', letterSpacing: '1px', marginBottom: '9px' }}>POS Quick Status</div>
              <div style={{ display: 'grid', gridTemplateColumns: 'repeat(4, minmax(0, 1fr))', gap: '10px' }}>
                {[
                  { icon: '📦', label: 'Products', value: products.length, bg: '#ecfdf5', fg: '#047857' },
                  { icon: '⏸️', label: 'Held Bills', value: heldBills.length, bg: '#fff7ed', fg: '#c2410c' },
                  { icon: '🧾', label: 'Sales Records', value: salesHistory.length, bg: '#eff6ff', fg: '#1d4ed8' },
                  { icon: '⚠️', label: 'Low Stock', value: (reportsData.lowStock || []).length, bg: '#fef2f2', fg: '#b91c1c' }
                ].map(stat => (
                  <div key={stat.label} style={{ background: stat.bg, border: `1px solid ${stat.fg}22`, borderRadius: '12px', padding: '12px 14px', display: 'flex', alignItems: 'center', gap: '10px' }}>
                    <span style={{ fontSize: '25px' }}>{stat.icon}</span>
                    <div><div style={{ fontSize: '11px', color: '#64748b', fontWeight: '800' }}>{stat.label}</div><div style={{ fontSize: '20px', color: stat.fg, fontWeight: '900' }}>{stat.value}</div></div>
                  </div>
                ))}
              </div>
            </div>
          </div>
        )}

        {activeTab === 'pos' && (
          completedSale ? (
            <div style={{ display: 'contents' }}>
            <div className="no-print" style={{
                display: 'flex',
                alignItems: 'center',
                justifyContent: 'space-between',
                gap: '14px',
                flexWrap: 'wrap',
                backgroundColor: '#e0f2fe',
                padding: '12px 16px',
                borderRadius: '10px',
                border: '1px solid #7dd3fc',
                boxShadow: '0 2px 8px rgba(2,132,199,0.12)',
                marginBottom: '15px'
              }}>
                <div style={{ display: 'flex', alignItems: 'center', gap: '12px', minWidth: 0, flex: '1 1 360px' }}>
                  {receiptSettings.logoUrl ? (
                    <img src={receiptSettings.logoUrl} alt="Shop Logo" style={{ width: '42px', height: '42px', objectFit: 'cover', borderRadius: '50%', border: '2px solid #3b82f6', flexShrink: 0 }} />
                  ) : (
                    <div style={{ width: '42px', height: '42px', backgroundColor: '#3b82f6', color: '#fff', borderRadius: '50%', display: 'flex', alignItems: 'center', justifyContent: 'center', fontWeight: 'bold', fontSize: '13px', flexShrink: 0 }}>POS</div>
                  )}
                  <div style={{ minWidth: 0 }}>
                    <div style={{ fontSize: '18px', fontWeight: '900', color: '#1f2937', overflow: 'hidden', textOverflow: 'ellipsis', whiteSpace: 'nowrap' }}>
                      {receiptSettings.shopName || 'POSHITHA POS SYSTEM'}
                    </div>
                    <div style={{ fontSize: '12px', color: '#475569', overflowWrap: 'anywhere', marginTop:'2px', whiteSpace:'pre-line' }}>
                      📍 {receiptSettings.address || '-'}
                    </div>
                    <div style={{ fontSize: '12px', color: '#475569', overflowWrap: 'anywhere' }}>
                      📞 {receiptSettings.phone || '-'} &nbsp;•&nbsp; ✉️ {receiptSettings.email || '-'}
                    </div>
                  </div>
                </div>
                <div style={{ display: 'flex', alignItems: 'center', gap: '9px', flexWrap: 'wrap', justifyContent: 'flex-end' }}>
                  <div style={{
                    fontSize: '12px',
                    backgroundColor: '#f8fafc',
                    color: '#334155',
                    padding: '7px 11px',
                    borderRadius: '7px',
                    fontWeight: '800',
                    border: '1px solid #e2e8f0',
                    textAlign: 'right',
                    lineHeight: 1.35
                  }}>
                    <div>📅 {currentDateTime.toLocaleDateString()}</div>
                    <div style={{ color: '#2563eb' }}>🕒 {currentDateTime.toLocaleTimeString()}</div>
                  </div>
                  <button
                    type="button"
                    onClick={() => { setInventoryShortcutMode(null); setActiveTab('home'); }}
                    style={{ padding: '9px 14px', background: '#2563eb', color: '#fff', border: 'none', borderRadius: '7px', cursor: 'pointer', fontWeight: '800' }}
                  >🏠 Home</button>
                  <button
                    type="button"
                    onClick={() => setIsLoggedIn(false)}
                    style={{ padding: '9px 14px', background: '#dc2626', color: '#fff', border: 'none', borderRadius: '7px', cursor: 'pointer', fontWeight: '800' }}
                  >🚪 {t.logout}</button>
                </div>
              </div>
            <div className="billing-workspace" style={{ display: 'flex', flexDirection: 'column', alignItems: 'center', justifyContent: 'flex-start', minHeight: '100%', height: 'auto', overflow: 'visible' }}>
              <div style={{ backgroundColor: '#fff', padding: '20px', borderRadius: '10px', boxShadow: '0 4px 15px rgba(0,0,0,0.1)', border: '1px solid #e5e7eb', width: '340px' }}>
                <div style={{ position: 'sticky', top: 0, zIndex: 30, display: 'flex', gap: '10px', marginBottom: '14px', padding: '8px 0', background: '#fff' }}>
                  <button
                    ref={completedPrintRef}
                    type="button"
                    data-bill-result-nav="true"
                    onClick={handlePrintReceipt}
                    style={{ flex: 1, padding: '12px', backgroundColor: '#2563eb', color: '#fff', border: 'none', borderRadius: '6px', fontWeight: 'bold', cursor: 'pointer', fontSize: '14px' }}
                  >🖨️ Receipt Print</button>
                  <button
                    ref={newBillRef}
                    type="button"
                    data-bill-result-nav="true"
                    onClick={handleStartNewBill}
                    style={{ flex: 1, padding: '12px', backgroundColor: '#059669', color: '#fff', border: 'none', borderRadius: '6px', fontWeight: 'bold', cursor: 'pointer', fontSize: '14px' }}
                  >➕ New Bill</button>
                </div>
                <h3 style={{ textAlign: 'center', color: '#059669', marginTop: 0 }}>🎉 Bill Completed Successfully!</h3>
                
                <div style={{ border: '1px dashed #9ca3af', padding: '12px', fontFamily: receiptSettings.fontFamily, fontSize: receiptSettings.fontSize, color: '#000', backgroundColor: '#fff', marginBottom: '15px', border: '1px solid #dbeafe', boxShadow: '0 8px 24px rgba(15,23,42,.08)', borderRadius: '10px', padding: '12px', ...receiptElementStyle('invoice', { fontSize: '14px', fontWeight: 700, textAlign: 'left' }) }}>
                  {receiptSettings.logoUrl && (
                    <div style={{ textAlign: 'center', marginBottom: '4px' }}>
                      <img src={receiptSettings.logoUrl} alt="Logo" style={{ width: '35px', height: '35px', objectFit: 'cover', borderRadius: '50%' }} />
                    </div>
                  )}
                  <div style={{ textAlign: 'center', fontWeight: '900', fontSize: '1.5em' }}>{receiptSettings.shopName}</div>
                  <div style={{ textAlign: 'center', fontWeight: '900', whiteSpace: 'pre-line' }}>{receiptSettings.address}</div>
                  <div style={{ textAlign: 'center', fontWeight: '900' }}>Tel: {receiptSettings.phone}</div>
                  <div style={{ borderTop: '2px dashed #000', margin: '6px 0' }}></div>
                  <div style={{ display: 'flex', justifyContent: 'space-between' }}><span>Invoice No:</span><span>{completedSale.invoiceNo}</span></div>
                  <div style={{ display: 'flex', justifyContent: 'space-between' }}><span>Date:</span><span>{completedSale.dateFormatted}</span></div>
                   <div style={{ display: 'flex', justifyContent: 'space-between', fontSize: '10px' }}><span>Price Type:</span><span>{completedSale.priceLevelLabel || getPriceLevelLabel(completedSale.priceLevel || 'normal')}</span></div>
                  <div style={{ borderTop: '2px dashed #000', margin: '6px 0' }}></div>
                  {completedSale.items.map((item, idx) => (
                    <div key={idx} style={{ marginBottom: '5px', fontWeight: '900' }}>
                      <div style={{ display: 'flex', justifyContent: 'space-between', wordBreak: 'break-all' }}>
                        <span>{item.name} x {item.qty}</span>
                        <span>Rs. {(parseFloat(item.finalPrice) || 0).toFixed(2)}</span>
                      </div>
                      <div style={{ fontSize: '10px', paddingLeft: '8px' }}>
                         {item.priceLevel && item.priceLevel !== 'normal' && (
                           <span style={{ marginRight: '5px' }}>Retail: <span style={{ textDecoration: 'line-through' }}>Rs. {Number(item.retailPrice ?? item.price ?? 0).toFixed(2)}</span></span>
                         )}
                         <span>{getPriceLevelLabel(item.priceLevel || 'normal')}: Rs. {Number(item.unitPrice ?? item.price ?? 0).toFixed(2)}</span>
                         {item.offerLabel ? ` • ${item.offerLabel}` : ''}
                       </div>
                      {Number(item.discount || 0) > 0 && <div style={{ fontSize: '10px', paddingLeft: '8px' }}>Discount: -Rs. {Number(item.discount).toFixed(2)}</div>}
                    </div>
                  ))}
                  <div style={{ borderTop: '2px dashed #000', margin: '6px 0' }}></div>
                  <div style={{ display: 'flex', justifyContent: 'space-between', fontWeight: '700' }}><span>Subtotal:</span><span>Rs. {Number(completedSale.subtotal ?? completedSale.total).toFixed(2)}</span></div>
                  <div style={{ display: 'flex', justifyContent: 'space-between', fontWeight: '700', color: '#dc2626' }}><span>Discount:</span><span>-Rs. {Number(completedSale.discount ?? 0).toFixed(2)}</span></div>
                  {Number(completedSale.billDiscountAmount || 0) > 0 && <div style={{ display:'flex', justifyContent:'space-between', fontSize:'11px', color:'#ea580c' }}><span>{completedSale.billDiscountType === 'lkr' ? 'Whole Bill Discount (Direct LKR):' : `Whole Bill Discount (${Number(completedSale.billDiscountPercent || 0).toFixed(2)}%):`}</span><span>-Rs. {Number(completedSale.billDiscountAmount || 0).toFixed(2)}</span></div>}
                  <div style={{ display: 'flex', justifyContent: 'space-between', fontWeight: '900', fontSize: '1.3em' }}><span>GRAND TOTAL:</span><span>Rs. {completedSale.total.toFixed(2)}</span></div>
                  <div style={{ textAlign: 'center', marginTop: '8px', fontWeight: '900', fontSize: '1.2em' }}>
                    <div>{receiptSettings.footerMsg}</div>
                    {receiptSettings.footerComment && (
                      <div style={{ marginTop: '6px', whiteSpace: 'pre-line', overflowWrap: 'anywhere', fontSize: '0.9em' }}>{receiptSettings.footerComment}</div>
                    )}
                  </div>
                </div>

              </div>
            </div>
            </div>
          ) : (
            <div style={{ display: 'flex', flexDirection: 'column', height: '100%', gap: '15px' }}>
              
<div className="no-print" style={{
                display: 'flex',
                alignItems: 'center',
                justifyContent: 'space-between',
                gap: '14px',
                flexWrap: 'wrap',
                backgroundColor: '#e0f2fe',
                padding: '12px 16px',
                borderRadius: '10px',
                border: '1px solid #7dd3fc',
                boxShadow: '0 2px 8px rgba(2,132,199,0.12)',
                marginBottom: '15px'
              }}>
                <div style={{ display: 'flex', alignItems: 'center', gap: '12px', minWidth: 0, flex: '1 1 360px' }}>
                  {receiptSettings.logoUrl ? (
                    <img src={receiptSettings.logoUrl} alt="Shop Logo" style={{ width: '42px', height: '42px', objectFit: 'cover', borderRadius: '50%', border: '2px solid #3b82f6', flexShrink: 0 }} />
                  ) : (
                    <div style={{ width: '42px', height: '42px', backgroundColor: '#3b82f6', color: '#fff', borderRadius: '50%', display: 'flex', alignItems: 'center', justifyContent: 'center', fontWeight: 'bold', fontSize: '13px', flexShrink: 0 }}>POS</div>
                  )}
                  <div style={{ minWidth: 0 }}>
                    <div style={{ fontSize: '18px', fontWeight: '900', color: '#1f2937', overflow: 'hidden', textOverflow: 'ellipsis', whiteSpace: 'nowrap' }}>
                      {receiptSettings.shopName || 'POSHITHA POS SYSTEM'}
                    </div>
                    <div style={{ fontSize: '12px', color: '#475569', overflowWrap: 'anywhere', marginTop:'2px', whiteSpace:'pre-line' }}>
                      📍 {receiptSettings.address || '-'}
                    </div>
                    <div style={{ fontSize: '12px', color: '#475569', overflowWrap: 'anywhere' }}>
                      📞 {receiptSettings.phone || '-'} &nbsp;•&nbsp; ✉️ {receiptSettings.email || '-'}
                    </div>
                  </div>
                </div>
                <div style={{ display: 'flex', alignItems: 'center', gap: '9px', flexWrap: 'wrap', justifyContent: 'flex-end' }}>
                  <div style={{
                    fontSize: '12px',
                    backgroundColor: '#f8fafc',
                    color: '#334155',
                    padding: '7px 11px',
                    borderRadius: '7px',
                    fontWeight: '800',
                    border: '1px solid #e2e8f0',
                    textAlign: 'right',
                    lineHeight: 1.35
                  }}>
                    <div>📅 {currentDateTime.toLocaleDateString()}</div>
                    <div style={{ color: '#2563eb' }}>🕒 {currentDateTime.toLocaleTimeString()}</div>
                  </div>
                  <button
                    type="button"
                    onClick={() => { setInventoryShortcutMode(null); setActiveTab('home'); }}
                    style={{ padding: '9px 14px', background: '#2563eb', color: '#fff', border: 'none', borderRadius: '7px', cursor: 'pointer', fontWeight: '800' }}
                  >🏠 Home</button>
                  <button
                    type="button"
                    onClick={() => setIsLoggedIn(false)}
                    style={{ padding: '9px 14px', background: '#dc2626', color: '#fff', border: 'none', borderRadius: '7px', cursor: 'pointer', fontWeight: '800' }}
                  >🚪 {t.logout}</button>
                </div>
              </div>

              {errorMessage && (
                <div style={{ backgroundColor: '#fee2e2', color: '#dc2626', padding: '12px 15px', borderRadius: '8px', fontWeight: 'bold', display: 'flex', justifyContent: 'space-between', alignItems: 'center', border: '1px solid #fecaca', boxShadow: '0 2px 5px rgba(0,0,0,0.05)' }}>
                  <span>⚠️ {errorMessage}</span>
                  <button onClick={() => setErrorMessage('')} style={{ background: 'none', border: 'none', color: '#dc2626', fontWeight: 'bold', cursor: 'pointer', fontSize: '16px' }}>✕</button>
                </div>
              )}

              <div style={{ display: 'flex', gap: '20px', flex: 1, overflow: 'hidden' }}>
              
                <div style={{ flex: 1.5, backgroundColor: '#fff', padding: '20px', borderRadius: '8px', border: '1px solid #e5e7eb', display: 'flex', flexDirection: 'column' }}>
                  
                  <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', marginBottom: '10px' }}>
                    <h3 style={{ margin: 0, color: '#1f2937' }}>🔍 {ui.scanBarcode}</h3>
                    {cart.length > 0 && (
                      <button onClick={handleHoldBill} style={{ padding: '6px 14px', backgroundColor: '#f59e0b', color: '#fff', border: 'none', borderRadius: '6px', cursor: 'pointer', fontWeight: 'bold' }}>
                        ⏸️ Hold Bill
                      </button>
                    )}
                  </div>

                  {heldBills.length > 0 && (
                    <div style={{ backgroundColor: '#fef3c7', padding: '10px', borderRadius: '6px', marginBottom: '15px', display: 'flex', gap: '10px', alignItems: 'center', overflowX: 'auto' }}>
                      <span style={{ fontSize: '12px', fontWeight: 'bold', color: '#92400e' }}>Held Bills ({heldBills.length}):</span>
                      {heldBills.map((hb) => (
                        <button key={hb.id} onClick={() => handleResumeBill(hb)} style={{ padding: '4px 10px', backgroundColor: '#b45309', color: '#fff', border: 'none', borderRadius: '4px', cursor: 'pointer', fontSize: '12px', fontWeight: 'bold' }}>
                          ▶️ Resume ({hb.customerName || 'Cust'} - {hb.time})
                        </button>
                      ))}
                    </div>
                  )}

                  <form onSubmit={handleBarcodeSubmit} style={{ display: 'flex', gap: '10px', marginBottom: '15px' }}>
                    <input 
                      ref={barcodeInputRef}
                      data-pos-nav="true"
                      type="text" 
                      placeholder="Scan Barcode Here & Press Enter..." 
                      value={barcodeInput}
                      onChange={(e) => setBarcodeInput(e.target.value)}
                      onKeyDown={(e) => {
                        if (e.key === 'Enter') {
                          e.preventDefault();
                          handleBarcodeSubmit(e);
                        }
                      }}
                      style={{ flex: 1, padding: '14px', borderRadius: '6px', border: '2px solid #3b82f6', fontSize: '18px', outline: 'none' }}
                    />
                    <button type="submit" data-pos-nav="true" style={{ padding: '0 25px', backgroundColor: '#3b82f6', color: '#fff', border: 'none', borderRadius: '6px', fontWeight: 'bold', cursor: 'pointer' }}>Add</button>
                  </form>

                  <div style={{ marginBottom: '15px' }}>
                    <select
                      ref={productSelectRef}
                      data-pos-nav="true"
                      onChange={(e) => {
                        const prod = products.find(p => p.id === parseInt(e.target.value));
                        if (prod) addToCart(prod);
                        e.target.value = '';
                      }}
                      style={{ width: '100%', padding: '10px', borderRadius: '6px', border: '1px solid #d1d5db', backgroundColor: '#f9fafb' }}
                    >
                      <option value="">-- OR Select Product From Inventory List --</option>
                      {products.map(p => (
                        <option key={p.id} value={p.id}>
                          {p.name} {p.lot_number ? `- Lot: ${p.lot_number}` : ''} - (Rs. {p.price}.00) - [Stock: {p.stock} {p.unit}]
                        </option>
                      ))}
                    </select>
                  </div>

                  <div style={{ display:'flex', alignItems:'center', justifyContent:'space-between', gap:'8px', marginBottom:'10px' }}>
                    <h4 style={{ margin:0, color:'#4b5563' }}>Items in Bill</h4>
                    <div aria-label="Current bill total" style={{ display:'inline-flex', alignItems:'center', gap:'5px', padding:'4px 8px', borderRadius:'6px', background:'#ecfdf5', border:'1px solid #a7f3d0', color:'#047857', fontSize:'12px', fontWeight:900, whiteSpace:'nowrap' }}>
                      <span>Total</span><span>Rs. {total.toFixed(2)}</span>
                    </div>
                  </div>
                  <div style={{ flex: 1, overflowY: 'auto' }}>
                    <table style={{ width: '100%', borderCollapse: 'collapse' }}>
                      <thead>
                        <tr style={{ backgroundColor: '#2563eb', color: '#ffffff', textAlign: 'left', borderBottom: '1px solid #1d4ed8' }}>
                          <th style={{ padding: '10px' }}>{ui.itemName}</th>
                          <th style={{ padding: '10px' }}>{ui.price}</th>
                          <th style={{ padding: '10px' }}>Qty (Pcs/kg)</th>
                          <th style={{ padding: '10px' }}>{ui.offer}</th>
                          <th style={{ padding: '10px' }}>Total</th>
                          <th style={{ padding: '10px' }}>{ui.action}</th>
                        </tr>
                      </thead>
                      <tbody>
                        {cart.map((item) => (
                          <tr key={item.id} style={{ borderBottom: '1px solid #f3f4f6' }}>
                            <td style={{ padding: '10px', fontWeight: 'bold' }}>
                              {item.name}
                              {(item.barcode || item.item_code || item.code) && <div style={{ marginTop: '2px', fontSize: '10px', color: '#9ca3af', fontWeight: '500' }}>{item.barcode || item.item_code || item.code}</div>}
                              {item.lot_number && <div style={{ marginTop: '3px', fontSize: '11px', color: '#7c3aed' }}>Lot: {item.lot_number}</div>}
                            </td>
                            <td style={{ padding: '10px' }}>
                              Rs. {Number(item.offer_type === 'bulk_price' ? (item.effectiveUnitPrice ?? item.price ?? 0) : (item.price ?? 0)).toFixed(2)}
                              {item.offer_type === 'bulk_price' && Number(item.qty) >= Math.max(2, Number(item.bulk_min_qty) || Number(item.offer_buy_qty) || 2) && (
                                <div style={{ fontSize: '10px', color: '#15803d', fontWeight: 'bold', marginTop: '2px' }}>Bulk Unit Price</div>
                              )}
                            </td>
                            <td style={{ padding: '10px' }}>
                              <div style={{ display: 'flex', alignItems: 'center', gap: '5px' }}>
                                <input 
                                  type="number" step="any" value={item.qty}
                                  onChange={(e) => handleQtyInputChange(item.id, e.target.value)}
                                  style={{ width: '70px', padding: '6px', borderRadius: '4px', border: '1px solid #3b82f6', fontWeight: 'bold', textAlign: 'center' }}
                                />
                                <span style={{ fontSize: '12px', color: '#6b7280', fontWeight: 'bold' }}>{item.unit}</span>
                              </div>
                            </td>
                            <td style={{ padding: '10px' }}>
                              {item.offerLabel ? (
                                <span style={{ display: 'inline-block', padding: '5px 8px', borderRadius: '5px', backgroundColor: '#f3e8ff', color: '#7e22ce', fontSize: '12px', fontWeight: 'bold', border: '1px solid #d8b4fe', whiteSpace: 'nowrap' }}>
                                  🎁 {item.offerLabel}
                                </span>
                              ) : (
                                <span style={{ color: '#9ca3af', fontSize: '12px' }}>— No Offer</span>
                              )}
                            </td>
                            <td style={{ padding: '10px', fontWeight: 'bold', color: '#059669' }}>
                              Rs. {(parseFloat(item.finalPrice) || 0).toFixed(2)}
                              {Number(item.discount || 0) > 0 && (
                                <div style={{ fontSize: '11px', color: '#dc2626', marginTop: '3px' }}>Discount: -Rs. {Number(item.discount).toFixed(2)}</div>
                              )}
                            </td>
                            <td style={{ padding: '10px' }}>
                              <button data-pos-nav="true" onClick={() => removeFromCart(item.id)} style={{ padding: '4px 8px', backgroundColor: '#ef4444', color: '#fff', border: 'none', borderRadius: '4px', cursor: 'pointer', fontSize: '12px' }}>❌ {ui.remove}</button>
                            </td>
                          </tr>
                        ))}
                        {cart.length === 0 && (
                          <tr>
                            <td colSpan="6" style={{ textAlign: 'center', padding: '30px', color: '#9ca3af' }}>No items scanned yet. Scan a barcode or select above!</td>
                          </tr>
                        )}
                      </tbody>
                    </table>
                  </div>
                </div>

                <div style={{ flex: 1, backgroundColor: '#fff', padding: '20px', borderRadius: '8px', border: '1px solid #e5e7eb', display: 'flex', flexDirection: 'column', justifyContent: 'space-between', overflowY: 'auto' }}>
                  <div>
                    <h3 style={{ marginTop: 0 }}>{ui.paymentCustomer}</h3>

                    <div style={{ marginBottom: '15px', display: 'flex', flexDirection: 'column', gap: '8px' }}>
                      <label style={{ fontSize: '13px', fontWeight: 'bold', color: '#374151' }}>{ui.customerInfo}:</label>
                      <input data-pos-nav="true" type="text" placeholder="Customer Name" value={customerName} onChange={(e) => setCustomerName(e.target.value)} style={{ padding: '8px', borderRadius: '6px', border: '1px solid #ccc' }} />
                      <input data-pos-nav="true" type="text" placeholder="Phone Number" value={customerPhone} onChange={(e) => setCustomerPhone(e.target.value)} style={{ padding: '8px', borderRadius: '6px', border: '1px solid #ccc' }} />
                    </div>
                    
                    <div style={{ marginBottom: '15px' }}>
                      <label style={{ display: 'block', fontSize: '13px', fontWeight: 'bold', color: '#374151', marginBottom: '6px' }}>{ui.paymentMethod}</label>
                      <div
                        ref={paymentMethodRef}
                        tabIndex={-1}
                        aria-label={ui.paymentMethod}
                        style={{ display: 'grid', gridTemplateColumns: 'repeat(4, 1fr)', gap: '8px', outline: 'none' }}
                      >
                        <button data-pos-nav="true" onClick={() => setPaymentMethod('Cash')} style={{ padding: '8px', borderRadius: '6px', border: '2px solid', borderColor: paymentMethod === 'Cash' ? '#10b981' : '#e5e7eb', backgroundColor: paymentMethod === 'Cash' ? '#ecfdf5' : '#fff', color: paymentMethod === 'Cash' ? '#047857' : '#374151', fontWeight: 'bold', cursor: 'pointer' }}>💵 {ui.cash}</button>
                        <button data-pos-nav="true" onClick={() => setPaymentMethod('Card')} style={{ padding: '8px', borderRadius: '6px', border: '2px solid', borderColor: paymentMethod === 'Card' ? '#3b82f6' : '#e5e7eb', backgroundColor: paymentMethod === 'Card' ? '#eff6ff' : '#fff', color: paymentMethod === 'Card' ? '#1d4ed8' : '#374151', fontWeight: 'bold', cursor: 'pointer' }}>💳 {ui.card}</button>
                        <button data-pos-nav="true" onClick={() => setPaymentMethod('Online')} style={{ padding: '8px', borderRadius: '6px', border: '2px solid', borderColor: paymentMethod === 'Online' ? '#8b5cf6' : '#e5e7eb', backgroundColor: paymentMethod === 'Online' ? '#f5f3ff' : '#fff', color: paymentMethod === 'Online' ? '#6d28d9' : '#374151', fontWeight: 'bold', cursor: 'pointer' }}>📱 Online / QR</button>
                        <button data-pos-nav="true" onClick={() => { setPaymentMethod('Credit'); setPaidAmount('0'); if (!registeredCustomers.length) setShowCustomerRegisterModal(true); }} style={{ padding: '8px', borderRadius: '6px', border: '2px solid', borderColor: paymentMethod === 'Credit' ? '#f59e0b' : '#e5e7eb', backgroundColor: paymentMethod === 'Credit' ? '#fffbeb' : '#fff', color: paymentMethod === 'Credit' ? '#b45309' : '#374151', fontWeight: 'bold', cursor: 'pointer' }}>🧾 Credit / ණයට</button>
                      </div>
                    </div>

                    <button data-pos-nav="true" onClick={handleOpenCreditPayment} style={{ width: '100%', padding: '9px', marginBottom: '8px', background: '#7c3aed', color: '#fff', border: 'none', borderRadius: '6px', fontWeight: 'bold', cursor: 'pointer' }}>💰 Receive Credit Payment / ණය ගෙවීම</button>
                    <button data-pos-nav="true" onClick={() => setShowCustomerDisplay(v => !v)} style={{ width:'100%', padding:'9px', marginBottom:'12px', background:'#0f766e', color:'#fff', border:0, borderRadius:'6px', fontWeight:'bold', cursor:'pointer' }}>🖥️ {showCustomerDisplay ? 'Hide' : 'Show'} Customer Display</button>

                    <div style={{ marginBottom: '12px', padding: '10px', background: '#eff6ff', border: '1px solid #bfdbfe', borderRadius: '8px' }}>
                      <label style={{ display: 'block', fontSize: '13px', fontWeight: 'bold', marginBottom: '6px' }}>Customer Price Type / මිල වර්ගය</label>
                      <select data-pos-nav="true" value={priceLevel} onChange={e => applyPriceLevelToCart(e.target.value)} style={{ width: '100%', padding: '9px', border: '1px solid #60a5fa', borderRadius: '6px', fontWeight: 'bold' }}>
                        <option value="normal">Normal Customer — Retail Price</option>
                        <option value="wholesale">Wholesale Customer — Wholesale Price</option>
                        <option value="special">Special Customer — Special Price</option>
                      </select>
                    </div>

                    {paymentMethod === 'Credit' && (
                      <div style={{ marginBottom: '15px', padding: '10px', background: '#fffbeb', border: '1px solid #fde68a', borderRadius: '8px' }}>
                        <div style={{ display: 'flex', gap: '8px', alignItems: 'center' }}>
                          <select data-pos-nav="true" value={selectedCreditCustomerId} onChange={e => handleSelectCreditCustomer(e.target.value)} style={{ flex: 1, padding: '9px', borderRadius: '6px', border: '1px solid #f59e0b' }}>
                            <option value="">Select Registered Customer</option>
                            {registeredCustomers.map(c => <option key={c.id} value={c.id}>{c.name} - {c.phone} • Due Rs. {getCustomerBalance(c.id).toFixed(2)}</option>)}
                          </select>
                          <button data-pos-nav="true" onClick={() => setShowCustomerRegisterModal(true)} style={{ padding: '9px 12px', background: '#f59e0b', color: '#fff', border: 'none', borderRadius: '6px', fontWeight: 'bold', cursor: 'pointer' }}>➕ Register</button>
                        </div>
                        {getSelectedCreditCustomer() && <div style={{ marginTop: '7px', fontSize: '12px', fontWeight: 'bold' }}>Current Outstanding: Rs. {getCustomerBalance(getSelectedCreditCustomer().id).toFixed(2)} • New Due: Rs. {Math.max(0, total - numericPaid).toFixed(2)}</div>}
                      </div>
                    )}

                    <div style={{ marginBottom: '15px' }}>
                      <label style={{ display: 'block', fontSize: '13px', fontWeight: 'bold', color: '#374151', marginBottom: '4px' }}>Paid Amount (ගෙවූ මුදල):</label>
                      <input ref={paidAmountRef} data-pos-nav="true" type="number" placeholder="e.g. 1000" value={paidAmount} onChange={(e) => setPaidAmount(e.target.value)} style={{ width: '100%', padding: '10px', borderRadius: '6px', border: '2px solid #10b981', fontSize: '16px', fontWeight: 'bold', outline: 'none', backgroundColor: '#fff', color: '#000' }} />
                    </div>

                    <div style={{ marginBottom:'12px', padding:'10px', background:'#fff7ed', border:'1px solid #fdba74', borderRadius:'8px' }}>
                      <label style={{ display:'block', fontSize:'13px', fontWeight:'bold', marginBottom:'7px' }}>Whole Bill Discount (Item offers වලට අමතරව)</label>
                      <div style={{ display:'grid', gridTemplateColumns:'1fr 1fr', gap:'8px' }}>
                        <div>
                          <div style={{ fontSize:'12px', fontWeight:'bold', marginBottom:'4px' }}>Discount %</div>
                          <input data-pos-nav="true" type="number" min="0" max="100" step="0.01" value={billDiscountPercent} onChange={e => { setBillDiscountPercent(e.target.value); if ((Number(e.target.value) || 0) > 0) setBillDiscountLkr('0'); }} style={{ width:'100%', padding:'9px', border:'2px solid #f97316', borderRadius:'6px', fontWeight:'bold', boxSizing:'border-box' }} />
                        </div>
                        <div>
                          <div style={{ fontSize:'12px', fontWeight:'bold', marginBottom:'4px' }}>Direct Discount (LKR)</div>
                          <input data-pos-nav="true" type="number" min="0" step="0.01" value={billDiscountLkr} onChange={e => { setBillDiscountLkr(e.target.value); if ((Number(e.target.value) || 0) > 0) setBillDiscountPercent('0'); }} style={{ width:'100%', padding:'9px', border:'2px solid #f97316', borderRadius:'6px', fontWeight:'bold', boxSizing:'border-box' }} />
                        </div>
                      </div>
                      <div style={{ marginTop:'6px', fontSize:'12px' }}>Item-offer total: Rs. {itemOfferTotal.toFixed(2)} • Bill discount: -Rs. {billDiscountAmount.toFixed(2)} {billDiscountType === 'lkr' ? '(Direct LKR)' : billDiscountType === 'percent' ? `(${normalizedBillDiscountPercent.toFixed(2)}%)` : ''}</div>
                    </div>

                    <div style={{ backgroundColor: '#f9fafb', padding: '12px', borderRadius: '6px', border: '1px solid #f3f4f6' }}>
                      <div style={{ display: 'flex', justifyContent: 'space-between', marginBottom: '6px', fontSize: '14px' }}><span>Total Items:</span><b>{cart.length}</b></div>
                      <div style={{ display: 'flex', justifyContent: 'space-between', marginBottom: '6px', fontSize: '14px' }}><span>Paid Amount:</span><b>Rs. {numericPaid.toFixed(2)}</b></div>
                      <div style={{ display: 'flex', justifyContent: 'space-between', fontSize: '15px', color: '#dc2626', fontWeight: 'bold' }}><span>Balance Change (ඉතිරි මුදල):</span><span>Rs. {changeAmount.toFixed(2)}</span></div>
                    </div>
                  </div>

                  <div style={{ borderTop: '2px solid #f3f4f6', paddingTop: '15px', marginTop: '15px' }}>
                    <div style={{ display: 'flex', justifyContent: 'space-between', fontSize: '20px', fontWeight: 'bold', marginBottom: '12px' }}>
                      <span>Total Amount:</span>
                      <span style={{ color: '#059669' }}>Rs. {total.toFixed(2)}</span>
                    </div>
                    <button ref={payPrintRef} data-pos-nav="true" onClick={handleOpenPreview} disabled={cart.length === 0} style={{ width: '100%', padding: '14px', backgroundColor: cart.length > 0 ? '#059669' : '#9ca3af', color: '#fff', border: 'none', borderRadius: '6px', fontSize: '18px', fontWeight: 'bold', cursor: cart.length > 0 ? 'pointer' : 'not-allowed' }}>
                      Pay & Print Receipt
                    </button>
                  </div>
                  {showCustomerDisplay && (
                    <div style={{ marginTop:'15px', padding:'14px', background:'#0f172a', color:'#fff', borderRadius:'10px', maxHeight:'360px', overflowY:'auto' }}>
                      <h3 style={{ margin:'0 0 10px', textAlign:'center' }}>🖥️ Customer Display</h3>
                      {cart.map((item, i) => <div key={`${item.id}-${i}`} style={{ display:'grid', gridTemplateColumns:'1fr 55px 90px 90px', gap:'6px', padding:'6px 0', borderBottom:'1px solid #334155', fontSize:'12px' }}><span>{item.name}</span><span>x{item.qty}</span><span>Rs. {Number(item.unitPrice ?? item.price ?? 0).toFixed(2)}</span><span>Rs. {Number(item.finalPrice || 0).toFixed(2)}</span></div>)}
                      <div style={{ marginTop:'10px', display:'grid', gap:'4px', fontWeight:'bold' }}>
                        <div style={{ display:'flex', justifyContent:'space-between' }}><span>Discount</span><span>-Rs. {(cartDiscount + billDiscountAmount).toFixed(2)}</span></div>
                        <div style={{ display:'flex', justifyContent:'space-between', fontSize:'20px' }}><span>Total</span><span>Rs. {total.toFixed(2)}</span></div>
                        <div style={{ display:'flex', justifyContent:'space-between' }}><span>Paid</span><span>Rs. {numericPaid.toFixed(2)}</span></div>
                        <div style={{ display:'flex', justifyContent:'space-between' }}><span>Balance</span><span>Rs. {Math.max(0, total - numericPaid).toFixed(2)}</span></div>
                      </div>
                    </div>
                  )}
                </div>

              </div>
            </div>
          )
        )}

        {activeTab === 'inventory' && (
          <div>
            <div style={{ backgroundColor: '#1e293b', color: '#fff', padding: '20px', borderRadius: '10px', marginBottom: '20px', boxShadow: '0 4px 6px rgba(0,0,0,0.1)' }}>
              <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', marginBottom: '15px' }}>
                <h3 style={{ margin: 0, color: '#38bdf8', fontSize: '18px' }}>📦 TOTAL INVENTORY STOCK SUMMARY (සමස්ත ගබඩා වටිනාකම)</h3>
                <div style={{ display: 'flex', alignItems: 'center', gap: '10px' }}>
                  {receiptSettings.logoUrl && <img src={receiptSettings.logoUrl} alt="Logo" style={{ width: '35px', height: '35px', objectFit: 'cover', borderRadius: '50%', border: '1px solid #38bdf8' }} />}
                  <span style={{ fontWeight: 'bold', fontSize: '15px', color: '#f8fafc' }}>{receiptSettings.shopName}</span>
                </div>
              </div>
              <div style={{ display: 'grid', gridTemplateColumns: 'repeat(3, 1fr)', gap: '15px', marginTop: '10px' }}>
                <div style={{ backgroundColor: '#0f172a', padding: '12px', borderRadius: '8px' }}>
                  <span style={{ fontSize: '12px', color: '#94a3b8' }}>🏢 TOTAL STOCK COST (ගැනුම් වටිනාකම)</span>
                  <h2 style={{ margin: '5px 0 0 0', color: '#38bdf8' }}>Rs. {totalStockCostValue.toFixed(2)}</h2>
                </div>
                <div style={{ backgroundColor: '#0f172a', padding: '12px', borderRadius: '8px' }}>
                  <span style={{ fontSize: '12px', color: '#94a3b8' }}>🏷️ EXPECTED SELLING VALUE (විකුණුම් වටිනාකම)</span>
                  <h2 style={{ margin: '5px 0 0 0', color: '#4ade80' }}>Rs. {totalStockSellingValue.toFixed(2)}</h2>
                </div>
                <div style={{ backgroundColor: '#0f172a', padding: '12px', borderRadius: '8px' }}>
                  <span style={{ fontSize: '12px', color: '#94a3b8' }}>📊 TOTAL ITEMS (මුළු ප්‍රමාණය)</span>
                  <h2 style={{ margin: '5px 0 0 0', color: '#facc15' }}>{totalStockItemsCount} Barcodes</h2>
                </div>
              </div>
            </div>

            <div style={{ position: 'sticky', top: 0, zIndex: 25, backgroundColor: '#fff', padding: '12px', borderRadius: '8px', border: '1px solid #e5e7eb', boxShadow: '0 2px 8px rgba(0,0,0,0.08)', marginBottom: '10px' }}>
              <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', marginBottom: '10px' }}>
                <h2>{editingId ? '✏️ Edit Product' : '➕ Add New Product'}</h2>
                <div style={{ display: 'flex', gap: '8px' }}>
                  <label style={{ padding: '8px 12px', backgroundColor: '#8b5cf6', color: '#fff', borderRadius: '6px', cursor: 'pointer', fontWeight: 'bold', display: 'flex', alignItems: 'center', gap: '5px' }}>
                    📥 Import Excel
                    <input type="file" accept=".xlsx, .xls" onChange={handleExcelImport} style={{ display: 'none' }} />
                  </label>
                  <button onClick={exportToExcel} style={{ padding: '8px 12px', backgroundColor: '#10b981', color: '#fff', border: 'none', borderRadius: '6px', cursor: 'pointer', fontWeight: 'bold' }}>📊 Export Excel</button>
                  <button onClick={exportToPDF} style={{ padding: '8px 12px', backgroundColor: '#6366f1', color: '#fff', border: 'none', borderRadius: '6px', cursor: 'pointer', fontWeight: 'bold' }}>📄 Export PDF</button>
                  {products.length > 0 && <button onClick={handleClearAllProducts} style={{ padding: '8px 12px', backgroundColor: '#dc2626', color: '#fff', border: 'none', borderRadius: '6px', cursor: 'pointer', fontWeight: 'bold' }}>🗑️ Clear All</button>}
                  <button type="button" onClick={() => { setEditingId(null); resetForm(); setShowProductForm(true); window.setTimeout(() => document.getElementById('inventory-product-form')?.scrollIntoView({ behavior:'smooth', block:'start' }), 50); }} style={{ padding: '8px 12px', backgroundColor: '#059669', color: '#fff', border: 'none', borderRadius: '6px', cursor: 'pointer', fontWeight: 'bold' }}>➕ Add New Product</button>
                </div>
              </div>

              <div style={{ display:'flex', gap:'8px', alignItems:'center', marginBottom:'10px', padding:'10px', background:'#eff6ff', border:'1px solid #bfdbfe', borderRadius:'8px' }}>
                <span style={{ fontWeight:'bold', whiteSpace:'nowrap' }}>📷 Scan New Product Barcode:</span>
                <form onSubmit={handleInventoryBarcodeScan} style={{ display:'flex', gap:'8px', flex:1 }}>
                  <input ref={inventoryScanRef} type="text" autoComplete="off" placeholder="Scan barcode here and press Enter" value={inventoryScanCode} onChange={e => setInventoryScanCode(e.target.value)} style={{ flex:1, padding:'9px', border:'2px solid #2563eb', borderRadius:'6px', fontWeight:'bold' }} />
                  <button type="submit" style={{ padding:'9px 14px', border:0, borderRadius:'6px', background:'#2563eb', color:'#fff', fontWeight:'bold', cursor:'pointer' }}>Use Barcode</button>
                </form>
              </div>
              {showProductForm && (
              <form id="inventory-product-form" onSubmit={handleSaveProduct} style={{ display: 'grid', gridTemplateColumns: 'repeat(4, minmax(0, 1fr))', gap: '10px', backgroundColor: '#fff' }}>
                <input ref={inventoryNameRef} type="text" placeholder="Product Name" value={newProd.name} onChange={(e) => setNewProd({ ...newProd, name: e.target.value })} required style={{ padding: '8px' }} />
                <div style={{ display: 'flex', gap: '4px' }}>
                  <input type="text" placeholder="Barcode (Or Auto)" value={newProd.barcode} onChange={(e) => setNewProd({ ...newProd, barcode: e.target.value })} style={{ padding: '8px', flex: 1 }} />
                  <button type="button" onClick={generateBarcode} title="Generate Random Barcode" style={{ padding: '8px', backgroundColor: '#8b5cf6', color: '#fff', border: 'none', borderRadius: '4px', cursor: 'pointer' }}>⚡</button>
                </div>
                <input type="text" placeholder="Lot Number (e.g. LOT-001)" value={newProd.lot_number} onChange={(e) => setNewProd({ ...newProd, lot_number: e.target.value })} style={{ padding: '8px', border: '2px solid #8b5cf6' }} />
                <input type="number" placeholder="GRN Rate" value={newProd.grn_rate} onChange={(e) => setNewProd({ ...newProd, grn_rate: e.target.value })} style={{ padding: '8px' }} />


                <select value={newProd.unit} onChange={(e) => setNewProd({ ...newProd, unit: e.target.value })} style={{ padding: '8px' }} title="Product Unit">
                  {productUnits.map((unit) => (
                    <option key={unit} value={unit}>{unit}</option>
                  ))}
                </select>
                <input type="number" placeholder="Buy Price" value={newProd.buying_price} onChange={(e) => setNewProd({ ...newProd, buying_price: e.target.value })} style={{ padding: '8px' }} />
                <input type="number" placeholder="Sell Price" value={newProd.price} onChange={(e) => setNewProd({ ...newProd, price: e.target.value })} required style={{ padding: '8px' }} />
                <input type="number" placeholder="Wholesale Price" value={newProd.wholesale_price} onChange={(e) => setNewProd({ ...newProd, wholesale_price: e.target.value })} style={{ padding: '8px', border: '2px solid #2563eb' }} />
                <input type="number" placeholder="Special Price" value={newProd.special_price} onChange={(e) => setNewProd({ ...newProd, special_price: e.target.value })} style={{ padding: '8px', border: '2px solid #8b5cf6' }} />
                <input type="number" placeholder="Stock Qty" value={newProd.stock} onChange={(e) => setNewProd({ ...newProd, stock: e.target.value })} style={{ padding: '8px' }} />

                <input type="number" placeholder="Low Stock Number" value={newProd.min_stock_alert} onChange={(e) => setNewProd({ ...newProd, min_stock_alert: e.target.value })} required style={{ padding: '8px', border: '1px solid #f87171' }} />
                <input type="date" placeholder="Expiry Date" title="Expiry Date" value={newProd.expiry_date} onChange={(e) => setNewProd({ ...newProd, expiry_date: e.target.value })} style={{ padding: '8px', border: '2px solid #f59e0b' }} />
                <select value={newProd.supplier} onChange={(e) => setNewProd({ ...newProd, supplier: e.target.value })} style={{ padding: '8px' }}>
                  <option value="">Select Supplier</option>
                  {suppliers.map(s => <option key={s.id} value={s.name}>{s.name}</option>)}
                </select>
                <select value={newProd.offer_type} onChange={(e) => setNewProd({ ...newProd, offer_type: e.target.value })} style={{ padding: '8px' }}>
                  <option value="none">No Offer</option><option value="percent">Percentage Discount (%)</option><option value="b1g1">Buy X Get Y Free</option><option value="bulk_price">Bulk / Quantity Price</option>
                </select>

                {newProd.offer_type === 'percent' && <input type="number" min="0" max="100" placeholder="Discount %" value={newProd.offer_value} onChange={(e) => setNewProd({ ...newProd, offer_value: e.target.value })} style={{ padding: '8px' }} />}
                {newProd.offer_type === 'b1g1' && (
                  <div style={{ display: 'flex', gap: '6px', alignItems: 'center', padding: '4px', backgroundColor: '#eff6ff', border: '1px solid #bfdbfe', borderRadius: '6px' }}>
                    <span style={{ fontSize: '12px', fontWeight: 'bold', color: '#1d4ed8', whiteSpace: 'nowrap' }}>Buy</span>
                    <input type="number" min="1" step="1" value={newProd.offer_buy_qty} onChange={(e) => setNewProd({ ...newProd, offer_buy_qty: e.target.value })} style={{ width: '55px', padding: '7px', border: '1px solid #93c5fd', borderRadius: '4px' }} />
                    <span style={{ fontSize: '12px', fontWeight: 'bold', color: '#1d4ed8', whiteSpace: 'nowrap' }}>Get</span>
                    <input type="number" min="1" step="1" value={newProd.offer_free_qty} onChange={(e) => setNewProd({ ...newProd, offer_free_qty: e.target.value })} style={{ width: '55px', padding: '7px', border: '1px solid #93c5fd', borderRadius: '4px' }} />
                    <span style={{ fontSize: '12px', fontWeight: 'bold', color: '#059669', whiteSpace: 'nowrap' }}>Free</span>
                  </div>
                )}
                {newProd.offer_type === 'bulk_price' && (
                  <div style={{ display: 'flex', gap: '8px', alignItems: 'center', padding: '6px', backgroundColor: '#f0fdf4', border: '1px solid #bbf7d0', borderRadius: '6px' }}>
                    <span style={{ fontSize: '12px', fontWeight: 'bold', color: '#166534', whiteSpace: 'nowrap' }}>Min Qty</span>
                    <input type="number" min="2" step="1" value={newProd.bulk_min_qty} onChange={(e) => setNewProd({ ...newProd, bulk_min_qty: e.target.value })} style={{ width: '70px', padding: '7px', border: '1px solid #86efac', borderRadius: '4px' }} />
                    <span style={{ fontSize: '12px', fontWeight: 'bold', color: '#166534', whiteSpace: 'nowrap' }}>Unit Price</span>
                    <input type="number" min="0" step="0.01" placeholder="e.g. 150" value={newProd.bulk_price} onChange={(e) => setNewProd({ ...newProd, bulk_price: e.target.value })} style={{ width: '95px', padding: '7px', border: '1px solid #86efac', borderRadius: '4px' }} />
                    <span style={{ fontSize: '11px', color: '#15803d', whiteSpace: 'nowrap' }}>Qty 2+ → Rs. 150 each</span>
                  </div>
                )}

                <div style={{ gridColumn: 'span 4', display: 'flex', gap: '10px' }}>
                  <button type="submit" style={{ padding: '10px 20px', backgroundColor: editingId ? '#f59e0b' : '#2563eb', color: '#fff', border: 'none', borderRadius: '4px', cursor: 'pointer', fontWeight: 'bold' }}>{editingId ? 'Update Product' : 'Save Product'}</button>
                  {editingId && <button type="button" onClick={resetForm} style={{ padding: '10px 15px', backgroundColor: '#6b7280', color: '#fff', border: 'none', borderRadius: '4px', cursor: 'pointer' }}>Cancel</button>}
                </div>
              </form>
              )}
            </div>

            {/* ==================== EXPIRY REPORT ==================== */}
            <div style={{ backgroundColor: '#fff', border: '1px solid #e5e7eb', borderRadius: '10px', padding: '14px', marginBottom: '10px', boxShadow: '0 2px 5px rgba(0,0,0,0.06)' }}>
              <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', gap: '10px', flexWrap: 'wrap' }}>
                <div>
                  <h3 style={{ margin: 0 }}>⏳ Expiry Management</h3>
                  <div style={{ marginTop: '5px', color: '#6b7280', fontSize: '12px' }}>
                    🔴 Expired: <b style={{ color: '#dc2626' }}>{expirySummary.expired}</b>
                    &nbsp; | &nbsp;
                    🟡 Near Expiry (≤ {EXPIRY_WARNING_DAYS} days): <b style={{ color: '#d97706' }}>{expirySummary.nearExpiry}</b>
                    &nbsp; | &nbsp;
                    🟢 Valid with Expiry: <b style={{ color: '#059669' }}>{expirySummary.validWithExpiry}</b>
                  </div>
                </div>
                <button
                  onClick={handleDownloadExpiryExcel}
                  disabled={!expiryReportProducts.length}
                  style={{
                    padding: '9px 13px',
                    background: expiryReportProducts.length ? '#10b981' : '#9ca3af',
                    color: '#fff',
                    border: 'none',
                    borderRadius: '6px',
                    cursor: expiryReportProducts.length ? 'pointer' : 'not-allowed',
                    fontWeight: 'bold'
                  }}
                >
                  📊 Download Expired / Near Expiry Excel
                </button>
              </div>

              {expiryReportProducts.length > 0 && (
                <div style={{ overflowX: 'auto', marginTop: '12px' }}>
                  <table style={{ width: '100%', borderCollapse: 'collapse', fontSize: '13px' }}>
                    <thead>
                      <tr style={{ backgroundColor: '#fff7ed', textAlign: 'left' }}>
                        <th style={{ padding: '8px' }}>Barcode</th>
                        <th style={{ padding: '8px' }}>Product</th>
                        <th style={{ padding: '8px' }}>Stock</th>
                        <th style={{ padding: '8px' }}>Expiry Date</th>
                        <th style={{ padding: '8px' }}>Days</th>
                        <th style={{ padding: '8px' }}>Status</th>
                      </tr>
                    </thead>
                    <tbody>
                      {expiryReportProducts.slice(0, 50).map(({ product, info }) => (
                        <tr
                          key={`expiry-${product.id}`}
                          style={{
                            borderBottom: '1px solid #eee',
                            backgroundColor: info.isExpired ? '#fef2f2' : '#fffbeb'
                          }}
                        >
                          <td style={{ padding: '8px', fontFamily: 'monospace' }}>{product.barcode || '-'}</td>
                          <td style={{ padding: '8px', fontWeight: 'bold' }}>{product.name}</td>
                          <td style={{ padding: '8px' }}>{product.stock} {product.unit || 'Pcs'}</td>
                          <td style={{ padding: '8px' }}>{product.expiry_date || '-'}</td>
                          <td style={{ padding: '8px', fontWeight: 'bold', color: info.isExpired ? '#dc2626' : '#d97706' }}>
                            {info.isExpired ? `${Math.abs(info.daysRemaining)} days overdue` : `${info.daysRemaining} days`}
                          </td>
                          <td style={{ padding: '8px', fontWeight: 'bold', color: info.isExpired ? '#dc2626' : '#d97706' }}>
                            {info.isExpired ? '🔴 EXPIRED' : '🟡 NEAR EXPIRY'}
                          </td>
                        </tr>
                      ))}
                    </tbody>
                  </table>
                  {expiryReportProducts.length > 50 && (
                    <div style={{ marginTop: '8px', color: '#6b7280', fontSize: '12px' }}>
                      First 50 items are shown here. Download Excel to get the complete list ({expiryReportProducts.length} items).
                    </div>
                  )}
                </div>
              )}

              {expiryReportProducts.length === 0 && (
                <div style={{ marginTop: '10px', padding: '10px', background: '#ecfdf5', borderRadius: '6px', color: '#047857', fontSize: '13px', fontWeight: 'bold' }}>
                  ✅ දැනට Expired හෝ Near Expiry products නොමැත.
                </div>
              )}
            </div>

            <div style={{ position: 'sticky', top: '125px', zIndex: 24, backgroundColor: '#fff', padding: '12px', borderRadius: '8px', border: '1px solid #e5e7eb', marginBottom: '10px', display: 'flex', gap: '8px', alignItems: 'center', flexWrap: 'wrap', boxShadow: '0 2px 5px rgba(0,0,0,0.08)' }}>
              <b style={{ color: '#1f2937' }}>🔎 Search Inventory</b>
              <input type="text" value={inventorySearch} onChange={e => { setInventorySearch(e.target.value); setInventoryPage(1); }} placeholder={`${uiMore.searchInventory || ui.searchInventory} — Barcode, Item Name or Supplier`} style={{ flex: 1, minWidth: '260px', padding: '10px', border: '2px solid #3b82f6', borderRadius: '6px' }} />
              <button onClick={() => { setInventorySearch(''); setInventoryPage(1); }} style={{ padding: '10px 14px', background: '#6b7280', color: '#fff', border: 'none', borderRadius: '6px', cursor: 'pointer' }}>{ui.clear}</button>
              <span style={{ fontSize: '12px', color: '#6b7280' }}>{filteredInventory.length} items found</span>
            </div>
            <div id="inventory-table-section">
              <h3>{ui.inventoryTable}</h3>
              <table style={{ width: '100%', backgroundColor: '#fff', borderCollapse: 'collapse', marginTop: '10px' }}>
              <thead>
                <tr style={{ backgroundColor: '#e5e7eb', textAlign: 'left' }}>
                  <th style={{ padding: '10px' }}>Barcode</th><th style={{ padding: '10px' }}>Lot</th><th style={{ padding: '10px' }}>Name</th><th style={{ padding: '10px' }}>{uiMore.buyPrice}</th><th style={{ padding: '10px' }}>Retail Price</th><th style={{ padding: '10px' }}>Wholesale Price</th><th style={{ padding: '10px' }}>Special Price</th><th style={{ padding: '10px' }}>{uiMore.stock}</th><th style={{ padding: '10px' }}>{uiMore.unit}</th><th style={{ padding: '10px' }}>{ui.offer}</th><th style={{ padding: '10px' }}>{ui.supplier}</th><th style={{ padding: '10px' }}>{uiMore.expiry}</th><th style={{ padding: '10px' }}>Actions</th>
                </tr>
              </thead>
              <tbody>
                {pagedInventory.map((p) => {
                  const itemStock = parseFloat(p.stock) || 0;
                  const alertLimit = parseFloat(p.min_stock_alert) || 5;
                  const isLowStock = itemStock <= alertLimit;
                  const expiryInfo = getExpiryInfo(p);
                  const isExpired = expiryInfo.isExpired;
                  const isNearExpiry = expiryInfo.isNearExpiry;

                  return (
                    <tr key={p.id} style={{ borderBottom: '1px solid #eee', backgroundColor: isExpired ? '#fef2f2' : isNearExpiry ? '#fffbeb' : isLowStock ? '#fef2f2' : 'transparent' }}>
                      <td style={{ padding: '10px', fontFamily: 'monospace' }}>{p.barcode}</td>
                      <td style={{ padding: '10px', fontWeight: 'bold', color: '#7c3aed' }}>{p.lot_number || '—'}</td>
                      <td style={{ padding: '10px', fontWeight: 'bold' }}>{p.name}</td>
                      <td style={{ padding: '10px', color: '#6b7280' }}>Rs. {p.buying_price || 0}.00</td>
                      <td style={{ padding: '10px', color: '#059669', fontWeight: 'bold' }}>Rs. {Number(p.price || 0).toFixed(2)}</td>
                      <td style={{ padding: '10px', color: '#2563eb', fontWeight: 'bold' }}>Rs. {Number(p.wholesale_price || 0).toFixed(2)}</td>
                      <td style={{ padding: '10px', color: '#7c3aed', fontWeight: 'bold' }}>Rs. {Number(p.special_price || 0).toFixed(2)}</td>
                      <td style={{ padding: '10px', color: isLowStock ? '#dc2626' : '#000', fontWeight: isLowStock ? 'bold' : 'normal' }}>{p.stock} {isLowStock && '⚠️'}</td>
                      <td style={{ padding: '10px', fontWeight: 'bold' }}>{p.unit || 'Pcs'}</td>
                      <td style={{ padding: '10px', color: p.offer_type !== 'none' ? '#7c3aed' : '#6b7280', fontWeight: p.offer_type !== 'none' ? 'bold' : 'normal' }}>{p.offer_type === 'percent' ? `${p.offer_value || 0}% OFF` : p.offer_type === 'b1g1' ? `BUY ${p.offer_buy_qty || (Number(p.offer_value) >= 1001 ? Math.floor(Number(p.offer_value) / 1000) : 1)} GET ${p.offer_free_qty || (Number(p.offer_value) >= 1001 ? Number(p.offer_value) % 1000 : 1)} FREE` : p.offer_type === 'bulk_price' ? `Rs. ${p.bulk_price || 0} EACH FROM ${p.bulk_min_qty || 2}` : '—'}</td>
                      <td style={{ padding: '10px', color: '#2563eb' }}>{p.supplier || '—'}</td>
                      <td style={{ padding: '10px', color: isExpired ? '#dc2626' : isNearExpiry ? '#d97706' : '#111827', fontWeight: isExpired || isNearExpiry ? 'bold' : 'normal' }}>
                        {p.expiry_date || 'N/A'}
                        {isExpired && <div style={{ fontSize: '11px', color: '#dc2626' }}>🔴 EXPIRED</div>}
                        {!isExpired && isNearExpiry && <div style={{ fontSize: '11px', color: '#d97706' }}>🟡 {expiryInfo.daysRemaining} days left</div>}
                      </td>
                      <td style={{ padding: '10px', display: 'flex', gap: '6px' }}>
                        <button onClick={() => handleEditClick(p)} style={{ padding: '4px 8px', backgroundColor: '#f59e0b', color: '#fff', border: 'none', borderRadius: '4px', cursor: 'pointer' }}>{uiMore.edit}</button>
                        <button onClick={() => openBarcodePrintModal(p)} title="Print Barcode" style={{ padding: '4px 9px', backgroundColor: '#2563eb', color: '#fff', border: 'none', borderRadius: '4px', cursor: 'pointer', fontSize: '15px', lineHeight: 1 }}>🖨️</button>
                        <button onClick={() => handleDeleteProduct(p.id)} style={{ padding: '4px 8px', backgroundColor: '#ef4444', color: '#fff', border: 'none', borderRadius: '4px', cursor: 'pointer' }}>Delete</button>
                      </td>
                    </tr>
                  );
                })}
              </tbody>
            </table>
            <PaginationControls page={inventoryPage} pageCount={filteredInventoryPageCount} setPage={setInventoryPage} totalRows={filteredInventory.length} label="Inventory Items" />
            </div>
          </div>
        )}

        {activeTab === 'suppliers' && currentUserRole === 'admin' && (
          <div>
            <h2>🚚 {ui.supplierDashboard}</h2>
            
            <div style={{ backgroundColor: '#fff', padding: '20px', borderRadius: '8px', marginBottom: '20px', border: '1px solid #e5e7eb', boxShadow: '0 2px 5px rgba(0,0,0,0.05)' }}>
              <h3 style={{ marginTop: 0, color: '#1f2937' }}>{editingSupplierId ? '✏️ Edit Supplier' : '➕ Add New Supplier'}</h3>
              <form onSubmit={handleSaveSupplier} style={{ display: 'grid', gridTemplateColumns: 'repeat(3, 1fr)', gap: '12px' }}>
                <input type="text" placeholder="Supplier Name *" value={newSupplier.name} onChange={(e) => setNewSupplier({ ...newSupplier, name: e.target.value })} required style={{ padding: '10px', borderRadius: '6px', border: '1px solid #ccc' }} />
                <input type="text" placeholder="Company / Agency Name" value={newSupplier.company} onChange={(e) => setNewSupplier({ ...newSupplier, company: e.target.value })} style={{ padding: '10px', borderRadius: '6px', border: '1px solid #ccc' }} />
                <input type="text" placeholder="Phone Number" value={newSupplier.phone} onChange={(e) => setNewSupplier({ ...newSupplier, phone: e.target.value })} style={{ padding: '10px', borderRadius: '6px', border: '1px solid #ccc' }} />
                <input type="email" placeholder="Email Address" value={newSupplier.email} onChange={(e) => setNewSupplier({ ...newSupplier, email: e.target.value })} style={{ padding: '10px', borderRadius: '6px', border: '1px solid #ccc' }} />
                <input type="text" placeholder="Address" value={newSupplier.address} onChange={(e) => setNewSupplier({ ...newSupplier, address: e.target.value })} style={{ padding: '10px', borderRadius: '6px', border: '1px solid #ccc', gridColumn: 'span 2' }} />
                <div style={{ gridColumn: 'span 3', display: 'flex', gap: '10px', marginTop: '5px' }}>
                  <button type="submit" style={{ padding: '10px 20px', backgroundColor: editingSupplierId ? '#f59e0b' : '#2563eb', color: '#fff', border: 'none', borderRadius: '6px', fontWeight: 'bold', cursor: 'pointer' }}>
                    {editingSupplierId ? 'Update Supplier' : 'Save Supplier'}
                  </button>
                  {editingSupplierId && (
                    <button type="button" onClick={() => { setEditingSupplierId(null); setNewSupplier({ name: '', phone: '', email: '', address: '', company: '' }); }} style={{ padding: '10px 15px', backgroundColor: '#6b7280', color: '#fff', border: 'none', borderRadius: '6px', cursor: 'pointer' }}>
                      Cancel
                    </button>
                  )}
                </div>
              </form>
            </div>

            <h3>Suppliers Directory</h3>
            <table style={{ width: '100%', backgroundColor: '#fff', borderCollapse: 'collapse', marginBottom: '25px', borderRadius: '8px', overflow: 'hidden' }}>
              <thead>
                <tr style={{ backgroundColor: '#e5e7eb', textAlign: 'left' }}>
                  <th style={{ padding: '10px' }}>Name</th><th style={{ padding: '10px' }}>Company</th><th style={{ padding: '10px' }}>Phone</th><th style={{ padding: '10px' }}>Email</th><th style={{ padding: '10px' }}>Address</th><th style={{ padding: '10px' }}>Actions</th>
                </tr>
              </thead>
              <tbody>
                {pagedSuppliers.map(sup => (
                  <tr key={sup.id} style={{ borderBottom: '1px solid #eee' }}>
                    <td style={{ padding: '10px', fontWeight: 'bold' }}>{sup.name}</td>
                    <td style={{ padding: '10px' }}>{sup.company || '-'}</td>
                    <td style={{ padding: '10px' }}>{sup.phone || '-'}</td>
                    <td style={{ padding: '10px' }}>{sup.email || '-'}</td>
                    <td style={{ padding: '10px' }}>{sup.address || '-'}</td>
                    <td style={{ padding: '10px', display: 'flex', gap: '6px' }}>
                      <button onClick={() => handleEditSupplier(sup)} style={{ padding: '4px 8px', backgroundColor: '#f59e0b', color: '#fff', border: 'none', borderRadius: '4px', cursor: 'pointer' }}>{uiMore.edit}</button>
                      <button onClick={() => handleDeleteSupplier(sup.id)} style={{ padding: '4px 8px', backgroundColor: '#ef4444', color: '#fff', border: 'none', borderRadius: '4px', cursor: 'pointer' }}>Delete</button>
                    </td>
                  </tr>
                ))}
              </tbody>
            </table>
            <PaginationControls page={supplierPage} pageCount={supplierPageCount} setPage={setSupplierPage} totalRows={suppliers.length} label="Suppliers" />

            <h3 style={{ marginTop: '30px', borderTop: '2px solid #e5e7eb', paddingTop: '20px' }}>📦 Low Stock Items & Supplier Order Requests</h3>
            <div style={{ backgroundColor: '#fff', padding: '15px', borderRadius: '8px', marginBottom: '20px', border: '1px solid #e5e7eb' }}>
              <div style={{ display: 'flex', justifyContent: 'space-between', gap: '10px', flexWrap: 'wrap', alignItems: 'center' }}>
                <div><h4 style={{ margin: 0 }}>Low Stock Items — Order Builder</h4><p style={{ margin: '5px 0 12px', color: '#6b7280', fontSize: '12px' }}>Low stock items 50 බැගින් pages ලෙස පෙන්වයි. අවශ්‍ය items select කර Qty සහ Supplier තෝරන්න.</p></div>
                <input value={supplierLowStockSearch} onChange={e => { setSupplierLowStockSearch(e.target.value); setSupplierLowStockPage(1); }} placeholder="Search low stock by name/barcode/supplier..." style={{ padding: '9px', minWidth: '280px', border: '1px solid #3b82f6', borderRadius: '6px' }} />
              </div>
              {lowStockList.length === 0 ? (
                <div style={{ color: '#059669', fontWeight: 'bold', padding: '15px 0' }}>✅ දැනට Low Stock items නොමැත.</div>
              ) : (
                <>
                  <div style={{ display: 'flex', gap: '8px', flexWrap: 'wrap', margin: '10px 0' }}>
                    <button onClick={handleSelectAllLowStock} style={{ padding: '8px 12px', background: '#2563eb', color: '#fff', border: 'none', borderRadius: '6px', cursor: 'pointer', fontWeight: 'bold' }}>☑️ Select All</button>
                    <button onClick={handleClearLowStockSelection} style={{ padding: '8px 12px', background: '#6b7280', color: '#fff', border: 'none', borderRadius: '6px', cursor: 'pointer', fontWeight: 'bold' }}>Clear Selection</button>
                    <button onClick={handleStartOrderGeneration} style={{ padding: '8px 14px', background: '#059669', color: '#fff', border: 'none', borderRadius: '6px', cursor: 'pointer', fontWeight: 'bold' }}>✅ ALL CONFIRM — Generate Order</button>
                    <button onClick={handleDownloadLowStockExcel} style={{ padding: '8px 12px', background: '#10b981', color: '#fff', border: 'none', borderRadius: '6px', cursor: 'pointer', fontWeight: 'bold' }}>📊 {ui.downloadExcel}</button>
                    {(lowStockList.length > 0) && <button onClick={handleDeleteAllLowStockItems} style={{ padding: '8px 12px', background: '#dc2626', color: '#fff', border: 'none', borderRadius: '6px', cursor: 'pointer', fontWeight: 'bold' }}>🗑️ {ui.deleteAll}</button>}
                  </div>
                  <div style={{ overflowX: 'auto' }}><table style={{ width: '100%', backgroundColor: '#fff', borderCollapse: 'collapse', fontSize: '13px' }}>
                    <thead><tr style={{ backgroundColor: '#e5e7eb', textAlign: 'left' }}><th style={{ padding: '8px' }}>Select</th><th style={{ padding: '8px' }}>Item</th><th style={{ padding: '8px' }}>Barcode</th><th style={{ padding: '8px' }}>Current</th><th style={{ padding: '8px' }}>{uiMore.buyPrice}</th><th style={{ padding: '8px' }}>{uiMore.unit}</th><th style={{ padding: '8px' }}>Alert</th><th style={{ padding: '8px' }}>{ui.orderQty}</th><th style={{ padding: '8px' }}>{ui.supplier}</th><th style={{ padding: '8px' }}>Delete</th></tr></thead>
                    <tbody>{pagedSupplierLowStock.map(item => {
                      const qty = customOrderInputs[item.id] ?? getSuggestedOrderQty(item);
                      const supplierValue = orderSupplierInputs[item.id] ?? item.supplier ?? '';
                      return <tr key={item.id} style={{ borderBottom: '1px solid #eee', background: selectedLowStockItems[item.id] ? '#ecfdf5' : '#fff' }}>
                        <td style={{ padding: '8px' }}><input type="checkbox" checked={!!selectedLowStockItems[item.id]} onChange={() => handleToggleLowStockSelection(item.id)} /></td>
                        <td style={{ padding: '8px', fontWeight: 'bold' }}>{item.name}</td><td style={{ padding: '8px', fontFamily: 'monospace' }}>{item.barcode}</td><td style={{ padding: '8px', color: '#dc2626', fontWeight: 'bold' }}>{item.stock}</td>
                        <td style={{ padding: '8px', fontWeight: 'bold' }}>Rs. {(Number(item.buying_price) || 0).toFixed(2)}</td><td style={{ padding: '8px' }}>{item.unit || 'Pcs'}</td><td style={{ padding: '8px' }}>{item.min_stock_alert || 5}</td>
                        <td style={{ padding: '8px' }}><input type="number" min="1" value={qty} onChange={e => handleCustomOrderInput(item.id, e.target.value)} style={{ width: '85px', padding: '6px', border: '1px solid #3b82f6', borderRadius: '5px', fontWeight: 'bold' }} /></td>
                        <td style={{ padding: '8px' }}><select value={supplierValue} onChange={e => handleCustomSupplierInput(item.id, e.target.value)} style={{ padding: '6px', minWidth: '170px', border: '1px solid #cbd5e1', borderRadius: '5px' }}><option value="">{uiMore.selectSupplier}</option>{suppliers.map(sup => <option key={sup.id} value={sup.name}>{sup.name}</option>)}</select></td>
                        <td style={{ padding: '8px' }}><button onClick={() => handleDeleteLowStockItem(item)} style={{ padding: '6px 9px', background: '#ef4444', color: '#fff', border: 'none', borderRadius: '5px', cursor: 'pointer' }}>🗑️ Delete</button></td>
                      </tr>;
                    })}</tbody>
                  </table></div>
                  <PaginationControls page={supplierLowStockPage} pageCount={supplierLowStockPageCount} setPage={setSupplierLowStockPage} totalRows={filteredSupplierLowStockList.length} label="Low Stock Items" />
                </>
              )}
            </div>

          </div>
        )}

        {showOrderGenerateModal && orderDraft && (
          <div className="no-print" style={{ position: 'fixed', inset: 0, background: 'rgba(0,0,0,.65)', zIndex: 5000, display: 'flex', alignItems: 'center', justifyContent: 'center', padding: '20px' }}>
            <div style={{ background: '#fff', width: 'min(1050px, 96vw)', maxHeight: '92vh', overflowY: 'auto', borderRadius: '14px', padding: '22px', boxShadow: '0 20px 60px rgba(0,0,0,.3)' }}>
              <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'flex-start', gap: '15px', marginBottom: '15px' }}><div><h2 style={{ margin: 0 }}>🧾 Generate Supplier Order</h2><div style={{ color: '#6b7280', marginTop: '5px' }}>Order No: <b>{orderDraft.orderNumber}</b></div></div><button onClick={() => { setShowOrderGenerateModal(false); setOrderDraft(null); }} style={{ border: 'none', background: '#fee2e2', color: '#b91c1c', borderRadius: '6px', padding: '8px 12px', cursor: 'pointer' }}>✕</button></div>
              <div style={{ display: 'grid', gridTemplateColumns: 'repeat(3, 1fr)', gap: '10px', background: '#f8fafc', padding: '14px', borderRadius: '10px', marginBottom: '15px' }}>
                <div><b>{uiMore.company}</b><div>{receiptSettings.shopName}</div></div><div><b>Address:</b><div>{receiptSettings.address}</div></div><div><b>Phone:</b><div>{receiptSettings.phone}</div></div>
                <div><b>Supplier:</b><div>{orderDraft.supplierName}</div></div><div><b>{uiMore.supplierCompany}</b><div>{orderDraft.supplierCompany || '-'}</div></div><div><b>{uiMore.supplierPhone}</b><div>{orderDraft.supplierPhone || '-'}</div></div>
              </div>
              <div style={{ overflowX: 'auto' }}><table style={{ width: '100%', borderCollapse: 'collapse' }}><thead><tr style={{ background: '#1f2937', color: '#fff', textAlign: 'left' }}><th style={{ padding: '9px' }}>#</th><th style={{ padding: '9px' }}>Item</th><th style={{ padding: '9px' }}>Barcode</th><th style={{ padding: '9px' }}>{uiMore.currentStock}</th><th style={{ padding: '9px' }}>{uiMore.buyPrice}</th><th style={{ padding: '9px' }}>{ui.orderQty}</th><th style={{ padding: '9px' }}>{uiMore.unit}</th><th style={{ padding: '9px' }}>{uiMore.cost}</th></tr></thead><tbody>{orderDraft.items.map((it, idx) => <tr key={it.id || idx} style={{ borderBottom: '1px solid #e5e7eb' }}><td style={{ padding: '9px' }}>{idx + 1}</td><td style={{ padding: '9px', fontWeight: 'bold' }}>{it.name}</td><td style={{ padding: '9px', fontFamily: 'monospace' }}>{it.barcode}</td><td style={{ padding: '9px' }}>{it.currentStock}</td><td style={{ padding: '9px', fontWeight: 'bold' }}>Rs. {(Number(it.buyingPrice) || 0).toFixed(2)}</td><td style={{ padding: '9px' }}><input type="number" min="1" value={it.orderQty} onChange={e => handleDraftOrderItemChange(idx, 'orderQty', e.target.value)} style={{ width: '90px', padding: '7px', border: '2px solid #3b82f6', borderRadius: '5px', fontWeight: 'bold' }} /></td><td style={{ padding: '9px' }}>{it.unit}</td><td style={{ padding: '9px', fontWeight: 'bold', color: '#059669' }}>Rs. {((Number(it.buyingPrice) || 0) * (Number(it.orderQty) || 0)).toFixed(2)}</td></tr>)}</tbody></table></div>
              <div style={{ textAlign: 'right', fontSize: '18px', fontWeight: 'bold', marginTop: '12px', color: '#059669' }}>TOTAL COST: Rs. {Number(orderDraft.totalCost || 0).toFixed(2)}</div>
              <div style={{ display: 'flex', justifyContent: 'flex-end', gap: '10px', marginTop: '18px', flexWrap: 'wrap' }}><button onClick={() => handleDownloadOrderExcel(orderDraft)} style={{ padding: '11px 16px', background: '#10b981', color: '#fff', border: 'none', borderRadius: '6px', fontWeight: 'bold', cursor: 'pointer' }}>📊 {ui.downloadExcel}</button><button onClick={() => { const preview = { ...orderDraft, id: orderDraft.orderNumber }; handleDownloadOrderPDF(preview); }} style={{ padding: '11px 16px', background: '#6366f1', color: '#fff', border: 'none', borderRadius: '6px', fontWeight: 'bold', cursor: 'pointer' }}>📄 PDF</button><button onClick={handleConfirmGeneratedOrder} style={{ padding: '11px 20px', background: '#059669', color: '#fff', border: 'none', borderRadius: '6px', fontWeight: 'bold', cursor: 'pointer' }}>✅ OK — Save Order</button></div>
            </div>
          </div>
        )}

        {activeTab === 'shopOrderHistory' && currentUserRole === 'admin' && (
          <div>
            <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', flexWrap: 'wrap', gap: '10px', marginBottom: '15px' }}>
              <div><h2>🧾 {ui.shopOrderHistory}</h2><p style={{ margin: 0, color: '#6b7280' }}>ALL CONFIRM කළ Shop Orders Order Number එක අනුව මෙහි ස්ථිරව save වේ.</p></div>
              <div style={{ display: 'flex', gap: '8px' }}>
                <button onClick={handleDownloadShopOrderHistoryExcel} style={{ padding: '9px 13px', background: '#10b981', color: '#fff', border: 'none', borderRadius: '6px', cursor: 'pointer', fontWeight: 'bold' }}>📊 {ui.downloadExcel}</button>
                {shopOrderHistory.length > 0 && <button onClick={handleDeleteAllShopOrders} style={{ padding: '9px 13px', background: '#dc2626', color: '#fff', border: 'none', borderRadius: '6px', cursor: 'pointer', fontWeight: 'bold' }}>🗑️ {ui.deleteAll}</button>}
              </div>
            </div>
            {shopOrderHistory.length === 0 ? <div style={{ background:'#fff', padding:'30px', borderRadius:'10px', color:'#6b7280' }}>{uiMore.noShopOrders}</div> : shopOrderHistory.map(ord => (
              <div key={ord.id} style={{ background:'#fff', padding:'16px', borderRadius:'10px', border:'1px solid #e5e7eb', marginBottom:'15px' }}>
                <div style={{ display:'flex', justifyContent:'space-between', alignItems:'center', flexWrap:'wrap', gap:'10px' }}>
                  <div><h3 style={{ margin:'0 0 5px' }}>{ord.orderNumber}</h3><div style={{ fontSize:'12px', color:'#6b7280' }}>Date: {ord.date} • Status: {ord.status}</div></div>
                  <div style={{ display:'flex', gap:'7px', alignItems:'center', flexWrap:'wrap' }}><select value={ord.status || 'Pending'} onChange={e => setShopOrderHistory(prev => prev.map(o => o.id === ord.id ? { ...o, status: e.target.value } : o))} style={{ padding:'7px', borderRadius:'5px', fontWeight:'bold' }}><option>{ui.pending}</option><option>{ui.ongoing}</option><option>{ui.complete}</option></select><button onClick={() => handleDownloadOrderExcel(ord)} style={{ padding:'7px 10px', background:'#10b981', color:'#fff', border:'none', borderRadius:'5px', cursor:'pointer' }}>📊 Excel</button><button onClick={() => handleDownloadOrderPDF(ord)} style={{ padding:'7px 10px', background:'#6366f1', color:'#fff', border:'none', borderRadius:'5px', cursor:'pointer' }}>📄 PDF</button><button onClick={() => handleDeleteShopOrder(ord.id)} style={{ padding:'7px 10px', background:'#ef4444', color:'#fff', border:'none', borderRadius:'5px', cursor:'pointer' }}>🗑️ Delete</button></div>
                </div>
                <div style={{ overflowX:'auto', marginTop:'12px' }}><table style={{ width:'100%', borderCollapse:'collapse', fontSize:'13px' }}><thead><tr style={{ background:'#f3f4f6', textAlign:'left' }}><th style={{padding:'8px'}}>#</th><th style={{padding:'8px'}}>{ui.supplier}</th><th style={{padding:'8px'}}>Item</th><th style={{padding:'8px'}}>Barcode</th><th style={{padding:'8px'}}>{uiMore.buyPrice}</th><th style={{padding:'8px'}}>Qty</th><th style={{padding:'8px'}}>{uiMore.unit}</th><th style={{padding:'8px'}}>{uiMore.cost}</th></tr></thead><tbody>{(ord.items||[]).map((it,i)=><tr key={i} style={{borderBottom:'1px solid #eee'}}><td style={{padding:'8px'}}>{i+1}</td><td style={{padding:'8px'}}>{it.supplier}</td><td style={{padding:'8px',fontWeight:'bold'}}>{it.name}</td><td style={{padding:'8px',fontFamily:'monospace'}}>{it.barcode}</td><td style={{padding:'8px'}}>Rs. {(Number(it.buyingPrice)||0).toFixed(2)}</td><td style={{padding:'8px'}}>{it.orderQty}</td><td style={{padding:'8px'}}>{it.unit}</td><td style={{padding:'8px',fontWeight:'bold'}}>Rs. {((Number(it.buyingPrice)||0)*(Number(it.orderQty)||0)).toFixed(2)}</td></tr>)}</tbody></table></div>
                <div style={{ textAlign:'right', marginTop:'10px', fontSize:'17px', fontWeight:'bold', color:'#059669' }}>TOTAL COST: Rs. {Number(ord.totalCost||0).toFixed(2)}</div>
              </div>
            ))}
          </div>
        )}

        {activeTab === 'customers' && currentUserRole === 'admin' && (
          <div>
            <h2>👥 {ui.customerDashboard}</h2>
            <p style={{ marginTop: '-8px', marginBottom: '15px', color: '#6b7280' }}>Customer purchase history සහ bill records මෙතැනින් බලන්න. Credit customers management සඳහා 📘 Credit Book භාවිතා කරන්න. Selected date: <b>{selectedReportDate}</b></p>

            <div style={{ background:'#fff', border:'1px solid #c7d2fe', borderRadius:'10px', padding:'15px', marginBottom:'15px' }}>
              <h3 style={{ margin:'0 0 12px' }}>⭐ Priority 2 — Customer Loyalty, Discounts & Credit</h3>
              <div style={{ display:'grid', gridTemplateColumns:'repeat(2,minmax(0,1fr))', gap:'12px' }}>
                <div style={{ padding:'12px', background:'#f8fafc', borderRadius:'8px' }}>
                  <b>Customer Loyalty Points</b>
                  <div style={{ fontSize:'12px', color:'#64748b', margin:'4px 0 8px' }}>Purchase Rs.100 = 1 point. Balance / redeem / lifetime report.</div>
                  <div style={{ display:'flex', gap:'6px' }}><select value={loyaltyRedeemCustomerId} onChange={e=>setLoyaltyRedeemCustomerId(e.target.value)} style={{ flex:1, padding:'7px' }}><option value="">Select Customer</option>{registeredCustomers.map(c=><option key={c.id} value={c.id}>{c.name} • {Number(c.loyaltyPoints||0)} pts</option>)}</select><input type="number" min="1" placeholder="Points" value={loyaltyRedeemPoints} onChange={e=>setLoyaltyRedeemPoints(e.target.value)} style={{ width:'90px', padding:'7px' }}/><button onClick={handleRedeemLoyaltyPoints} style={{ padding:'7px 10px', background:'#7c3aed', color:'#fff', border:0, borderRadius:'5px' }}>Redeem</button></div>
                </div>
                <div style={{ padding:'12px', background:'#f8fafc', borderRadius:'8px' }}>
                  <b>Customer Discounts / Membership</b>
                  <div style={{ fontSize:'12px', color:'#64748b', margin:'4px 0 8px' }}>Customer-specific discount, membership level, wholesale / retail pricing profile.</div>
                  <div style={{ fontSize:'12px' }}>Edit values directly in the customer table below; settings are saved with the customer profile.</div>
                </div>
              </div>
              {registeredCustomers.length > 0 && <div style={{ overflowX:'auto', marginTop:'12px' }}><table style={{ width:'100%', borderCollapse:'collapse', fontSize:'12px' }}><thead><tr style={{ background:'#eef2ff' }}><th style={{padding:'7px'}}>Customer</th><th>Points</th><th>Membership</th><th>Discount %</th><th>Price Type</th><th>Due / Status</th><th>Reminder</th></tr></thead><tbody>{registeredCustomers.map(c=><tr key={`priority-${c.id}`} style={{ borderBottom:'1px solid #e5e7eb' }}><td style={{padding:'7px'}}><b>{c.name}</b><div>{c.phone}</div></td><td>{Number(c.loyaltyPoints||0)}<div style={{color:'#64748b'}}>Lifetime {Number(c.lifetimePoints||0)}</div></td><td><select value={c.membershipLevel||'Regular'} onChange={e=>updateCustomerSettings(c.id,{membershipLevel:e.target.value})}><option>Regular</option><option>Silver</option><option>Gold</option><option>Platinum</option><option>Wholesale</option></select></td><td><input type="number" min="0" max="100" value={Number(c.customerDiscount||0)} onChange={e=>updateCustomerSettings(c.id,{customerDiscount:Number(e.target.value)||0})} style={{width:'65px'}} /></td><td><select value={c.priceType||'Retail'} onChange={e=>updateCustomerSettings(c.id,{priceType:e.target.value})}><option>Retail</option><option>Wholesale</option></select></td><td>Rs. {getCustomerBalance(c.id).toFixed(2)}<div style={{fontWeight:'bold',color:isCustomerOverdue(c.id)?'#dc2626':'#059669'}}>{isCustomerOverdue(c.id)?'OVERDUE':'OK'}</div></td><td><button onClick={()=>sendCustomerReminder(c)} disabled={getCustomerBalance(c.id)<=0} style={{padding:'5px 7px',background:getCustomerBalance(c.id)>0?'#16a34a':'#9ca3af',color:'#fff',border:0,borderRadius:'4px'}}>WhatsApp</button></td></tr>)}</tbody></table></div>}
            </div>
            
            <div style={{ display: 'flex', gap: '10px', marginBottom: '15px', backgroundColor: '#fff', padding: '15px', borderRadius: '8px', border: '1px solid #e5e7eb' }}>
              <input 
                type="text" 
                placeholder="Search Customer by Phone Number..." 
                value={customerSearchPhone}
                onChange={(e) => setCustomerSearchPhone(e.target.value)}
                style={{ flex: 1, padding: '10px', borderRadius: '6px', border: '1px solid #ccc', fontSize: '14px' }}
              />
              <input type="date" value={selectedReportDate} onChange={e=>setSelectedReportDate(e.target.value)} title="Customer Dashboard Date" style={{ padding: '9px', borderRadius: '6px', border: '1px solid #ccc' }} />
              <button onClick={() => { const d=new Date(); setSelectedReportDate(`${d.getFullYear()}-${String(d.getMonth()+1).padStart(2,'0')}-${String(d.getDate()).padStart(2,'0')}`); }} style={{ padding: '10px 15px', backgroundColor: '#2563eb', color: '#fff', border: 'none', borderRadius: '6px', cursor: 'pointer', fontWeight:'bold' }}>📅 Today</button>
              <button onClick={() => setSelectedReportDate('')} style={{ padding: '10px 15px', backgroundColor: '#0f766e', color: '#fff', border: 'none', borderRadius: '6px', cursor: 'pointer', fontWeight:'bold' }}>📚 All Dates</button>
              <button onClick={() => setCustomerSearchPhone('')} style={{ padding: '10px 15px', backgroundColor: '#6b7280', color: '#fff', border: 'none', borderRadius: '6px', cursor: 'pointer' }}>{ui.clear}</button>
              <button onClick={exportCustomerHistoryToExcel} style={{ padding: '10px 15px', backgroundColor: '#10b981', color: '#fff', border: 'none', borderRadius: '6px', cursor: 'pointer', fontWeight: 'bold' }}>📊 {ui.downloadExcel}</button>
              {salesHistory.length > 0 && <button onClick={handleDeleteAllCustomerHistory} style={{ padding: '10px 15px', backgroundColor: '#dc2626', color: '#fff', border: 'none', borderRadius: '6px', cursor: 'pointer', fontWeight: 'bold' }}>🗑️ {ui.deleteAll}</button>}
            </div>

            <table style={{ width: '100%', backgroundColor: '#fff', borderCollapse: 'collapse', borderRadius: '8px', overflow: 'hidden' }}>
              <thead>
                <tr style={{ backgroundColor: '#e5e7eb', textAlign: 'left' }}>
                  <th style={{ padding: '10px' }}>{uiMore.invoice}</th>
                  <th style={{ padding: '10px' }}>{uiMore.dateTime}</th>
                  <th style={{ padding: '10px' }}>{uiMore.customerName}</th>
                  <th style={{ padding: '10px' }}>{uiMore.phoneNumber}</th>
                  <th style={{ padding: '10px' }}>Payment</th>
                  <th style={{ padding: '10px' }}>{uiMore.totalAmount}</th>
                  <th style={{ padding: '10px' }}>Actions</th>
                </tr>
              </thead>
              <tbody>
                {pagedCustomerHistory.map(sale => (
                    <tr key={getSaleInvoiceNo(sale)} style={{ borderBottom: '1px solid #eee' }}>
                      <td style={{ padding: '10px', fontWeight: 'bold' }}>#{getSaleInvoiceNo(sale)}</td>
                      <td style={{ padding: '10px' }}>{(parseSaleDate(getSaleDateValue(sale)) || new Date()).toLocaleString()}</td>
                      <td style={{ padding: '10px', fontWeight: 'bold', color: '#2563eb' }}>{sale.customer_name || sale.customerName || 'General Customer'}</td>
                      <td style={{ padding: '10px' }}>{sale.customer_phone || sale.customerPhone || '-'}</td>
                      <td style={{ padding: '10px' }}>{sale.payment_method || sale.paymentMethod || '-'}</td>
                      <td style={{ padding: '10px', fontWeight: 'bold', color: '#059669' }}>Rs. {(parseFloat(sale.total) || 0).toFixed(2)}</td>
                      <td style={{ padding: '10px', display: 'flex', gap: '6px' }}>
                        <button onClick={() => downloadCustomerBillPDF(sale)} style={{ padding: '4px 8px', backgroundColor: '#3b82f6', color: '#fff', border: 'none', borderRadius: '4px', cursor: 'pointer' }}>📄 Bill PDF</button>
                        <button onClick={() => handlePrintCreditBookBill(sale)} style={{ padding: '4px 8px', backgroundColor: '#7c3aed', color: '#fff', border: 'none', borderRadius: '4px', cursor: 'pointer', fontWeight: 'bold' }}>🖨️ Print Bill</button>
                        <button onClick={() => handleDeleteSaleItem(getSaleInvoiceNo(sale))} title="Delete this customer purchase record" style={{ padding: '4px 8px', backgroundColor: '#ef4444', color: '#fff', border: 'none', borderRadius: '4px', cursor: 'pointer', fontWeight: 'bold' }}>🗑️ Delete</button>
                      </td>
                    </tr>
                  ))}
                {filteredCustomerHistory.length === 0 && (
                  <tr>
                    <td colSpan="7" style={{ textAlign: 'center', padding: '30px', color: '#9ca3af' }}>{uiMore.noCustomerHistory}</td>
                  </tr>
                )}
              </tbody>
            </table>
            <PaginationControls page={customerPage} pageCount={customerPageCount} setPage={setCustomerPage} totalRows={filteredCustomerHistory.length} label="Customer Records" />
          </div>
        )}

        {activeTab === 'creditBook' && currentUserRole === 'admin' && (
          <div>
            <div style={{ display:'flex', justifyContent:'space-between', alignItems:'center', flexWrap:'wrap', gap:'10px', marginBottom:'15px' }}>
              <div>
                <h2 style={{ margin:'0 0 4px' }}>📘 {ui.creditBook}</h2>
                <p style={{ margin:0, color:'#6b7280' }}>Registered Credit Customers සහ ඔවුන්ගේ outstanding credit මෙතැනින් වෙනම කළමනාකරණය කරන්න.</p>
              </div>
            </div>

            <div style={{ display:'flex', gap:'10px', marginBottom:'15px', background:'#fff', padding:'15px', borderRadius:'8px', border:'1px solid #e5e7eb', alignItems:'center', flexWrap:'wrap' }}>
              <button onClick={() => { setCustomerForm({name:'',phone:'',address:'',creditLimit:''}); setShowCustomerRegisterModal(true); }} style={{ padding:'10px 15px', background:'#2563eb', color:'#fff', border:0, borderRadius:'6px', fontWeight:'bold' }}>➕ {ui.registerCustomer}</button>
              <button onClick={handleOpenCreditPayment} style={{ padding:'10px 15px', background:'#7c3aed', color:'#fff', border:0, borderRadius:'6px', fontWeight:'bold' }}>💰 {ui.receivePayment}</button>
              <button onClick={exportRegisteredCustomersToExcel} style={{ padding:'10px 15px', background:'#10b981', color:'#fff', border:0, borderRadius:'6px', fontWeight:'bold' }}>📊 {ui.customersExcel}</button>
              <span style={{ marginLeft:'auto', fontWeight:'bold' }}>Registered Customers: {registeredCustomers.length} • Total Outstanding: Rs. {registeredCustomers.reduce((sum,c)=>sum+getCustomerBalance(c.id),0).toFixed(2)}</span>
            </div>

            <div style={{ background:'#fff', padding:'15px', borderRadius:'8px', border:'1px solid #e5e7eb', marginBottom:'15px', overflowX:'auto' }}>
              <h3 style={{ marginTop:0 }}>👥 {ui.registeredCreditCustomers}</h3>
              {registeredCustomers.length === 0 ? (
                <div style={{ padding:'25px', textAlign:'center', color:'#9ca3af' }}>{uiMore.noCreditCustomers}</div>
              ) : (
                <table style={{ width:'100%', borderCollapse:'collapse' }}><thead><tr style={{ background:'#f1f5f9', textAlign:'left' }}><th style={{padding:'9px'}}>ID</th><th style={{padding:'9px'}}>{uiMore.customer}</th><th style={{padding:'9px'}}>Phone</th><th style={{padding:'9px'}}>Credit Limit</th><th style={{padding:'9px'}}>Outstanding</th><th style={{padding:'9px'}}>Actions</th></tr></thead><tbody>
                  {registeredCustomers.map(c=><tr key={c.id} style={{borderBottom:'1px solid #eee'}}><td style={{padding:'9px'}}>{c.id}</td><td style={{padding:'9px',fontWeight:'bold'}}>{c.name}</td><td style={{padding:'9px'}}>{c.phone}</td><td style={{padding:'9px'}}>Rs. {Number(c.creditLimit||0).toFixed(2)}</td><td style={{padding:'9px',fontWeight:'bold',color:getCustomerBalance(c.id)>0?'#dc2626':'#059669'}}>Rs. {getCustomerBalance(c.id).toFixed(2)}</td><td style={{padding:'9px',display:'flex',gap:'6px'}}><button onClick={()=>{setCreditPaymentSearch(c.phone);setCreditPaymentCustomer(c);setCreditPaymentAmount('');setShowCreditPaymentModal(true);}} style={{padding:'5px 8px',background:'#7c3aed',color:'#fff',border:0,borderRadius:'4px'}}>{ui.payment}</button><button onClick={()=>handleDeleteRegisteredCustomer(c.id)} style={{padding:'5px 8px',background:'#dc2626',color:'#fff',border:0,borderRadius:'4px'}}>🗑️ Delete</button></td></tr>)}
                </tbody></table>
              )}
            </div>

            <div style={{ background:'#fff', padding:'15px', borderRadius:'8px', border:'1px solid #e5e7eb', marginBottom:'15px' }}>
              <h3 style={{ marginTop:0 }}>📄 Customer Statements / Payment History / Due Dates</h3>
              {registeredCustomers.map(c => {
                const history = customerPaymentHistory(c.id);
                return <details key={`statement-${c.id}`} style={{ borderBottom:'1px solid #e5e7eb', padding:'8px 0' }}><summary style={{ cursor:'pointer', fontWeight:'bold' }}>{c.name} ({c.phone}) — Due Rs. {getCustomerBalance(c.id).toFixed(2)} {isCustomerOverdue(c.id)?'⚠️ OVERDUE':''}</summary><div style={{ padding:'8px 0 0 15px' }}>{history.length===0?<span style={{color:'#64748b'}}>No payment/credit history.</span>:history.map(h=><div key={h.id} style={{ display:'grid', gridTemplateColumns:'120px 90px 1fr 130px', gap:'8px', padding:'4px 0', fontSize:'12px' }}><span>{h.date?new Date(h.date).toLocaleDateString():'-'}</span><span>{h.type}</span><span>Ref: {h.reference||h.billNo||'-'}</span><span>Rs. {Number(h.amount||0).toFixed(2)}{h.dueDate?` • Due ${h.dueDate}`:''}</span></div>)}</div></details>;
              })}
            </div>
          </div>
        )}

        {activeTab === 'reports' && currentUserRole === 'admin' && (
          <div>
            <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', marginBottom: '15px' }}>
              <h2>📊 {ui.reportsDashboard}</h2><div style={{ display: 'flex', gap: '6px', marginLeft: '15px' }}><button onClick={() => setReportView('overview')} style={{ padding: '7px 10px', background: reportView === 'overview' ? '#2563eb' : '#e5e7eb', color: reportView === 'overview' ? '#fff' : '#1f2937', border: 'none', borderRadius: '6px', cursor: 'pointer', fontWeight: 'bold' }}>📊 {ui.dashboard}</button><button onClick={() => setReportView('monthly')} style={{ padding: '7px 10px', background: reportView === 'monthly' ? '#7c3aed' : '#e5e7eb', color: reportView === 'monthly' ? '#fff' : '#1f2937', border: 'none', borderRadius: '6px', cursor: 'pointer', fontWeight: 'bold' }}>📅 {ui.monthlyReport}</button></div>
              <div style={{ display: 'flex', gap: '8px' }}>
                <button onClick={exportAnalyticsToExcel} style={{ padding: '8px 12px', backgroundColor: '#10b981', color: '#fff', border: 'none', borderRadius: '6px', cursor: 'pointer', fontWeight: 'bold' }}>📊 {ui.exportSalesExcel}</button>
                <button onClick={exportAnalyticsToPDF} style={{ padding: '8px 12px', backgroundColor: '#6366f1', color: '#fff', border: 'none', borderRadius: '6px', cursor: 'pointer', fontWeight: 'bold' }}>📄 {ui.exportSalesPdf}</button>
                {salesHistory.length > 0 && <button onClick={handleClearAllSales} style={{ padding: '8px 12px', backgroundColor: '#dc2626', color: '#fff', border: 'none', borderRadius: '6px', cursor: 'pointer', fontWeight: 'bold' }}>🗑️ {ui.clearAllSales}</button>}
              </div>
            </div>

            <div style={{ display: 'grid', gridTemplateColumns: 'repeat(5, minmax(120px, 1fr))', gap: '8px', marginBottom: '18px', position: 'sticky', top: 0, zIndex: 5, background: '#f8fafc', padding: '8px', borderRadius: '10px', border: '1px solid #e5e7eb' }}>
              {[['overview','📊 Dashboard'],['sales','🧾 Sales Report'],['daily','📅 Daily Sales'],['monthly','📅 Monthly Report'],['products','📦 Product Report'],['lowstock','⚠️ Low Stock'],['payments','💳 Payments'],['customers','👥 Customers'],['discounts','🏷️ Discounts'],['profitlost','📉 GIVE PROFIT LOST'],['returns','↩️ Returns'],['credit','💰 Credit'],['cashier','👨‍💼 Cashier'],['suppliers','🚚 Suppliers'],['stock','📦 Stock']].map(([key,label]) => (
                <button key={key} onClick={() => setReportView(key)} style={{ padding: '10px 8px', border: 'none', borderRadius: '7px', cursor: 'pointer', fontWeight: 'bold', background: reportView === key ? '#2563eb' : '#fff', color: reportView === key ? '#fff' : '#1f2937', boxShadow: reportView === key ? '0 2px 6px rgba(37,99,235,.25)' : 'none' }}>{label}</button>
              ))}
            </div>

            {reportView === 'overview' && <div style={{ display: 'block' }}>
            <div style={{ background:'#fff', border:'1px solid #e5e7eb', borderRadius:'10px', padding:'14px', marginBottom:'15px', display:'flex', gap:'12px', alignItems:'center', flexWrap:'wrap' }}>
              <div><b>📅 Selected Date Sales & Profit</b><div style={{fontSize:'12px',color:'#64748b'}}>ඕනෑම පරණ දවසක් තෝරලා එදාගේ Sales / Profit බලන්න. Invoice records delete වෙන්නේ නැහැ.</div></div>
              <input type="date" value={selectedReportDate} onChange={e=>setSelectedReportDate(e.target.value)} style={{padding:'9px',border:'1px solid #cbd5e1',borderRadius:'6px'}} />
              <div style={{fontWeight:'800'}}>Sales: Rs. {selectedDateSales.sales.toFixed(2)}</div>
              <div style={{fontWeight:'800',color:'#059669'}}>Profit: Rs. {selectedDateSales.profit.toFixed(2)}</div>
            </div>
            <div style={{ display: 'grid', gridTemplateColumns: 'repeat(3, 1fr)', gap: '15px', marginBottom: '25px' }}>
              <div style={{ backgroundColor: '#fff', padding: '20px', borderRadius: '8px', border: '1px solid #e5e7eb', boxShadow: '0 2px 4px rgba(0,0,0,0.05)' }}>
                <span style={{ color: '#6b7280', fontSize: '13px', fontWeight: 'bold' }}>📅 {ui.todaySalesProfit}</span>
                <h2 style={{ color: '#2563eb', margin: '8px 0 4px 0' }}>Rs. {todaySales.toFixed(2)}</h2>
                <span style={{ color: '#059669', fontSize: '13px', fontWeight: 'bold' }}>Profit: Rs. {todayProfit.toFixed(2)}</span>
              </div>
              <div style={{ backgroundColor: '#fff', padding: '20px', borderRadius: '8px', border: '1px solid #e5e7eb', boxShadow: '0 2px 4px rgba(0,0,0,0.05)' }}>
                <span style={{ color: '#6b7280', fontSize: '13px', fontWeight: 'bold' }}>📆 {ui.monthSalesProfit}</span>
                <h2 style={{ color: '#7c3aed', margin: '8px 0 4px 0' }}>Rs. {monthSales.toFixed(2)}</h2>
                <span style={{ color: '#059669', fontSize: '13px', fontWeight: 'bold' }}>Profit: Rs. {monthProfit.toFixed(2)}</span>
              </div>
              <div style={{ backgroundColor: '#fff', padding: '20px', borderRadius: '8px', border: '1px solid #e5e7eb', boxShadow: '0 2px 4px rgba(0,0,0,0.05)' }}>
                <span style={{ color: '#6b7280', fontSize: '13px', fontWeight: 'bold' }}>📈 {ui.yearSalesProfit}</span>
                <h2 style={{ color: '#059669', margin: '8px 0 4px 0' }}>Rs. {yearSales.toFixed(2)}</h2>
                <span style={{ color: '#059669', fontSize: '13px', fontWeight: 'bold' }}>Profit: Rs. {yearProfit.toFixed(2)}</span>
              </div>
            </div>
            </div>}

            {reportView === 'monthly' && <div style={{ display: 'block' }}>
            <div style={{ backgroundColor: '#fff', border: '1px solid #e5e7eb', borderRadius: '12px', padding: '18px', marginBottom: '20px' }}>
              <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', gap: '10px', flexWrap: 'wrap' }}>
                <div><h2 style={{ margin: 0 }}>📅 {ui.monthlySalesProfitAnalysis} — {selectedReportYear}</h2><p style={{ margin: '5px 0 0', color: '#6b7280', fontSize: '13px' }}>මාසෙන් මාසෙට Sales සහ Profit එක වෙනම table එකකින් බලන්න. පැරණි අවුරුදු data delete නොවේ.</p></div>
                <div style={{ display: 'flex', gap: '7px', flexWrap: 'wrap', alignItems:'center' }}>
                  <label style={{fontSize:'12px',fontWeight:'bold'}}>Year:</label>
                  <select value={selectedReportYear} onChange={e=>{setSelectedReportYear(Number(e.target.value)); setMonthlyAnalyticsData([]);}} style={{padding:'8px',border:'1px solid #cbd5e1',borderRadius:'6px'}}>
                    {Array.from(new Set([new Date().getFullYear(), ...salesHistory.map(s=>{const d=parseSaleDate(s.date); return d?.getFullYear();}).filter(Boolean)])).sort((a,b)=>b-a).map(y=><option key={y} value={y}>{y}</option>)}
                  </select>
                  <button onClick={refreshMonthlyAnalytics} style={{ padding: '8px 12px', background: '#2563eb', color: '#fff', border: 'none', borderRadius: '6px', cursor: 'pointer', fontWeight: 'bold' }}>🔄 {ui.fillRefresh}</button>
                  <button onClick={exportMonthlyAnalyticsToExcel} disabled={!monthlyAnalyticsData.length} style={{ padding: '8px 12px', background: '#10b981', color: '#fff', border: 'none', borderRadius: '6px', cursor: 'pointer', fontWeight: 'bold' }}>📊 {ui.downloadExcel}</button>
                  <button onClick={handleDeleteMonthlyAnalytics} disabled={!monthlyAnalyticsData.length} style={{ padding: '8px 12px', background: '#dc2626', color: '#fff', border: 'none', borderRadius: '6px', cursor: 'pointer', fontWeight: 'bold' }}>🗑️ Delete</button>
                </div>
              </div>
              <div style={{ marginTop: '18px', padding: '15px', border: '1px solid #e5e7eb', borderRadius: '10px', background: '#fafafa', overflowX: 'auto' }}>
                <h3 style={{ marginTop: 0 }}>📈 {ui.fullYearChart}</h3>
                <div style={{ minWidth: '760px', height: '330px', display: 'flex', alignItems: 'flex-end', gap: '10px', padding: '10px 10px 40px', borderBottom: '2px solid #94a3b8' }}>
                  {(monthlyAnalyticsData.length ? monthlyAnalyticsData : Array.from({length:12}, (_,i)=>({monthNo:i+1, month:new Date(selectedReportYear,i,1).toLocaleString(undefined,{month:'short'}), sales:0, profit:0}))).map(r => {
                    const max = Math.max(...(monthlyAnalyticsData.length ? monthlyAnalyticsData : [{sales:0,profit:0}]).map(x => Math.max(Number(x.sales)||0, Number(x.profit)||0)), 1);
                    return <div key={r.monthNo} style={{ flex: 1, minWidth: '45px', height: '100%', display: 'flex', alignItems: 'flex-end', justifyContent: 'center', gap: '3px', position: 'relative' }}>
                      <div title={`${r.month}: Sales Rs. ${Number(r.sales||0).toFixed(2)}`} style={{ width: '18px', height: `${Math.max(3,(Number(r.sales)||0)/max*260)}px`, background: '#3b82f6', borderRadius: '4px 4px 0 0' }} />
                      <div title={`${r.month}: Profit Rs. ${Number(r.profit||0).toFixed(2)}`} style={{ width: '18px', height: `${Math.max(3,(Number(r.profit)||0)/max*260)}px`, background: '#10b981', borderRadius: '4px 4px 0 0' }} />
                      <div style={{ position: 'absolute', bottom: '-28px', fontSize: '10px', fontWeight: 'bold', transform: 'rotate(-35deg)', whiteSpace: 'nowrap' }}>{String(r.month).slice(0,3)}</div>
                    </div>;
                  })}
                </div>
                <div style={{ marginTop: '35px', fontSize: '12px', color: '#475569', fontWeight: 'bold' }}>🔵 Sales &nbsp;&nbsp; 🟢 Profit — January to December</div>
              </div>
              {monthlyAnalyticsData.length ? <div style={{ overflowX: 'auto', marginTop: '15px' }}><table style={{ width: '100%', borderCollapse: 'collapse' }}><thead><tr style={{ background: '#0f172a', color: '#fff' }}><th style={{ padding: '9px', textAlign: 'left' }}>{ui.month}</th><th style={{ padding: '9px' }}>{ui.sales}</th><th style={{ padding: '9px' }}>{ui.profit}</th><th style={{ padding: '9px' }}>{ui.profitMargin}</th></tr></thead><tbody>{monthlyAnalyticsData.map(r => <tr key={`${r.year}-${r.monthNo}`} style={{ borderBottom: '1px solid #e5e7eb' }}><td style={{ padding: '9px', fontWeight: 'bold' }}>{r.month}</td><td style={{ padding: '9px', textAlign: 'right' }}>Rs. {Number(r.sales || 0).toFixed(2)}</td><td style={{ padding: '9px', textAlign: 'right', color: '#059669', fontWeight: 'bold' }}>Rs. {Number(r.profit || 0).toFixed(2)}</td><td style={{ padding: '9px', textAlign: 'right' }}>{Number(r.sales) > 0 ? ((Number(r.profit) / Number(r.sales)) * 100).toFixed(2) : '0.00'}%</td></tr>)}<tr style={{ background: '#f8fafc', fontWeight: 'bold' }}><td style={{ padding: '9px' }}>{ui.yearTotal}</td><td style={{ padding: '9px', textAlign: 'right' }}>Rs. {monthlyAnalyticsData.reduce((a,b)=>a+Number(b.sales||0),0).toFixed(2)}</td><td style={{ padding: '9px', textAlign: 'right', color: '#059669' }}>Rs. {monthlyAnalyticsData.reduce((a,b)=>a+Number(b.profit||0),0).toFixed(2)}</td><td style={{ padding: '9px', textAlign: 'right' }}>{(() => { const s=monthlyAnalyticsData.reduce((a,b)=>a+Number(b.sales||0),0); const p=monthlyAnalyticsData.reduce((a,b)=>a+Number(b.profit||0),0); return s ? ((p/s)*100).toFixed(2) : '0.00'; })()}%</td></tr></tbody></table></div> : <div style={{ marginTop: '15px', padding: '22px', textAlign: 'center', background: '#f8fafc', borderRadius: '8px', color: '#6b7280' }}>🔄 Fill / Refresh button එක click කරලා monthly analysis table එක පුරවන්න.</div>}
            </div>
            </div>}

            {reportView === 'profitlost' && <div style={{ display: 'block' }}>
              <div style={{ backgroundColor: '#fff', border: '1px solid #e5e7eb', borderRadius: '12px', padding: '18px', marginBottom: '20px' }}>
                <h2 style={{ marginTop: 0 }}>📉 GIVE PROFIT LOST</h2>
                <p style={{ color: '#6b7280', fontSize: '13px' }}>Item Offers සහ Whole Bill Discount නිසා අඩු වූ profit amount එක මෙතනින් වෙනම බලන්න. Discount නැති bills වල normal profit calculation එක වෙනස් නොවේ.</p>
                <div style={{ display: 'grid', gridTemplateColumns: 'repeat(3, 1fr)', gap: '12px' }}>
                  <div style={{ padding: '15px', borderRadius: '8px', background: '#fff7ed' }}><b>Item Offers Profit Lost</b><div style={{ marginTop: '6px', fontSize: '24px', fontWeight: '900' }}>Rs. {profitLostReport.reduce((sum,row)=>sum+Number(row['Item Offers Lost']||0),0).toFixed(2)}</div></div>
                  <div style={{ padding: '15px', borderRadius: '8px', background: '#fef2f2' }}><b>Whole Bill Discount Profit Lost</b><div style={{ marginTop: '6px', fontSize: '24px', fontWeight: '900' }}>Rs. {profitLostReport.reduce((sum,row)=>sum+Number(row['Whole Bill Discount Lost']||0),0).toFixed(2)}</div></div>
                  <div style={{ padding: '15px', borderRadius: '8px', background: '#f8fafc' }}><b>Total Profit Lost</b><div style={{ marginTop: '6px', fontSize: '24px', fontWeight: '900' }}>Rs. {profitLostReport.reduce((sum,row)=>sum+Number(row['Total Profit Lost']||0),0).toFixed(2)}</div></div>
                </div>
              </div>
            </div>}

            {reportView === 'products' && <div style={{ display: 'block' }}>
            {/* PRODUCT PERFORMANCE CHARTS - Charts and separate tables are both kept */}
            <div style={{ backgroundColor: '#fff', border: '1px solid #e5e7eb', borderRadius: '12px', padding: '18px', marginBottom: '20px' }}>
              <h2 style={{ marginTop: 0 }}>📊 {ui.productPerformance}</h2>
              <p style={{ color: '#6b7280', fontSize: '13px' }}>These charts use the same live sales data as the reports below. The tables are also kept separately.</p>
              <div style={{ display: 'grid', gridTemplateColumns: 'repeat(2, minmax(0, 1fr))', gap: '20px' }}>
                <div style={{ border: '1px solid #d1fae5', borderRadius: '10px', padding: '15px', minHeight: '420px', overflowX: 'auto' }}>
                  <h3 style={{ marginTop: 0 }}>🏆 {ui.bestSellingChart}</h3>
                  {bestSellingProducts.length > 0 ? <div style={{ minWidth: '620px', height: '340px', display: 'flex', alignItems: 'flex-end', gap: '8px', padding: '15px 10px 35px', borderBottom: '2px solid #9ca3af', position: 'relative' }}>
                    {bestSellingProducts.map((item, index) => {
                      const maxSold = Math.max(...bestSellingProducts.map(x => Number(x.totalSold) || 0), 1);
                      const barHeight = Math.max(12, ((Number(item.totalSold) || 0) / maxSold) * 260);
                      return <div key={`best-chart-${item.key}`} title={`${item.name}: ${item.totalSold} ${item.unit} (${item.salesPercentage.toFixed(2)}%)`} style={{ flex: 1, minWidth: '24px', height: '100%', display: 'flex', flexDirection: 'column', justifyContent: 'flex-end', alignItems: 'center' }}>
                        <div style={{ fontSize: '10px', fontWeight: 'bold', marginBottom: '5px' }}>{item.totalSold}</div>
                        <div style={{ width: '100%', maxWidth: '42px', height: `${barHeight}px`, backgroundColor: '#10b981', borderRadius: '5px 5px 0 0', cursor: 'pointer' }} />
                        <div style={{ fontSize: '9px', width: '70px', marginTop: '7px', textAlign: 'center', transform: 'rotate(-45deg)', transformOrigin: 'top center', whiteSpace: 'nowrap', overflow: 'hidden', textOverflow: 'ellipsis' }}>{item.name}</div>
                      </div>;
                    })}
                  </div> : <div style={{ padding: '40px', textAlign: 'center', color: '#6b7280' }}>No sales data available for chart.</div>}
                </div>
                <div style={{ border: '1px solid #fed7aa', borderRadius: '10px', padding: '15px', minHeight: '420px', overflowX: 'auto' }}>
                  <h3 style={{ marginTop: 0 }}>📉 {ui.slowMovingChart}</h3>
                  {slowMovingProducts.length > 0 ? <div style={{ minWidth: '620px', height: '340px', display: 'flex', alignItems: 'flex-end', gap: '8px', padding: '15px 10px 35px', borderBottom: '2px solid #9ca3af' }}>
                    {slowMovingProducts.map((item, index) => {
                      const maxSold = Math.max(...slowMovingProducts.map(x => Number(x.totalSold) || 0), 1);
                      const barHeight = Math.max(12, ((Number(item.totalSold) || 0) / maxSold) * 260);
                      return <div key={`slow-chart-${item.key}`} title={`${item.name}: ${item.totalSold} ${item.unit} (${item.salesPercentage.toFixed(2)}%)`} style={{ flex: 1, minWidth: '24px', height: '100%', display: 'flex', flexDirection: 'column', justifyContent: 'flex-end', alignItems: 'center' }}>
                        <div style={{ fontSize: '10px', fontWeight: 'bold', marginBottom: '5px' }}>{item.totalSold}</div>
                        <div style={{ width: '100%', maxWidth: '42px', height: `${barHeight}px`, backgroundColor: '#f59e0b', borderRadius: '5px 5px 0 0', cursor: 'pointer' }} />
                        <div style={{ fontSize: '9px', width: '70px', marginTop: '7px', textAlign: 'center', transform: 'rotate(-45deg)', transformOrigin: 'top center', whiteSpace: 'nowrap', overflow: 'hidden', textOverflow: 'ellipsis' }}>{item.name}</div>
                      </div>;
                    })}
                  </div> : <div style={{ padding: '40px', textAlign: 'center', color: '#6b7280' }}>No sales data available for chart.</div>}
                </div>
              </div>
            </div>

            <div style={{ display: 'grid', gridTemplateColumns: 'repeat(2, minmax(0, 1fr))', gap: '20px', marginBottom: '25px' }}>
              <div style={{ backgroundColor: '#fff', border: '1px solid #e5e7eb', borderRadius: '10px', padding: '15px', overflowX: 'auto' }}>
                <h3 style={{ marginTop: 0 }}>🏆 {ui.bestSelling}</h3>
                <p style={{ marginTop: 0, color: '#6b7280', fontSize: '12px' }}>Percentage = item sold quantity ÷ total sold quantity × 100</p>
                <table style={{ width: '100%', borderCollapse: 'collapse', fontSize: '13px' }}>
                  <thead><tr style={{ backgroundColor: '#ecfdf5' }}><th style={{ padding: '8px' }}>#</th><th style={{ padding: '8px', textAlign: 'left' }}>Product</th><th style={{ padding: '8px' }}>Sold</th><th style={{ padding: '8px' }}>%</th></tr></thead>
                  <tbody>{bestSellingProducts.map((item, index) => <tr key={item.key} style={{ borderBottom: '1px solid #eee' }}><td style={{ padding: '8px', textAlign: 'center' }}>{index + 1}</td><td style={{ padding: '8px', fontWeight: 'bold' }}>{item.name}<div style={{ color: '#6b7280', fontSize: '11px' }}>{item.barcode}</div></td><td style={{ padding: '8px', textAlign: 'center' }}>{item.totalSold} {item.unit}</td><td style={{ padding: '8px', textAlign: 'center', fontWeight: 'bold', color: '#059669' }}>{item.salesPercentage.toFixed(2)}%</td></tr>)}{bestSellingProducts.length === 0 && <tr><td colSpan="4" style={{ padding: '18px', textAlign: 'center', color: '#6b7280' }}>No sold items yet.</td></tr>}</tbody>
                </table>
              </div>
              <div style={{ backgroundColor: '#fff', border: '1px solid #e5e7eb', borderRadius: '10px', padding: '15px', overflowX: 'auto' }}>
                <h3 style={{ marginTop: 0 }}>📉 {ui.slowMoving}</h3>
                <p style={{ marginTop: 0, color: '#6b7280', fontSize: '12px' }}>Sold items only, sorted from lowest sold quantity.</p>
                <table style={{ width: '100%', borderCollapse: 'collapse', fontSize: '13px' }}>
                  <thead><tr style={{ backgroundColor: '#fff7ed' }}><th style={{ padding: '8px' }}>#</th><th style={{ padding: '8px', textAlign: 'left' }}>Product</th><th style={{ padding: '8px' }}>Sold</th><th style={{ padding: '8px' }}>%</th></tr></thead>
                  <tbody>{slowMovingProducts.map((item, index) => <tr key={item.key} style={{ borderBottom: '1px solid #eee' }}><td style={{ padding: '8px', textAlign: 'center' }}>{index + 1}</td><td style={{ padding: '8px', fontWeight: 'bold' }}>{item.name}<div style={{ color: '#6b7280', fontSize: '11px' }}>{item.barcode}</div></td><td style={{ padding: '8px', textAlign: 'center' }}>{item.totalSold} {item.unit}</td><td style={{ padding: '8px', textAlign: 'center', fontWeight: 'bold', color: '#d97706' }}>{item.salesPercentage.toFixed(2)}%</td></tr>)}{slowMovingProducts.length === 0 && <tr><td colSpan="4" style={{ padding: '18px', textAlign: 'center', color: '#6b7280' }}>No sold items yet.</td></tr>}</tbody>
                </table>
              </div>
            </div>
            <div style={{ backgroundColor: '#fff', border: '1px solid #e5e7eb', borderRadius: '10px', padding: '15px', marginBottom: '25px', overflowX: 'auto' }}>
              <h3 style={{ marginTop: 0 }}>🚫 {ui.neverSold} ({neverSoldProducts.length})</h3>
              <p style={{ color: '#6b7280', fontSize: '12px' }}>Products in the inventory that have not been sold yet.</p>
              <div style={{ display: 'flex', flexWrap: 'wrap', gap: '8px' }}>{neverSoldProducts.length > 0 ? neverSoldProducts.map(item => <span key={item.key} style={{ backgroundColor: '#fef2f2', color: '#991b1b', padding: '7px 10px', borderRadius: '6px', fontSize: '13px' }}>{item.name} ({item.barcode})</span>) : <span style={{ color: '#059669' }}>All inventory products have been sold at least once.</span>}</div>
            </div>
            </div>}

            {reportView === 'lowstock' && <div style={{ display: 'block' }}>
            <h3 style={{ marginTop: '20px' }}>⚠️ {ui.lowStock}</h3>
            <div style={{ display: 'flex', justifyContent: 'flex-end', marginBottom: '10px' }}>
              <div style={{ display: 'flex', gap: '8px' }}>
                <button onClick={handleDownloadLowStockExcel} style={{ padding: '6px 12px', backgroundColor: '#10b981', color: '#fff', border: 'none', borderRadius: '6px', cursor: 'pointer', fontWeight: 'bold', fontSize: '13px' }}>📊 {ui.downloadLowStock}</button>
                {(reportsData.lowStock || []).length > 0 && <button onClick={handleDeleteAllLowStockItems} style={{ padding: '6px 12px', backgroundColor: '#dc2626', color: '#fff', border: 'none', borderRadius: '6px', cursor: 'pointer', fontWeight: 'bold', fontSize: '13px' }}>🗑️ {ui.deleteAll}</button>}
              </div>
            </div>
            
            <table style={{ width: '100%', backgroundColor: '#fff', borderCollapse: 'collapse', marginBottom: '25px', borderRadius: '8px', overflow: 'hidden' }}>
              <thead>
                <tr style={{ backgroundColor: '#ffeeec', textAlign: 'left', color: '#b91c1c' }}>
                  <th style={{ padding: '10px' }}>Barcode</th>
                  <th style={{ padding: '10px' }}>{ui.productName}</th>
                  <th style={{ padding: '10px' }}>{ui.remainingStock}</th>
                  <th style={{ padding: '10px' }}>{ui.alertLimit}</th>
                  <th style={{ padding: '10px' }}>{ui.statusAction}</th><th style={{ padding: '10px' }}>Delete</th>
                </tr>
              </thead>
              <tbody>
                {pagedLowStock.map(item => {
                  const currentStockVal = editedQuantities[item.barcode] !== undefined ? editedQuantities[item.barcode] : item.stock;
                  return (
                    <tr key={item.id} style={{ borderBottom: '1px solid #eee' }}>
                      <td style={{ padding: '10px', fontFamily: 'monospace' }}>{item.barcode}</td>
                      <td style={{ padding: '10px', fontWeight: 'bold' }}>{item.name}</td>
                      <td style={{ padding: '10px' }}>
                        {isEditing[item.barcode] ? (
                          <input 
                            type="number" 
                            value={currentStockVal} 
                            onChange={(e) => handleQuantityChange(item.barcode, e.target.value)}
                            style={{ width: '70px', padding: '4px', border: '1px solid #3b82f6', borderRadius: '4px' }} 
                          />
                        ) : (
                          <span style={{ fontWeight: 'bold', color: '#dc2626' }}>{currentStockVal} {item.unit}</span>
                        )}
                      </td>
                      <td style={{ padding: '10px' }}>{item.min_stock_alert || 5} {item.unit}</td>
                      <td style={{ padding: '10px' }}>
                        <button onClick={() => toggleEdit(item.barcode, item.stock)} style={{ padding: '4px 8px', backgroundColor: isEditing[item.barcode] ? '#059669' : '#3b82f6', color: '#fff', border: 'none', borderRadius: '4px', cursor: 'pointer', fontSize: '12px' }}>
                          {isEditing[item.barcode] ? 'Save Qty' : 'Edit Qty'}
                        </button>
                      </td>
                      <td style={{ padding: '10px' }}><button onClick={() => handleDeleteLowStockItem(item)} style={{ padding: '5px 9px', backgroundColor: '#ef4444', color: '#fff', border: 'none', borderRadius: '4px', cursor: 'pointer', fontSize: '12px' }}>🗑️ Delete</button></td>
                    </tr>
                  );
                })}
                {(!reportsData.lowStock || reportsData.lowStock.length === 0) && (
                  <tr>
                    <td colSpan="6" style={{ textAlign: 'center', padding: '20px', color: '#059669', fontWeight: 'bold' }}>All stock levels are optimal! No low stock items.</td>
                  </tr>
                )}
              </tbody>
            </table>
            <PaginationControls page={lowStockPage} pageCount={lowStockPageCount} setPage={setLowStockPage} totalRows={lowStockList.length} label="Low Stock Items" />
            </div>}

            {reportView === 'sales' && <div style={{ display: 'block' }}>
              <div style={{ backgroundColor: '#fff', border: '1px solid #e5e7eb', borderRadius: '10px', padding: '18px', marginBottom: '18px' }}>
                <h2 style={{ marginTop: 0 }}>🧾 Sales Report</h2>
                <p style={{ marginTop: 0, color: '#6b7280' }}>Billing එකෙන් සාර්ථකව save වන sales data එක මෙතනින් එකින් එක බලන්න. Existing sales history/report data එකම භාවිතා කරයි.</p>
                <div style={{ display: 'grid', gridTemplateColumns: 'repeat(3, 1fr)', gap: '12px', marginBottom: '18px' }}>
                  <div style={{ padding: '14px', background: '#eff6ff', borderRadius: '8px' }}><b>Total Bills</b><div style={{ fontSize: '24px', fontWeight: '900', marginTop: '5px' }}>{salesHistory.length}</div></div>
                  <div style={{ padding: '14px', background: '#f0fdf4', borderRadius: '8px' }}><b>Total Sales</b><div style={{ fontSize: '24px', fontWeight: '900', marginTop: '5px' }}>Rs. {salesHistory.reduce((sum,s) => sum + Number(s.total || s.grandTotal || 0), 0).toFixed(2)}</div></div>
                  <div style={{ padding: '14px', background: '#faf5ff', borderRadius: '8px' }}><b>Current Month</b><div style={{ fontSize: '24px', fontWeight: '900', marginTop: '5px' }}>Rs. {monthSales.toFixed(2)}</div></div>
                </div>
                <div style={{ overflowX: 'auto' }}>
                  <table style={{ width: '100%', borderCollapse: 'collapse', fontSize: '13px' }}>
                    <thead><tr style={{ background: '#0f172a', color: '#fff' }}><th style={{ padding: '9px', textAlign: 'left' }}>{ui.invoiceNo}</th><th style={{ padding: '9px', textAlign: 'left' }}>{ui.dateTime}</th><th style={{ padding: '9px', textAlign: 'left' }}>{ui.customerName}</th><th style={{ padding: '9px', textAlign: 'right' }}>{ui.totalAmount}</th></tr></thead>
                    <tbody>{salesHistory.slice().reverse().map((sale, index) => <tr key={sale.invoiceNo ?? sale.invoice_no ?? sale.id ?? index} style={{ borderBottom: '1px solid #e5e7eb' }}><td style={{ padding: '9px' }}>#{sale.invoiceNo ?? sale.invoice_no ?? sale.id ?? '-'}</td><td style={{ padding: '9px' }}>{sale.date || sale.dateFormatted || sale.created_at || '-'}</td><td style={{ padding: '9px' }}>{sale.customerName || sale.customer_name || 'General Customer'}</td><td style={{ padding: '9px', textAlign: 'right', fontWeight: 'bold' }}>Rs. {Number(sale.total || sale.grandTotal || 0).toFixed(2)}</td></tr>)}{salesHistory.length === 0 && <tr><td colSpan="4" style={{ padding: '22px', textAlign: 'center', color: '#6b7280' }}>No sales recorded yet.</td></tr>}</tbody>
                  </table>
                </div>
              </div>
            </div>}


            {/* Advanced live report controls/charts. Existing report pages above remain unchanged. */}
            <div style={{marginTop:'18px',padding:'14px',background:'#f8fafc',border:'1px solid #e5e7eb',borderRadius:'10px'}}>
              <div style={{display:'flex',justifyContent:'space-between',alignItems:'center',gap:'8px',flexWrap:'wrap'}}>
                <h3 style={{margin:0}}>📈 Live Charts — {reportPageTitle[reportView] || 'Report'}</h3>
                <div style={{display:'flex',gap:'7px',flexWrap:'wrap'}}>
                  <button onClick={()=>refreshReportPage(reportView)} style={{padding:'8px 12px',background:'#2563eb',color:'#fff',border:0,borderRadius:'6px',fontWeight:'bold'}}>🔄 Refresh</button>
                  <button onClick={()=>downloadReportExcel(`${reportPageTitle[reportView]||'Report'}_Report`,reportPageRows(reportView))} style={{padding:'8px 12px',background:'#10b981',color:'#fff',border:0,borderRadius:'6px',fontWeight:'bold'}}>📊 Download Excel</button>
                  <button onClick={()=>markReportDeleted(reportView)} style={{padding:'8px 12px',background:'#dc2626',color:'#fff',border:0,borderRadius:'6px',fontWeight:'bold'}}>🗑️ Delete Report</button>
                </div>
              </div>
              {!reportDeletedFlags[reportView] && <div style={{display:'grid',gridTemplateColumns:'repeat(3,minmax(0,1fr))',gap:'15px',marginTop:'14px'}}>
                <BarChart data={reportChartData(reportView)} title={`${reportPageTitle[reportView]||'Report'} - Bar Chart`}/>
                <LineChart data={reportChartData(reportView)} title={`${reportPageTitle[reportView]||'Report'} - Line Chart`}/>
                <PieChart data={reportChartData(reportView)} title={`${reportPageTitle[reportView]||'Report'} - Pie Chart`}/>
              </div>}
            </div>
            {['daily','payments','customers','discounts','profitlost','returns','credit','cashier','suppliers','stock'].includes(reportView) && !reportDeletedFlags[reportView] && <div style={{marginTop:'15px',background:'#fff',border:'1px solid #e5e7eb',borderRadius:'10px',padding:'14px',overflowX:'auto'}}><h3 style={{marginTop:0}}>📋 {reportPageTitle[reportView]} Data</h3><table style={{width:'100%',borderCollapse:'collapse'}}><thead><tr>{Object.keys(reportPageRows(reportView)[0]||{Report:'No data'}).map(k=><th key={k} style={{padding:'9px',background:'#f1f5f9',textAlign:'left'}}>{k}</th>)}</tr></thead><tbody>{reportPageRows(reportView).slice(0,200).map((r,i)=><tr key={i}>{Object.keys(r).map(k=><td key={k} style={{padding:'9px',borderBottom:'1px solid #e5e7eb'}}>{typeof r[k]==='number'?Number(r[k]).toFixed(2):String(r[k]??'')}</td>)}</tr>)}</tbody></table></div>}

          </div>
        )}

        {activeTab === 'settings' && currentUserRole === 'admin' && (
          <div>
            <h2>⚙️ {ui.shopProfile}</h2>

            {/* Dashboard navigation customization — existing Shop Profile options remain unchanged. */}
            <div style={{ backgroundColor: '#fff', padding: '20px', borderRadius: '8px', marginBottom: '20px', border: '1px solid #e5e7eb', boxShadow: '0 2px 5px rgba(0,0,0,0.05)' }}>
              <h3 style={{ marginTop: 0, color: '#1f2937' }}>🧭 Dashboard Style / Dashboard පෙනුම</h3>
              <p style={{ fontSize: '13px', color: '#64748b', marginTop: 0 }}>Login වූ පසු navigation එක ඔබට කැමති විදිහට තෝරන්න. දැනට ඇති Billing, Inventory, Reports, Product Table සහ අනෙකුත් options කිසිවක් ඉවත් නොවේ.</p>
              <div style={{ display: 'grid', gridTemplateColumns: 'repeat(4, minmax(0, 1fr))', gap: '10px' }}>
                {[
                  { value: 'home', icon: '🏠', title: 'Home Dashboard', text: 'විශාල icons/cards සහිත Home Page' },
                  { value: 'sidebar', icon: '📚', title: 'Sidebar', text: 'Sidebar navigation එක පමණක්' },
                  { value: 'both', icon: '🔄', title: 'Both', text: 'Home Dashboard + Sidebar දෙකම' },
                  { value: 'classic', icon: '🖥️', title: 'Classic / Original', text: 'Home Page add කිරීමට කලින් තිබුණු පරණ collapsible dashboard layout එක' }
                ].map(option => (
                  <button key={option.value} type="button" onClick={() => setDashboardStyle(option.value)} style={{
                    padding: '14px', textAlign: 'left', borderRadius: '9px', cursor: 'pointer',
                    border: dashboardStyle === option.value ? '2px solid #2563eb' : '1px solid #d1d5db',
                    background: dashboardStyle === option.value ? '#eff6ff' : '#fff'
                  }}>
                    <div style={{ fontSize: '22px' }}>{option.icon}</div>
                    <div style={{ fontWeight: '900', marginTop: '5px', color: '#111827' }}>{option.title}</div>
                    <div style={{ fontSize: '11px', color: '#64748b', marginTop: '3px' }}>{option.text}</div>
                  </button>
                ))}
              </div>
            </div>
            
            {/* 💾 DATA BACKUP & RESTORE SECTION ADDED HERE */}
            <div style={{ backgroundColor: '#fff', padding: '20px', borderRadius: '8px', marginBottom: '20px', border: '1px solid #e5e7eb', boxShadow: '0 2px 5px rgba(0,0,0,0.05)' }}>
              <h3 style={{ marginTop: 0, color: '#1f2937' }}>💾 {ui.databaseBackup}</h3>
              <p style={{ fontSize: '13px', color: '#6b7280', margin: '0 0 15px 0' }}>
                ඔබගේ කඩයේ සියලුම දත්ත (Products, Sales History, Suppliers, Settings) ආරක්ෂිතව Backup එකක් ලෙස බාගත කරගන්න හෝ කලින් සාදන ලද Backup එකක් නැවත ප්‍රතිස්ථාපනය (Restore) කරන්න.
              </p>
              <div style={{ display: 'flex', gap: '15px', alignItems: 'center' }}>
                <button onClick={handleDownloadBackup} style={{ padding: '10px 20px', backgroundColor: '#059669', color: '#fff', border: 'none', borderRadius: '6px', fontWeight: 'bold', cursor: 'pointer', display: 'flex', alignItems: 'center', gap: '6px' }}>
                  📥 Download Backup (JSON)
                </button>
                <label style={{ padding: '10px 20px', backgroundColor: '#d97706', color: '#fff', borderRadius: '6px', cursor: 'pointer', fontWeight: 'bold', display: 'flex', alignItems: 'center', gap: '6px' }}>
                  📤 Restore Backup
                  <input type="file" accept=".json" onChange={handleRestoreBackup} style={{ display: 'none' }} />
                </label>
              </div>
            </div>

            <div style={{ display: 'grid', gridTemplateColumns: 'repeat(2, 1fr)', gap: '20px' }}>
              
              <div style={{ backgroundColor: '#fff', padding: '20px', borderRadius: '8px', border: '1px solid #e5e7eb' }}>
                <h3 style={{ marginTop: 0 }}>{ui.shopProfileDetails}</h3>
                <div style={{ display: 'flex', flexDirection: 'column', gap: '12px' }}>
                  <div>
                    <label style={{ display: 'block', fontSize: '12px', fontWeight: 'bold', color: '#374151', marginBottom: '4px' }}>Shop Name:</label>
                    <input type="text" value={receiptSettings.shopName} onChange={(e) => setReceiptSettings({ ...receiptSettings, shopName: e.target.value })} style={{ width: '100%', padding: '8px', borderRadius: '6px', border: '1px solid #ccc' }} />
                  </div>
                  <div>
                    <label style={{ display: 'block', fontSize: '12px', fontWeight: 'bold', color: '#374151', marginBottom: '4px' }}>Address:</label>
                    <textarea value={receiptSettings.address} onChange={(e) => setReceiptSettings({ ...receiptSettings, address: e.target.value })} style={{ width: '100%', padding: '8px', borderRadius: '6px', border: '1px solid #ccc', height: '60px' }} />
                  </div>
                  <div>
                    <label style={{ display: 'block', fontSize: '12px', fontWeight: 'bold', color: '#374151', marginBottom: '4px' }}>Phone Number:</label>
                    <input type="text" value={receiptSettings.phone} onChange={(e) => setReceiptSettings({ ...receiptSettings, phone: e.target.value })} style={{ width: '100%', padding: '8px', borderRadius: '6px', border: '1px solid #ccc' }} />
                  </div>
                  <div>
                    <label style={{ display: 'block', fontSize: '12px', fontWeight: 'bold', color: '#374151', marginBottom: '4px' }}>Email:</label>
                    <input type="email" value={receiptSettings.email} onChange={(e) => setReceiptSettings({ ...receiptSettings, email: e.target.value })} style={{ width: '100%', padding: '8px', borderRadius: '6px', border: '1px solid #ccc' }} />
                  </div>
                  <div>
                    <label style={{ display: 'block', fontSize: '12px', fontWeight: 'bold', color: '#374151', marginBottom: '4px' }}>Shop Logo (Upload):</label>
                    <input type="file" accept="image/*" onChange={handleLogoUpload} style={{ width: '100%', padding: '6px' }} />
                    {receiptSettings.logoUrl && (
                      <div style={{ marginTop: '8px' }}>
                        <img src={receiptSettings.logoUrl} alt="Logo Preview" style={{ width: '70px', height: '70px', objectFit: 'contain', borderRadius: '0', border: '1px solid #d1d5db', background: '#fff' }} />
                      </div>
                    )}
                    <div style={{ marginTop: '10px' }}>
                      <label style={{ display: 'block', fontSize: '12px', fontWeight: 'bold', color: '#374151', marginBottom: '4px' }}>{ui.receiptLogoSize}</label>
                      <select value={receiptSettings.logoSize || '40mm'} onChange={(e) => setReceiptSettings({ ...receiptSettings, logoSize: e.target.value })} style={{ width: '100%', padding: '8px', borderRadius: '6px', border: '1px solid #ccc' }}>
                        <option value="20mm">Small — 2 × 2 cm</option>
                        <option value="30mm">Medium — 3 × 3 cm</option>
                        <option value="40mm">Large — 4 × 4 cm (Recommended)</option>
                        <option value="50mm">Extra Large — 5 × 5 cm</option>
                        <option value="60mm">Maximum — 6 × 6 cm</option>
                      </select>
                      <div style={{ marginTop: '5px', fontSize: '11px', color: '#64748b', lineHeight: 1.4 }}>
                        Receipt Preview සහ Print දෙකේම Logo එක මේ ප්‍රමාණයට square ලෙස print වේ.
                      </div>
                    </div>
                  </div>
                </div>
              </div>

              <div style={{ display: 'flex', flexDirection: 'column', gap: '20px' }}>
                <div style={{ backgroundColor: '#fff', padding: '20px', borderRadius: '8px', border: '1px solid #e5e7eb' }}>
                  <h3 style={{ marginTop: 0 }}>{ui.receiptSettings}</h3>
                  <div style={{ display: 'flex', flexDirection: 'column', gap: '12px' }}>
                    <div>
                      <label style={{ display: 'block', fontSize: '12px', fontWeight: 'bold', color: '#374151', marginBottom: '4px' }}>{ui.terminal}</label>
                      <input type="text" value={receiptSettings.terminal} onChange={(e) => setReceiptSettings({ ...receiptSettings, terminal: e.target.value })} style={{ width: '100%', padding: '8px', borderRadius: '6px', border: '1px solid #ccc' }} />
                    </div>
                    <div>
                      <label style={{ display: 'block', fontSize: '12px', fontWeight: 'bold', color: '#374151', marginBottom: '4px' }}>{ui.footer}</label>
                      <input type="text" value={receiptSettings.footerMsg} onChange={(e) => setReceiptSettings({ ...receiptSettings, footerMsg: e.target.value })} style={{ width: '100%', padding: '8px', borderRadius: '6px', border: '1px solid #ccc' }} />
                      <div style={{ marginTop: '10px' }}>
                        <label style={{ display: 'block', fontSize: '12px', fontWeight: 'bold', color: '#374151', marginBottom: '4px' }}>
                          {lang === 'si' ? 'බිල්පතේ අවසාන Comment / පණිවිඩය:' : lang === 'ta' ? 'ரசீது இறுதி Comment / செய்தி:' : 'Receipt Final Comment / Message:'}
                        </label>
                        <textarea
                          value={receiptSettings.footerComment || ''}
                          onChange={(e) => setReceiptSettings({ ...receiptSettings, footerComment: e.target.value })}
                          placeholder={lang === 'si' ? 'උදා: දින 7ක් ඇතුළත නැවත පැමිණෙන්න...' : lang === 'ta' ? 'உதா: 7 நாட்களுக்குள் மீண்டும் வரவும்...' : 'Example: Come back within 7 days...'}
                          rows={3}
                          style={{ width: '100%', padding: '8px', borderRadius: '6px', border: '1px solid #ccc', resize: 'vertical', boxSizing: 'border-box' }}
                        />
                      </div>
                    </div>
                    <div style={{ background: '#f8fafc', padding: '12px', borderRadius: '8px', border: '1px solid #e2e8f0' }}>
                      <label style={{ display: 'flex', alignItems: 'center', gap: '8px', fontSize: '13px', fontWeight: 900, marginBottom: '8px' }}>
                        <input
                          type="checkbox"
                          checked={receiptSettings.poweredByEnabled === true}
                          onChange={(e) => setReceiptSettings({ ...receiptSettings, poweredByEnabled: e.target.checked })}
                        />
                        {lang === 'si' ? 'Receipt අවසානයේ කුඩා Powered by නම පෙන්වන්න' : lang === 'ta' ? 'ரசீதின் முடிவில் சிறிய Powered by பெயரை காட்டவும்' : 'Show small Powered by text at receipt bottom'}
                      </label>
                      <input
                        type="text"
                        value={receiptSettings.poweredByText || 'Powered by Positha POS System'}
                        onChange={(e) => setReceiptSettings({ ...receiptSettings, poweredByText: e.target.value })}
                        placeholder="Powered by Positha POS System"
                        disabled={receiptSettings.poweredByEnabled !== true}
                        style={{ width: '100%', padding: '8px', borderRadius: '6px', border: '1px solid #ccc', boxSizing: 'border-box' }}
                      />
                    </div>
                    <div style={{ display: 'grid', gridTemplateColumns: 'repeat(2, minmax(0, 1fr))', gap: '10px' }}>
                      <div>
                        <label style={{ display: 'block', fontSize: '12px', fontWeight: 'bold', color: '#374151', marginBottom: '4px' }}>{ui.paperWidth}</label>
                        <select value={receiptSettings.paperWidth} onChange={(e) => setReceiptSettings({ ...receiptSettings, paperWidth: e.target.value })} style={{ width: '100%', padding: '8px', borderRadius: '6px', border: '1px solid #ccc' }}>
                          <option value="80mm">80mm Thermal</option>
                          <option value="58mm">58mm Thermal</option>
                        </select>
                      </div>
                      <div>
                        <label style={{ display: 'block', fontSize: '12px', fontWeight: 'bold', color: '#374151', marginBottom: '4px' }}>{ui.fontSize}</label>
                        <select value={receiptSettings.fontSize} onChange={(e) => setReceiptSettings({ ...receiptSettings, fontSize: e.target.value })} style={{ width: '100%', padding: '8px', borderRadius: '6px', border: '1px solid #ccc' }}>
                          <option value="12px">Compact (12px)</option>
                          <option value="14px">Normal (14px)</option>
                          <option value="16px">Large (16px)</option>
                          <option value="18px">Extra Large (18px)</option>
                        </select>
                      </div>
                      <div>
                        <label style={{ display: 'block', fontSize: '12px', fontWeight: 'bold', color: '#374151', marginBottom: '4px' }}>{ui.receiptFont}</label>
                        <select value={receiptSettings.fontFamily} onChange={(e) => setReceiptSettings({ ...receiptSettings, fontFamily: e.target.value })} style={{ width: '100%', padding: '8px', borderRadius: '6px', border: '1px solid #ccc' }}>
                          <option value="'Courier New', Courier, monospace">Courier New</option>
                          <option value="Arial, Helvetica, sans-serif">Arial</option>
                          <option value="Verdana, Geneva, sans-serif">Verdana</option>
                          <option value="Tahoma, Arial, sans-serif">Tahoma</option>
                          <option value="'Trebuchet MS', Arial, sans-serif">Trebuchet MS</option>
                          <option value="Georgia, 'Times New Roman', serif">Georgia</option>
                        </select>
                      </div>
                      <div>
                        <label style={{ display: 'block', fontSize: '12px', fontWeight: 'bold', color: '#374151', marginBottom: '4px' }}>{ui.receiptDesign}</label>
                        <select value={receiptSettings.receiptDesign || 'classic'} onChange={(e) => setReceiptSettings({ ...receiptSettings, receiptDesign: e.target.value })} style={{ width: '100%', padding: '8px', borderRadius: '6px', border: '1px solid #ccc' }}>
                          <option value="classic">1. Classic POS</option>
                          <option value="modern">2. Modern Clean</option>
                          <option value="bold">3. Bold Total</option>
                          <option value="compact">4. Compact Thermal</option>
                          <option value="elegant">5. Elegant</option>
                        </select>
                      </div>
                    </div>
                    <div style={{ background: '#f8fafc', padding: '14px', borderRadius: '8px', border: '1px solid #e2e8f0' }}>
                      <h4 style={{ margin: '0 0 10px' }}>🖨️ Receipt Text Style — එක් එක් කොටස වෙන වෙනම</h4>
                      <div style={{ display: 'grid', gridTemplateColumns: 'repeat(2, minmax(0, 1fr))', gap: '10px' }}>
                        {[
                          ['shopName', 'Shop Name / ප්‍රධාන නම'],
                          ['address', 'Address / ලිපිනය'],
                          ['phone', 'Phone / දුරකථන අංකය'],
                          ['email', 'Email'],
                          ['meta', 'Bill Details / බිල් විස්තර'],
                          ['items', 'Items / භාණ්ඩ'],
                          ['grandTotal', 'Grand Total / මුළු එකතුව'],
                          ['footer', 'Thank You / අවසාන පණිවිඩය'],
                          ['footerComment', 'Final Comment / අවසාන පණිවිඩය'],
                          ['poweredBy', 'Powered By / අවසාන කුඩා නම'],
                          ['invoice', 'Invoice / අවසන් Invoice']
                        ].map(([key, label]) => {
                          const st = receiptSettings.elementStyles?.[key] || {};
                          const update = (patch) => setReceiptSettings(prev => ({ ...prev, elementStyles: { ...prev.elementStyles, [key]: { ...(prev.elementStyles?.[key] || {}), ...patch } } }));
                          return (
                            <div key={key} style={{ background: '#fff', border: '1px solid #dbe3ef', borderRadius: '7px', padding: '10px' }}>
                              <div style={{ fontWeight: 900, marginBottom: '7px' }}>{label}</div>
                              <div style={{ display: 'grid', gridTemplateColumns: '1fr 1fr', gap: '7px' }}>
                                <select value={st.fontSize || '14px'} onChange={e => update({ fontSize: e.target.value })} style={{ width: '100%', padding: '7px' }}>
                                  {['10px','12px','14px','16px','18px','20px','22px','24px','26px','28px','30px'].map(v => <option key={v} value={v}>{v}</option>)}
                                </select>
                                <select value={String(st.fontWeight ?? 900)} onChange={e => update({ fontWeight: Number(e.target.value) })} style={{ width: '100%', padding: '7px' }}>
                                  <option value="400">Normal</option><option value="500">Medium</option><option value="600">Semi Bold</option><option value="700">Bold</option><option value="800">Extra Bold</option><option value="900">Black</option>
                                </select>
                                <select value={st.textAlign || 'left'} onChange={e => update({ textAlign: e.target.value })} style={{ width: '100%', padding: '7px' }}>
                                  <option value="left">Left</option><option value="center">Center</option><option value="right">Right</option>
                                </select>
                                <input type="color" value={st.color || '#000000'} onChange={e => update({ color: e.target.value })} title="Text color" style={{ width: '100%', height: '34px', padding: '2px' }} />
                              </div>
                              <div style={{ fontSize: '11px', color: '#64748b', marginTop: '6px' }}>Size • Bold • Align • Color</div>
                            </div>
                          );
                        })}
                      </div>
                    </div>

                    <div style={{ padding: '10px 12px', background: '#f8fafc', border: '1px solid #e2e8f0', borderRadius: '8px', fontSize: '12px', color: '#475569', lineHeight: 1.5 }}>
                      <b>Print-safe:</b> 58mm and 80mm widths are handled separately. Long receipts can grow vertically, so 50–60+ items continue onto the same thermal roll without a fixed page height.
                    </div>
                  </div>
                </div>

                <div style={{ backgroundColor: '#fff', padding: '20px', borderRadius: '8px', border: '1px solid #e5e7eb', gridColumn: '1 / -1' }}>
                  <div style={{ display:'flex', justifyContent:'space-between', gap:'12px', alignItems:'center', flexWrap:'wrap' }}>
                    <div>
                      <h3 style={{ margin:'0 0 6px' }}>
                        📜 {lang === 'si' ? 'කඩයේ කොන්දේසි සහ සේවා නියමයන්' : lang === 'ta' ? 'கடை நிபந்தனைகள் மற்றும் சேவை விதிமுறைகள்' : 'Shop Conditions & Service Terms'}
                      </h3>
                      <div style={{ fontSize:'12px', color:'#64748b' }}>
                        {lang === 'si' ? 'තෝරා ඇති භාෂාවෙන් බලන්න සහ Excel ලෙස බාගන්න.' :
                         lang === 'ta' ? 'தேர்ந்தெடுத்த மொழியில் பார்த்து Excel ஆக பதிவிறக்கவும்.' :
                         'View the conditions in the selected language and download them as Excel.'}
                      </div>
                    </div>
                    <button onClick={downloadShopConditionsExcel} style={{ padding:'10px 15px', background:'#059669', color:'#fff', border:0, borderRadius:'7px', fontWeight:'800', cursor:'pointer' }}>
                      📥 {lang === 'si' ? 'කොන්දේසි Excel බාගන්න' : lang === 'ta' ? 'நிபந்தனைகள் Excel பதிவிறக்கம்' : 'Download Conditions Excel'}
                    </button>
                  </div>
                  <div style={{ marginTop:'14px', background:'#f8fafc', border:'1px solid #e2e8f0', borderRadius:'9px', padding:'14px', lineHeight:1.7, maxHeight:'360px', overflowY:'auto' }}>
                    {shopConditions[lang].map((condition, index) => (
                      <div key={index} style={{ padding:'6px 0', borderBottom:index < shopConditions[lang].length - 1 ? '1px solid #e5e7eb' : 'none', fontSize:'13px', fontWeight:600 }}>
                        {condition}
                      </div>
                    ))}
                  </div>
                </div>

                <div style={{ backgroundColor: '#fff', padding: '20px', borderRadius: '8px', border: '1px solid #e5e7eb' }}>
                  <h3 style={{ marginTop: 0 }}>{ui.cashierAccounts}</h3>
                  <div style={{ display: 'flex', flexDirection: 'column', gap: '12px' }}>
                    <div>
                      <label style={{ display: 'block', fontSize: '12px', fontWeight: 'bold', color: '#374151', marginBottom: '4px' }}>{ui.selectCashier}</label>
                      <select 
                        value={settingsActiveCashierId} 
                        onChange={(e) => {
                          const cId = parseInt(e.target.value);
                          setSettingsActiveCashierId(cId);
                          const found = cashiers.find(c => c.id === cId);
                          if (found) {
                            setSettingsCashierUsername(found.username);
                            setSettingsCashierPassword(found.password);
                          }
                        }}
                        style={{ width: '100%', padding: '8px', borderRadius: '6px', border: '1px solid #ccc' }}
                      >
                        {cashiers.map(c => <option key={c.id} value={c.id}>Cashier {c.id} ({c.username})</option>)}
                      </select>
                    </div>
                    <div>
                      <label style={{ display: 'block', fontSize: '12px', fontWeight: 'bold', color: '#374151', marginBottom: '4px' }}>Username:</label>
                      <input type="text" value={settingsCashierUsername} onChange={(e) => setSettingsCashierUsername(e.target.value)} style={{ width: '100%', padding: '8px', borderRadius: '6px', border: '1px solid #ccc' }} />
                    </div>
                    <div>
                      <label style={{ display: 'block', fontSize: '12px', fontWeight: 'bold', color: '#374151', marginBottom: '4px' }}>Password:</label>
                      <input type="text" value={settingsCashierPassword} onChange={(e) => setSettingsCashierPassword(e.target.value)} style={{ width: '100%', padding: '8px', borderRadius: '6px', border: '1px solid #ccc' }} />
                    </div>
                    <button 
                      onClick={() => {
                        const updated = cashiers.map(c => c.id === settingsActiveCashierId ? { ...c, username: settingsCashierUsername, password: settingsCashierPassword } : c);
                        setCashiers(updated);
                        showNotification('අයකැමි ගිණුම සාර්ථකව යාවත්කාලීන කරන ලදී!');
                      }}
                      style={{ padding: '10px', backgroundColor: '#2563eb', color: '#fff', border: 'none', borderRadius: '6px', fontWeight: 'bold', cursor: 'pointer' }}
                    >
                      Update Cashier Credentials
                    </button>
                  </div>
                </div>
              </div>

            </div>
          </div>
        )}

      </div>
    </div>
  );
}


export default function AppWithErrorBoundary() {
  return (
    <AppErrorBoundary>
      <App />
    </AppErrorBoundary>
  )
}
