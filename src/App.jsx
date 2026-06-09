import { useState, useEffect } from "react";
import * as XLSX from "xlsx";

const PLATFORM_COMMISSIONS = {
  smarttender: { label: "SmartTender", rate: 0.01 },
  prozorro: { label: "Prozorro", rate: 0.005 },
  other: { label: "Інша", rate: 0 },
};

const VEHICLE_TYPES = {
  truck: { label: "Фура", icon: "🚛", consumption: 35 },
  gazelle: { label: "ГАЗель", icon: "🚐", consumption: 15 },
  truck_small: { label: "Вантажівка (середня)", icon: "🚚", consumption: 22 },
};

const FUEL_PRICE_DEFAULT = 52;

function fmt(val) {
  if (val === "" || val === null || val === undefined || isNaN(Number(val))) return 0;
  return Number(val);
}

function formatUAH(val) {
  if (!val && val !== 0) return "—";
  return Number(val).toLocaleString("uk-UA", { minimumFractionDigits: 2, maximumFractionDigits: 2 }) + " ₴";
}

function NumInput({ label, value, onChange, hint, prefix, step = "any", suffix }) {
  return (
    <div style={{ marginBottom: 12 }}>
      <label style={{ display: "block", fontSize: 12, color: "#8a9bb8", marginBottom: 4, fontWeight: 600, letterSpacing: "0.04em", textTransform: "uppercase" }}>{label}</label>
      <div style={{ position: "relative", display: "flex", alignItems: "center" }}>
        {prefix && <span style={{ position: "absolute", left: 12, color: "#4e6a9a", fontWeight: 700, fontSize: 14, pointerEvents: "none", zIndex: 1 }}>{prefix}</span>}
        <input type="number" min={0} step={step} value={value} onChange={e => onChange(e.target.value)}
          style={{ width: "100%", padding: prefix ? "10px 12px 10px 28px" : suffix ? "10px 44px 10px 12px" : "10px 12px", background: "#1e2d45", border: "1.5px solid #2a3f5f", borderRadius: 8, color: "#e8f0ff", fontSize: 15, outline: "none", boxSizing: "border-box" }} />
        {suffix && <span style={{ position: "absolute", right: 12, color: "#4e6a9a", fontWeight: 600, fontSize: 13, pointerEvents: "none" }}>{suffix}</span>}
      </div>
      {hint && <div style={{ fontSize: 11, color: "#4e6a9a", marginTop: 3 }}>{hint}</div>}
    </div>
  );
}

function Section({ title, children, color = "#2a7fff" }) {
  return (
    <div style={{ background: "#131e30", borderRadius: 12, padding: "18px 20px", marginBottom: 16, border: "1px solid #1e2d45" }}>
      <div style={{ fontSize: 13, fontWeight: 700, color, letterSpacing: "0.06em", textTransform: "uppercase", marginBottom: 14, display: "flex", alignItems: "center", gap: 8 }}>
        <span style={{ display: "inline-block", width: 3, height: 16, background: color, borderRadius: 2 }} />
        {title}
      </div>
      {children}
    </div>
  );
}

function ResultRow({ label, value, color, bold = false, sub = false }) {
  return (
    <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", padding: sub ? "4px 0 4px 12px" : "7px 0" }}>
      <span style={{ fontSize: sub ? 12 : 13, color: bold ? "#c9d8f5" : sub ? "#4e6a9a" : "#7a90b8", fontWeight: bold ? 700 : 400 }}>{label}</span>
      <span style={{ fontSize: bold ? 16 : sub ? 12 : 13, color: color || (bold ? "#e8f0ff" : "#c9d8f5"), fontWeight: bold ? 700 : 500 }}>{value}</span>
    </div>
  );
}

export default function App() {
  const [tenderName, setTenderName] = useState("");
  const [tenderAmount, setTenderAmount] = useState("");
  const [platform, setPlatform] = useState("smarttender");
  const [customCommission, setCustomCommission] = useState("");
  const [goods, setGoods] = useState([{ name: "", qty: "", unit: "кг", buyPrice: "" }]);
  const [drivers, setDrivers] = useState([{ salary: "", trips: "" }]);
  const [fuelEnabled, setFuelEnabled] = useState(true);
  const [routeKm, setRouteKm] = useState("");
  const [roundTrip, setRoundTrip] = useState(true);
  const [vehicleType, setVehicleType] = useState("truck");
  const [customConsumption, setCustomConsumption] = useState("");
  const [fuelPrice, setFuelPrice] = useState(FUEL_PRICE_DEFAULT);
  const [extraCosts, setExtraCosts] = useState([{ name: "", amount: "" }]);
  const [taxType, setTaxType] = useState("fop2");
  const [fop2Monthly, setFop2Monthly] = useState("4500");
  const [fop3Rate, setFop3Rate] = useState("6");
  const [customTaxRate, setCustomTaxRate] = useState("");

  // Історія тендерів
  const [history, setHistory] = useState([]);
  const [addedMsg, setAddedMsg] = useState(false);

  // Завантажити Google Identity Services
  useEffect(() => {
    if (!document.getElementById("google-gsi")) {
      const s = document.createElement("script");
      s.id = "google-gsi";
      s.src = "https://accounts.google.com/gsi/client";
      s.async = true;
      document.head.appendChild(s);
    }
  }, []);

  // Google Sheets
  const GOOGLE_CLIENT_ID = "982995692115-ouvkd2f7sem4umc86m3dr6svkk49lmnp.apps.googleusercontent.com";
  const SCOPES = "https://www.googleapis.com/auth/spreadsheets https://www.googleapis.com/auth/drive.file";
  const [sheetId, setSheetId] = useState(() => localStorage.getItem("tender_sheet_id") || "");
  const [sheetMsg, setSheetMsg] = useState("");
  const [sheetLoading, setSheetLoading] = useState(false);
  const [showSheetInput, setShowSheetInput] = useState(false);

  const getGoogleToken = () => new Promise((resolve, reject) => {
    if (!window.google) { reject(new Error("Google API не завантажено")); return; }
    const client = window.google.accounts.oauth2.initTokenClient({
      client_id: GOOGLE_CLIENT_ID,
      scope: SCOPES,
      callback: (resp) => resp.error ? reject(resp) : resolve(resp.access_token),
    });
    client.requestAccessToken();
  });

  const saveToSheets = async () => {
    if (history.length === 0) { setSheetMsg("❗ Спочатку додай тендер до таблиці"); setTimeout(()=>setSheetMsg(""),3000); return; }
    setSheetLoading(true);
    setSheetMsg("Авторизація...");
    try {
      const token = await getGoogleToken();
      const headers = ["Дата","Назва тендера","Товар","Сума тендера","Собівартість","Комісія","Зарплата водіїв","Паливо","Інші витрати","Загальні витрати","Прибуток до податку","Податок","Тип ФОП","Чистий заробіток","Маржа %"];
      const rows = history.map(r => [r.date,r.name,r.goods,r.amount,r.goodsTotal,r.commissionAmt,r.driverTotal,r.fuelTotal,r.extraTotal,r.totalCosts,r.grossProfit,r.taxAmt,r.taxType,r.netProfit,parseFloat(r.margin.toFixed(2))]);

      let sid = sheetId;

      if (!sid) {
        // Створити нову таблицю
        setSheetMsg("Створення таблиці...");
        const createRes = await fetch("https://sheets.googleapis.com/v4/spreadsheets", {
          method: "POST",
          headers: { Authorization: `Bearer ${token}`, "Content-Type": "application/json" },
          body: JSON.stringify({ properties: { title: "Аналітика тендерів" }, sheets: [{ properties: { title: "Тендери" } }] })
        });
        const created = await createRes.json();
        sid = created.spreadsheetId;
        setSheetId(sid);
        localStorage.setItem("tender_sheet_id", sid);

        // Записати шапку
        await fetch(`https://sheets.googleapis.com/v4/spreadsheets/${sid}/values/Тендери!A1:O1?valueInputOption=USER_ENTERED`, {
          method: "PUT",
          headers: { Authorization: `Bearer ${token}`, "Content-Type": "application/json" },
          body: JSON.stringify({ values: [headers] })
        });
      }

      // Знайти перший порожній рядок
      setSheetMsg("Запис даних...");
      const getRes = await fetch(`https://sheets.googleapis.com/v4/spreadsheets/${sid}/values/Тендери!A:A`, {
        headers: { Authorization: `Bearer ${token}` }
      });
      const getData = await getRes.json();
      const nextRow = (getData.values ? getData.values.length : 0) + 1;

      // Дописати рядки
      await fetch(`https://sheets.googleapis.com/v4/spreadsheets/${sid}/values/Тендери!A${nextRow}:O${nextRow + rows.length - 1}?valueInputOption=USER_ENTERED`, {
        method: "PUT",
        headers: { Authorization: `Bearer ${token}`, "Content-Type": "application/json" },
        body: JSON.stringify({ values: rows })
      });

      setSheetMsg(`✅ Збережено! Тендерів: ${rows.length}`);
      setTimeout(() => setSheetMsg(""), 4000);
    } catch(e) {
      console.error(e);
      setSheetMsg("❗ Помилка. Спробуй ще раз.");
      setTimeout(() => setSheetMsg(""), 4000);
    }
    setSheetLoading(false);
  };

  const openSheet = () => {
    if (sheetId) window.open(`https://docs.google.com/spreadsheets/d/${sheetId}`, "_blank");
  };

  // ---- Розрахунки ----
  const totalTrips = drivers.reduce((sum, d) => sum + (parseInt(d.trips) || 0), 0) || 1;
  const kmPerTrip = fmt(routeKm);
  const kmTotal = kmPerTrip * (roundTrip ? 2 : 1) * totalTrips;
  const consumption = fmt(customConsumption) || VEHICLE_TYPES[vehicleType].consumption;
  const litersTotal = (kmTotal * consumption) / 100;
  const fuelTotal = fuelEnabled ? litersTotal * fmt(fuelPrice) : 0;

  const amount = fmt(tenderAmount);
  const commissionRate = platform === "other" ? fmt(customCommission) / 100 : PLATFORM_COMMISSIONS[platform].rate;
  const commissionAmt = amount * commissionRate;
  const goodsTotal = goods.reduce((sum, g) => sum + fmt(g.qty) * fmt(g.buyPrice), 0);
  const driverTotal = drivers.reduce((sum, d) => sum + fmt(d.salary) * (parseInt(d.trips) || 1), 0);
  const extraTotal = extraCosts.reduce((sum, e) => sum + fmt(e.amount), 0);
  const totalCosts = goodsTotal + commissionAmt + driverTotal + fuelTotal + extraTotal;
  const grossProfit = amount - totalCosts;

  let taxAmt = 0, taxLabel = "";
  if (taxType === "fop2") {
    taxAmt = fmt(fop2Monthly);
    taxLabel = `ФОП 2 гр. — фіксовано ${taxAmt.toLocaleString("uk-UA")} ₴/міс`;
  } else if (taxType === "fop3") {
    const rate = fmt(fop3Rate) || 6;
    taxAmt = amount * (rate / 100);
    taxLabel = `ФОП 3 гр. — ${rate}% від обороту`;
  } else {
    const rate = fmt(customTaxRate);
    taxAmt = grossProfit > 0 ? grossProfit * (rate / 100) : 0;
    taxLabel = `Податок — ${rate}% від прибутку`;
  }

  const netProfit = grossProfit - taxAmt;
  const margin = amount > 0 ? (netProfit / amount) * 100 : 0;
  const profitColor = netProfit >= 0 ? "#3ddc84" : "#ff5a5a";
  const marginColor = margin >= 15 ? "#3ddc84" : margin >= 5 ? "#ffc947" : "#ff5a5a";

  // Helpers
  const addGood = () => setGoods(g => [...g, { name: "", qty: "", unit: "кг", buyPrice: "" }]);
  const removeGood = i => setGoods(g => g.filter((_, idx) => idx !== i));
  const updateGood = (i, k, v) => setGoods(g => g.map((item, idx) => idx === i ? { ...item, [k]: v } : item));
  const addDriver = () => setDrivers(d => [...d, { salary: "", trips: "" }]);
  const removeDriver = i => setDrivers(d => d.filter((_, idx) => idx !== i));
  const updateDriver = (i, k, v) => setDrivers(d => d.map((item, idx) => idx === i ? { ...item, [k]: v } : item));
  const addExtra = () => setExtraCosts(e => [...e, { name: "", amount: "" }]);
  const removeExtra = i => setExtraCosts(e => e.filter((_, idx) => idx !== i));
  const updateExtra = (i, k, v) => setExtraCosts(e => e.map((item, idx) => idx === i ? { ...item, [k]: v } : item));

  const reset = () => {
    setTenderName(""); setTenderAmount(""); setPlatform("smarttender"); setCustomCommission("");
    setGoods([{ name: "", qty: "", unit: "кг", buyPrice: "" }]);
    setDrivers([{ salary: "", trips: "" }]);
    setRouteKm(""); setRoundTrip(true); setVehicleType("truck"); setCustomConsumption(""); setFuelPrice(FUEL_PRICE_DEFAULT);
    setExtraCosts([{ name: "", amount: "" }]);
    setTaxType("fop2"); setFop2Monthly("4500"); setFop3Rate("6"); setCustomTaxRate("");
  };

  // ---- Додати тендер в історію ----
  const addToHistory = () => {
    if (!amount) return;
    const goodsStr = goods.filter(g => g.name).map(g => `${g.name} ${g.qty}${g.unit}`).join(", ");
    const entry = {
      date: new Date().toLocaleDateString("uk-UA"),
      name: tenderName || "Без назви",
      goods: goodsStr,
      amount,
      goodsTotal,
      commissionAmt,
      driverTotal,
      fuelTotal: fuelEnabled ? fuelTotal : 0,
      extraTotal,
      totalCosts,
      grossProfit,
      taxAmt,
      netProfit,
      margin,
      taxType: taxType === "fop2" ? "ФОП 2" : taxType === "fop3" ? "ФОП 3" : "Інше",
    };
    setHistory(h => [...h, entry]);
    setAddedMsg(true);
    setTimeout(() => setAddedMsg(false), 2500);
  };

  // ---- Скачати Excel (зі стилями) ----
  const downloadExcel = () => {
    if (history.length === 0) return;

    const headers = [
      "Дата", "Назва тендера", "Товар", "Сума тендера ₴", "Собівартість товару ₴",
      "Комісія платформи ₴", "Зарплата водіїв ₴", "Паливо ₴", "Інші витрати ₴",
      "Загальні витрати ₴", "Прибуток до податку ₴", "Податок ₴", "Тип ФОП",
      "Чистий заробіток ₴", "Маржа %"
    ];

    // Числові колонки (індекси 3-11, 13-14) — грошовий формат
    const moneyFmt = '#,##0.00\\ "₴"';
    const pctFmt   = '0.00"%"';

    const rows = history.map(r => [
      r.date, r.name, r.goods,
      r.amount, r.goodsTotal, r.commissionAmt, r.driverTotal,
      r.fuelTotal, r.extraTotal, r.totalCosts, r.grossProfit,
      r.taxAmt, r.taxType, r.netProfit,
      parseFloat(r.margin.toFixed(2))
    ]);

    const sumRow = [
      "РАЗОМ", "", "",
      history.reduce((s,r)=>s+r.amount,0),
      history.reduce((s,r)=>s+r.goodsTotal,0),
      history.reduce((s,r)=>s+r.commissionAmt,0),
      history.reduce((s,r)=>s+r.driverTotal,0),
      history.reduce((s,r)=>s+r.fuelTotal,0),
      history.reduce((s,r)=>s+r.extraTotal,0),
      history.reduce((s,r)=>s+r.totalCosts,0),
      history.reduce((s,r)=>s+r.grossProfit,0),
      history.reduce((s,r)=>s+r.taxAmt,0),
      "",
      history.reduce((s,r)=>s+r.netProfit,0),
      parseFloat((history.reduce((s,r)=>s+r.margin,0)/history.length).toFixed(2))
    ];

    const totalRows = 1 + rows.length + 1 + 1; // header + data + empty + sum
    const sumRowIdx = rows.length + 2; // 0-based: header=0, data=1..N, empty=N+1, sum=N+2

    const ws = XLSX.utils.aoa_to_sheet([headers, ...rows, [], sumRow]);

    // ---- Ширина колонок ----
    ws["!cols"] = [
      {wch:12},{wch:32},{wch:36},{wch:18},{wch:20},{wch:20},{wch:18},
      {wch:15},{wch:15},{wch:18},{wch:20},{wch:15},{wch:10},{wch:20},{wch:10}
    ];

    // ---- Висота шапки ----
    ws["!rows"] = [{ hpt: 22 }];

    // Стилі через cell-by-cell
    const numCols  = [3,4,5,6,7,8,9,10,11,13]; // грошові
    const pctCol   = 14;
    const totalCols = headers.length;

    const colLetter = n => {
      let s = "";
      n++;
      while (n > 0) { s = String.fromCharCode(65+(n-1)%26) + s; n = Math.floor((n-1)/26); }
      return s;
    };

    // --- Шапка (рядок 1) ---
    for (let c = 0; c < totalCols; c++) {
      const addr = colLetter(c) + "1";
      if (!ws[addr]) ws[addr] = { v: headers[c], t: "s" };
      ws[addr].s = {
        fill: { fgColor: { rgb: "1F3864" } },
        font: { bold: true, color: { rgb: "FFFFFF" }, sz: 11 },
        alignment: { horizontal: "center", vertical: "center", wrapText: true },
        border: {
          bottom: { style: "medium", color: { rgb: "4472C4" } },
          right:  { style: "thin",   color: { rgb: "2E4A7A" } },
        }
      };
    }

    // --- Рядки даних ---
    rows.forEach((row, ri) => {
      const excelRow = ri + 2; // 1-based, row 1 = header
      const isEven = ri % 2 === 1;
      for (let c = 0; c < totalCols; c++) {
        const addr = colLetter(c) + excelRow;
        if (!ws[addr]) continue;
        const isNum = numCols.includes(c);
        const isPct = c === pctCol;
        const isProfitCol = c === 13; // Чистий заробіток
        const marginVal = row[14];

        let fontColor = "333333";
        if (isProfitCol) fontColor = row[13] >= 0 ? "1A6B3C" : "B71C1C";
        if (isPct) fontColor = marginVal >= 15 ? "1A6B3C" : marginVal >= 5 ? "7B5800" : "B71C1C";

        ws[addr].s = {
          fill: { fgColor: { rgb: isEven ? "EEF2F8" : "FFFFFF" } },
          font: { color: { rgb: fontColor }, sz: 10, bold: isProfitCol },
          numFmt: isNum ? moneyFmt : isPct ? pctFmt : undefined,
          alignment: { horizontal: isNum || isPct ? "right" : "left", vertical: "center" },
          border: {
            bottom: { style: "thin", color: { rgb: "D0D7E3" } },
            right:  { style: "thin", color: { rgb: "D0D7E3" } },
          }
        };
        if (isNum) ws[addr].z = moneyFmt;
        if (isPct) ws[addr].z = pctFmt;
      }
    });

    // --- Підсумковий рядок РАЗОМ ---
    const sri = sumRowIdx + 1; // 1-based excel row
    for (let c = 0; c < totalCols; c++) {
      const addr = colLetter(c) + sri;
      if (!ws[addr]) continue;
      const isNum = numCols.includes(c);
      const isPct = c === pctCol;
      ws[addr].s = {
        fill: { fgColor: { rgb: "1F3864" } },
        font: { bold: true, color: { rgb: "FFFFFF" }, sz: 11 },
        numFmt: isNum ? moneyFmt : isPct ? pctFmt : undefined,
        alignment: { horizontal: isNum || isPct ? "right" : "left", vertical: "center" },
        border: {
          top:    { style: "medium", color: { rgb: "4472C4" } },
          bottom: { style: "medium", color: { rgb: "4472C4" } },
          right:  { style: "thin",   color: { rgb: "2E4A7A" } },
        }
      };
      if (isNum) ws[addr].z = moneyFmt;
      if (isPct) ws[addr].z = pctFmt;
    }

    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "Тендери");
    const now = new Date();
    const fname = `тендери_${now.getFullYear()}-${String(now.getMonth()+1).padStart(2,"0")}-${String(now.getDate()).padStart(2,"0")}.xlsx`;
    XLSX.writeFile(wb, fname);
  };

  const inp = { padding: "9px 10px", background: "#1e2d45", border: "1.5px solid #2a3f5f", borderRadius: 8, color: "#e8f0ff", fontSize: 13, outline: "none", boxSizing: "border-box", width: "100%" };
  const lbl = { fontSize: 11, color: "#8a9bb8", marginBottom: 4, fontWeight: 600, letterSpacing: "0.04em", textTransform: "uppercase", display: "block" };

  return (
    <div style={{ fontFamily: "'Inter','Segoe UI',sans-serif", background: "#0b1525", minHeight: "100vh", color: "#c9d8f5", paddingBottom: 40 }}>
      {/* Header */}
      <div style={{ background: "linear-gradient(135deg,#0f1e35 0%,#0b2a4a 100%)", padding: "28px 24px 22px", borderBottom: "1px solid #1a2e4a", marginBottom: 20 }}>
        <div style={{ maxWidth: 720, margin: "0 auto", display: "flex", justifyContent: "space-between", alignItems: "flex-start", flexWrap: "wrap", gap: 12 }}>
          <div>
            <div style={{ fontSize: 11, letterSpacing: "0.15em", color: "#4e7ab5", fontWeight: 700, textTransform: "uppercase", marginBottom: 6 }}>Prozorro · SmartTender</div>
            <h1 style={{ margin: 0, fontSize: 26, fontWeight: 800, color: "#e8f0ff", letterSpacing: "-0.02em" }}>Аналітика тендерів</h1>
            <div style={{ fontSize: 13, color: "#5a7aaa", marginTop: 6 }}>Фрукти та овочі — розрахунок реального заробітку</div>
          </div>
          {history.length > 0 && (
            <div style={{ display: "flex", alignItems: "center", gap: 10 }}>
              <div style={{ background: "#0d2a55", borderRadius: 10, padding: "8px 14px", textAlign: "center" }}>
                <div style={{ fontSize: 11, color: "#4e7ab5" }}>ТЕНДЕРІВ</div>
                <div style={{ fontSize: 20, fontWeight: 800, color: "#4a8aff" }}>{history.length}</div>
              </div>
              <div style={{ background: "#0a2e1e", borderRadius: 10, padding: "8px 14px", textAlign: "center" }}>
                <div style={{ fontSize: 11, color: "#4e6a9a" }}>ЗАГАЛЬНИЙ ДОХІД</div>
                <div style={{ fontSize: 16, fontWeight: 800, color: "#3ddc84" }}>{formatUAH(history.reduce((s,r)=>s+r.netProfit,0))}</div>
              </div>
            </div>
          )}
        </div>
      </div>

      <div style={{ maxWidth: 720, margin: "0 auto", padding: "0 16px" }}>

        {/* ТЕНДЕР */}
        <Section title="Тендер" color="#4a8aff">
          <div style={{ display: "flex", gap: 12, flexWrap: "wrap" }}>
            <div style={{ flex: "2 1 200px" }}>
              <label style={lbl}>Назва тендера</label>
              <input value={tenderName} onChange={e => setTenderName(e.target.value)} placeholder="напр. Постачання яблук — школа №5" style={{ ...inp, fontSize: 14 }} />
            </div>
            <div style={{ flex: "1 1 140px" }}>
              <label style={lbl}>Сума тендера (₴)</label>
              <input type="number" min={0} value={tenderAmount} onChange={e => setTenderAmount(e.target.value)} placeholder="0" style={{ ...inp, fontSize: 14 }} />
            </div>
          </div>
          <div style={{ marginTop: 12 }}>
            <label style={lbl}>Платформа</label>
            <div style={{ display: "flex", gap: 8, flexWrap: "wrap" }}>
              {Object.entries(PLATFORM_COMMISSIONS).map(([key, { label }]) => (
                <button key={key} onClick={() => setPlatform(key)} style={{ padding: "8px 16px", borderRadius: 8, border: "1.5px solid", borderColor: platform === key ? "#4a8aff" : "#2a3f5f", background: platform === key ? "#0d2a55" : "#1e2d45", color: platform === key ? "#4a8aff" : "#7a90b8", fontWeight: platform === key ? 700 : 400, cursor: "pointer", fontSize: 13 }}>
                  {label}{key !== "other" && <span style={{ marginLeft: 5, fontSize: 11, opacity: 0.7 }}>{(PLATFORM_COMMISSIONS[key].rate*100).toFixed(1)}%</span>}
                </button>
              ))}
            </div>
            {platform === "other" && <div style={{ marginTop: 10, maxWidth: 200 }}><NumInput label="Комісія (%)" value={customCommission} onChange={setCustomCommission} /></div>}
            {amount > 0 && <div style={{ marginTop: 8, fontSize: 12, color: "#4e6a9a" }}>Комісія: {(commissionRate*100).toFixed(2)}% = <span style={{ color: "#ffc947" }}>{formatUAH(commissionAmt)}</span></div>}
          </div>
        </Section>

        {/* ТОВАР */}
        <Section title="Товар — Собівартість закупівлі" color="#38bdf8">
          <div style={{ display: "flex", gap: 8, marginBottom: 6 }}>
            <div style={{ flex: 3 }}><span style={lbl}>Назва товару</span></div>
            <div style={{ flex: 2 }}><span style={lbl}>Кількість</span></div>
            <div style={{ flex: 1 }}><span style={lbl}>Од.</span></div>
            <div style={{ flex: 2 }}><span style={lbl}>Закуп. ціна ₴</span></div>
            <div style={{ flex: 2 }}><span style={lbl}>Сума ₴</span></div>
            <div style={{ width: 36 }} />
          </div>
          {goods.map((g, i) => {
            const rowTotal = fmt(g.qty) * fmt(g.buyPrice);
            return (
              <div key={i} style={{ display: "flex", gap: 8, marginBottom: 8, alignItems: "center" }}>
                <div style={{ flex: 3 }}><input value={g.name} onChange={e => updateGood(i,"name",e.target.value)} placeholder="Яблука, морква…" style={inp} /></div>
                <div style={{ flex: 2 }}><input type="number" min={0} step="any" value={g.qty} onChange={e => updateGood(i,"qty",e.target.value)} placeholder="0" style={inp} /></div>
                <div style={{ flex: 1 }}>
                  <select value={g.unit} onChange={e => updateGood(i,"unit",e.target.value)} style={{ ...inp, padding: "9px 6px", cursor: "pointer" }}>
                    <option>кг</option><option>т</option><option>шт</option><option>ящ</option><option>л</option>
                  </select>
                </div>
                <div style={{ flex: 2 }}><input type="number" min={0} step="any" value={g.buyPrice} onChange={e => updateGood(i,"buyPrice",e.target.value)} placeholder="0" style={inp} /></div>
                <div style={{ flex: 2, padding: "9px 10px", background: "#0d1a2e", border: "1.5px solid #1e2d45", borderRadius: 8, fontSize: 13, color: rowTotal > 0 ? "#ffc947" : "#2a3f5f", textAlign: "right" }}>
                  {rowTotal > 0 ? rowTotal.toLocaleString("uk-UA", { maximumFractionDigits: 2 }) : "—"}
                </div>
                {goods.length > 1 ? <button onClick={() => removeGood(i)} style={{ width:36,height:38,background:"#2a1a1a",border:"1.5px solid #5a2020",borderRadius:8,color:"#ff6b6b",cursor:"pointer",fontSize:17,flexShrink:0 }}>×</button> : <div style={{ width: 36 }} />}
              </div>
            );
          })}
          <button onClick={addGood} style={{ background:"#0d1f35",border:"1.5px dashed #2a3f5f",borderRadius:8,color:"#38bdf8",padding:"8px 16px",cursor:"pointer",fontSize:13,width:"100%",marginTop:4 }}>+ Додати позицію</button>
          {goodsTotal > 0 && <div style={{ marginTop:10,fontSize:13,color:"#4e6a9a",textAlign:"right" }}>Загальна собівартість: <span style={{ color:"#ff8a5a",fontWeight:700 }}>{formatUAH(goodsTotal)}</span></div>}
        </Section>

        {/* ВОДІЇ */}
        <Section title="Водії" color="#a78bfa">
          <div style={{ display: "flex", gap: 8, marginBottom: 6 }}>
            <div style={{ flex: 2 }}><span style={lbl}>Денна ставка ₴</span></div>
            <div style={{ flex: 2 }}><span style={lbl}>Кількість поїздок</span></div>
            <div style={{ flex: 2 }}><span style={lbl}>Зарплата ₴</span></div>
            <div style={{ width: 36 }} />
          </div>
          {drivers.map((d, i) => {
            const rowTotal = fmt(d.salary) * (parseInt(d.trips) || 1);
            return (
              <div key={i} style={{ display:"flex",gap:8,marginBottom:8,alignItems:"center" }}>
                <div style={{ flex:2 }}><input type="number" min={0} value={d.salary} onChange={e=>updateDriver(i,"salary",e.target.value)} placeholder="0" style={inp} /></div>
                <div style={{ flex:2 }}><input type="number" min={1} value={d.trips} onChange={e=>updateDriver(i,"trips",e.target.value)} placeholder="1" style={inp} /></div>
                <div style={{ flex:2,padding:"9px 10px",background:"#0d1a2e",border:"1.5px solid #1e2d45",borderRadius:8,fontSize:13,color:rowTotal>0?"#ffc947":"#2a3f5f",textAlign:"right" }}>
                  {rowTotal > 0 ? rowTotal.toLocaleString("uk-UA",{maximumFractionDigits:2}) : "—"}
                </div>
                {drivers.length > 1 ? <button onClick={()=>removeDriver(i)} style={{ width:36,height:38,background:"#2a1a1a",border:"1.5px solid #5a2020",borderRadius:8,color:"#ff6b6b",cursor:"pointer",fontSize:17,flexShrink:0 }}>×</button> : <div style={{ width:36 }} />}
              </div>
            );
          })}
          <button onClick={addDriver} style={{ background:"#0d1f35",border:"1.5px dashed #2a3f5f",borderRadius:8,color:"#a78bfa",padding:"8px 16px",cursor:"pointer",fontSize:13,width:"100%",marginTop:4 }}>+ Додати водія</button>
          {driverTotal > 0 && <div style={{ marginTop:8,fontSize:12,color:"#4e6a9a",textAlign:"right" }}>Всього зарплат: <span style={{ color:"#ff8a5a",fontWeight:700 }}>{formatUAH(driverTotal)}</span></div>}
        </Section>

        {/* ПАЛИВО */}
        <Section title="Паливо" color="#34d399">
          <div style={{ display:"flex",alignItems:"center",gap:10,marginBottom:16 }}>
            <div onClick={()=>setFuelEnabled(v=>!v)} style={{ width:40,height:22,borderRadius:11,background:fuelEnabled?"#0a4a2a":"#1e2d45",border:`1.5px solid ${fuelEnabled?"#34d399":"#2a3f5f"}`,cursor:"pointer",position:"relative",transition:"all 0.2s",flexShrink:0 }}>
              <div style={{ position:"absolute",top:2,left:fuelEnabled?18:2,width:16,height:16,borderRadius:"50%",background:fuelEnabled?"#34d399":"#4e6a9a",transition:"left 0.2s" }} />
            </div>
            <span style={{ fontSize:13,color:fuelEnabled?"#34d399":"#4e6a9a" }}>{fuelEnabled?"Враховується":"Не враховувати"}</span>
          </div>
          {fuelEnabled && (<>
            <div style={{ marginBottom:14 }}>
              <label style={lbl}>Тип транспорту</label>
              <div style={{ display:"flex",gap:8,flexWrap:"wrap" }}>
                {Object.entries(VEHICLE_TYPES).map(([key,{label,icon,consumption}]) => (
                  <button key={key} onClick={()=>setVehicleType(key)} style={{ padding:"9px 14px",borderRadius:8,border:"1.5px solid",borderColor:vehicleType===key?"#34d399":"#2a3f5f",background:vehicleType===key?"#0a2e1e":"#1e2d45",color:vehicleType===key?"#34d399":"#7a90b8",fontWeight:vehicleType===key?700:400,cursor:"pointer",fontSize:13,display:"flex",flexDirection:"column",alignItems:"center",gap:2,minWidth:90 }}>
                    <span style={{ fontSize:20 }}>{icon}</span><span>{label}</span><span style={{ fontSize:11,opacity:0.7 }}>{consumption}л/100км</span>
                  </button>
                ))}
              </div>
            </div>
            <div style={{ marginBottom:14 }}>
              <label style={lbl}>Свій середній розхід (л/100км) — необов'язково</label>
              <input type="number" min={1} step="any" value={customConsumption} onChange={e=>setCustomConsumption(e.target.value)} placeholder={`За замовчуванням: ${VEHICLE_TYPES[vehicleType].consumption} л/100км`} style={{ ...inp, maxWidth:280 }} />
            </div>
            <div style={{ display:"flex",gap:12,flexWrap:"wrap",alignItems:"flex-start" }}>
              <div style={{ flex:"1 1 140px" }}>
                <NumInput label="Км одного рейсу (в одну сторону)" value={routeKm} onChange={setRouteKm} suffix="км" hint="Від бази до точки розвантаження" />
              </div>
              <div style={{ flex:"0 0 auto",paddingTop:4 }}>
                <label style={lbl}>Маршрут</label>
                <div style={{ display:"flex",gap:8 }}>
                  {[{val:true,label:"Туди + Назад"},{val:false,label:"Тільки туди"}].map(opt=>(
                    <button key={String(opt.val)} onClick={()=>setRoundTrip(opt.val)} style={{ padding:"9px 12px",borderRadius:8,border:"1.5px solid",borderColor:roundTrip===opt.val?"#34d399":"#2a3f5f",background:roundTrip===opt.val?"#0a2e1e":"#1e2d45",color:roundTrip===opt.val?"#34d399":"#7a90b8",fontWeight:roundTrip===opt.val?700:400,cursor:"pointer",fontSize:12 }}>{opt.label}</button>
                  ))}
                </div>
              </div>
              <div style={{ flex:"1 1 120px" }}>
                <NumInput label="Ціна палива (₴/л)" value={fuelPrice} onChange={setFuelPrice} prefix="₴" />
              </div>
            </div>
            {kmPerTrip > 0 && (
              <div style={{ background:"#0d1a2e",borderRadius:10,padding:"12px 16px",marginTop:4,display:"flex",gap:16,flexWrap:"wrap" }}>
                <div style={{ flex:1 }}>
                  <div style={{ fontSize:11,color:"#4e6a9a",marginBottom:3 }}>ЗАГАЛЬНИЙ ПРОБІГ</div>
                  <div style={{ fontSize:16,fontWeight:700,color:"#e8f0ff" }}>{kmTotal.toLocaleString("uk-UA")} км</div>
                  <div style={{ fontSize:11,color:"#4e6a9a",marginTop:2 }}>{kmPerTrip}км × {roundTrip?"2":"1"} × {totalTrips} поїздок</div>
                </div>
                <div style={{ flex:1 }}>
                  <div style={{ fontSize:11,color:"#4e6a9a",marginBottom:3 }}>ВИТРАТА ПАЛИВА</div>
                  <div style={{ fontSize:16,fontWeight:700,color:"#e8f0ff" }}>{litersTotal.toFixed(1)} л</div>
                  <div style={{ fontSize:11,color:"#4e6a9a",marginTop:2 }}>{consumption}л/100км</div>
                </div>
                <div style={{ flex:1 }}>
                  <div style={{ fontSize:11,color:"#4e6a9a",marginBottom:3 }}>ВАРТІСТЬ ПАЛИВА</div>
                  <div style={{ fontSize:16,fontWeight:700,color:"#ffc947" }}>{formatUAH(fuelTotal)}</div>
                  <div style={{ fontSize:11,color:"#4e6a9a",marginTop:2 }}>{litersTotal.toFixed(1)}л × {fuelPrice}₴</div>
                </div>
              </div>
            )}
          </>)}
        </Section>

        {/* ІНШІ ВИТРАТИ */}
        <Section title="Інші витрати" color="#f472b6">
          {extraCosts.map((e, i) => (
            <div key={i} style={{ display:"flex",gap:8,marginBottom:8 }}>
              <input value={e.name} onChange={ev=>updateExtra(i,"name",ev.target.value)} placeholder="Упаковка, логістика, дозвіл…" style={{ ...inp,flex:3 }} />
              <input type="number" min={0} step="any" value={e.amount} onChange={ev=>updateExtra(i,"amount",ev.target.value)} placeholder="0 ₴" style={{ ...inp,flex:2 }} />
              <button onClick={()=>removeExtra(i)} style={{ width:36,height:38,background:"#2a1a1a",border:"1.5px solid #5a2020",borderRadius:8,color:"#ff6b6b",cursor:"pointer",fontSize:17,flexShrink:0 }}>×</button>
            </div>
          ))}
          <button onClick={addExtra} style={{ background:"#0d1f35",border:"1.5px dashed #2a3f5f",borderRadius:8,color:"#f472b6",padding:"8px 16px",cursor:"pointer",fontSize:13,width:"100%" }}>+ Додати витрату</button>
          {extraTotal > 0 && <div style={{ marginTop:8,fontSize:12,color:"#4e6a9a",textAlign:"right" }}>Всього: <span style={{ color:"#ff8a5a",fontWeight:700 }}>{formatUAH(extraTotal)}</span></div>}
        </Section>

        {/* ОПОДАТКУВАННЯ */}
        <Section title="Оподаткування" color="#fbbf24">
          <label style={lbl}>Система оподаткування</label>
          <div style={{ display:"flex",gap:8,flexWrap:"wrap",marginBottom:14 }}>
            {[{key:"fop2",label:"ФОП 2 група",sub:"фіксована сума/міс"},{key:"fop3",label:"ФОП 3 група",sub:"% від обороту"},{key:"custom",label:"Інше",sub:"% від прибутку"}].map(opt => (
              <button key={opt.key} onClick={()=>setTaxType(opt.key)} style={{ padding:"10px 16px",borderRadius:8,border:"1.5px solid",borderColor:taxType===opt.key?"#fbbf24":"#2a3f5f",background:taxType===opt.key?"#2a1f00":"#1e2d45",color:taxType===opt.key?"#fbbf24":"#7a90b8",fontWeight:taxType===opt.key?700:400,cursor:"pointer",fontSize:13,display:"flex",flexDirection:"column",alignItems:"flex-start",gap:2 }}>
                <span>{opt.label}</span><span style={{ fontSize:11,opacity:0.7,fontWeight:400 }}>{opt.sub}</span>
              </button>
            ))}
          </div>
          {taxType === "fop2" && (
            <div style={{ display:"flex",gap:12,alignItems:"flex-start",flexWrap:"wrap" }}>
              <div style={{ flex:"1 1 180px" }}><NumInput label="Фіксована сума на місяць (₴)" value={fop2Monthly} onChange={setFop2Monthly} prefix="₴" hint="Єдиний податок + ЄСВ" /></div>
              <div style={{ flex:"1 1 180px",background:"#0d1a2e",borderRadius:10,padding:"12px 14px" }}>
                <div style={{ fontSize:11,color:"#4e6a9a",marginBottom:4 }}>СУМА ПОДАТКУ</div>
                <div style={{ fontSize:18,fontWeight:700,color:"#fbbf24" }}>{formatUAH(fmt(fop2Monthly))}</div>
                <div style={{ fontSize:11,color:"#4e6a9a",marginTop:4 }}>Фіксована — не залежить від суми тендера</div>
              </div>
            </div>
          )}
          {taxType === "fop3" && (
            <div style={{ display:"flex",gap:12,alignItems:"flex-start",flexWrap:"wrap" }}>
              <div style={{ flex:"1 1 140px" }}><NumInput label="Ставка (%)" value={fop3Rate} onChange={setFop3Rate} suffix="%" hint="Зазвичай 5% або 6% від обороту" /></div>
              <div style={{ flex:"1 1 180px",background:"#0d1a2e",borderRadius:10,padding:"12px 14px" }}>
                <div style={{ fontSize:11,color:"#4e6a9a",marginBottom:4 }}>СУМА ПОДАТКУ</div>
                <div style={{ fontSize:18,fontWeight:700,color:"#fbbf24" }}>{formatUAH(amount*(fmt(fop3Rate)||6)/100)}</div>
                <div style={{ fontSize:11,color:"#4e6a9a",marginTop:4 }}>{fop3Rate}% від обороту {formatUAH(amount)}</div>
              </div>
            </div>
          )}
          {taxType === "custom" && (
            <div style={{ maxWidth:240 }}><NumInput label="Ставка % від прибутку" value={customTaxRate} onChange={setCustomTaxRate} suffix="%" /></div>
          )}
        </Section>

        {/* РЕЗУЛЬТАТ */}
        <div style={{ background:"linear-gradient(135deg,#0d1f38 0%,#0a1828 100%)",borderRadius:16,padding:"22px 24px",border:`2px solid ${netProfit>=0?"#1a4a2a":"#4a1a1a"}`,marginBottom:16 }}>
          <div style={{ fontSize:13,fontWeight:700,color:"#8a9bb8",letterSpacing:"0.06em",textTransform:"uppercase",marginBottom:16,display:"flex",alignItems:"center",gap:8,flexWrap:"wrap" }}>
            <span style={{ display:"inline-block",width:3,height:16,background:"#4a8aff",borderRadius:2 }} />
            Результат
            {tenderName && <span style={{ fontWeight:400,color:"#4e6a9a",textTransform:"none",letterSpacing:0 }}>— {tenderName}</span>}
          </div>

          <ResultRow label="Сума тендера (дохід)" value={formatUAH(amount)} color="#7dd3fc" />
          <div style={{ height:1,background:"#1e2d45",margin:"6px 0" }} />
          <ResultRow label="Собівартість товару" value={`− ${formatUAH(goodsTotal)}`} color="#ff8a5a" />
          <ResultRow label={`Комісія платформи (${(commissionRate*100).toFixed(2)}%)`} value={`− ${formatUAH(commissionAmt)}`} color="#ff8a5a" />
          <ResultRow label="Зарплати водіїв" value={`− ${formatUAH(driverTotal)}`} color="#ff8a5a" />
          {fuelEnabled && (<>
            <ResultRow label="Паливо" value={`− ${formatUAH(fuelTotal)}`} color="#ff8a5a" />
            {kmPerTrip > 0 && <ResultRow label={`  ${kmTotal} км · ${litersTotal.toFixed(1)}л · ${fuelPrice}₴/л`} value="" color="#4e6a9a" sub />}
          </>)}
          {extraTotal > 0 && <ResultRow label="Інші витрати" value={`− ${formatUAH(extraTotal)}`} color="#ff8a5a" />}
          <div style={{ height:1,background:"#1e2d45",margin:"8px 0" }} />
          <ResultRow label="Загальні витрати" value={`− ${formatUAH(totalCosts)}`} color="#ffc947" bold />
          <div style={{ height:1,background:"#1e2d45",margin:"8px 0" }} />
          <ResultRow label="Прибуток до оподаткування" value={formatUAH(grossProfit)} color={grossProfit>=0?"#7dd3fc":"#ff5a5a"} />
          {taxAmt > 0 && <ResultRow label={taxLabel} value={`− ${formatUAH(taxAmt)}`} color="#ff8a5a" />}
          <div style={{ height:2,background:"#1e2d45",margin:"12px 0" }} />
          <div style={{ display:"flex",justifyContent:"space-between",alignItems:"center" }}>
            <span style={{ fontSize:16,fontWeight:800,color:"#e8f0ff" }}>Чистий заробіток</span>
            <span style={{ fontSize:28,fontWeight:900,color:profitColor }}>{formatUAH(netProfit)}</span>
          </div>

          {amount > 0 && (
            <div style={{ marginTop:14,display:"flex",gap:10 }}>
              <div style={{ flex:1,background:"#0b1525",borderRadius:10,padding:"12px 14px",textAlign:"center" }}>
                <div style={{ fontSize:11,color:"#4e6a9a",marginBottom:4,textTransform:"uppercase",letterSpacing:"0.05em" }}>Маржа</div>
                <div style={{ fontSize:22,fontWeight:800,color:marginColor }}>{margin.toFixed(1)}%</div>
              </div>
              <div style={{ flex:1,background:"#0b1525",borderRadius:10,padding:"12px 14px",textAlign:"center" }}>
                <div style={{ fontSize:11,color:"#4e6a9a",marginBottom:4,textTransform:"uppercase",letterSpacing:"0.05em" }}>Витрати</div>
                <div style={{ fontSize:22,fontWeight:800,color:"#7dd3fc" }}>{amount>0?((totalCosts/amount)*100).toFixed(1):0}%</div>
              </div>
              <div style={{ flex:1,background:"#0b1525",borderRadius:10,padding:"12px 14px",textAlign:"center" }}>
                <div style={{ fontSize:11,color:"#4e6a9a",marginBottom:4,textTransform:"uppercase",letterSpacing:"0.05em" }}>Націнка</div>
                <div style={{ fontSize:22,fontWeight:800,color:"#a78bfa" }}>{goodsTotal>0?(((amount-goodsTotal)/goodsTotal)*100).toFixed(1)+"%":"—"}</div>
              </div>
            </div>
          )}
        </div>

        {/* КНОПКИ ДІЙ */}
        <div style={{ display:"flex",gap:10,marginBottom:12,flexWrap:"wrap" }}>
          <button onClick={addToHistory} disabled={!amount} style={{ flex:2,minWidth:180,padding:"14px 20px",borderRadius:10,border:"none",background:amount?"linear-gradient(135deg,#1a4a8a,#0d2a55)":"#1a2235",color:amount?"#7dd3fc":"#4e6a9a",cursor:amount?"pointer":"not-allowed",fontSize:14,fontWeight:700,display:"flex",alignItems:"center",justifyContent:"center",gap:8 }}>
            {addedMsg ? <><span>✓</span> Додано!</> : <><span style={{fontSize:18}}>📊</span> Додати до таблиці</>}
          </button>
          <button onClick={downloadExcel} disabled={history.length===0} style={{ flex:1,minWidth:130,padding:"14px 20px",borderRadius:10,border:"none",background:history.length>0?"linear-gradient(135deg,#1a3a1a,#0a2e1e)":"#1a2235",color:history.length>0?"#3ddc84":"#4e6a9a",cursor:history.length>0?"pointer":"not-allowed",fontSize:14,fontWeight:700,display:"flex",alignItems:"center",justifyContent:"center",gap:8 }}>
            <span style={{fontSize:18}}>⬇️</span> Excel {history.length>0&&`(${history.length})`}
          </button>
        </div>

        {/* Google Sheets кнопка */}
        <div style={{ marginBottom:16 }}>
          <div style={{ display:"flex",gap:10,flexWrap:"wrap" }}>
            <button onClick={saveToSheets} disabled={sheetLoading||history.length===0} style={{ flex:2,minWidth:180,padding:"14px 20px",borderRadius:10,border:"none",background:history.length>0?"linear-gradient(135deg,#1a3a20,#0d2e18)":"#1a2235",color:history.length>0?"#4ade80":"#4e6a9a",cursor:history.length>0&&!sheetLoading?"pointer":"not-allowed",fontSize:14,fontWeight:700,display:"flex",alignItems:"center",justifyContent:"center",gap:8 }}>
              <span style={{fontSize:18}}>🟢</span>
              {sheetLoading ? sheetMsg||"Завантаження..." : "Зберегти в Google Sheets"}
            </button>
            {sheetId && (
              <button onClick={openSheet} style={{ flex:1,minWidth:130,padding:"14px 20px",borderRadius:10,border:"1.5px solid #2a5a3a",background:"transparent",color:"#4ade80",cursor:"pointer",fontSize:14,fontWeight:700,display:"flex",alignItems:"center",justifyContent:"center",gap:8 }}>
                <span style={{fontSize:18}}>📋</span> Відкрити таблицю
              </button>
            )}
          </div>
          {sheetMsg && !sheetLoading && (
            <div style={{ marginTop:8,fontSize:13,color:sheetMsg.startsWith("✅")?"#4ade80":"#ffc947",padding:"8px 12px",background:"#0d1a2e",borderRadius:8 }}>{sheetMsg}</div>
          )}
          {sheetId && (
            <div style={{ marginTop:6,fontSize:11,color:"#4e6a9a" }}>
              Таблиця: <span style={{color:"#38bdf8",cursor:"pointer",textDecoration:"underline"}} onClick={openSheet}>docs.google.com/spreadsheets/d/{sheetId.slice(0,20)}…</span>
              <span style={{marginLeft:8,cursor:"pointer",color:"#ff6b6b"}} onClick={()=>{setSheetId("");localStorage.removeItem("tender_sheet_id");}}>✕ відключити</span>
            </div>
          )}
        </div>

        {/* Підключити іншу таблицю */}

        {/* Список тендерів в пам'яті */}
        {history.length > 0 && (
          <div style={{ background:"#131e30",borderRadius:12,padding:"16px 20px",marginBottom:16,border:"1px solid #1e2d45" }}>
            <div style={{ fontSize:13,fontWeight:700,color:"#8a9bb8",letterSpacing:"0.06em",textTransform:"uppercase",marginBottom:12 }}>📋 Збережені тендери ({history.length})</div>
            {history.map((r, i) => (
              <div key={i} style={{ display:"flex",justifyContent:"space-between",alignItems:"center",padding:"8px 0",borderBottom:i<history.length-1?"1px solid #1e2d45":"none",gap:12,flexWrap:"wrap" }}>
                <div style={{ flex:1,minWidth:120 }}>
                  <div style={{ fontSize:13,color:"#e8f0ff",fontWeight:600 }}>{r.name}</div>
                  <div style={{ fontSize:11,color:"#4e6a9a" }}>{r.date} · {r.goods}</div>
                </div>
                <div style={{ textAlign:"right" }}>
                  <div style={{ fontSize:13,color:"#7dd3fc" }}>{formatUAH(r.amount)}</div>
                  <div style={{ fontSize:13,fontWeight:700,color:r.netProfit>=0?"#3ddc84":"#ff5a5a" }}>{formatUAH(r.netProfit)}</div>
                </div>
                <button onClick={()=>setHistory(h=>h.filter((_,idx)=>idx!==i))} style={{ background:"#2a1a1a",border:"1.5px solid #5a2020",borderRadius:6,color:"#ff6b6b",padding:"4px 10px",cursor:"pointer",fontSize:12 }}>видалити</button>
              </div>
            ))}
            <div style={{ marginTop:12,paddingTop:12,borderTop:"2px solid #2a3f5f",display:"flex",justifyContent:"space-between" }}>
              <span style={{ fontSize:13,color:"#8a9bb8",fontWeight:700 }}>Загальний чистий дохід</span>
              <span style={{ fontSize:16,fontWeight:800,color:"#3ddc84" }}>{formatUAH(history.reduce((s,r)=>s+r.netProfit,0))}</span>
            </div>
          </div>
        )}

        {/* Очистити форму */}
        <button onClick={reset} style={{ width:"100%",padding:"12px",borderRadius:10,background:"transparent",border:"1.5px solid #2a3f5f",color:"#4e6a9a",cursor:"pointer",fontSize:13,fontWeight:600 }}>
          Очистити форму та почати новий розрахунок
        </button>

      </div>
    </div>
  );
}
