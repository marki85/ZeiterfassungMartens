import React, { useState, useEffect, useMemo, useCallback } from "react";
import * as XLSX from "xlsx";
import {
  Clock, Users, FileText, Settings as SettingsIcon, Plus, Trash2,
  Download, Printer, X, ChevronLeft, Hammer,
} from "lucide-react";

/* ============================================================
   Farbsystem als echtes CSS (keine Tailwind-Arbitrary-Values,
   die im Artefakt nicht rendern). color-scheme: light erzwingt
   helle Darstellung auch bei iPhone-Dark-Mode.
   ============================================================ */
const STYLE = `
:root { color-scheme: light only; }
* { -webkit-tap-highlight-color: transparent; box-sizing: border-box; }

.wz {
  --bg:#EDEFEA; --surface:#FFFFFF; --surface2:#F7F8F5;
  --ink:#16201A; --ink2:#525C55; --ink3:#8B938B;
  --brand:#1F3B2C; --brand2:#2C5140; --onbrand:#F1F5EE;
  --accent:#C07A22; --onaccent:#FFFFFF;
  --danger:#A5392C; --border:#DFE3DD; --borderStrong:#C7CDC4;
  background:var(--bg); color:var(--ink);
  min-height:100vh; padding-bottom:88px;
  font-family:-apple-system,BlinkMacSystemFont,"Segoe UI",Roboto,Helvetica,Arial,sans-serif;
}
.wz-num { font-variant-numeric:tabular-nums;
  font-family:"SF Mono",ui-monospace,Menlo,Consolas,monospace; }

.wz-header { background:var(--brand); color:var(--onbrand);
  padding:22px 20px 18px; }
.wz-header-in { max-width:560px; margin:0 auto; display:flex;
  align-items:center; gap:12px; }
.wz-logo { width:42px; height:42px; border-radius:11px; background:var(--accent);
  display:flex; align-items:center; justify-content:center; flex-shrink:0; }
.wz-title { font-weight:700; font-size:18px; letter-spacing:-0.02em; line-height:1; margin:0; }
.wz-sub { font-size:12px; color:#B9C6B7; margin:5px 0 0; }

.wz-main { max-width:560px; margin:0 auto; padding:18px 16px 0; }
.wz-h2 { font-weight:650; font-size:19px; color:var(--brand); margin:0; letter-spacing:-0.01em; }
.wz-row-head { display:flex; align-items:center; justify-content:space-between; margin-bottom:14px; }

.wz-fab { background:var(--accent); color:var(--onaccent); border:none;
  border-radius:999px; width:40px; height:40px; display:flex; align-items:center;
  justify-content:center; box-shadow:0 2px 6px rgba(31,59,44,.22); cursor:pointer; }
.wz-fab:active { transform:scale(.94); }

.wz-card { background:var(--surface); border:1px solid var(--border);
  border-radius:14px; overflow:hidden; }
.wz-list-item { display:flex; align-items:flex-start; justify-content:space-between;
  padding:13px 14px; border-bottom:1px solid var(--border); }
.wz-list-item:last-child { border-bottom:none; }

.wz-month { font-size:11px; font-weight:700; letter-spacing:.06em; text-transform:uppercase;
  color:var(--ink3); margin:0 0 8px 2px; }
.wz-name { font-weight:600; font-size:14px; color:var(--ink); margin:0; }
.wz-desc { font-size:12.5px; color:var(--ink2); margin:2px 0 0; }
.wz-meta { font-size:12px; color:var(--ink3); margin:5px 0 0; }
.wz-hours { font-weight:650; font-size:14px; color:var(--brand); margin:0; }
.wz-hours-helper { font-size:12px; color:var(--accent); margin:2px 0 0; font-weight:600; }

.wz-badge { display:inline-block; font-size:10px; font-weight:600; background:#DDE7D8;
  color:var(--brand); padding:2px 7px; border-radius:999px; vertical-align:middle; margin-left:6px; }

.wz-del { color:var(--danger); background:none; border:none; font-size:12px; cursor:pointer;
  display:flex; align-items:center; gap:4px; margin-left:auto; padding:2px 0; }

.wz-empty { text-align:center; padding:38px 20px; background:var(--surface2);
  border:1px dashed var(--borderStrong); border-radius:14px; }
.wz-empty-title { font-weight:600; color:var(--brand); margin:8px 0 0; }
.wz-empty-text { font-size:13.5px; color:var(--ink2); margin:5px 0 0; }

.wz-btn { width:100%; border:none; border-radius:12px; font-weight:650; font-size:15px;
  padding:13px; cursor:pointer; transition:transform .05s; }
.wz-btn:active { transform:scale(.99); }
.wz-btn-primary { background:var(--brand); color:var(--onbrand); }
.wz-btn-accent { background:var(--accent); color:var(--onaccent); }
.wz-btn[disabled] { opacity:.4; cursor:not-allowed; }
.wz-btn-ghost { width:100%; background:none; border:1.5px solid var(--brand); color:var(--brand);
  border-radius:12px; font-weight:600; font-size:14px; padding:11px; cursor:pointer;
  display:flex; align-items:center; justify-content:center; gap:8px; }

.wz-label { display:block; font-size:11px; font-weight:700; letter-spacing:.05em;
  text-transform:uppercase; color:var(--ink2); margin-bottom:6px; }
.wz-field { margin-bottom:16px; }
.wz-input { width:100%; background:var(--surface); border:1px solid var(--borderStrong);
  border-radius:11px; padding:11px 12px; font-size:15px; color:var(--ink);
  outline:none; font-family:inherit; }
.wz-input:focus { border-color:var(--accent); box-shadow:0 0 0 3px rgba(192,122,34,.18); }
textarea.wz-input { min-height:74px; resize:vertical; }
.wz-suffix { position:relative; }
.wz-suffix span { position:absolute; right:12px; top:11px; font-size:13px; color:var(--ink3); }
.wz-suffix .wz-input { padding-right:42px; }

.wz-seg { display:flex; gap:8px; margin-bottom:16px; }
.wz-seg button { flex:1; padding:9px; border-radius:999px; font-size:13px; font-weight:600;
  cursor:pointer; border:1px solid var(--borderStrong); background:var(--surface); color:var(--ink2); }
.wz-seg button.on { background:var(--brand); color:var(--onbrand); border-color:var(--brand); }

.wz-check { display:flex; align-items:center; gap:10px; font-size:14px; font-weight:600;
  color:var(--brand); cursor:pointer; margin-bottom:14px; }
.wz-check input { width:18px; height:18px; accent-color:var(--accent); }
.wz-helper-box { border-left:3px solid var(--accent); padding-left:12px; margin:0 0 14px 2px;
  display:grid; grid-template-columns:1fr 1fr; gap:12px; }
.wz-grid2 { display:grid; grid-template-columns:1fr 1fr; gap:12px; }

.wz-nav { position:fixed; bottom:0; left:0; right:0; background:var(--brand);
  border-top:1px solid #16281E; z-index:40; }
.wz-nav-in { max-width:560px; margin:0 auto; display:grid; grid-template-columns:repeat(4,1fr); }
.wz-nav button { display:flex; flex-direction:column; align-items:center; gap:3px;
  padding:11px 0 13px; font-size:11px; font-weight:600; background:none; border:none;
  color:#8FA089; cursor:pointer; }
.wz-nav button.on { color:var(--accent); }

.wz-sheet-bg { position:fixed; inset:0; z-index:50; background:rgba(10,18,13,.5);
  display:flex; align-items:flex-end; justify-content:center; }
.wz-sheet { background:var(--bg); width:100%; max-width:560px; border-radius:20px 20px 0 0;
  max-height:88vh; overflow-y:auto; padding:20px 18px 30px; }
.wz-sheet-head { display:flex; align-items:center; justify-content:space-between; margin-bottom:16px; }
.wz-sheet-head h3 { margin:0; font-weight:650; font-size:18px; color:var(--brand); }
.wz-x { background:none; border:none; color:var(--ink2); cursor:pointer; padding:4px; }

.wz-toast { position:fixed; bottom:96px; left:50%; transform:translateX(-50%);
  background:var(--brand); color:var(--onbrand); padding:10px 18px; border-radius:999px;
  font-size:13.5px; z-index:60; box-shadow:0 4px 14px rgba(0,0,0,.25); }

.wz-warn { font-size:13.5px; color:var(--danger); margin:0; }
.wz-hint { font-size:14px; color:var(--ink2); margin:8px 0; }
.wz-selrow { display:flex; align-items:center; gap:12px; background:var(--surface);
  border:1px solid var(--border); border-radius:10px; padding:10px; margin-bottom:6px; }
.wz-selrow input { width:18px; height:18px; accent-color:var(--accent); }
.wz-total { display:flex; justify-content:space-between; align-items:center;
  border-top:1px solid var(--border); padding-top:12px; margin-top:14px; }

.wz-inv-card { width:100%; text-align:left; background:var(--surface); border:1px solid var(--border);
  border-radius:12px; padding:13px; display:flex; justify-content:space-between; align-items:center;
  cursor:pointer; margin-bottom:8px; }

.wz-tip { font-size:12.5px; color:var(--ink2); line-height:1.55; padding-top:16px;
  margin-top:16px; border-top:1px solid var(--border); }

/* ---------- Rechnungs-Vorschau / Druck ---------- */
.wz-invwrap { background:var(--bg); min-height:100vh; }
.wz-invbar { max-width:560px; margin:0 auto; padding:22px 16px 10px;
  display:flex; align-items:center; justify-content:space-between; }
.wz-back { display:flex; align-items:center; gap:4px; background:none; border:none;
  color:var(--brand); font-weight:600; font-size:14px; cursor:pointer; }
.wz-print { display:flex; align-items:center; gap:8px; background:var(--brand); color:var(--onbrand);
  border:none; padding:9px 16px; border-radius:999px; font-size:14px; font-weight:600; cursor:pointer; }
.wz-paper { max-width:560px; margin:6px auto; background:#fff; color:#1a1a1a; padding:34px;
  box-shadow:0 1px 8px rgba(0,0,0,.08); }
.wz-paper h4 { margin:0; font-size:17px; }
.wz-paper table { width:100%; border-collapse:collapse; font-size:13.5px; margin:22px 0 10px; }
.wz-paper th { text-align:left; border-bottom:2px solid #1a1a1a; padding:7px 4px; font-size:12px; }
.wz-paper td { padding:6px 4px; border-bottom:1px solid #eee; }
.wz-r { text-align:right; }

@media print {
  body * { visibility:hidden; }
  #wz-print, #wz-print * { visibility:visible; }
  #wz-print { position:absolute; top:0; left:0; width:100%; box-shadow:none; margin:0; }
  .wz-noprint { display:none !important; }
}
`;

/* ---------- helpers ---------- */
const uid = () => Math.random().toString(36).slice(2, 10);
const todayISO = () => new Date().toISOString().slice(0, 10);
const eur = (n) => new Intl.NumberFormat("de-DE", { style: "currency", currency: "EUR" }).format(n || 0);
const fmtDate = (iso) => { if (!iso) return ""; const [y, m, d] = iso.split("-"); return `${d}.${m}.${y}`; };
const hoursBetween = (start, end) => {
  if (!start || !end) return 0;
  const [sh, sm] = start.split(":").map(Number);
  const [eh, em] = end.split(":").map(Number);
  let mins = eh * 60 + em - (sh * 60 + sm);
  if (mins < 0) mins += 1440;
  return Math.round((mins / 60) * 100) / 100;
};

const DEFAULT_DATA = {
  settings: {
    firma: "Martens Garten- & Handwerksservice",
    inhaber: "Johannes Martens",
    adresse: "",
    steuernummer: "",
    iban: "",
    sohnName: "Sohn",
    rateOwn: 35,
    rateHelper: 15,
    nextInvoiceNo: 1,
    invoicePrefix: new Date().getFullYear().toString(),
  },
  customers: [],
  entries: [],
  invoices: [],
};
const STORAGE_KEY = "werkzeit-data-v2";

export default function App() {
  const [data, setData] = useState(DEFAULT_DATA);
  const [loaded, setLoaded] = useState(false);
  const [tab, setTab] = useState("zeit");
  const [invoicePreview, setInvoicePreview] = useState(null);
  const [toast, setToast] = useState("");

  useEffect(() => {
    (async () => {
      try {
        const res = await window.storage.get(STORAGE_KEY, false);
        if (res && res.value) {
          const p = JSON.parse(res.value);
          setData({ ...DEFAULT_DATA, ...p, settings: { ...DEFAULT_DATA.settings, ...p.settings } });
        }
      } catch (e) {}
      setLoaded(true);
    })();
  }, []);

  const save = useCallback(async (next) => {
    setData(next);
    try { await window.storage.set(STORAGE_KEY, JSON.stringify(next), false); }
    catch (e) { setToast("Speichern fehlgeschlagen"); setTimeout(() => setToast(""), 3000); }
  }, []);

  const showToast = (m) => { setToast(m); setTimeout(() => setToast(""), 2200); };

  if (!loaded) {
    return (
      <div className="wz" style={{ display: "flex", alignItems: "center", justifyContent: "center" }}>
        <style>{STYLE}</style>
        <span style={{ color: "#1F3B2C", fontWeight: 600 }}>Lade Daten…</span>
      </div>
    );
  }

  if (invoicePreview) {
    return (
      <div className="wz">
        <style>{STYLE}</style>
        <InvoicePreview invoice={invoicePreview} settings={data.settings}
          customers={data.customers} onBack={() => setInvoicePreview(null)} />
      </div>
    );
  }

  return (
    <div className="wz">
      <style>{STYLE}</style>
      <Header settings={data.settings} />
      <main className="wz-main">
        {tab === "zeit" && <ZeitTab data={data} save={save} showToast={showToast} />}
        {tab === "kunden" && <KundenTab data={data} save={save} showToast={showToast} />}
        {tab === "rechnungen" && <RechnungenTab data={data} save={save} showToast={showToast} openPreview={setInvoicePreview} />}
        {tab === "einstellungen" && <EinstellungenTab data={data} save={save} showToast={showToast} />}
      </main>
      <NavBar tab={tab} setTab={setTab} />
      {toast && <div className="wz-toast">{toast}</div>}
    </div>
  );
}

function Header({ settings }) {
  return (
    <header className="wz-header">
      <div className="wz-header-in">
        <div className="wz-logo"><Hammer size={21} color="#1F3B2C" /></div>
        <div>
          <p className="wz-title">{settings.firma || "Werkzeit"}</p>
          <p className="wz-sub">Zeiterfassung &amp; Rechnungen</p>
        </div>
      </div>
    </header>
  );
}

function NavBar({ tab, setTab }) {
  const items = [
    { id: "zeit", label: "Zeit", icon: Clock },
    { id: "kunden", label: "Kunden", icon: Users },
    { id: "rechnungen", label: "Rechnungen", icon: FileText },
    { id: "einstellungen", label: "Mehr", icon: SettingsIcon },
  ];
  return (
    <nav className="wz-nav wz-noprint">
      <div className="wz-nav-in">
        {items.map((it) => {
          const Icon = it.icon;
          return (
            <button key={it.id} className={tab === it.id ? "on" : ""} onClick={() => setTab(it.id)}>
              <Icon size={20} /> {it.label}
            </button>
          );
        })}
      </div>
    </nav>
  );
}

/* ---------- Zeit ---------- */
function ZeitTab({ data, save, showToast }) {
  const [showForm, setShowForm] = useState(false);
  const [edit, setEdit] = useState(null);
  const custName = (id) => data.customers.find((c) => c.id === id)?.name || "(kein Kunde)";

  const grouped = useMemo(() => {
    const sorted = [...data.entries].sort((a, b) => (a.date < b.date ? 1 : -1));
    const map = {};
    for (const e of sorted) { const k = e.date.slice(0, 7); (map[k] = map[k] || []).push(e); }
    return map;
  }, [data.entries]);

  const del = (id) => {
    if (data.entries.find((e) => e.id === id)?.invoiceId) { showToast("Abgerechnet – nicht löschbar"); return; }
    save({ ...data, entries: data.entries.filter((e) => e.id !== id) });
    showToast("Eintrag gelöscht");
  };

  return (
    <div>
      <div className="wz-row-head">
        <h2 className="wz-h2">Arbeitszeiten</h2>
        <button className="wz-fab" onClick={() => { setEdit(null); setShowForm(true); }}><Plus size={20} /></button>
      </div>

      {data.entries.length === 0 &&
        <EmptyState icon={Clock} title="Noch keine Zeit erfasst"
          text="Trag deinen ersten Arbeitseinsatz ein — Datum, Kunde, Dauer." />}

      {Object.entries(grouped).map(([month, entries]) => (
        <div key={month} style={{ marginBottom: 20 }}>
          <p className="wz-month">
            {new Date(month + "-01").toLocaleDateString("de-DE", { month: "long", year: "numeric" })}
          </p>
          <div className="wz-card">
            {entries.map((e) => (
              <div key={e.id} className="wz-list-item">
                <button style={{ textAlign: "left", flex: 1, background: "none", border: "none", padding: 0, cursor: "pointer" }}
                  onClick={() => {
                    if (e.invoiceId) { showToast("Abgerechnet – nicht bearbeitbar"); return; }
                    setEdit(e); setShowForm(true);
                  }}>
                  <p className="wz-name">
                    {custName(e.customerId)}
                    {e.invoiceId && <span className="wz-badge">abgerechnet</span>}
                  </p>
                  <p className="wz-desc">{e.description || "—"}</p>
                  <p className="wz-meta">{fmtDate(e.date)}</p>
                </button>
                <div style={{ textAlign: "right", marginLeft: 8, flexShrink: 0 }}>
                  <p className="wz-hours wz-num">{e.hoursOwn.toFixed(2)} h</p>
                  {e.hoursHelper > 0 &&
                    <p className="wz-hours-helper wz-num">+{e.hoursHelper.toFixed(2)} h ({data.settings.sohnName})</p>}
                  {!e.invoiceId &&
                    <button className="wz-del" onClick={() => del(e.id)}><Trash2 size={12} /> löschen</button>}
                </div>
              </div>
            ))}
          </div>
        </div>
      ))}

      {showForm &&
        <EntryForm entry={edit} data={data} onClose={() => setShowForm(false)}
          onSave={(entry) => {
            const exists = data.entries.some((e) => e.id === entry.id);
            const entries = exists ? data.entries.map((e) => (e.id === entry.id ? entry : e)) : [...data.entries, entry];
            save({ ...data, entries }); setShowForm(false);
            showToast(exists ? "Aktualisiert" : "Gespeichert");
          }} />}
    </div>
  );
}

function EntryForm({ entry, data, onClose, onSave }) {
  const [date, setDate] = useState(entry?.date || todayISO());
  const [customerId, setCustomerId] = useState(entry?.customerId || data.customers[0]?.id || "");
  const [description, setDescription] = useState(entry?.description || "");
  const [start, setStart] = useState(entry?.start || "08:00");
  const [end, setEnd] = useState(entry?.end || "12:00");
  const [manualHours, setManualHours] = useState(entry ? entry.hoursOwn : hoursBetween("08:00", "12:00"));
  const [useManual, setUseManual] = useState(!!entry?.manual);
  const [rateOwn, setRateOwn] = useState(entry?.rateOwn ?? data.settings.rateOwn);
  const [helperOn, setHelperOn] = useState(entry ? entry.hoursHelper > 0 : false);
  const [hoursHelper, setHoursHelper] = useState(entry?.hoursHelper || hoursBetween("08:00", "12:00"));
  const [rateHelper, setRateHelper] = useState(entry?.rateHelper ?? data.settings.rateHelper);

  const computed = useManual ? Number(manualHours) || 0 : hoursBetween(start, end);

  const handleSave = () => {
    if (!customerId) return;
    onSave({
      id: entry?.id || uid(), date, customerId, description,
      start: useManual ? null : start, end: useManual ? null : end, manual: useManual,
      hoursOwn: computed, rateOwn: Number(rateOwn) || 0,
      hoursHelper: helperOn ? Number(hoursHelper) || 0 : 0, rateHelper: Number(rateHelper) || 0,
      invoiceId: entry?.invoiceId || null,
    });
  };

  return (
    <Sheet title={entry ? "Eintrag bearbeiten" : "Neuer Eintrag"} onClose={onClose}>
      <Field label="Datum"><input type="date" className="wz-input" value={date} onChange={(e) => setDate(e.target.value)} /></Field>

      <Field label="Kunde">
        {data.customers.length === 0
          ? <p className="wz-warn">Erst einen Kunden anlegen (Tab „Kunden").</p>
          : <select className="wz-input" value={customerId} onChange={(e) => setCustomerId(e.target.value)}>
              {data.customers.map((c) => <option key={c.id} value={c.id}>{c.name}</option>)}
            </select>}
      </Field>

      <Field label="Tätigkeit">
        <input className="wz-input" value={description} onChange={(e) => setDescription(e.target.value)}
          placeholder="z.B. Hecke schneiden, Terrasse pflastern…" />
      </Field>

      <div className="wz-seg">
        <button className={!useManual ? "on" : ""} onClick={() => setUseManual(false)}>Start / Ende</button>
        <button className={useManual ? "on" : ""} onClick={() => setUseManual(true)}>Stunden direkt</button>
      </div>

      {useManual
        ? <Field label="Stunden"><input type="number" step="0.25" min="0" className="wz-input"
            value={manualHours} onChange={(e) => setManualHours(e.target.value)} /></Field>
        : <div className="wz-grid2">
            <Field label="Start"><input type="time" className="wz-input" value={start} onChange={(e) => setStart(e.target.value)} /></Field>
            <Field label="Ende"><input type="time" className="wz-input" value={end} onChange={(e) => setEnd(e.target.value)} /></Field>
          </div>}

      <p className="wz-hint">Ergibt <b className="wz-num">{computed.toFixed(2)} h</b></p>

      <Field label="Stundenlohn (du)">
        <div className="wz-suffix"><input type="number" step="0.5" className="wz-input"
          value={rateOwn} onChange={(e) => setRateOwn(e.target.value)} /><span>€/h</span></div>
      </Field>

      <label className="wz-check">
        <input type="checkbox" checked={helperOn} onChange={(e) => setHelperOn(e.target.checked)} />
        {data.settings.sohnName} hat mitgeholfen
      </label>

      {helperOn &&
        <div className="wz-helper-box">
          <Field label={`Stunden (${data.settings.sohnName})`}>
            <input type="number" step="0.25" min="0" className="wz-input"
              value={hoursHelper} onChange={(e) => setHoursHelper(e.target.value)} />
          </Field>
          <Field label="Stundenlohn">
            <div className="wz-suffix"><input type="number" step="0.5" className="wz-input"
              value={rateHelper} onChange={(e) => setRateHelper(e.target.value)} /><span>€/h</span></div>
          </Field>
        </div>}

      <button className="wz-btn wz-btn-accent" disabled={!customerId} onClick={handleSave}>Speichern</button>
    </Sheet>
  );
}

/* ---------- Kunden ---------- */
function KundenTab({ data, save, showToast }) {
  const [showForm, setShowForm] = useState(false);
  const [edit, setEdit] = useState(null);
  const del = (id) => {
    if (data.entries.some((e) => e.customerId === id)) { showToast("Kunde hat Einträge – nicht löschbar"); return; }
    save({ ...data, customers: data.customers.filter((c) => c.id !== id) });
  };
  return (
    <div>
      <div className="wz-row-head">
        <h2 className="wz-h2">Kunden</h2>
        <button className="wz-fab" onClick={() => { setEdit(null); setShowForm(true); }}><Plus size={20} /></button>
      </div>
      {data.customers.length === 0 &&
        <EmptyState icon={Users} title="Noch keine Kunden" text="Leg deinen ersten Kunden an, um Zeiten zuzuordnen." />}
      <div>
        {data.customers.map((c) => (
          <div key={c.id} className="wz-card" style={{ padding: 13, display: "flex", justifyContent: "space-between", alignItems: "flex-start", marginBottom: 8 }}>
            <button style={{ textAlign: "left", background: "none", border: "none", padding: 0, cursor: "pointer" }}
              onClick={() => { setEdit(c); setShowForm(true); }}>
              <p className="wz-name">{c.name}</p>
              <p className="wz-desc" style={{ whiteSpace: "pre-line" }}>{c.address}</p>
            </button>
            <button className="wz-del" style={{ marginLeft: 8 }} onClick={() => del(c.id)}><Trash2 size={16} /></button>
          </div>
        ))}
      </div>
      {showForm &&
        <CustomerForm customer={edit} onClose={() => setShowForm(false)}
          onSave={(c) => {
            const exists = data.customers.some((x) => x.id === c.id);
            const customers = exists ? data.customers.map((x) => (x.id === c.id ? c : x)) : [...data.customers, c];
            save({ ...data, customers }); setShowForm(false);
          }} />}
    </div>
  );
}

function CustomerForm({ customer, onClose, onSave }) {
  const [name, setName] = useState(customer?.name || "");
  const [address, setAddress] = useState(customer?.address || "");
  return (
    <Sheet title={customer ? "Kunde bearbeiten" : "Neuer Kunde"} onClose={onClose}>
      <Field label="Name"><input className="wz-input" value={name} onChange={(e) => setName(e.target.value)} placeholder="Familie Müller" /></Field>
      <Field label="Adresse"><textarea className="wz-input" value={address} onChange={(e) => setAddress(e.target.value)} placeholder={"Musterstraße 1\n12345 Musterstadt"} /></Field>
      <button className="wz-btn wz-btn-accent" disabled={!name.trim()}
        onClick={() => onSave({ id: customer?.id || uid(), name: name.trim(), address })}>Speichern</button>
    </Sheet>
  );
}

/* ---------- Rechnungen ---------- */
function RechnungenTab({ data, save, showToast, openPreview }) {
  const [showForm, setShowForm] = useState(false);
  const exportExcel = () => {
    if (data.entries.length === 0) { showToast("Keine Daten zum Export"); return; }
    const rows = data.entries.map((e) => ({
      Datum: fmtDate(e.date),
      Kunde: data.customers.find((c) => c.id === e.customerId)?.name || "",
      Tätigkeit: e.description,
      "Stunden (du)": e.hoursOwn, "Satz (du) €/h": e.rateOwn,
      [`Stunden (${data.settings.sohnName})`]: e.hoursHelper, "Satz Helfer €/h": e.rateHelper,
      "Summe €": +(e.hoursOwn * e.rateOwn + e.hoursHelper * e.rateHelper).toFixed(2),
      Abgerechnet: e.invoiceId ? "ja" : "nein",
    }));
    const ws = XLSX.utils.json_to_sheet(rows);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "Zeiterfassung");
    XLSX.writeFile(wb, `zeiterfassung_${todayISO()}.xlsx`);
    showToast("Excel wird heruntergeladen…");
  };
  return (
    <div>
      <div className="wz-row-head">
        <h2 className="wz-h2">Rechnungen</h2>
        <button className="wz-fab" onClick={() => setShowForm(true)}><Plus size={20} /></button>
      </div>
      <div style={{ marginBottom: 16 }}>
        <button className="wz-btn-ghost" onClick={exportExcel}><Download size={16} /> Alle Zeiten als Excel exportieren</button>
      </div>
      {data.invoices.length === 0 &&
        <EmptyState icon={FileText} title="Noch keine Rechnung" text="Wähle Kunde und Einträge, um eine Rechnung zu erstellen." />}
      <div>
        {[...data.invoices].reverse().map((inv) => (
          <button key={inv.id} className="wz-inv-card" onClick={() => openPreview(inv)}>
            <div>
              <p className="wz-name">{data.customers.find((c) => c.id === inv.customerId)?.name || "Kunde gelöscht"}</p>
              <p className="wz-meta wz-num">Rechnung {inv.number} · {fmtDate(inv.createdAt)}</p>
            </div>
            <p className="wz-hours wz-num" style={{ fontSize: 15 }}>{eur(inv.total)}</p>
          </button>
        ))}
      </div>
      {showForm &&
        <NewInvoiceForm data={data} onClose={() => setShowForm(false)}
          onCreate={(invoice, updatedEntries) => {
            save({ ...data, entries: updatedEntries, invoices: [...data.invoices, invoice],
              settings: { ...data.settings, nextInvoiceNo: data.settings.nextInvoiceNo + 1 } });
            setShowForm(false); openPreview(invoice);
          }} />}
    </div>
  );
}

function NewInvoiceForm({ data, onClose, onCreate }) {
  const [customerId, setCustomerId] = useState(data.customers[0]?.id || "");
  const [selected, setSelected] = useState({});
  const openEntries = data.entries.filter((e) => e.customerId === customerId && !e.invoiceId);
  useEffect(() => {
    const s = {}; openEntries.forEach((e) => (s[e.id] = true)); setSelected(s);
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [customerId]);
  const chosen = openEntries.filter((e) => selected[e.id]);
  const total = chosen.reduce((s, e) => s + e.hoursOwn * e.rateOwn + e.hoursHelper * e.rateHelper, 0);

  const create = () => {
    if (chosen.length === 0) return;
    const number = `${data.settings.invoicePrefix}-${String(data.settings.nextInvoiceNo).padStart(3, "0")}`;
    const invoice = {
      id: uid(), number, customerId, createdAt: todayISO(), entryIds: chosen.map((e) => e.id),
      lines: chosen.map((e) => ({ date: e.date, description: e.description, hoursOwn: e.hoursOwn,
        rateOwn: e.rateOwn, hoursHelper: e.hoursHelper, rateHelper: e.rateHelper })),
      total,
    };
    const updated = data.entries.map((e) => (selected[e.id] ? { ...e, invoiceId: invoice.id } : e));
    onCreate(invoice, updated);
  };

  return (
    <Sheet title="Neue Rechnung" onClose={onClose}>
      <Field label="Kunde">
        {data.customers.length === 0
          ? <p className="wz-warn">Erst Kunden anlegen.</p>
          : <select className="wz-input" value={customerId} onChange={(e) => setCustomerId(e.target.value)}>
              {data.customers.map((c) => <option key={c.id} value={c.id}>{c.name}</option>)}
            </select>}
      </Field>

      {openEntries.length === 0
        ? <p className="wz-warn">Keine offenen Einträge für diesen Kunden.</p>
        : <div>
            <p className="wz-month">Offene Einträge</p>
            {openEntries.map((e) => (
              <label key={e.id} className="wz-selrow">
                <input type="checkbox" checked={!!selected[e.id]}
                  onChange={(ev) => setSelected({ ...selected, [e.id]: ev.target.checked })} />
                <div style={{ flex: 1 }}>
                  <p className="wz-name">{fmtDate(e.date)} — {e.description || "Arbeitszeit"}</p>
                  <p className="wz-meta wz-num">{e.hoursOwn.toFixed(2)} h{e.hoursHelper > 0 ? ` + ${e.hoursHelper.toFixed(2)} h Helfer` : ""}</p>
                </div>
                <p className="wz-num" style={{ fontSize: 13.5 }}>{eur(e.hoursOwn * e.rateOwn + e.hoursHelper * e.rateHelper)}</p>
              </label>
            ))}
          </div>}

      <div className="wz-total">
        <span style={{ fontWeight: 650 }}>Gesamt</span>
        <span className="wz-num" style={{ fontWeight: 700, fontSize: 18, color: "#1F3B2C" }}>{eur(total)}</span>
      </div>
      <div style={{ marginTop: 14 }}>
        <button className="wz-btn wz-btn-accent" disabled={chosen.length === 0} onClick={create}>Rechnung erstellen</button>
      </div>
    </Sheet>
  );
}

function InvoicePreview({ invoice, settings, customers, onBack }) {
  const customer = customers.find((c) => c.id === invoice.customerId);
  return (
    <div className="wz-invwrap">
      <div className="wz-invbar wz-noprint">
        <button className="wz-back" onClick={onBack}><ChevronLeft size={18} /> Zurück</button>
        <button className="wz-print" onClick={() => window.print()}><Printer size={16} /> Als PDF speichern</button>
      </div>
      <div id="wz-print" className="wz-paper">
        <div style={{ display: "flex", justifyContent: "space-between", alignItems: "flex-start", marginBottom: 26 }}>
          <div>
            <h4>{settings.firma || settings.inhaber}</h4>
            <p style={{ margin: "4px 0 0", fontSize: 13, whiteSpace: "pre-line", color: "#444" }}>{settings.adresse}</p>
          </div>
          <div style={{ textAlign: "right", fontSize: 13 }}>
            <p className="wz-num" style={{ margin: 0, fontWeight: 700, fontSize: 15 }}>Rechnung {invoice.number}</p>
            <p style={{ margin: "3px 0 0", color: "#444" }}>Datum: {fmtDate(invoice.createdAt)}</p>
          </div>
        </div>
        <div style={{ marginBottom: 22 }}>
          <p style={{ fontSize: 11, textTransform: "uppercase", color: "#888", margin: "0 0 3px" }}>Rechnung an</p>
          <p style={{ margin: 0, fontWeight: 600 }}>{customer?.name}</p>
          <p style={{ margin: "2px 0 0", fontSize: 13, whiteSpace: "pre-line", color: "#444" }}>{customer?.address}</p>
        </div>
        <table>
          <thead>
            <tr><th>Datum</th><th>Leistung</th><th className="wz-r">Std.</th><th className="wz-r">Satz</th><th className="wz-r">Betrag</th></tr>
          </thead>
          <tbody>
            {invoice.lines.map((l, i) => (
              <React.Fragment key={i}>
                <tr>
                  <td>{fmtDate(l.date)}</td><td>{l.description || "Arbeitszeit"}</td>
                  <td className="wz-r wz-num">{l.hoursOwn.toFixed(2)}</td>
                  <td className="wz-r wz-num">{eur(l.rateOwn)}</td>
                  <td className="wz-r wz-num">{eur(l.hoursOwn * l.rateOwn)}</td>
                </tr>
                {l.hoursHelper > 0 &&
                  <tr style={{ color: "#666" }}>
                    <td></td><td style={{ fontStyle: "italic" }}>davon Helfer ({settings.sohnName})</td>
                    <td className="wz-r wz-num">{l.hoursHelper.toFixed(2)}</td>
                    <td className="wz-r wz-num">{eur(l.rateHelper)}</td>
                    <td className="wz-r wz-num">{eur(l.hoursHelper * l.rateHelper)}</td>
                  </tr>}
              </React.Fragment>
            ))}
          </tbody>
        </table>
        <div style={{ display: "flex", justifyContent: "flex-end", marginBottom: 26 }}>
          <div style={{ width: 220 }}>
            <div style={{ display: "flex", justifyContent: "space-between", fontWeight: 700, fontSize: 15, borderTop: "2px solid #1a1a1a", paddingTop: 8 }}>
              <span>Gesamtbetrag</span><span className="wz-num">{eur(invoice.total)}</span>
            </div>
          </div>
        </div>
        <p style={{ fontSize: 12, color: "#666", marginBottom: 22 }}>
          Gemäß § 19 UStG wird keine Umsatzsteuer berechnet und ausgewiesen (Kleinunternehmerregelung).
        </p>
        <div style={{ fontSize: 12, color: "#666", borderTop: "1px solid #eee", paddingTop: 14, lineHeight: 1.6 }}>
          {settings.steuernummer && <p style={{ margin: 0 }}>Steuernummer: {settings.steuernummer}</p>}
          {settings.iban && <p style={{ margin: 0 }}>IBAN: {settings.iban}</p>}
          {settings.inhaber && <p style={{ margin: 0 }}>{settings.inhaber}</p>}
        </div>
      </div>
    </div>
  );
}

/* ---------- Einstellungen ---------- */
function EinstellungenTab({ data, save, showToast }) {
  const [s, setS] = useState(data.settings);
  const set = (k) => (e) => setS({ ...s, [k]: e.target.value });
  const persist = () => {
    save({ ...data, settings: { ...s, rateOwn: Number(s.rateOwn), rateHelper: Number(s.rateHelper), nextInvoiceNo: Number(s.nextInvoiceNo) } });
    showToast("Einstellungen gespeichert");
  };
  return (
    <div style={{ marginBottom: 24 }}>
      <h2 className="wz-h2" style={{ marginBottom: 14 }}>Einstellungen</h2>
      <Field label="Firmenname (auf Rechnung)"><input className="wz-input" value={s.firma} onChange={set("firma")} /></Field>
      <Field label="Name (Inhaber)"><input className="wz-input" value={s.inhaber} onChange={set("inhaber")} /></Field>
      <Field label="Adresse"><textarea className="wz-input" value={s.adresse} onChange={set("adresse")} placeholder={"Straße Nr.\nPLZ Ort"} /></Field>
      <Field label="Steuernummer"><input className="wz-input" value={s.steuernummer} onChange={set("steuernummer")} /></Field>
      <Field label="IBAN"><input className="wz-input" value={s.iban} onChange={set("iban")} /></Field>
      <div className="wz-grid2">
        <Field label="Dein Stundenlohn"><div className="wz-suffix"><input type="number" step="0.5" className="wz-input" value={s.rateOwn} onChange={set("rateOwn")} /><span>€/h</span></div></Field>
        <Field label="Lohn Helfer"><div className="wz-suffix"><input type="number" step="0.5" className="wz-input" value={s.rateHelper} onChange={set("rateHelper")} /><span>€/h</span></div></Field>
      </div>
      <Field label="Name des Helfers"><input className="wz-input" value={s.sohnName} onChange={set("sohnName")} placeholder="z.B. Sohn / Max" /></Field>
      <div className="wz-grid2">
        <Field label="Rechnungspräfix"><input className="wz-input" value={s.invoicePrefix} onChange={set("invoicePrefix")} /></Field>
        <Field label="Nächste Nr."><input type="number" className="wz-input" value={s.nextInvoiceNo} onChange={set("nextInvoiceNo")} /></Field>
      </div>
      <button className="wz-btn wz-btn-primary" onClick={persist}>Speichern</button>
      <p className="wz-tip">
        Tipp fürs iPhone: über „Teilen" → „Zum Home-Bildschirm" bekommst du ein eigenes App-Symbol —
        ganz ohne App Store oder Entwickler-Account.
      </p>
    </div>
  );
}

/* ---------- shared ---------- */
function Field({ label, children }) {
  return <div className="wz-field"><label className="wz-label">{label}</label>{children}</div>;
}
function Sheet({ title, children, onClose }) {
  return (
    <div className="wz-sheet-bg" onClick={onClose}>
      <div className="wz-sheet" onClick={(e) => e.stopPropagation()}>
        <div className="wz-sheet-head"><h3>{title}</h3><button className="wz-x" onClick={onClose}><X size={20} /></button></div>
        {children}
      </div>
    </div>
  );
}
function EmptyState({ icon: Icon, title, text }) {
  return (
    <div className="wz-empty">
      <Icon size={28} color="#9AA396" style={{ margin: "0 auto" }} />
      <p className="wz-empty-title">{title}</p>
      <p className="wz-empty-text">{text}</p>
    </div>
  );
}
