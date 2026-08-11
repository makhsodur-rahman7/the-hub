import React, { useState, useEffect, useRef, useCallback } from "react";
import * as XLSX from "xlsx";
import { supabase } from "./utils/supabase.js";

/* ================================================================
   THE HUB — DIGITAL WEDDING SEATING PLAN
   ----------------------------------------------------------------
   Customer view : fills in a guest name per seat, submits the plan.
   Team view     : reviews submissions, downloads Excel spreadsheets.
   Team access code (change as needed):
*/
const TEAM_ACCESS_CODE = "HUB-TEAM";
/* ================================================================ */

const VENUE = [
  {
    area: "Main Hall",
    blurb: "Five feature tables at the heart of the venue.",
    tables: [
      { id: "table-1", name: "Table One", capacity: 9, dims: "2800mm × 700mm", mm: [2800, 700], layout: { top: 4, bottom: 4, left: 1, right: 0 } },
      { id: "table-2", name: "Table Two", capacity: 9, dims: "2800mm × 700mm", mm: [2800, 700], layout: { top: 4, bottom: 4, left: 0, right: 1 } },
      { id: "table-3", name: "Table Three", capacity: 10, dims: "2700mm × 900mm", mm: [2700, 900], layout: { top: 5, bottom: 5, left: 0, right: 0 } },
      { id: "table-4", name: "Table Four", capacity: 6, dims: "1800mm × 900mm", mm: [1800, 900], layout: { top: 3, bottom: 3, left: 0, right: 0 } },
      { id: "table-5", name: "Table Five", capacity: 9, dims: "2600mm × 660mm", mm: [2600, 660], layout: { top: 4, bottom: 4, left: 0, right: 1 } },
    ],
  },
  {
    area: "Nursery & Concourse",
    blurb: "Beside the open kitchen — watch the chefs at work.",
    hasKitchen: true,
    tables: [
      { id: "nursery-1", name: "Nursery Table One", capacity: 8, dims: "2800mm × 800mm", mm: [2800, 800], layout: { top: 4, bottom: 4, left: 0, right: 0 } },
      { id: "nursery-2", name: "Nursery Table Two", capacity: 8, dims: "2800mm × 800mm", mm: [2800, 800], layout: { top: 4, bottom: 4, left: 0, right: 0 } },
      { id: "concourse", name: "Concourse Table", capacity: 14, dims: "3600mm × 600mm", mm: [3600, 600], layout: { top: 7, bottom: 7, left: 0, right: 0 } },
    ],
  },
  {
    area: "Mezzanine",
    blurb: "An intimate upper level of benches and high tables.",
    tables: [
      { id: "mezz-1", name: "Mezz Table One", capacity: 8, dims: "2200mm × 790mm", mm: [2200, 790], note: "Benches", bench: true },
      { id: "mezz-2", name: "Mezz Table Two", capacity: 8, dims: "2200mm × 790mm", mm: [2200, 790], note: "Benches", bench: true },
      { id: "mezz-3", name: "Mezz Table Three", capacity: 4, dims: "1400mm × 700mm", mm: [1400, 700], layout: { top: 2, bottom: 2, left: 0, right: 0 } },
      { id: "mezz-4", name: "Mezz Table Four", capacity: 4, dims: "1400mm × 700mm", mm: [1400, 700], layout: { top: 2, bottom: 2, left: 0, right: 0 } },
      { id: "mezz-5", name: "Mezz Table Five", capacity: 4, dims: "1200mm × 600mm", mm: [1200, 600], note: "High table", layout: { top: 2, bottom: 2, left: 0, right: 0 } },
      { id: "mezz-6", name: "Mezz Table Six", capacity: 4, dims: "1200mm × 600mm", mm: [1200, 600], note: "High table", layout: { top: 2, bottom: 2, left: 0, right: 0 } },
    ],
  },
];

const ALL_TABLES = VENUE.flatMap((a) => a.tables.map((t) => ({ ...t, area: a.area })));
const TOTAL_SEATS = ALL_TABLES.reduce((s, t) => s + t.capacity, 0);

const emptyNames = () => {
  const o = {};
  ALL_TABLES.forEach((t) => { o[t.id] = Array(t.capacity).fill(""); });
  return o;
};

/* ---------- storage helpers (fail-safe) ---------- */
// Fallback for window.storage since we're in a standard React app
const storageGet = async (key, shared = false) => {
  try { 
    if (window.storage && window.storage.get) {
      const r = await window.storage.get(key, shared); return r ? JSON.parse(r.value) : null; 
    } else {
      const r = localStorage.getItem(key);
      return r ? JSON.parse(r) : null;
    }
  }
  catch { return null; }
};
const storageSet = async (key, value, shared = false) => {
  try { 
    if (window.storage && window.storage.set) {
      await window.storage.set(key, JSON.stringify(value), shared); return true; 
    } else {
      localStorage.setItem(key, JSON.stringify(value));
      return true;
    }
  }
  catch { return false; }
};

/* ---------- Excel export ---------- */
function buildWorkbook(sub) {
  const wb = XLSX.utils.book_new();
  const summary = [
    ["THE HUB — WEDDING SEATING PLAN"],
    [],
    ["Couple", sub.coupleName],
    ["Wedding date", sub.eventDate || "—"],
    ["Contact email", sub.email || "—"],
    ["Submitted", new Date(sub.submittedAt).toLocaleString("en-GB")],
    ["Guests seated", `${sub.filledCount} of ${TOTAL_SEATS}`],
    [],
    ["Special requests / further information"],
    [sub.specialRequests && sub.specialRequests.trim() ? sub.specialRequests : "(none given)"],
    [],
    ["Table", "Seats used", "Capacity"],
    ...ALL_TABLES.map((t) => [
      t.name,
      (sub.names[t.id] || []).filter((n) => n.trim()).length,
      t.capacity,
    ]),
  ];
  const ws1 = XLSX.utils.aoa_to_sheet(summary);
  ws1["!cols"] = [{ wch: 24 }, { wch: 28 }, { wch: 10 }];
  XLSX.utils.book_append_sheet(wb, ws1, "Summary");

  const rows = [["Area", "Table", "Table size", "Seat", "Guest name"]];
  ALL_TABLES.forEach((t) => {
    (sub.names[t.id] || []).forEach((n, i) => {
      rows.push([t.area, t.name, t.dims, i + 1, n.trim() || "(empty)"]);
    });
  });
  const ws2 = XLSX.utils.aoa_to_sheet(rows);
  ws2["!cols"] = [{ wch: 22 }, { wch: 22 }, { wch: 16 }, { wch: 6 }, { wch: 30 }];
  XLSX.utils.book_append_sheet(wb, ws2, "Seating plan");
  return wb;
}
function downloadExcel(sub) {
  const safe = (sub.coupleName || "seating-plan").replace(/[^a-z0-9]+/gi, "-").toLowerCase();
  XLSX.writeFile(buildWorkbook(sub), `the-hub-${safe}.xlsx`);
}

/* ---------- seat input ---------- */
function Seat({ value, onChange, seatNo, readOnly }) {
  return (
    <div className="seat">
      <span className="seat-no">{seatNo}</span>
      {readOnly ? (
        <span className={"seat-name" + (value ? "" : " empty")}>{value || "—"}</span>
      ) : (
        <input
          className="seat-input"
          type="text"
          value={value}
          maxLength={40}
          placeholder="Name"
          onChange={(e) => onChange(e.target.value)}
          aria-label={`Seat ${seatNo}`}
        />
      )}
    </div>
  );
}

/* ---------- one table ---------- */
function TableCard({ table, names, setName, readOnly }) {
  const filled = names.filter((n) => n.trim()).length;
  let seatIdx = 0;
  const take = (n) => Array.from({ length: n }, () => seatIdx++);
  /* Table bars mirror real dimensions: lengths at 0.16px/mm, depths at 0.06px/mm
     (depth uses a compressed scale so the form stays usable). */
  const barLength = Math.round(table.mm[0] * 0.16);
  const barDepth = Math.max(34, Math.round(table.mm[1] * 0.06));

  let body;
  if (table.bench) {
    const left = take(4), right = take(4);
    body = (
      <div className="bench-wrap">
        <div className="bench-col">{left.map((i) => <Seat key={i} seatNo={i + 1} value={names[i]} readOnly={readOnly} onChange={(v) => setName(i, v)} />)}</div>
        <div className="bench-bar" style={{ width: `${barDepth}px` }}><span>{table.name.toUpperCase()}</span></div>
        <div className="bench-col">{right.map((i) => <Seat key={i} seatNo={i + 1} value={names[i]} readOnly={readOnly} onChange={(v) => setName(i, v)} />)}</div>
      </div>
    );
  } else {
    const { top, bottom, left, right } = table.layout;
    const topSeats = take(top), bottomSeats = take(bottom), leftSeats = take(left), rightSeats = take(right);
    body = (
      <div className="table-wrap">
        {leftSeats.length > 0 && <div className="end-col">{leftSeats.map((i) => <Seat key={i} seatNo={i + 1} value={names[i]} readOnly={readOnly} onChange={(v) => setName(i, v)} />)}</div>}
        <div className="table-mid" style={{ "--tw": `${barLength}px`, "--th": `${barDepth}px` }}>
          <div className="seat-row">{topSeats.map((i) => <Seat key={i} seatNo={i + 1} value={names[i]} readOnly={readOnly} onChange={(v) => setName(i, v)} />)}</div>
          <div className="table-bar"><span>{table.name.toUpperCase()}</span></div>
          <div className="seat-row">{bottomSeats.map((i) => <Seat key={i} seatNo={i + 1} value={names[i]} readOnly={readOnly} onChange={(v) => setName(i, v)} />)}</div>
        </div>
        {rightSeats.length > 0 && <div className="end-col">{rightSeats.map((i) => <Seat key={i} seatNo={i + 1} value={names[i]} readOnly={readOnly} onChange={(v) => setName(i, v)} />)}</div>}
      </div>
    );
  }

  return (
    <section className="card" id={table.id} aria-label={table.name}>
      <header className="card-head">
        <div>
          <h3>{table.name}</h3>
          <p className="meta">
            Seats {table.capacity} · {table.dims}{table.note ? ` · ${table.note}` : ""}
          </p>
        </div>
        <span className={"count" + (filled === table.capacity ? " full" : "")}>{filled}/{table.capacity}</span>
      </header>
      {body}
    </section>
  );
}

/* ---------- single-view floor plan ---------- */
/* One scale for every table: 0.09px per mm, so all sizes are true to each other. */
const FP_SCALE = 0.09;
const FP_MAP = {
  /* Main Hall — as PDF page 1 */
  "table-1": { x: 60, y: 96 },
  "table-2": { x: 560, y: 96 },
  "table-3": { x: 60, y: 196 },
  "table-4": { x: 560, y: 196 },
  "table-5": { x: 60, y: 302 },
  /* Nursery & Concourse — concourse top edge meets Nursery Two's bottom edge (y 538 + 72 = 610) */
  "nursery-1": { x: 60, y: 446 },
  "nursery-2": { x: 60, y: 538 },
  concourse: { x: 560, y: 610 },
  /* Mezzanine (upper level) — now a full-width band below Nursery & Concourse,
     offset to the right to show its real position relative to the areas beneath it.
     All tables lie horizontally: high tables Five/Six left column, Three/Four middle,
     benches One/Two right. Tables Two, Four and Six share a common top edge (y 855). */
  "mezz-1": { x: 734, y: 764 },
  "mezz-2": { x: 734, y: 855 },
  "mezz-3": { x: 568, y: 764 },
  "mezz-4": { x: 568, y: 855 },
  "mezz-5": { x: 420, y: 764 },
  "mezz-6": { x: 420, y: 855 },
};
const FP_PANELS = [
  { label: "MAIN HALL", x: 16, y: 46, w: 968, h: 330 },
  { label: "NURSERY & CONCOURSE", x: 16, y: 396, w: 968, h: 288 },
  { label: "MEZZANINE", note: "UPPER LEVEL · SITS ABOVE THE CONCOURSE, EXTENDING RIGHT", x: 380, y: 704, w: 604, h: 246 },
];

function FloorPlan({ names, onJump }) {
  return (
    <svg className="floor-svg" viewBox="0 0 1000 966" role="img"
      aria-label="Floor plan of The Hub showing all tables to scale across the Main Hall, Nursery and Concourse, and the Mezzanine upper level">
      {FP_PANELS.map((p) => (
        <g key={p.label}>
          <rect x={p.x} y={p.y} width={p.w} height={p.h} rx="14" fill="#FDFCFA" stroke="#E1DCD0" />
          <text x={p.x + 18} y={p.y + 26} className="fp-area">{p.label}</text>
          {p.note && <text x={p.x + 18} y={p.y + 42} className="fp-note">{p.note}</text>}
        </g>
      ))}
      {/* open kitchen */}
      <g>
        <rect x="560" y="446" width="300" height="120" rx="10" fill="none" stroke="#A3823F" strokeDasharray="6 5" />
        <text x="710" y="512" textAnchor="middle" className="fp-kitchen">OPEN KITCHEN</text>
      </g>
      {/* key — stacked in the space left of the Mezzanine */}
      <g aria-hidden="true">
        <text x="40" y="740" className="fp-area">KEY</text>
        <rect x="40" y="756" width="16" height="16" rx="4" fill="#EAF0EB" stroke="#CFDCD2" />
        <text x="66" y="769" className="fp-key">Not started</text>
        <rect x="40" y="786" width="16" height="16" rx="4" fill="#EAF0EB" stroke="#A3823F" strokeWidth="2" />
        <text x="66" y="799" className="fp-key">In progress</text>
        <rect x="40" y="816" width="16" height="16" rx="4" fill="#43604D" />
        <text x="66" y="829" className="fp-key">Complete</text>
      </g>
      {ALL_TABLES.map((t) => {
        const m = FP_MAP[t.id];
        const w = (m.vertical ? t.mm[1] : t.mm[0]) * FP_SCALE;
        const h = (m.vertical ? t.mm[0] : t.mm[1]) * FP_SCALE;
        const filled = (names[t.id] || []).filter((n) => n.trim()).length;
        const full = filled === t.capacity;
        const cx = m.x + w / 2, cy = m.y + h / 2;
        const short = t.name.replace("Table ", "").replace("Mezz ", "M· ").toUpperCase();
        const dimsShort = `${t.mm[0]}×${t.mm[1]}mm`;
        return (
          <g key={t.id} className="fp-table" onClick={() => onJump && onJump(t.id)}
            tabIndex={0} role="button" aria-label={`${t.name}, ${t.dims}, ${filled} of ${t.capacity} seats named. Activate to jump to this table.`}
            onKeyDown={(e) => (e.key === "Enter" || e.key === " ") && onJump && onJump(t.id)}>
            <rect x={m.x} y={m.y} width={w} height={h} rx="8"
              fill={full ? "#43604D" : "#EAF0EB"} stroke={filled > 0 && !full ? "#A3823F" : "#CFDCD2"}
              strokeWidth={filled > 0 && !full ? 2 : 1} />
            {m.vertical ? (
              <g transform={`rotate(-90 ${cx} ${cy})`}>
                <text x={cx} y={cy - 3} textAnchor="middle" className={"fp-name" + (full ? " on" : "")}>
                  {short} · {filled}/{t.capacity}
                </text>
                <text x={cx} y={cy + 12} textAnchor="middle" className={"fp-dims" + (full ? " on" : "")}>{dimsShort}</text>
              </g>
            ) : (
              <>
                <text x={cx} y={cy - 7} textAnchor="middle" className={"fp-name" + (full ? " on" : "")}>{short}</text>
                <text x={cx} y={cy + 6} textAnchor="middle" className={"fp-count" + (full ? " on" : "")}>{filled}/{t.capacity}</text>
                <text x={cx} y={m.y + h - 6} textAnchor="middle" className={"fp-dims" + (full ? " on" : "")}>{dimsShort}</text>
              </>
            )}
          </g>
        );
      })}
    </svg>
  );
}

/* ---------- main app ---------- */
export default function TheHubSeatingPlan() {
  const [view, setView] = useState("customer"); // customer | success | teamLogin | team | teamDetail
  const [names, setNames] = useState(emptyNames);
  const [coupleName, setCoupleName] = useState("");
  const [eventDate, setEventDate] = useState("");
  const [email, setEmail] = useState("");
  const [specialRequests, setSpecialRequests] = useState("");
  const [error, setError] = useState("");
  const [saveState, setSaveState] = useState("");
  const [submitting, setSubmitting] = useState(false);
  const [lastSubmission, setLastSubmission] = useState(null);
  const [codeInput, setCodeInput] = useState("");
  const [submissions, setSubmissions] = useState([]);
  const [detail, setDetail] = useState(null);
  const [loadingSubs, setLoadingSubs] = useState(false);
  const [confirmPartial, setConfirmPartial] = useState(false);
  const saveTimer = useRef(null);
  const restored = useRef(false);

  const filledCount = ALL_TABLES.reduce(
    (s, t) => s + names[t.id].filter((n) => n.trim()).length, 0
  );

  /* restore draft */
  useEffect(() => {
    (async () => {
      const draft = await storageGet("hub-seating-draft");
      if (draft) {
        if (draft.names) setNames((prev) => ({ ...emptyNames(), ...draft.names }));
        if (draft.coupleName) setCoupleName(draft.coupleName);
        if (draft.eventDate) setEventDate(draft.eventDate);
        if (draft.email) setEmail(draft.email);
        if (draft.specialRequests) setSpecialRequests(draft.specialRequests);
        setSaveState("Draft restored");
      }
      restored.current = true;
    })();
  }, []);

  /* autosave draft */
  useEffect(() => {
    if (!restored.current) return;
    clearTimeout(saveTimer.current);
    saveTimer.current = setTimeout(async () => {
      const ok = await storageSet("hub-seating-draft", { names, coupleName, eventDate, email, specialRequests });
      setSaveState(ok ? "Draft saved" : "");
    }, 900);
    return () => clearTimeout(saveTimer.current);
  }, [names, coupleName, eventDate, email, specialRequests]);

  const setName = useCallback((tableId, idx, value) => {
    setNames((prev) => {
      const next = { ...prev, [tableId]: [...prev[tableId]] };
      next[tableId][idx] = value;
      return next;
    });
  }, []);

  /* submit */
  const submit = async () => {
    setError("");
    if (!coupleName.trim()) { setError("Please add the couple's names so we know whose plan this is."); return; }
    if (!email.trim()) { setError("Please add a contact email so the bookings team can reach you."); return; }
    if (filledCount === 0) { setError("The plan is empty — add your guests before sending it."); return; }
    if (filledCount < TOTAL_SEATS && !confirmPartial) {
      setConfirmPartial(true);
      setError(`${TOTAL_SEATS - filledCount} of the ${TOTAL_SEATS} seats are still blank — blank seats are simply left unused on the day. Press "Send anyway" if the plan is complete, or keep adding names.`);
      return;
    }
    setConfirmPartial(false);
    setSubmitting(true);
    const sub = {
      coupleName: coupleName.trim(),
      eventDate, email: email.trim(),
      specialRequests: specialRequests.trim(),
      names, filledCount,
      submittedAt: new Date().toISOString(),
    };
    
    const { error: insertError } = await supabase
      .from('the-hub-booking')
      .insert({
        email: email.trim(),
        'booking-data': sub,
        created_at: new Date().toISOString()
      });

    sub.stored = !insertError;
    if (insertError) console.error("Error saving to Supabase:", insertError);

    setLastSubmission(sub);
    await storageSet("hub-seating-draft", { names, coupleName, eventDate, email, specialRequests });
    setSubmitting(false);
    setView("success");
    window.scrollTo(0, 0);
  };

  /* team */
  const loadSubmissions = async () => {
    setLoadingSubs(true);
    const { data, error } = await supabase
      .from('the-hub-booking')
      .select('*')
      .order('created_at', { ascending: false });

    if (error) {
      console.error("Error loading submissions", error);
      setSubmissions([]);
    } else {
      const subs = data.map(row => ({
        ...row['booking-data'],
        id: row.id
      }));
      setSubmissions(subs);
    }
    setLoadingSubs(false);
  };
  const [deleteId, setDeleteId] = useState(null);
  const deleteSubmission = async (id) => {
    if (deleteId !== id) { setDeleteId(id); return; } // first click arms, second confirms
    await supabase.from('the-hub-booking').delete().eq('id', id);
    setDeleteId(null);
    loadSubmissions();
  };
  const tryCode = () => {
    if (codeInput.trim().toUpperCase() === TEAM_ACCESS_CODE) {
      setCodeInput(""); setError(""); setView("team"); loadSubmissions();
    } else setError("That access code isn't recognised.");
  };

  const css = (
    <style>{`
      @import url('https://fonts.googleapis.com/css2?family=Cormorant+Garamond:ital,wght@0,500;0,600;1,500&family=Archivo:wght@400;500;600&display=swap');
      .hub-root { --paper:#F6F4EF; --panel:#FFFFFF; --ink:#24221C; --muted:#716C60; --line:#E1DCD0;
        --green:#43604D; --green-soft:#EAF0EB; --brass:#A3823F;
        background:var(--paper); color:var(--ink); min-height:100vh;
        font-family:'Archivo',system-ui,sans-serif; }
      .hub-root * { box-sizing:border-box; }
      .shell { max-width:1060px; margin:0 auto; padding:0 20px 80px; }
      .masthead { text-align:center; padding:52px 16px 8px; }
      .masthead .eyebrow { letter-spacing:.32em; font-size:11px; font-weight:600; color:var(--brass); text-transform:uppercase; }
      .masthead h1 { font-family:'Cormorant Garamond',serif; font-weight:600; font-size:clamp(34px,5vw,52px); margin:10px 0 6px; }
      .masthead p.sub { color:var(--muted); max-width:560px; margin:0 auto; font-size:15px; line-height:1.55; }
      .rule { width:64px; height:1px; background:var(--brass); margin:22px auto; }
      .progress-band { position:sticky; top:0; z-index:20; background:var(--paper); border-bottom:1px solid var(--line);
        padding:10px 20px; display:flex; gap:14px; align-items:center; justify-content:center; flex-wrap:wrap; }
      .progress-band .bar { flex:1; max-width:420px; height:6px; background:var(--line); border-radius:99px; overflow:hidden; }
      .progress-band .fill { height:100%; background:var(--green); border-radius:99px; transition:width .3s; }
      .progress-band .label { font-size:13px; color:var(--muted); }
      .progress-band .savenote { font-size:12px; color:var(--brass); min-width:86px; text-align:right; }
      .details { background:var(--panel); border:1px solid var(--line); border-radius:14px; padding:22px; margin:28px 0;
        display:grid; grid-template-columns:repeat(auto-fit,minmax(220px,1fr)); gap:16px; }
      .field label { display:block; font-size:12px; font-weight:600; letter-spacing:.06em; text-transform:uppercase; color:var(--muted); margin-bottom:6px; }
      .field input { width:100%; padding:10px 12px; border:1px solid var(--line); border-radius:8px; font-size:15px; background:#FDFCFA; font-family:inherit; }
      .field input:focus, .seat-input:focus { outline:2px solid var(--green); outline-offset:1px; border-color:var(--green); }
      .area-head { margin:42px 0 6px; display:flex; align-items:baseline; gap:14px; flex-wrap:wrap; }
      .area-head h2 { font-family:'Cormorant Garamond',serif; font-size:30px; font-weight:600; margin:0; }
      .area-head .blurb { color:var(--muted); font-size:14px; }
      .kitchen { border:1px dashed var(--brass); color:var(--brass); border-radius:10px; padding:10px 18px;
        display:inline-block; margin:10px 0 4px; font-size:12px; letter-spacing:.22em; font-weight:600; }
      .card { background:var(--panel); border:1px solid var(--line); border-radius:14px; padding:18px 18px 22px; margin:18px 0; }
      .card-head { display:flex; justify-content:space-between; align-items:flex-start; gap:12px; margin-bottom:14px; }
      .card-head h3 { font-family:'Cormorant Garamond',serif; font-size:22px; font-weight:600; margin:0; }
      .card-head .meta { color:var(--muted); font-size:13px; margin:3px 0 0; }
      .count { font-size:13px; font-weight:600; color:var(--muted); background:var(--paper); border:1px solid var(--line);
        border-radius:99px; padding:4px 12px; white-space:nowrap; }
      .count.full { color:#fff; background:var(--green); border-color:var(--green); }
      .table-wrap { display:flex; align-items:center; justify-content:center; gap:10px; }
      .table-mid { flex:0 1 auto; width:100%; max-width:var(--tw, 100%); min-width:0; }
      .seat-row { display:flex; flex-wrap:nowrap; gap:8px; margin:8px 0; }
      .seat-row .seat { flex:1 1 0; min-width:0; max-width:none; }
      .end-col { display:flex; flex-direction:column; gap:8px; }
      .end-col .seat { flex:0 0 auto; width:130px; }
      .table-bar, .bench-bar { background:var(--green-soft); border:1px solid #CFDCD2; color:var(--green);
        border-radius:8px; display:flex; align-items:center; justify-content:center;
        font-size:11px; letter-spacing:.18em; font-weight:600; }
      .table-bar { height:var(--th, 34px); padding:0 10px; text-align:center; }
      .bench-wrap { display:flex; gap:10px; justify-content:center; }
      .bench-col { display:flex; flex-direction:column; gap:8px; }
      .bench-bar { writing-mode:vertical-rl; padding:14px 0; }
      .seat { display:flex; align-items:center; gap:6px; background:var(--paper); border:1px solid var(--line);
        border-radius:8px; padding:5px 8px; }
      .bench-col .seat { flex:0 0 auto; width:150px; }
      .seat-no { font-size:10px; font-weight:600; color:var(--brass); min-width:14px; text-align:center; }
      .seat-input { border:none; background:transparent; width:100%; font-size:13.5px; font-family:inherit; padding:4px 2px; }
      .seat-name { font-size:13.5px; padding:4px 2px; } .seat-name.empty { color:var(--muted); }
      .submit-zone { text-align:center; margin-top:46px; }
      .btn { font-family:inherit; font-size:15px; font-weight:600; border-radius:99px; padding:13px 34px; cursor:pointer; border:1px solid var(--green); }
      .btn-primary { background:var(--green); color:#fff; }
      .btn-primary:hover { background:#375140; }
      .btn-primary:disabled { opacity:.55; cursor:default; }
      .btn-ghost { background:transparent; color:var(--green); margin-left:10px; }
      .btn-danger { background:#8C3B2E; border-color:#8C3B2E; color:#fff; margin-left:10px; }
      .btn-small { padding:8px 18px; font-size:13px; }
      .error { color:#8C3B2E; background:#F7E9E5; border:1px solid #E4C4BB; border-radius:8px; padding:10px 16px; margin:14px auto; max-width:520px; font-size:14px; }
      .footer { text-align:center; margin-top:60px; color:var(--muted); font-size:12.5px; }
      .footer button { background:none; border:none; color:var(--muted); text-decoration:underline; cursor:pointer; font-size:12.5px; font-family:inherit; }
      .panel-narrow { max-width:520px; margin:60px auto; background:var(--panel); border:1px solid var(--line); border-radius:14px; padding:34px; text-align:center; }
      .panel-narrow h2 { font-family:'Cormorant Garamond',serif; font-size:30px; margin:0 0 10px; }
      .panel-narrow p { color:var(--muted); font-size:14.5px; line-height:1.6; }
      .sub-list { margin-top:20px; }
      .sub-row { background:var(--panel); border:1px solid var(--line); border-radius:12px; padding:16px 18px; margin:10px 0;
        display:flex; justify-content:space-between; align-items:center; gap:14px; flex-wrap:wrap; }
      .sub-row .who { font-weight:600; }
      .sub-row .when { color:var(--muted); font-size:13px; }
      .floor-card { background:var(--panel); border:1px solid var(--line); border-radius:14px; padding:18px 18px 12px; margin:0 0 10px; }
      .floor-card h3 { font-family:'Cormorant Garamond',serif; font-size:22px; font-weight:600; margin:0 0 2px; }
      .floor-card .meta { color:var(--muted); font-size:13px; margin:0 0 10px; }
      .floor-svg { width:100%; height:auto; display:block; }
      .floor-svg .fp-area { font-family:'Archivo',sans-serif; font-size:13px; font-weight:600; letter-spacing:.22em; fill:var(--brass); }
      .floor-svg .fp-kitchen { font-family:'Archivo',sans-serif; font-size:12px; font-weight:600; letter-spacing:.2em; fill:var(--brass); }
      .floor-svg .fp-name { font-family:'Archivo',sans-serif; font-size:12px; font-weight:600; letter-spacing:.08em; fill:var(--green); }
      .floor-svg .fp-count { font-family:'Archivo',sans-serif; font-size:11px; fill:var(--muted); }
      .floor-svg .fp-name.on, .floor-svg .fp-count.on { fill:#fff; }
      .floor-svg .fp-dims { font-family:'Archivo',sans-serif; font-size:8.5px; letter-spacing:.04em; fill:#9A958A; }
      .floor-svg .fp-dims.on { fill:rgba(255,255,255,.75); }
      .floor-svg .fp-note { font-family:'Archivo',sans-serif; font-size:9.5px; letter-spacing:.14em; fill:var(--muted); }
      .floor-svg .fp-key { font-family:'Archivo',sans-serif; font-size:13px; fill:var(--muted); }
      .floor-svg .fp-table { cursor:pointer; }
      .floor-svg .fp-table:hover rect, .floor-svg .fp-table:focus rect { stroke:var(--green); stroke-width:2.5; outline:none; }
      .legend { display:flex; gap:18px; flex-wrap:wrap; justify-content:center; color:var(--muted); font-size:12.5px; padding:8px 0 4px; }
      .legend span { display:inline-flex; align-items:center; gap:6px; }
      .legend i { width:14px; height:14px; border-radius:4px; display:inline-block; }
      .requests-card { background:var(--panel); border:1px solid var(--line); border-radius:14px; padding:22px; margin:34px 0 0; }
      .requests-box { width:100%; padding:10px 12px; border:1px solid var(--line); border-radius:8px; font-size:14.5px;
        background:#FDFCFA; font-family:inherit; line-height:1.5; resize:vertical; min-height:110px; }
      .requests-box:focus { outline:2px solid var(--green); outline-offset:1px; border-color:var(--green); }
      .requests-hint { color:var(--muted); font-size:12px; margin:6px 0 0; }
      .requests-view { background:var(--panel); border:1px solid var(--line); border-radius:14px; padding:18px 22px; margin:14px 0;
        font-size:14.5px; line-height:1.6; white-space:pre-wrap; }
      .requests-view .rq-label { display:block; font-size:12px; font-weight:600; letter-spacing:.06em; text-transform:uppercase; color:var(--muted); margin-bottom:6px; }
      @media (max-width:640px){
        .table-wrap{flex-direction:column;} .end-col{flex-direction:row;} .bench-wrap{flex-wrap:wrap;}
        .table-mid{max-width:100%;}
        .seat-row{flex-wrap:wrap;}
        .seat-row .seat{flex:1 1 110px; min-width:110px;}
        .table-bar{height:auto; min-height:34px; padding:8px 10px;}
      }
      @media (prefers-reduced-motion: reduce){ .progress-band .fill{transition:none;} }
    `}</style>
  );

  /* ---------- SUCCESS ---------- */
  if (view === "success" && lastSubmission) {
    return (
      <div className="hub-root">{css}
        <div className="shell">
          <div className="panel-narrow">
            <div className="masthead" style={{ padding: 0 }}>
              <span className="eyebrow">The Hub</span>
              <h1 style={{ fontSize: 34 }}>Plan received</h1>
            </div>
            <div className="rule" />
            <p>
              Thank you, <strong>{lastSubmission.coupleName}</strong>. Your seating plan
              ({lastSubmission.filledCount} guests) has been sent to The Hub bookings team.
              {!lastSubmission.stored && " We couldn't reach our system just now, so please download the spreadsheet below and email it to the bookings team."}
            </p>
            <p style={{ marginTop: 18 }}>
              <button className="btn btn-primary btn-small" onClick={() => downloadExcel(lastSubmission)}>Download a copy (Excel)</button>
              <button className="btn btn-ghost btn-small" onClick={() => setView("customer")}>Back to the plan</button>
            </p>
          </div>
        </div>
      </div>
    );
  }

  /* ---------- TEAM LOGIN ---------- */
  if (view === "teamLogin") {
    return (
      <div className="hub-root">{css}
        <div className="shell">
          <div className="panel-narrow">
            <span className="eyebrow" style={{ letterSpacing: ".32em", fontSize: 11, fontWeight: 600, color: "var(--brass)", textTransform: "uppercase" }}>The Hub</span>
            <h2>Bookings team access</h2>
            <p>Enter the team access code to review submitted seating plans.</p>
            <div className="field" style={{ margin: "16px 0", textAlign: "left" }}>
              <label htmlFor="code">Access code</label>
              <input id="code" type="password" value={codeInput} onChange={(e) => setCodeInput(e.target.value)}
                onKeyDown={(e) => e.key === "Enter" && tryCode()} />
            </div>
            {error && <div className="error">{error}</div>}
            <button className="btn btn-primary" onClick={tryCode}>Open dashboard</button>
            <button className="btn btn-ghost" onClick={() => { setError(""); setView("customer"); }}>Cancel</button>
          </div>
        </div>
      </div>
    );
  }

  /* ---------- TEAM DASHBOARD ---------- */
  if (view === "team" || view === "teamDetail") {
    return (
      <div className="hub-root">{css}
        <div className="shell">
          <div className="masthead">
            <span className="eyebrow">The Hub · Bookings Team</span>
            <h1>Submitted seating plans</h1>
            <div className="rule" />
          </div>
          {view === "team" && (
            <>
              <div style={{ textAlign: "center" }}>
                <button className="btn btn-ghost btn-small" onClick={loadSubmissions}>{loadingSubs ? "Refreshing…" : "Refresh"}</button>
                <button className="btn btn-ghost btn-small" onClick={() => setView("customer")}>Exit dashboard</button>
              </div>
              <div className="sub-list">
                {submissions.length === 0 && (
                  <div className="panel-narrow"><p>No seating plans have been submitted yet. When a couple presses “Send seating plan to The Hub”, their plan appears here.</p></div>
                )}
                {submissions.map((s) => (
                  <div className="sub-row" key={s.id}>
                    <div>
                      <div className="who">{s.coupleName}</div>
                      <div className="when">
                        {s.eventDate ? `Wedding ${new Date(s.eventDate + "T00:00").toLocaleDateString("en-GB")} · ` : ""}
                        Submitted {new Date(s.submittedAt).toLocaleString("en-GB")} · {s.filledCount}/{TOTAL_SEATS} guests · {s.email}
                        {s.specialRequests ? " · Has special requests" : ""}
                      </div>
                    </div>
                    <div>
                      <button className="btn btn-ghost btn-small" onClick={() => { setDetail(s); setView("teamDetail"); window.scrollTo(0, 0); }}>View plan</button>
                      <button className="btn btn-primary btn-small" onClick={() => downloadExcel(s)}>Download Excel</button>
                      <button className={"btn btn-small " + (deleteId === s.id ? "btn-danger" : "btn-ghost")}
                        onClick={() => deleteSubmission(s.id)}
                        onBlur={() => deleteId === s.id && setDeleteId(null)}>
                        {deleteId === s.id ? "Confirm delete?" : "Delete"}
                      </button>
                    </div>
                  </div>
                ))}
              </div>
            </>
          )}
          {view === "teamDetail" && detail && (
            <>
              <div style={{ textAlign: "center", marginBottom: 8 }}>
                <button className="btn btn-ghost btn-small" onClick={() => setView("team")}>← All submissions</button>
                <button className="btn btn-primary btn-small" onClick={() => downloadExcel(detail)}>Download Excel</button>
              </div>
              <p style={{ textAlign: "center", color: "var(--muted)", fontSize: 14 }}>
                <strong style={{ color: "var(--ink)" }}>{detail.coupleName}</strong> · {detail.filledCount}/{TOTAL_SEATS} guests · {detail.email}
              </p>
              {detail.specialRequests && (
                <div className="requests-view">
                  <span className="rq-label">Special requests / further information</span>
                  {detail.specialRequests}
                </div>
              )}
              <div className="floor-card">
                <h3>Floor plan overview</h3>
                <p className="meta">Green tables are fully seated. Tap a table to jump to its guest list.</p>
                <FloorPlan
                  names={detail.names}
                  onJump={(id) => {
                    const el = document.getElementById(id);
                    if (el) el.scrollIntoView({ behavior: "smooth", block: "center" });
                  }}
                />
              </div>
              {VENUE.map((areaDef) => (
                <div key={areaDef.area}>
                  <div className="area-head"><h2>{areaDef.area}</h2><span className="blurb">{areaDef.blurb}</span></div>
                  {areaDef.hasKitchen && <div className="kitchen">OPEN KITCHEN</div>}
                  {areaDef.tables.map((t) => (
                    <TableCard key={t.id} table={t} names={detail.names[t.id] || Array(t.capacity).fill("")} readOnly setName={() => {}} />
                  ))}
                </div>
              ))}
            </>
          )}
        </div>
      </div>
    );
  }

  /* ---------- CUSTOMER VIEW ---------- */
  return (
    <div className="hub-root">{css}
      <div className="masthead">
        <span className="eyebrow">The Hub · Weddings</span>
        <h1>Your seating plan</h1>
        <p className="sub">
          Add a guest's name to each seat below. The plan mirrors the venue exactly — {TOTAL_SEATS} seats
          across the Main Hall, the Nursery &amp; Concourse and the Mezzanine. Leave any seats you don't
          need blank. Your work saves automatically as you type.
        </p>
        <div className="rule" />
      </div>

      <div className="progress-band">
        <span className="label">{filledCount} of {TOTAL_SEATS} seats named</span>
        <div className="bar"><div className="fill" style={{ width: `${(filledCount / TOTAL_SEATS) * 100}%` }} /></div>
        <span className="savenote">{saveState}</span>
      </div>

      <div className="shell">
        <div className="details">
          <div className="field">
            <label htmlFor="couple">Couple's names *</label>
            <input id="couple" type="text" placeholder="e.g. Amy & Jordan Baker" value={coupleName} onChange={(e) => setCoupleName(e.target.value)} />
          </div>
          <div className="field">
            <label htmlFor="date">Wedding date</label>
            <input id="date" type="date" value={eventDate} onChange={(e) => setEventDate(e.target.value)} />
          </div>
          <div className="field">
            <label htmlFor="email">Contact email *</label>
            <input id="email" type="email" placeholder="you@example.com" value={email} onChange={(e) => setEmail(e.target.value)} />
          </div>
        </div>

        <div className="floor-card">
          <h3>The venue at a glance</h3>
          <p className="meta">
            All {ALL_TABLES.length} tables, drawn to scale in their real positions. Tap any table to jump
            straight to its seats — tables turn green once every seat is named.
          </p>
          <FloorPlan
            names={names}
            onJump={(id) => {
              const el = document.getElementById(id);
              if (el) el.scrollIntoView({ behavior: "smooth", block: "center" });
            }}
          />
        </div>

        {VENUE.map((areaDef) => (
          <div key={areaDef.area}>
            <div className="area-head"><h2>{areaDef.area}</h2><span className="blurb">{areaDef.blurb}</span></div>
            {areaDef.hasKitchen && <div className="kitchen">OPEN KITCHEN</div>}
            {areaDef.tables.map((t) => (
              <TableCard key={t.id} table={t} names={names[t.id]}
                setName={(idx, v) => setName(t.id, idx, v)} />
            ))}
          </div>
        ))}

        {error && <div className="error">{error}</div>}
        <div className="requests-card">
          <div className="field">
            <label htmlFor="requests">Special requests or further information</label>
            <textarea id="requests" className="requests-box" rows={5} maxLength={2000}
              placeholder="e.g. Can we move some tables around?"
              value={specialRequests} onChange={(e) => setSpecialRequests(e.target.value)} />
            <p className="requests-hint">{specialRequests.length}/2000 · Optional — sent to the bookings team with your plan.</p>
          </div>
        </div>
        <div className="submit-zone">
          <button className="btn btn-primary" disabled={submitting} onClick={submit}>
            {submitting ? "Sending…" : confirmPartial ? "Send anyway" : "Send seating plan to The Hub"}
          </button>
          {confirmPartial && (
            <button className="btn btn-ghost" onClick={() => { setConfirmPartial(false); setError(""); }}>
              Keep editing
            </button>
          )}
          <p style={{ color: "var(--muted)", fontSize: 13, marginTop: 12 }}>
            You'll be able to download a copy for yourselves after sending.
          </p>
          <p style={{ color: "var(--muted)", fontSize: 12.5, marginTop: 10, maxWidth: 560, marginLeft: "auto", marginRight: "auto", lineHeight: 1.55 }}>
            <strong>Privacy:</strong> the details you enter here (guest names, your contact details and any
            notes) are stored so The Hub's bookings team can view your plan, and are used only to arrange
            your event. This is a pilot version of the page — please don't include sensitive information
            such as medical or health details. If you'd like your plan removed, contact The Hub bookings team.
          </p>
        </div>

        <div className="footer">
          The Hub · wedding events ·{" "}
          <button onClick={() => { setError(""); setView("teamLogin"); }}>Bookings team access</button>
        </div>
      </div>
    </div>
  );
}
