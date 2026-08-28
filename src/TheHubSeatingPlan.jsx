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
const TEAM_ACCESS_CODE = "HUB-ADMIN-4297";
/* ================================================================ */

const VENUE = [
  {
    area: "Main Restaurant",
    images: ["/images/main-restaurant.jpg"],
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
    images: ["/images/Nursery.png", "/images/Concourse.png"],
    tables: [
      { id: "nursery-1", name: "Nursery Table One", capacity: 8, dims: "2800mm × 800mm", mm: [2800, 800], layout: { top: 4, bottom: 4, left: 0, right: 0 } },
      { id: "nursery-2", name: "Nursery Table Two", capacity: 8, dims: "2800mm × 800mm", mm: [2800, 800], layout: { top: 4, bottom: 4, left: 0, right: 0 } },
      { id: "concourse", name: "Concourse Table", capacity: 14, dims: "3600mm × 600mm", mm: [3600, 600], layout: { top: 7, bottom: 7, left: 0, right: 0 } },
    ],
  },
  {
    area: "Mezzanine",
    images: ["/images/Mezzanine.png"],
    tables: [
      { id: "mezz-1", name: "Mezz Table One", note: "Benches", bench: true, capacity: 8, dims: "2200mm × 790mm", mm: [2200, 790], layout: { top: 4, bottom: 4, left: 0, right: 0 } },
      { id: "mezz-2", name: "Mezz Table Two", note: "Benches", bench: true, capacity: 8, dims: "2200mm × 790mm", mm: [2200, 790], layout: { top: 4, bottom: 4, left: 0, right: 0 } },
      { id: "mezz-3", name: "Mezz Table Three", capacity: 4, dims: "1400mm × 700mm", mm: [1400, 700], layout: { top: 2, bottom: 2, left: 0, right: 0 } },
      { id: "mezz-4", name: "Mezz Table Four", capacity: 4, dims: "1400mm × 700mm", mm: [1400, 700], layout: { top: 2, bottom: 2, left: 0, right: 0 } },
      { id: "mezz-5", name: "Mezz Table Five", note: "High table", capacity: 4, dims: "1200mm × 600mm", mm: [1200, 600], layout: { top: 2, bottom: 2, left: 0, right: 0 } },
      { id: "mezz-6", name: "Mezz Table Six", note: "High table", capacity: 4, dims: "1200mm × 600mm", mm: [1200, 600], layout: { top: 2, bottom: 2, left: 0, right: 0 } },
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

/* ---------- one table (original layout – used in team detail view) ---------- */
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
    const half = Math.ceil(table.capacity / 2);
    const left = take(half), right = take(table.capacity - half);
    body = (
      <div className="bench-wrap">
        <div className="bench-col">{left.map((i) => <Seat key={i} seatNo={i + 1} value={names[i]} readOnly={readOnly} onChange={(v) => setName(i, v)} />)}</div>
        <div className="bench-bar" style={{ width: `${barDepth}px` }}><span>{table.name.toUpperCase()}</span></div>
        <div className="bench-col">{right.map((i) => <Seat key={i} seatNo={i + 1} value={names[i]} readOnly={readOnly} onChange={(v) => setName(i, v)} />)}</div>
      </div>
    );
  } else {
    const { top, bottom, left, right } = table.layout;
    const topSeats = take(top);
    const leftSeats = take(left);
    const bottomSeats = take(bottom);
    const rightSeats = take(right);
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
            Seats {table.capacity}{table.note ? ` · ${table.note}` : ""}
          </p>
        </div>
        <span className={"count" + (filled === table.capacity ? " full" : "")}>{filled}/{table.capacity}</span>
      </header>
      {body}
    </section>
  );
}

/* ---------- compact table card with pill-style inputs (customer Seat Guests view) ---------- */
function TableCardCompact({ table, names, setName }) {
  const seats = Array.from({ length: table.capacity }).map((_, i) => names[i] || "");
  const filled = seats.filter((n) => n.trim()).length;

  return (
    <div className="tc-compact" id={table.id} aria-label={table.name}>
      <div className="tc-head">
        <div className="tc-title">
          <span className="tc-name">{table.name}</span>
          <span className="tc-meta">Seats {table.capacity}{table.note ? ` · ${table.note}` : ""}</span>
        </div>
        <span className="tc-count">{filled} guests</span>
        <button className="tc-menu-btn" aria-label="Table options">
          <svg viewBox="0 0 24 24" fill="currentColor">
            <circle cx="12" cy="5" r="2" />
            <circle cx="12" cy="12" r="2" />
            <circle cx="12" cy="19" r="2" />
          </svg>
        </button>
      </div>
      <div className="tc-pills" style={{ gridTemplateColumns: `repeat(${table.layout?.top || 4}, 1fr)` }}>
        {seats.map((n, i) => (
          <input
            key={i}
            className={"tc-pill" + (n.trim() ? " filled" : "")}
            type="text"
            value={n}
            maxLength={40}
            placeholder={n.trim() ? "" : `Seat ${i + 1}`}
            onChange={(e) => setName(i, e.target.value)}
            aria-label={`${table.name} seat ${i + 1}`}
          />
        ))}
      </div>
    </div>
  );
}

function DecorativeRule() {
  return (
    <div className="deco-rule">
      <div className="deco-line"></div>
      <svg className="deco-svg" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="1.5" strokeLinecap="round" strokeLinejoin="round">
        <path d="M7 21a5 5 0 0 1-4.7-3.6c-.3-1 .2-2.1 1.2-2.4a5 5 0 0 1 5.9 3.2" />
        <path d="M11 16a5 5 0 0 1-4.7-3.6c-.3-1 .2-2.1 1.2-2.4a5 5 0 0 1 5.9 3.2" />
        <path d="M15 11a5 5 0 0 1-4.7-3.6c-.3-1 .2-2.1 1.2-2.4a5 5 0 0 1 5.9 3.2" />
        <path d="M5 22L20 4" />
        <path d="M10 20a5 5 0 0 0 5-3.2c.4-.9-.1-2-1.1-2.4A5 5 0 0 0 8 17.6" />
        <path d="M14 15a5 5 0 0 0 5-3.2c.4-.9-.1-2-1.1-2.4A5 5 0 0 0 12 12.6" />
        <path d="M18 10a5 5 0 0 0 5-3.2c.4-.9-.1-2-1.1-2.4A5 5 0 0 0 16 7.6" />
      </svg>
      <div className="deco-line"></div>
    </div>
  );
}

/* ---------- area icon SVG ---------- */
function AreaIcon({ area }) {
  if (area.includes("Restaurant")) {
    return (
      <svg className="area-icon-svg" viewBox="0 0 24 24" fill="currentColor" style={{ color: '#A87A5B' }}>
        <path d="M11,9H9V2H7v7H5V2H3v7c0,2.12,1.66,3.84,3.75,3.97V22h2.5v-9.03C11.34,12.84,13,11.12,13,9V2h-2V9z M16,6v8h2.5v8H21V2 C18.24,2,16,4.24,16,6z" />
      </svg>
    );
  }
  if (area.includes("Nursery") || area.includes("Concourse")) {
    return (
      <svg className="area-icon-svg" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="1.5" strokeLinecap="round" strokeLinejoin="round" style={{ color: '#43604D' }}>
        <polygon points="12 2 20 8 17 21 7 21 4 8" />
        <line x1="6.1" y1="17" x2="17.9" y2="17" />
        <line x1="12" y1="17" x2="12" y2="2" />
        <line x1="12" y1="14" x2="17" y2="10" />
        <line x1="12" y1="11.5" x2="7" y2="7.5" />
        <line x1="12" y1="8.5" x2="16" y2="5.3" />
        <line x1="12" y1="5.5" x2="10.1" y2="4" />
      </svg>
    );
  }
  // Mezzanine
  return (
    <div className="area-icon-svg" style={{
      backgroundColor: '#A87A5B',
      WebkitMask: 'url(/images/mezzanine-icon.png) no-repeat center / contain',
      mask: 'url(/images/mezzanine-icon.png) no-repeat center / contain'
    }} />
  );
}

/* ---------- image lightbox ---------- */
function Lightbox({ images, alt, onClose }) {
  const [index, setIndex] = useState(0);

  useEffect(() => {
    const handleKey = (e) => {
      if (e.key === "Escape") onClose();
      if (e.key === "ArrowRight") setIndex((i) => (i + 1) % images.length);
      if (e.key === "ArrowLeft") setIndex((i) => (i - 1 + images.length) % images.length);
    };
    document.addEventListener("keydown", handleKey);
    return () => document.removeEventListener("keydown", handleKey);
  }, [onClose, images.length]);

  return (
    <div className="lightbox-overlay" onClick={onClose}>
      <div className="lightbox-inner" onClick={(e) => e.stopPropagation()}>
        <button className="lightbox-close" onClick={onClose} aria-label="Close image">×</button>
        <img src={images[index]} alt={`${alt} - ${index + 1}`} className="lightbox-img" />
        {images.length > 1 && (
          <>
            <button className="lightbox-prev" onClick={(e) => { e.stopPropagation(); setIndex((i) => (i - 1 + images.length) % images.length); }}>
              <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round"><polyline points="15 18 9 12 15 6"></polyline></svg>
            </button>
            <button className="lightbox-next" onClick={(e) => { e.stopPropagation(); setIndex((i) => (i + 1) % images.length); }}>
              <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round"><polyline points="9 18 15 12 9 6"></polyline></svg>
            </button>
            <div className="lightbox-dots">
              {images.map((_, i) => (
                <span key={i} className={`lightbox-dot ${i === index ? "active" : ""}`} onClick={(e) => { e.stopPropagation(); setIndex(i); }} />
              ))}
            </div>
          </>
        )}
      </div>
    </div>
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
  /* Mezzanine (upper level) — now a full-width band below Nursery & Conservatory,
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
  { label: "MAIN RESTAURANT", x: 16, y: 46, w: 968, h: 330 },
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
        return (
          <g key={t.id} className="fp-table" onClick={() => onJump && onJump(t.id)}
            tabIndex={0} role="button" aria-label={`${t.name}, ${filled} of ${t.capacity} seats named. Activate to jump to this table.`}
            onKeyDown={(e) => (e.key === "Enter" || e.key === " ") && onJump && onJump(t.id)}>
            <rect x={m.x} y={m.y} width={w} height={h} rx="8"
              fill={full ? "#43604D" : "#EAF0EB"} stroke={filled > 0 && !full ? "#A3823F" : "#CFDCD2"}
              strokeWidth={filled > 0 && !full ? 2 : 1} />
            {m.vertical ? (
              <g transform={`rotate(-90 ${cx} ${cy})`}>
                <text x={cx} y={cy + 2} textAnchor="middle" className={"fp-name" + (full ? " on" : "")}>
                  {short} · {filled}/{t.capacity}
                </text>
              </g>
            ) : (
              <>
                <text x={cx} y={cy - 3} textAnchor="middle" className={"fp-name" + (full ? " on" : "")}>{short}</text>
                <text x={cx} y={cy + 10} textAnchor="middle" className={"fp-count" + (full ? " on" : "")}>{filled}/{t.capacity}</text>
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
  const [tab, setTab] = useState("seats"); // seats | floorplan
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
  const [lightboxImg, setLightboxImg] = useState(null);
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
      .shell { max-width:1280px; margin:0 auto; padding:0 24px 80px; }
      .masthead { text-align:center; padding:52px 16px 8px; }
      .masthead .eyebrow { letter-spacing:.32em; font-size:11px; font-weight:600; color:var(--brass); text-transform:uppercase; }
      .masthead h1 { font-family:'Cormorant Garamond',serif; font-weight:600; font-size:clamp(34px,5vw,52px); margin:10px 0 6px; }
      .masthead p.sub { color:var(--muted); max-width:560px; margin:0 auto; font-size:15px; line-height:1.55; }
      .deco-rule { display:flex; align-items:center; justify-content:center; gap:16px; margin:28px 0 20px; }
      .deco-line { width:52px; height:1px; background:#D1B89B; }
      .deco-svg { width:28px; height:28px; color:#A87A5B; }
      .rule { width:64px; height:1px; background:var(--brass); margin:22px auto; }

      /* ---------- view toggle ---------- */
      .view-toggle { display:flex; justify-content:center; margin:0 auto 6px; gap:0; background:var(--paper); border:1px solid var(--line); border-radius:99px; padding:4px; max-width:440px; }
      .view-toggle button { flex:1; padding:11px 20px; border:none; border-radius:99px; font-family:inherit; font-size:14px; font-weight:600; cursor:pointer; display:flex; align-items:center; justify-content:center; gap:8px; transition:all .25s ease; background:transparent; color:var(--muted); white-space:nowrap; outline:none; }
      .view-toggle button:focus-visible { outline:2px solid var(--green); outline-offset:2px; }
      .view-toggle button.active { background:var(--green); color:#fff; box-shadow:0 2px 8px rgba(67,96,77,.25); }
      .view-toggle button:not(.active):hover { color:var(--ink); background:rgba(0,0,0,.03); }
      .view-toggle .toggle-icon { width:16px; height:16px; flex-shrink:0; }

      .progress-band { position:sticky; top:0; z-index:20; background:var(--paper); border-bottom:1px solid var(--line);
        padding:10px 20px; display:flex; gap:14px; align-items:center; justify-content:center; flex-wrap:wrap; }
      .progress-band .bar { flex:1; max-width:420px; height:6px; background:var(--line); border-radius:99px; overflow:hidden; }
      .progress-band .fill { height:100%; background:var(--green); border-radius:99px; transition:width .3s; }
      .progress-band .label { font-size:13px; color:var(--muted); }
      .progress-band .savenote { font-size:12px; color:var(--brass); min-width:86px; text-align:right; }

      /* ---------- guest summary bar ---------- */
      .guest-summary { background:var(--panel); border:1px solid var(--line); border-radius:14px; padding:16px 22px; margin:20px 0 8px;
        display:flex; align-items:center; gap:0; flex-wrap:wrap; }
      .guest-summary .gs-section { flex:1; min-width:140px; padding:4px 16px; }
      .guest-summary .gs-section:not(:last-child) { border-right:1px solid var(--line); }
      .guest-summary .gs-label { font-size:11px; font-weight:600; letter-spacing:.1em; text-transform:uppercase; color:var(--muted); margin-bottom:4px; }
      .guest-summary .gs-value { font-size:14px; font-weight:500; }
      .guest-summary .gs-value strong { font-weight:600; }
      .guest-summary .legend-dots { display:flex; gap:16px; flex-wrap:wrap; align-items:center; }
      .guest-summary .legend-dot { display:inline-flex; align-items:center; gap:5px; font-size:13px; color:var(--muted); }
      .guest-summary .legend-dot i { width:10px; height:10px; border-radius:50%; display:inline-block; }
      @media (max-width:600px) {
        .guest-summary { flex-direction:column; gap:10px; }
        .guest-summary .gs-section { border-right:none !important; border-bottom:1px solid var(--line); padding:8px 0; width:100%; }
        .guest-summary .gs-section:last-child { border-bottom:none; }
      }

      /* ---------- top nav ---------- */
      .top-nav { display:flex; justify-content:space-between; align-items:center; padding:16px 24px; border-bottom:1px solid var(--line); background:#fff; position:sticky; top:0; z-index:100; }
      .nav-left { display:flex; align-items:center; gap:16px; }
      .nav-logo { display:flex; align-items:center; gap:8px; font-family:'Cormorant Garamond',serif; color:#A87A5B; }
      .logo-lw { font-size:24px; font-style:italic; font-weight:600; }
      .logo-text { font-size:11px; font-weight:600; letter-spacing:.15em; font-family:'Archivo',sans-serif; color:var(--ink); }
      .nav-divider { width:1px; height:24px; background:var(--line); }
      .nav-couple { color:var(--muted); font-size:14px; font-weight:500; }
      .nav-right { display:flex; align-items:center; gap:12px; }
      .nav-btn { display:flex; align-items:center; gap:6px; background:transparent; border:none; cursor:pointer; font-family:inherit; font-size:13px; font-weight:500; color:var(--ink); padding:6px 12px; border-radius:6px; }
      .nav-btn:hover { background:rgba(0,0,0,.03); }
      .nav-btn svg { width:16px; height:16px; }
      .nav-save-btn { background:var(--green); color:#fff; padding:8px 16px; border:none; border-radius:6px; font-family:inherit; font-size:13px; font-weight:600; cursor:pointer; }
      .nav-save-btn:hover { background:#375140; }
      
      @media (max-width: 600px) {
        .top-nav { flex-direction: column; gap: 16px; padding: 16px; align-items: stretch; text-align: center; }
        .nav-left { flex-direction: column; gap: 8px; }
        .nav-divider { display: none; }
        .nav-right { justify-content: center; flex-wrap: wrap; }
      }

      .details { background:var(--panel); border:1px solid var(--line); border-radius:14px; padding:22px; margin:28px 0;
        display:grid; grid-template-columns:repeat(auto-fit,minmax(220px,1fr)); gap:16px; }
      .field label { display:block; font-size:12px; font-weight:600; letter-spacing:.06em; text-transform:uppercase; color:var(--muted); margin-bottom:6px; }
      .field input { width:100%; padding:10px 12px; border:1px solid var(--line); border-radius:8px; font-size:15px; background:#FDFCFA; font-family:inherit; }
      .field input:focus, .seat-input:focus, .tc-pill:focus { outline:2px solid var(--green); outline-offset:1px; border-color:var(--green); }

      /* ---------- area layout with header & grid ---------- */
      .area-card { background:#ffffff; border:1px solid var(--line); border-radius:12px; margin:24px 0; overflow:hidden; }
      .area-header { display:flex; align-items:center; justify-content:space-between; gap:16px; padding:10px 24px; }
      .area-divider { height:1px; background:var(--line); margin:0; border:none; }
      .area-header-left { display:flex; align-items:center; gap:14px; }
      .area-icon { width:48px; height:48px; border-radius:50%; background:#FAF5F0; display:flex; align-items:center; justify-content:center; flex-shrink:0; }
      .area-icon-svg { width:26px; height:26px; color:var(--green); }
      .area-header-info h2 { font-family:'Cormorant Garamond',serif; font-size:26px; font-weight:600; margin:0; line-height:1.2; }
      .area-header-info .area-stats { color:var(--muted); font-size:13px; margin:2px 0 0; }
      .area-thumb { position:relative; width:380px; height:112px; border-radius:8px; overflow:hidden; flex-shrink:0; cursor:pointer; border:1px solid var(--line); }
      .area-thumb img { width:100%; height:100%; object-fit:cover; display:block; transition:transform .3s ease; }
      .area-thumb:hover img { transform:scale(1.05); }
      .area-thumb-overlay { position:absolute; right:8px; bottom:8px; background:rgba(255,255,255,0.9);
        backdrop-filter:blur(2px); color:var(--ink); padding:6px 10px; border-radius:6px;
        display:flex; align-items:center; justify-content:center; gap:6px;
        font-size:11px; font-weight:600; box-shadow:0 2px 4px rgba(0,0,0,.15); transition:transform .2s ease; }
      .area-thumb:hover .area-thumb-overlay { transform:scale(1.05); }
      .area-thumb-overlay svg { width:13px; height:13px; }

      /* ---------- compact table grid ---------- */
      .table-grid { display:grid; grid-template-columns:repeat(auto-fill, minmax(400px, 1fr)); gap:16px; padding:24px; }
      .tc-compact { background:#ffffff; border:1px solid var(--line); border-radius:8px; padding:16px; display:flex; flex-direction:column; gap:16px; }
      .tc-head { display:flex; justify-content:space-between; align-items:center; margin-bottom:0; }
      .tc-title { display:flex; flex-direction:column; align-items:flex-start; gap:2px; }
      .tc-name { font-family:'Cormorant Garamond',serif; font-size:18px; font-weight:600; }
      .tc-meta { font-size:13px; color:var(--muted); }
      .tc-count { font-size:12px; font-weight:600; color:var(--muted); margin-left: auto; }
      .tc-menu-btn { background:transparent; border:none; color:var(--muted); cursor:pointer; display:flex; align-items:center; justify-content:center; padding:4px; border-radius:50%; margin-left: 8px; }
      .tc-menu-btn:hover { background:var(--paper); }
      .tc-menu-btn svg { width:16px; height:16px; }
      .tc-pills { display:grid; grid-template-columns:repeat(4, 1fr); gap:8px; }
      .tc-pill { width:100%; min-width:0; padding:8px 6px; border-radius:4px; border:1px solid transparent; background:var(--green-soft); font-family:inherit; font-size:12px;
        text-align:center; transition:all .2s ease; cursor:text; color:var(--ink); box-sizing:border-box; }
      .tc-pill::placeholder { color:var(--muted); opacity:0.7; }
      .tc-pill.filled { background:#e8efe9; border-color:#d5e0d8; color:var(--green); font-weight:500; }
      .tc-pill:focus { outline:none; border-color:var(--green); background:#ffffff; box-shadow:0 0 0 2px rgba(67,96,77,.15); }

      /* ---------- old table card styles (used in team detail) ---------- */
      .area-head { margin:42px 0 6px; display:flex; align-items:baseline; gap:14px; flex-wrap:wrap; }
      .area-head h2 { font-family:'Cormorant Garamond',serif; font-size:30px; font-weight:600; margin:0; }
      .area-head .blurb { color:var(--muted); font-size:14px; }
      .card { background:var(--panel); border:1px solid var(--line); border-radius:14px; padding:18px 18px 22px; margin:0; }
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
      .end-col .seat { flex:0 1 auto; width:100px; }
      .table-bar, .bench-bar { background:var(--green-soft); border:1px solid #CFDCD2; color:var(--green);
        border-radius:8px; display:flex; align-items:center; justify-content:center;
        font-size:11px; letter-spacing:.18em; font-weight:600; }
      .table-bar { height:var(--th, 34px); padding:0 10px; text-align:center; }
      .bench-wrap { display:flex; gap:10px; justify-content:center; }
      .bench-col { display:flex; flex-direction:column; gap:8px; }
      .bench-bar { writing-mode:vertical-rl; padding:14px 0; }
      .seat { display:flex; align-items:center; gap:6px; background:var(--paper); border:1px solid var(--line);
        border-radius:8px; padding:5px 8px; }
      .bench-col .seat { flex:1 1 0%; min-width:90px; max-width:140px; width:100%; }
      .seat-no { font-size:10px; font-weight:600; color:var(--brass); min-width:14px; text-align:center; }
      .seat-input { border:none; background:transparent; width:100%; font-size:13.5px; font-family:inherit; padding:4px 2px; }
      .seat-name { font-size:13.5px; padding:4px 2px; } .seat-name.empty { color:var(--muted); }
      .submit-zone { text-align:center; margin-top:46px; }
      .btn { font-family:inherit; font-size:15px; font-weight:600; border-radius:99px; padding:13px 34px; cursor:pointer; border:1px solid var(--green); white-space:nowrap; }
      .btn-primary { background:var(--green); color:#fff; }
      .btn-primary:hover { background:#375140; }
      .btn-primary:disabled { opacity:.55; cursor:default; }
      .btn-ghost { background:transparent; color:var(--green); margin-left:10px; }
      .btn-danger { background:#8C3B2E; border-color:#8C3B2E; color:#fff; margin-left:10px; }
      .btn-small { padding:8px 18px; font-size:13px; }
      .modal-overlay { position:fixed; top:0; left:0; right:0; bottom:0; background:rgba(0,0,0,0.4); backdrop-filter:blur(4px); display:flex; align-items:center; justify-content:center; z-index:1000; padding:24px; animation:fade-in 0.2s ease-out; }
      .modal { background:var(--panel); border-radius:16px; box-shadow:0 10px 40px rgba(0,0,0,0.1); padding:32px; max-width:400px; width:100%; text-align:center; animation:modal-pop 0.3s cubic-bezier(0.175, 0.885, 0.32, 1.275); }
      @keyframes fade-in { 0% { opacity:0; } 100% { opacity:1; } }
      @keyframes modal-pop { 0% { opacity:0; transform:scale(0.95); } 100% { opacity:1; transform:scale(1); } }
      .modal-icon { width:48px; height:48px; border-radius:50%; background:#F7E9E5; color:#8C3B2E; display:flex; align-items:center; justify-content:center; margin:0 auto 16px; }
      .modal-icon svg { width:24px; height:24px; }
      .modal-text { font-size:15px; color:var(--text); line-height:1.5; margin-bottom:24px; }
      .modal-actions { display:flex; gap:12px; justify-content:center; }
      .footer { text-align:center; margin-top:60px; color:var(--muted); font-size:12.5px; }
      .footer button { background:none; border:none; color:var(--muted); text-decoration:underline; cursor:pointer; font-size:12.5px; font-family:inherit; }
      .panel-narrow { max-width:520px; margin:60px auto; background:var(--panel); border:1px solid var(--line); border-radius:14px; padding:34px; text-align:center; }
      .panel-narrow h2 { font-family:'Cormorant Garamond',serif; font-size:30px; margin:0 0 10px; }
      .panel-narrow p { color:var(--muted); font-size:14.5px; line-height:1.6; }
      .sub-list { margin-top:20px; }
      .sub-row { background:var(--panel); border:1px solid var(--line); border-radius:12px; padding:16px 18px; margin:10px 0;
        display:flex; justify-content:space-between; align-items:center; gap:14px; flex-wrap:wrap; }
      .sub-row .who { font-weight:600; font-size:16px; margin-bottom:4px; }
      .sub-row .when { color:var(--muted); font-size:13.5px; line-height:1.5; overflow-wrap:anywhere; }
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
      .requests-header { display:flex; align-items:center; gap:8px; font-size:12px; font-weight:600; letter-spacing:.06em; text-transform:uppercase; color:var(--muted); margin-bottom:12px; }
      .requests-header svg { width:16px; height:16px; }
      .requests-box { width:100%; padding:10px 12px; border:1px solid var(--line); border-radius:8px; font-size:14.5px;
        background:#FDFCFA; font-family:inherit; line-height:1.5; resize:vertical; min-height:110px; }
      .requests-box:focus { outline:2px solid var(--green); outline-offset:1px; border-color:var(--green); }
      .requests-hint { color:var(--muted); font-size:12px; margin:6px 0 0; }
      .requests-view { background:var(--panel); border:1px solid var(--line); border-radius:14px; padding:18px 22px; margin:14px 0;
        font-size:14.5px; line-height:1.6; white-space:pre-wrap; }
      .requests-view .rq-label { display:block; font-size:12px; font-weight:600; letter-spacing:.06em; text-transform:uppercase; color:var(--muted); margin-bottom:6px; }

      /* ---------- lightbox ---------- */
      .lightbox-overlay { position:fixed; inset:0; z-index:1000; background:rgba(0,0,0,.7); display:flex; align-items:center; justify-content:center; padding:24px; backdrop-filter:blur(4px); animation:lb-in .2s ease; }
      .lightbox-inner { position:relative; max-width:900px; max-height:90vh; border-radius:14px; overflow:hidden; background:var(--panel); box-shadow:0 20px 60px rgba(0,0,0,.4); }
      .lightbox-img { display:block; width:100%; height:auto; max-height:85vh; object-fit:contain; }
      .lightbox-close { position:absolute; top:12px; right:12px; z-index:10; width:36px; height:36px; border-radius:50%; border:none; background:rgba(0,0,0,.5); color:#fff; font-size:22px; cursor:pointer; display:flex; align-items:center; justify-content:center; transition:background .2s; }
      .lightbox-close:hover { background:rgba(0,0,0,.75); }
      .lightbox-prev, .lightbox-next { position:absolute; top:50%; transform:translateY(-50%); width:48px; height:48px; border-radius:50%; border:none; background:rgba(0,0,0,.5); color:#fff; cursor:pointer; display:flex; align-items:center; justify-content:center; transition:background .2s; }
      .lightbox-prev:hover, .lightbox-next:hover { background:rgba(0,0,0,.75); }
      .lightbox-prev { left:12px; }
      .lightbox-next { right:12px; }
      .lightbox-prev svg, .lightbox-next svg { width:24px; height:24px; }
      .lightbox-dots { position:absolute; bottom:16px; left:0; right:0; display:flex; justify-content:center; gap:8px; }
      .lightbox-dot { width:8px; height:8px; border-radius:50%; background:rgba(255,255,255,.5); cursor:pointer; transition:background .2s; }
      .lightbox-dot.active { background:#fff; transform:scale(1.2); }
      @keyframes lb-in { from { opacity:0; transform:scale(.95); } to { opacity:1; transform:scale(1); } }

      /* ---------- auto-save note ---------- */
      .autosave-note { text-align:center; padding:12px 0 0; display:flex; align-items:center; justify-content:center; gap:6px; color:var(--muted); font-size:12.5px; }
      .autosave-note svg { width:14px; height:14px; fill:currentColor; }

      .btn-submit-icon { display:inline-flex; align-items:center; justify-content:center; gap:8px; }
      .btn-submit-icon svg { width:16px; height:16px; }

      @media (max-width:640px){
        .table-wrap{flex-direction:column;} .end-col{flex-direction:row;} .bench-wrap{flex-wrap:nowrap;}
        .table-mid{max-width:100%;}
        .seat-row{flex-wrap:wrap;}
        .seat-row .seat{flex:1 1 110px; min-width:110px;}
        .table-bar{height:auto; min-height:34px; padding:8px 10px;}
        .area-header { flex-direction:column; gap:12px; }
        .area-thumb { width:100%; height:120px; }
        .table-grid { grid-template-columns:1fr; }
        .view-toggle { max-width:100%; }
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
        {error && (
          <div className="modal-overlay">
            <div className="modal">
              <div className="modal-icon">
                <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round">
                  <circle cx="12" cy="12" r="10" /><line x1="12" y1="8" x2="12" y2="12" /><line x1="12" y1="16" x2="12.01" y2="16" />
                </svg>
              </div>
              <div className="modal-text">{error}</div>
              <div className="modal-actions">
                <button className="btn btn-primary" onClick={() => setError("")}>OK, got it</button>
              </div>
            </div>
          </div>
        )}
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
            <div style={{ display: "flex", justifyContent: "center", gap: 12, marginTop: 16, flexWrap: "wrap" }}>
              <button className="btn btn-primary" onClick={tryCode}>Open dashboard</button>
              <button className="btn btn-ghost" onClick={() => { setError(""); setView("customer"); }}>Cancel</button>
            </div>
          </div>
        </div>
      </div>
    );
  }

  /* ---------- TEAM DASHBOARD ---------- */
  if (view === "team" || view === "teamDetail") {
    return (
      <div className="hub-root">{css}
        {lightboxImg && <Lightbox images={lightboxImg.images} alt={lightboxImg.alt} onClose={() => setLightboxImg(null)} />}
        <div className="shell">
          <div className="masthead">
            <span className="eyebrow">The Hub · Bookings Team</span>
            <h1>Submitted seating plans</h1>
            <div className="rule" />
          </div>
          {view === "team" && (
            <>
              <div style={{ display: "flex", justifyContent: "center", gap: 12 }}>
                <button className="btn btn-ghost btn-small" onClick={loadSubmissions}>{loadingSubs ? "Refreshing…" : "Refresh"}</button>
                <button className="btn btn-ghost btn-small" onClick={() => setView("customer")}>Exit dashboard</button>
              </div>
              <div className="sub-list">
                {submissions.length === 0 && (
                  <div className="panel-narrow"><p>No seating plans have been submitted yet. When a couple presses \u201cSend seating plan to The Hub\u201d, their plan appears here.</p></div>
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
                    <div style={{ display: "flex", gap: 10, alignItems: "center", flexWrap: "wrap", marginTop: 12 }}>
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
              <div style={{ display: "flex", justifyContent: "center", gap: 12, marginBottom: 8 }}>
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

              <div className="view-toggle" style={{ marginTop: 24, marginBottom: 24 }}>
                <button className={tab === "seats" ? "active" : ""} onClick={() => setTab("seats")}>
                  <svg className="toggle-icon" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="1.5" strokeLinecap="round" strokeLinejoin="round">
                    <path d="M16 21v-2a4 4 0 0 0-4-4H6a4 4 0 0 0-4 4v2" /><circle cx="9" cy="7" r="4" /><path d="M22 21v-2a4 4 0 0 0-3-3.87" /><path d="M16 3.13a4 4 0 0 1 0 7.75" />
                  </svg>
                  Seat guests
                </button>
                <button className={tab === "floorplan" ? "active" : ""} onClick={() => setTab("floorplan")}>
                  <svg className="toggle-icon" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round">
                    <polygon points="3 6 9 3 15 6 21 3 21 18 15 21 9 18 3 21" />
                    <line x1="9" y1="3" x2="9" y2="21" />
                    <line x1="15" y1="3" x2="15" y2="21" />
                  </svg>
                  Floor plan overview
                </button>
              </div>

              {tab === "floorplan" && (
                <div className="floor-card">
                  <h3>Floor plan overview</h3>
                  <p className="meta">Green tables are fully seated. Tap a table to jump to its guest list.</p>
                  <FloorPlan
                    names={detail.names}
                    onJump={(id) => {
                      setTab("seats");
                      setTimeout(() => {
                        const el = document.getElementById(id);
                        if (el) el.scrollIntoView({ behavior: "smooth", block: "center" });
                      }, 50);
                    }}
                  />
                </div>
              )}
              {tab === "seats" && (
                <>
                  {VENUE.map((areaDef) => {
                    const areaFilled = areaDef.tables.reduce((s, t) => s + (detail.names[t.id] || []).filter((n) => n.trim()).length, 0);
                    return (
                      <div key={areaDef.area} className="area-card">
                        <div className="area-header">
                          <div className="area-header-left">
                            <div className="area-icon"><AreaIcon area={areaDef.area} /></div>
                            <div className="area-header-info">
                              <h2>{areaDef.area}</h2>
                              <p className="area-stats">{areaDef.tables.length} tables · {areaFilled} guests assigned</p>
                            </div>
                          </div>
                          {areaDef.images && areaDef.images.length > 0 && (
                            <div className="area-thumb" onClick={() => setLightboxImg({ images: areaDef.images, alt: areaDef.area })}>
                              <img src={areaDef.images[0]} alt={areaDef.area} />
                              <div className="area-thumb-overlay">
                                <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round">
                                  <path d="M18 13v6a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2V8a2 2 0 0 1 2-2h6" />
                                  <polyline points="15 3 21 3 21 9" />
                                  <line x1="10" y1="14" x2="21" y2="3" />
                                </svg>
                                View photo
                              </div>
                            </div>
                          )}
                        </div>
                        <hr className="area-divider" />
                        <div className="table-grid">
                          {areaDef.tables.map((t) => (
                            <TableCard key={t.id} table={t} names={detail.names[t.id] || Array(t.capacity).fill("")} readOnly setName={() => { }} />
                          ))}
                        </div>
                      </div>
                    );
                  })}
                </>
              )}
            </>
          )}
        </div>
      </div>
    );
  }

  /* ---------- CUSTOMER VIEW ---------- */
  const assignedByArea = (areaDef) => areaDef.tables.reduce((s, t) => s + (names[t.id] || []).filter((n) => n.trim()).length, 0);

  return (
    <div className="hub-root">{css}
      {error && (
        <div className="modal-overlay">
          <div className="modal">
            <div className="modal-icon">
              <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round">
                <circle cx="12" cy="12" r="10" /><line x1="12" y1="8" x2="12" y2="12" /><line x1="12" y1="16" x2="12.01" y2="16" />
              </svg>
            </div>
            <div className="modal-text">{error}</div>
            <div className="modal-actions">
              {confirmPartial ? (
                <>
                  <button className="btn btn-ghost" onClick={() => { setConfirmPartial(false); setError(""); }}>Keep editing</button>
                  <button className="btn btn-primary" onClick={() => { setError(""); submit(); }}>Send anyway</button>
                </>
              ) : (
                <button className="btn btn-primary" onClick={() => setError("")}>OK, got it</button>
              )}
            </div>
          </div>
        </div>
      )}
      {lightboxImg && <Lightbox images={lightboxImg.images} alt={lightboxImg.alt} onClose={() => setLightboxImg(null)} />}

      {/* ---------- Top Nav ---------- */}
      <header className="top-nav">
        <div className="nav-left">
          <div className="nav-logo">
            <svg viewBox="0 0 332.4 332.4" style={{ height: 32, width: 'auto' }} fill="none" stroke="currentColor" strokeMiterlimit="10" strokeWidth="5">
              <circle cx="166.2" cy="166.2" r="160" />
              <circle cx="166.2" cy="166.99" r="145" />
              <line x1="41.75" y1="250.07" x2="290.52" y2="250.07" />
              <line x1="16.84" y1="166.99" x2="315.56" y2="166.99" />
              <line x1="42.06" y1="83.91" x2="290.79" y2="83.91" />
            </svg>
            <span className="logo-lw" style={{ fontStyle: "normal" }}>Wedding Seating Planner</span>
          </div>
          {coupleName && (
            <>
              <div className="nav-divider" />
              <div className="nav-couple">Wedding of {coupleName}</div>
            </>
          )}
        </div>
        <div className="nav-right">
          <button className="nav-save-btn" onClick={submit} disabled={submitting}>
            {submitting ? "Sending..." : "Save plan"}
          </button>
        </div>
      </header>

      <div className="masthead">
        <span className="eyebrow">The Hub · Weddings</span>
        <h1>Your seating plan</h1>
        <p className="sub">
          Add a guest's name to each seat below. Leave any seats you don't need blank. Your work is saved when you submit the form.
        </p>
        <DecorativeRule />
      </div>

      {/* ---------- Toggle ---------- */}
      <div className="view-toggle">
        <button className={tab === "seats" ? "active" : ""} onClick={() => setTab("seats")}>
          <svg className="toggle-icon" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="1.5" strokeLinecap="round" strokeLinejoin="round">
            <path d="M16 21v-2a4 4 0 0 0-4-4H6a4 4 0 0 0-4 4v2" /><circle cx="9" cy="7" r="4" /><path d="M22 21v-2a4 4 0 0 0-3-3.87" /><path d="M16 3.13a4 4 0 0 1 0 7.75" />
          </svg>
          Seat guests
        </button>
        <button className={tab === "floorplan" ? "active" : ""} onClick={() => setTab("floorplan")}>
          <svg className="toggle-icon" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round">
            <polygon points="3 6 9 3 15 6 21 3 21 18 15 21 9 18 3 21" />
            <line x1="9" y1="3" x2="9" y2="21" />
            <line x1="15" y1="3" x2="15" y2="21" />
          </svg>
          Floor plan overview
        </button>
      </div>

      <div className="progress-band">
        <span className="label">{filledCount} of {TOTAL_SEATS} seats named</span>
        <div className="bar"><div className="fill" style={{ width: `${(filledCount / TOTAL_SEATS) * 100}%` }} /></div>
        <span className="savenote">{saveState}</span>
      </div>

      <div className="shell">
        {/* ---------- Guest summary bar ---------- */}
        <div className="guest-summary">
          <div className="gs-section">
            <div className="gs-label">Guest Summary</div>
            <div className="gs-value">
              <strong>{TOTAL_SEATS}</strong> Capacity · <strong>{filledCount}</strong> Assigned · <strong>{TOTAL_SEATS - filledCount}</strong> Unassigned
            </div>
          </div>
          <div className="gs-section">
            <div className="gs-label">Tables</div>
            <div className="gs-value"><strong>{ALL_TABLES.length}</strong> Tables</div>
          </div>
          <div className="gs-section">
            <div className="gs-label">Legend</div>
            <div className="legend-dots">
              <span className="legend-dot"><i style={{ background: "var(--green)" }} /> Assigned</span>
              <span className="legend-dot"><i style={{ background: "var(--muted)" }} /> Unassigned</span>
            </div>
          </div>
        </div>

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

        {/* ============ SEAT GUESTS TAB ============ */}
        {tab === "seats" && (
          <>
            {VENUE.map((areaDef) => {
              const areaFilled = assignedByArea(areaDef);
              return (
                <div key={areaDef.area} className="area-card">
                  <div className="area-header">
                    <div className="area-header-left">
                      <div className="area-icon"><AreaIcon area={areaDef.area} /></div>
                      <div className="area-header-info">
                        <h2>{areaDef.area}</h2>
                        <p className="area-stats">{areaDef.tables.length} tables · {areaFilled} guests assigned</p>
                      </div>
                    </div>
                    {areaDef.images && areaDef.images.length > 0 && (
                      <div className="area-thumb" onClick={() => setLightboxImg({ images: areaDef.images, alt: areaDef.area })}>
                        <img src={areaDef.images[0]} alt={areaDef.area} />
                        <div className="area-thumb-overlay">
                          <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round">
                            <path d="M18 13v6a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2V8a2 2 0 0 1 2-2h6" />
                            <polyline points="15 3 21 3 21 9" />
                            <line x1="10" y1="14" x2="21" y2="3" />
                          </svg>
                          View photo
                        </div>
                      </div>
                    )}
                  </div>
                  <hr className="area-divider" />
                  <div className="table-grid">
                    {areaDef.tables.map((t) => (
                      <TableCard key={t.id} table={t} names={names[t.id]}
                        setName={(idx, v) => setName(t.id, idx, v)} />
                    ))}
                  </div>
                </div>
              );
            })}
          </>
        )}

        {/* ============ FLOOR PLAN OVERVIEW TAB ============ */}
        {tab === "floorplan" && (
          <div className="floor-card" style={{ marginTop: 20 }}>
            <h3>Floor plan overview</h3>
            <p className="meta">
              Use this view to remind yourself of the full layout. All {ALL_TABLES.length} tables are drawn to scale
              in their real positions. Tap any table to switch to Seat Guests and jump to that table.
            </p>
            <FloorPlan
              names={names}
              onJump={(id) => {
                setTab("seats");
                setTimeout(() => {
                  const el = document.getElementById(id);
                  if (el) el.scrollIntoView({ behavior: "smooth", block: "center" });
                }, 100);
              }}
            />
          </div>
        )}

        <div className="requests-card">
          <div className="field">
            <div className="requests-header">
              <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round">
                <path d="M12 20h9" /><path d="M16.5 3.5a2.121 2.121 0 0 1 3 3L7 19l-4 1 1-4L16.5 3.5z" />
              </svg>
              <label htmlFor="requests" style={{ marginBottom: 0 }}>Notes & special arrangements</label>
            </div>
            <textarea id="requests" className="requests-box" rows={5} maxLength={2000}
              placeholder="Add any notes about seating preferences, dietary requirements, or special arrangements…"
              value={specialRequests} onChange={(e) => setSpecialRequests(e.target.value)} />
            <p className="requests-hint">{specialRequests.length}/2000 · Optional — sent to the bookings team with your plan.</p>
          </div>
        </div>
        <div className="submit-zone">
          <button className="btn btn-primary btn-submit-icon" disabled={submitting} onClick={submit}>
            {submitting ? "Sending..." : "Review & share seating plan"}
            {!submitting && (
              <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round">
                <line x1="22" y1="2" x2="11" y2="13" /><polygon points="22 2 15 22 11 13 2 9 22 2" />
              </svg>
            )}
          </button>
          <div className="autosave-note">
            <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round">
              <circle cx="12" cy="12" r="10" /><line x1="12" y1="16" x2="12" y2="12" /><line x1="12" y1="8" x2="12.01" y2="8" />
            </svg>
            You'll be able to download a copy for yourselves after sending.
          </div>
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
