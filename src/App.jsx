import { useState, useCallback, useRef } from "react";
import * as XLSX from "xlsx";

// ─── Formatters ───────────────────────────────────────────────────────────────
const fmt = (n) =>
  typeof n === "number" && !isNaN(n)
    ? n.toLocaleString("en-IN", { maximumFractionDigits: 2 })
    : n ?? "—";

const fmtCur = (n) =>
  typeof n === "number" && !isNaN(n)
    ? "₹" + n.toLocaleString("en-IN", { maximumFractionDigits: 2, minimumFractionDigits: 2 })
    : "—";

// ─── Parser ───────────────────────────────────────────────────────────────────
function parseSheet(ws) {
  const raw = XLSX.utils.sheet_to_json(ws, { header: 1, defval: null });
  const headerRow = raw.find((r) => r && r[0] === "Sl.No");
  if (!headerRow) return { headers: [], rows: [] };
  const hi = raw.indexOf(headerRow);
  const headers = headerRow.map((h) => (h != null ? String(h).trim() : ""));
  const rows = [];
  for (let i = hi + 1; i < raw.length; i++) {
    const r = raw[i];
    if (!r || r[0] == null || String(r[0]).toLowerCase().includes("grand")) continue;
    const obj = {};
    headers.forEach((h, j) => { obj[h] = r[j]; });
    rows.push(obj);
  }
  return { headers, rows };
}

function sumField(rows, field) {
  return rows.reduce((s, r) => {
    const v = parseFloat(r[field]);
    return s + (isNaN(v) ? 0 : v);
  }, 0);
}

function buildSummary(workbook) {
  const sheets = workbook.SheetNames;
  const find = (kw) => sheets.find((s) => s.toLowerCase().includes(kw.toLowerCase()));
  const b2bName = find("b2b");
  const b2cName = find("b2c sales");
  const hsnB2BName = sheets.find((s) => s.toLowerCase().includes("hsn") && s.toLowerCase().includes("b2b"));
  const hsnB2CName = sheets.find((s) => s.toLowerCase().includes("hsn") && s.toLowerCase().includes("b2c"));

  const b2b = b2bName ? parseSheet(workbook.Sheets[b2bName]) : null;
  const b2c = b2cName ? parseSheet(workbook.Sheets[b2cName]) : null;
  const hsnB2B = hsnB2BName ? parseSheet(workbook.Sheets[hsnB2BName]) : null;
  const hsnB2C = hsnB2CName ? parseSheet(workbook.Sheets[hsnB2CName]) : null;

  const period = (() => {
    const ws = workbook.Sheets[b2bName || b2cName];
    if (!ws) return "";
    const first = XLSX.utils.sheet_to_json(ws, { header: 1, defval: null })[0];
    const cell = first && first.find((c) => c && String(c).toLowerCase().includes("sales reports"));
    return cell ? String(cell).replace(/sales reports/i, "").trim() : "";
  })();

  const TAX_SLABS = [0, 5, 12, 18, 28];
  function byTaxRate(rows) {
    const map = {};
    TAX_SLABS.forEach((s) => { map[s] = { taxableAmt: 0, cgst: 0, sgst: 0, igst: 0, invoices: new Set() }; });
    rows.forEach((r) => {
      const rate = parseFloat(r["RATE"]);
      const slab = TAX_SLABS.find((s) => s === rate) ?? 0;
      const bucket = map[slab] || (map[slab] = { taxableAmt: 0, cgst: 0, sgst: 0, igst: 0, invoices: new Set() });
      bucket.taxableAmt += parseFloat(r["T'BLE AMT"]) || 0;
      bucket.cgst += parseFloat(r["CGST"]) || 0;
      bucket.sgst += parseFloat(r["SGST"]) || 0;
      bucket.igst += parseFloat(r["IGST"]) || 0;
      if (r["INVOICE No"]) bucket.invoices.add(r["INVOICE No"]);
    });
    return Object.entries(map)
      .map(([rate, d]) => ({ rate: Number(rate), ...d, invoices: d.invoices.size }))
      .filter((d) => d.taxableAmt > 0 || d.cgst > 0 || d.sgst > 0);
  }

  const b2bSummary = b2b ? {
    invoices: new Set(b2b.rows.map((r) => r["INVOICE No"])).size,
    recipients: new Set(b2b.rows.map((r) => r["RECEIPIENT"])).size,
    taxableAmt: sumField(b2b.rows, "T'BLE AMT"),
    invoiceValue: sumField(b2b.rows, "INVOICE VALUE"),
    igst: sumField(b2b.rows, "IGST"),
    cgst: sumField(b2b.rows, "CGST"),
    sgst: sumField(b2b.rows, "SGST"),
    cess: sumField(b2b.rows, "CESS"),
    byRate: byTaxRate(b2b.rows),
    rows: b2b.rows,
  } : null;

  const b2cSummary = b2c ? {
    invoices: new Set(b2c.rows.map((r) => r["INVOICE No"])).size,
    taxableAmt: sumField(b2c.rows, "T'BLE AMT"),
    invoiceValue: sumField(b2c.rows, "INVOICE VALUE"),
    igst: sumField(b2c.rows, "IGST"),
    cgst: sumField(b2c.rows, "CGST"),
    sgst: sumField(b2c.rows, "SGST"),
    cess: sumField(b2c.rows, "CESS"),
    byRate: byTaxRate(b2c.rows),
    rows: b2c.rows,
  } : null;

  return { period, b2bSummary, b2cSummary, hsnB2B, hsnB2C, sheetNames: sheets };
}

// ─── Design tokens ────────────────────────────────────────────────────────────
const C = {
  brand: "#1D9E75",
  brandDark: "#0F6E56",
  brandLight: "#E1F5EE",
  bg: "#F7F6F3",
  surface: "#FFFFFF",
  surfaceHover: "#F0EEE8",
  border: "rgba(0,0,0,0.08)",
  borderStrong: "rgba(0,0,0,0.14)",
  text: "#1A1A18",
  textSub: "#5F5E5A",
  textMuted: "#9E9C97",
  blue: "#185FA5",
  blueBg: "#E6F1FB",
  amber: "#854F0B",
  amberBg: "#FAEEDA",
  red: "#A32D2D",
  redBg: "#FCEBEB",
  gray: "#5F5E5A",
  grayBg: "#F1EFE8",
  shadow: "0 1px 3px rgba(0,0,0,0.07), 0 4px 12px rgba(0,0,0,0.05)",
  shadowMd: "0 2px 8px rgba(0,0,0,0.09), 0 8px 24px rgba(0,0,0,0.06)",
};

const RATE_META = {
  0:  { bg: C.grayBg,  text: C.gray,  label: "0%"  },
  5:  { bg: C.brandLight, text: C.brandDark, label: "5%"  },
  12: { bg: C.blueBg,  text: C.blue,  label: "12%" },
  18: { bg: C.amberBg, text: C.amber, label: "18%" },
  28: { bg: C.redBg,   text: C.red,   label: "28%" },
};

// ─── Shared UI ────────────────────────────────────────────────────────────────
const css = String.raw;
const GLOBAL_STYLE = css`
  * { box-sizing: border-box; margin: 0; padding: 0; }
  body { background: ${C.bg}; font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', sans-serif; color: ${C.text}; }
  button { font-family: inherit; cursor: pointer; border: none; background: none; }
  input { font-family: inherit; }
  @keyframes fadeUp { from { opacity:0; transform:translateY(8px); } to { opacity:1; transform:translateY(0); } }
  @keyframes shimmer { 0%,100%{opacity:.4} 50%{opacity:.8} }
  @keyframes spin { to { transform: rotate(360deg); } }
  .fade-up { animation: fadeUp 0.3s ease forwards; }
  tr:hover td { background: ${C.surfaceHover}; transition: background 0.12s; }
`;

function RateBadge({ rate }) {
  const m = RATE_META[rate] || RATE_META[0];
  return (
    <span style={{ background: m.bg, color: m.text, fontSize: 11, fontWeight: 600, padding: "2px 7px", borderRadius: 20, letterSpacing: "0.02em" }}>
      {m.label}
    </span>
  );
}

function StatCard({ label, value, sub, accent, icon }) {
  return (
    <div style={{
      background: C.surface, borderRadius: 14, padding: "16px 18px",
      boxShadow: C.shadow, border: `1px solid ${C.border}`,
      borderLeft: accent ? `3px solid ${accent}` : undefined,
      minWidth: 0,
    }}>
      <div style={{ display: "flex", justifyContent: "space-between", alignItems: "flex-start" }}>
        <p style={{ fontSize: 11, color: C.textMuted, fontWeight: 500, textTransform: "uppercase", letterSpacing: "0.05em", marginBottom: 6 }}>{label}</p>
        {icon && <span style={{ fontSize: 16, opacity: 0.4 }}>{icon}</span>}
      </div>
      <p style={{ fontSize: 19, fontWeight: 600, color: C.text, lineHeight: 1.2 }}>{value}</p>
      {sub && <p style={{ fontSize: 11, color: C.textMuted, marginTop: 4 }}>{sub}</p>}
    </div>
  );
}

function SearchBar({ value, onChange, placeholder }) {
  return (
    <div style={{ position: "relative", flex: 1, minWidth: 160 }}>
      <span style={{ position: "absolute", left: 10, top: "50%", transform: "translateY(-50%)", color: C.textMuted, fontSize: 14, pointerEvents: "none" }}>🔍</span>
      <input
        value={value}
        onChange={(e) => onChange(e.target.value)}
        placeholder={placeholder || "Search…"}
        style={{
          width: "100%", padding: "7px 10px 7px 30px", fontSize: 13,
          border: `1px solid ${C.border}`, borderRadius: 8,
          background: C.surface, color: C.text, outline: "none",
          transition: "border 0.15s",
        }}
        onFocus={(e) => e.target.style.borderColor = C.brand}
        onBlur={(e) => e.target.style.borderColor = C.border}
      />
    </div>
  );
}

function Pagination({ page, pages, total, perPage, onPage }) {
  if (pages <= 1) return null;
  const start = page * perPage + 1;
  const end = Math.min((page + 1) * perPage, total);
  return (
    <div style={{ display: "flex", alignItems: "center", justifyContent: "space-between", marginTop: 12, padding: "8px 0" }}>
      <span style={{ fontSize: 12, color: C.textMuted }}>Showing {start}–{end} of {total}</span>
      <div style={{ display: "flex", gap: 4, alignItems: "center" }}>
        <PagBtn onClick={() => onPage(0)} disabled={page === 0} label="«" />
        <PagBtn onClick={() => onPage(page - 1)} disabled={page === 0} label="‹" />
        {Array.from({ length: Math.min(5, pages) }, (_, i) => {
          const p = Math.max(0, Math.min(pages - 5, page - 2)) + i;
          return (
            <button key={p} onClick={() => onPage(p)} style={{
              width: 28, height: 28, borderRadius: 6, fontSize: 12,
              background: p === page ? C.brand : "transparent",
              color: p === page ? "#fff" : C.textSub,
              fontWeight: p === page ? 600 : 400,
              transition: "all 0.12s",
            }}>{p + 1}</button>
          );
        })}
        <PagBtn onClick={() => onPage(page + 1)} disabled={page === pages - 1} label="›" />
        <PagBtn onClick={() => onPage(pages - 1)} disabled={page === pages - 1} label="»" />
      </div>
    </div>
  );
}

function PagBtn({ onClick, disabled, label }) {
  return (
    <button onClick={onClick} disabled={disabled} style={{
      width: 28, height: 28, borderRadius: 6, fontSize: 13,
      color: disabled ? C.textMuted : C.textSub,
      background: disabled ? "transparent" : C.surfaceHover,
      cursor: disabled ? "default" : "pointer",
      transition: "all 0.12s",
    }}>{label}</button>
  );
}

function SectionHeader({ title, count, children }) {
  return (
    <div style={{ display: "flex", alignItems: "center", justifyContent: "space-between", flexWrap: "wrap", gap: 10, marginBottom: 16 }}>
      <div style={{ display: "flex", alignItems: "center", gap: 8 }}>
        <h3 style={{ fontSize: 15, fontWeight: 600, color: C.text }}>{title}</h3>
        {count != null && (
          <span style={{ background: C.brandLight, color: C.brandDark, fontSize: 11, fontWeight: 600, padding: "2px 8px", borderRadius: 20 }}>{count}</span>
        )}
      </div>
      <div style={{ display: "flex", gap: 8, alignItems: "center" }}>{children}</div>
    </div>
  );
}

function TaxRateBreakdown({ byRate }) {
  if (!byRate || byRate.length === 0) return null;
  return (
    <div style={{ marginTop: 16, paddingTop: 14, borderTop: `1px dashed ${C.border}` }}>
      <p style={{ fontSize: 11, fontWeight: 600, color: C.textMuted, textTransform: "uppercase", letterSpacing: "0.06em", marginBottom: 10 }}>
        Breakdown by tax rate
      </p>
      <table style={{ width: "100%", borderCollapse: "collapse", fontSize: 12 }}>
        <thead>
          <tr>
            {["Rate", "Taxable Amt", "CGST", "SGST", "IGST", "Total Tax"].map((h) => (
              <th key={h} style={{ padding: "4px 8px", textAlign: h === "Rate" ? "left" : "right", fontWeight: 500, color: C.textMuted, borderBottom: `1px solid ${C.border}`, fontSize: 11 }}>{h}</th>
            ))}
          </tr>
        </thead>
        <tbody>
          {byRate.map((d) => {
            const totalTax = d.cgst + d.sgst + d.igst;
            return (
              <tr key={d.rate}>
                <td style={{ padding: "5px 8px" }}><RateBadge rate={d.rate} /></td>
                <td style={{ padding: "5px 8px", textAlign: "right", fontWeight: 500 }}>{fmtCur(d.taxableAmt)}</td>
                <td style={{ padding: "5px 8px", textAlign: "right", color: d.cgst > 0 ? C.text : C.textMuted }}>{d.cgst > 0 ? fmtCur(d.cgst) : "—"}</td>
                <td style={{ padding: "5px 8px", textAlign: "right", color: d.sgst > 0 ? C.text : C.textMuted }}>{d.sgst > 0 ? fmtCur(d.sgst) : "—"}</td>
                <td style={{ padding: "5px 8px", textAlign: "right", color: d.igst > 0 ? C.text : C.textMuted }}>{d.igst > 0 ? fmtCur(d.igst) : "—"}</td>
                <td style={{ padding: "5px 8px", textAlign: "right", fontWeight: 600, color: C.brandDark }}>{totalTax > 0 ? fmtCur(totalTax) : "—"}</td>
              </tr>
            );
          })}
        </tbody>
      </table>
    </div>
  );
}

// ─── Sections ─────────────────────────────────────────────────────────────────
function OverviewSection({ summary }) {
  const b = summary.b2bSummary;
  const c = summary.b2cSummary;
  const totalTaxable = (b?.taxableAmt || 0) + (c?.taxableAmt || 0);
  const totalTax = (b?.cgst||0)+(b?.sgst||0)+(b?.igst||0)+(c?.cgst||0)+(c?.sgst||0)+(c?.igst||0);

  return (
    <div className="fade-up">
      {/* Top KPI row */}
      <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fit, minmax(150px, 1fr))", gap: 12, marginBottom: 24 }}>
        <StatCard label="Total taxable value" value={fmtCur(totalTaxable)} accent={C.brand} icon="📊" />
        <StatCard label="Total tax liability" value={fmtCur(totalTax)} accent={C.amber} icon="💰" />
        <StatCard label="B2B invoices" value={fmt(b?.invoices)} sub={`${b?.recipients} recipients`} accent={C.blue} icon="🏢" />
        <StatCard label="B2C invoices" value={fmt(c?.invoices)} sub={`${(c?.rows?.length||0)} line items`} accent={C.brandDark} icon="🛒" />
        <StatCard label="Total CGST" value={fmtCur((b?.cgst||0)+(c?.cgst||0))} icon="📋" />
        <StatCard label="Total SGST" value={fmtCur((b?.sgst||0)+(c?.sgst||0))} icon="📋" />
      </div>

      {/* B2B + B2C cards */}
      <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 16 }}>
        {[
          { title: "B2B Sales", subtitle: "Registered recipients", d: b, accent: C.blue, accentBg: C.blueBg },
          { title: "B2C Sales", subtitle: "Unregistered recipients", d: c, accent: C.brand, accentBg: C.brandLight },
        ].map(({ title, subtitle, d, accent, accentBg }) => d && (
          <div key={title} style={{
            background: C.surface, borderRadius: 14, padding: "18px 20px",
            boxShadow: C.shadow, border: `1px solid ${C.border}`,
          }}>
            <div style={{ display: "flex", alignItems: "center", justifyContent: "space-between", marginBottom: 14 }}>
              <div>
                <p style={{ fontSize: 14, fontWeight: 600, color: C.text }}>{title}</p>
                <p style={{ fontSize: 11, color: C.textMuted, marginTop: 2 }}>{subtitle}</p>
              </div>
              <span style={{ background: accentBg, color: accent, fontSize: 11, fontWeight: 600, padding: "3px 10px", borderRadius: 20 }}>
                {d.invoices} invoices
              </span>
            </div>
            {[
              ["Invoice value", fmtCur(d.invoiceValue)],
              ["Taxable amount", fmtCur(d.taxableAmt)],
              ["CGST", fmtCur(d.cgst)],
              ["SGST", fmtCur(d.sgst)],
              ["IGST", fmtCur(d.igst)],
            ].map(([k, v]) => (
              <div key={k} style={{
                display: "flex", justifyContent: "space-between",
                padding: "7px 0", borderBottom: `1px solid ${C.border}`, fontSize: 13,
              }}>
                <span style={{ color: C.textSub }}>{k}</span>
                <span style={{ fontWeight: 600 }}>{v}</span>
              </div>
            ))}
            <TaxRateBreakdown byRate={d.byRate} />
          </div>
        ))}
      </div>

      {/* Footer info */}
      <div style={{ marginTop: 16, padding: "10px 14px", background: C.surface, borderRadius: 10, border: `1px solid ${C.border}`, fontSize: 12, color: C.textMuted, display: "flex", alignItems: "center", gap: 6 }}>
        <span>📄</span>
        <span>Sheets loaded: <strong style={{ color: C.textSub }}>{summary.sheetNames.join("  ·  ")}</strong></span>
      </div>
    </div>
  );
}

function B2BSection({ data }) {
  const [page, setPage] = useState(0);
  const [search, setSearch] = useState("");
  const perPage = 10;

  const filtered = data.rows.filter((r) => {
    const q = search.toLowerCase();
    return !q || [r["RECEIPIENT"], r["INVOICE No"], r["GSTIN/UN OF RECEIPIENT"]].some((v) => v && String(v).toLowerCase().includes(q));
  });
  const pages = Math.ceil(filtered.length / perPage);
  const slice = filtered.slice(page * perPage, (page + 1) * perPage);

  const handleSearch = (v) => { setSearch(v); setPage(0); };

  return (
    <div className="fade-up">
      <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fit, minmax(130px, 1fr))", gap: 12, marginBottom: 20 }}>
        <StatCard label="Total invoices" value={fmt(data.invoices)} accent={C.blue} />
        <StatCard label="Recipients" value={fmt(data.recipients)} accent={C.brand} />
        <StatCard label="Invoice value" value={fmtCur(data.invoiceValue)} accent={C.amber} />
        <StatCard label="Taxable amount" value={fmtCur(data.taxableAmt)} accent={C.brandDark} />
        <StatCard label="CGST" value={fmtCur(data.cgst)} />
        <StatCard label="SGST" value={fmtCur(data.sgst)} />
      </div>

      <SectionHeader title="Invoice list" count={filtered.length}>
        <SearchBar value={search} onChange={handleSearch} placeholder="Search recipient, invoice, GSTIN…" />
      </SectionHeader>

      <div style={{ overflowX: "auto", borderRadius: 12, border: `1px solid ${C.border}`, background: C.surface }}>
        <table style={{ width: "100%", borderCollapse: "collapse", fontSize: 13 }}>
          <thead>
            <tr style={{ background: C.bg }}>
              {["#", "Recipient", "GSTIN", "Invoice No", "Date", "Invoice Value", "Taxable Amt", "CGST", "SGST"].map((h, i) => (
                <th key={h} style={{ padding: "10px 12px", textAlign: i > 4 ? "right" : "left", fontWeight: 600, color: C.textMuted, fontSize: 11, textTransform: "uppercase", letterSpacing: "0.04em", borderBottom: `1px solid ${C.border}`, whiteSpace: "nowrap" }}>{h}</th>
              ))}
            </tr>
          </thead>
          <tbody>
            {slice.length === 0 ? (
              <tr><td colSpan={9} style={{ padding: "32px", textAlign: "center", color: C.textMuted, fontSize: 13 }}>No results found</td></tr>
            ) : slice.map((r, i) => (
              <tr key={i} style={{ borderBottom: `1px solid ${C.border}` }}>
                <td style={{ padding: "10px 12px", color: C.textMuted, fontSize: 12 }}>{page * perPage + i + 1}</td>
                <td style={{ padding: "10px 12px", fontWeight: 500, maxWidth: 140, overflow: "hidden", textOverflow: "ellipsis", whiteSpace: "nowrap" }}>{r["RECEIPIENT"]}</td>
                <td style={{ padding: "10px 12px", fontFamily: "monospace", fontSize: 11, color: C.textSub, maxWidth: 130, overflow: "hidden", textOverflow: "ellipsis" }}>{r["GSTIN/UN OF RECEIPIENT"] || "—"}</td>
                <td style={{ padding: "10px 12px", color: C.blue, fontWeight: 500 }}>{r["INVOICE No"]}</td>
                <td style={{ padding: "10px 12px", color: C.textMuted, whiteSpace: "nowrap" }}>{r["INVOICE DATE"]}</td>
                <td style={{ padding: "10px 12px", textAlign: "right" }}>{fmtCur(parseFloat(r["INVOICE VALUE"]))}</td>
                <td style={{ padding: "10px 12px", textAlign: "right", fontWeight: 500 }}>{fmtCur(parseFloat(r["T'BLE AMT"]))}</td>
                <td style={{ padding: "10px 12px", textAlign: "right", color: C.brandDark }}>{fmtCur(parseFloat(r["CGST"]))}</td>
                <td style={{ padding: "10px 12px", textAlign: "right", color: C.brandDark }}>{fmtCur(parseFloat(r["SGST"]))}</td>
              </tr>
            ))}
          </tbody>
        </table>
      </div>
      <Pagination page={page} pages={pages} total={filtered.length} perPage={perPage} onPage={setPage} />
    </div>
  );
}

function B2CSection({ data }) {
  const [page, setPage] = useState(0);
  const [search, setSearch] = useState("");
  const [rateFilter, setRateFilter] = useState("all");
  const perPage = 15;

  const allRates = [...new Set(data.rows.map((r) => parseFloat(r["RATE"])).filter((v) => !isNaN(v)))].sort((a, b) => a - b);

  const filtered = data.rows.filter((r) => {
    const q = search.toLowerCase();
    const matchQ = !q || [r["RECEIPIENT"], r["INVOICE No"]].some((v) => v && String(v).toLowerCase().includes(q));
    const matchRate = rateFilter === "all" || String(parseFloat(r["RATE"])) === rateFilter;
    return matchQ && matchRate;
  });
  const pages = Math.ceil(filtered.length / perPage);
  const slice = filtered.slice(page * perPage, (page + 1) * perPage);

  const handleSearch = (v) => { setSearch(v); setPage(0); };
  const handleRate = (v) => { setRateFilter(v); setPage(0); };

  return (
    <div className="fade-up">
      <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fit, minmax(130px, 1fr))", gap: 12, marginBottom: 20 }}>
        <StatCard label="Line items" value={fmt(data.rows.length)} accent={C.brand} />
        <StatCard label="Unique invoices" value={fmt(data.invoices)} accent={C.blue} />
        <StatCard label="Invoice value" value={fmtCur(data.invoiceValue)} accent={C.amber} />
        <StatCard label="Taxable amount" value={fmtCur(data.taxableAmt)} accent={C.brandDark} />
        <StatCard label="CGST" value={fmtCur(data.cgst)} />
        <StatCard label="SGST" value={fmtCur(data.sgst)} />
      </div>

      <SectionHeader title="Sales list" count={filtered.length}>
        {/* Rate filter pills */}
        <div style={{ display: "flex", gap: 6, flexWrap: "wrap" }}>
          {["all", ...allRates.map(String)].map((r) => (
            <button key={r} onClick={() => handleRate(r)} style={{
              padding: "4px 12px", borderRadius: 20, fontSize: 12, fontWeight: 500,
              background: rateFilter === r ? C.brand : C.surface,
              color: rateFilter === r ? "#fff" : C.textSub,
              border: `1px solid ${rateFilter === r ? C.brand : C.border}`,
              transition: "all 0.12s",
            }}>
              {r === "all" ? "All rates" : `${r}%`}
            </button>
          ))}
        </div>
        <SearchBar value={search} onChange={handleSearch} placeholder="Search recipient, invoice…" />
      </SectionHeader>

      <div style={{ overflowX: "auto", borderRadius: 12, border: `1px solid ${C.border}`, background: C.surface }}>
        <table style={{ width: "100%", borderCollapse: "collapse", fontSize: 13 }}>
          <thead>
            <tr style={{ background: C.bg }}>
              {["#", "Recipient", "Invoice No", "Date", "Rate", "Invoice Value", "Taxable Amt", "CGST", "SGST"].map((h, i) => (
                <th key={h} style={{ padding: "10px 12px", textAlign: i > 3 ? "right" : "left", fontWeight: 600, color: C.textMuted, fontSize: 11, textTransform: "uppercase", letterSpacing: "0.04em", borderBottom: `1px solid ${C.border}`, whiteSpace: "nowrap" }}>{h}</th>
              ))}
            </tr>
          </thead>
          <tbody>
            {slice.length === 0 ? (
              <tr><td colSpan={9} style={{ padding: "32px", textAlign: "center", color: C.textMuted }}>No results found</td></tr>
            ) : slice.map((r, i) => (
              <tr key={i} style={{ borderBottom: `1px solid ${C.border}` }}>
                <td style={{ padding: "10px 12px", color: C.textMuted, fontSize: 12 }}>{page * perPage + i + 1}</td>
                <td style={{ padding: "10px 12px", fontWeight: 500, maxWidth: 140, overflow: "hidden", textOverflow: "ellipsis", whiteSpace: "nowrap" }}>{r["RECEIPIENT"]}</td>
                <td style={{ padding: "10px 12px", color: C.blue, fontWeight: 500 }}>{r["INVOICE No"]}</td>
                <td style={{ padding: "10px 12px", color: C.textMuted, whiteSpace: "nowrap" }}>{r["INVOICE DATE"]}</td>
                <td style={{ padding: "10px 12px", textAlign: "right" }}><RateBadge rate={parseFloat(r["RATE"])} /></td>
                <td style={{ padding: "10px 12px", textAlign: "right" }}>{fmtCur(parseFloat(r["INVOICE VALUE"]))}</td>
                <td style={{ padding: "10px 12px", textAlign: "right", fontWeight: 500 }}>{fmtCur(parseFloat(r["T'BLE AMT"]))}</td>
                <td style={{ padding: "10px 12px", textAlign: "right", color: C.brandDark }}>{fmtCur(parseFloat(r["CGST"]))}</td>
                <td style={{ padding: "10px 12px", textAlign: "right", color: C.brandDark }}>{fmtCur(parseFloat(r["SGST"]))}</td>
              </tr>
            ))}
          </tbody>
        </table>
      </div>
      <Pagination page={page} pages={pages} total={filtered.length} perPage={perPage} onPage={setPage} />
    </div>
  );
}

function HSNSection({ data }) {
  const [search, setSearch] = useState("");
  if (!data || !data.rows.length) return (
    <div style={{ textAlign: "center", padding: "48px 0", color: C.textMuted }}>
      <p style={{ fontSize: 32, marginBottom: 8 }}>📭</p>
      <p>No HSN data found in this sheet.</p>
    </div>
  );

  const totalTaxable = sumField(data.rows, "T'BLE AMT");
  const totalQty = sumField(data.rows, "Tot.Qty");
  const totalVal = sumField(data.rows, "Tot Value");

  const filtered = data.rows.filter((r) => {
    const q = search.toLowerCase();
    return !q || [r["HSN CODE"], r["Description"]].some((v) => v && String(v).toLowerCase().includes(q));
  });

  return (
    <div className="fade-up">
      <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fit, minmax(130px, 1fr))", gap: 12, marginBottom: 20 }}>
        <StatCard label="HSN codes" value={new Set(data.rows.map((r) => r["HSN CODE"])).size} accent={C.brand} />
        <StatCard label="Total quantity" value={fmt(totalQty) + " kg"} accent={C.blue} />
        <StatCard label="Total value" value={fmtCur(totalVal)} accent={C.amber} />
        <StatCard label="Taxable amount" value={fmtCur(totalTaxable)} accent={C.brandDark} />
      </div>

      <SectionHeader title="HSN-wise summary" count={filtered.length}>
        <SearchBar value={search} onChange={setSearch} placeholder="Search HSN code or description…" />
      </SectionHeader>

      <div style={{ overflowX: "auto", borderRadius: 12, border: `1px solid ${C.border}`, background: C.surface }}>
        <table style={{ width: "100%", borderCollapse: "collapse", fontSize: 13 }}>
          <thead>
            <tr style={{ background: C.bg }}>
              {["#", "HSN Code", "Description", "UQC", "Total Qty", "Total Value", "Taxable Amt", "IGST", "CGST", "SGST", "Rate"].map((h, i) => (
                <th key={h} style={{ padding: "10px 12px", textAlign: i < 3 ? "left" : "right", fontWeight: 600, color: C.textMuted, fontSize: 11, textTransform: "uppercase", letterSpacing: "0.04em", borderBottom: `1px solid ${C.border}`, whiteSpace: "nowrap" }}>{h}</th>
              ))}
            </tr>
          </thead>
          <tbody>
            {filtered.length === 0 ? (
              <tr><td colSpan={11} style={{ padding: "32px", textAlign: "center", color: C.textMuted }}>No results found</td></tr>
            ) : filtered.map((r, i) => (
              <tr key={i} style={{ borderBottom: `1px solid ${C.border}` }}>
                <td style={{ padding: "10px 12px", color: C.textMuted, fontSize: 12 }}>{i + 1}</td>
                <td style={{ padding: "10px 12px", fontFamily: "monospace", fontWeight: 600, color: C.blue }}>{r["HSN CODE"]}</td>
                <td style={{ padding: "10px 12px", color: C.textSub }}>{r["Description"]}</td>
                <td style={{ padding: "10px 12px", textAlign: "right", color: C.textMuted }}>{r["UQC"]}</td>
                <td style={{ padding: "10px 12px", textAlign: "right" }}>{fmt(parseFloat(r["Tot.Qty"]))}</td>
                <td style={{ padding: "10px 12px", textAlign: "right" }}>{fmtCur(parseFloat(r["Tot Value"]))}</td>
                <td style={{ padding: "10px 12px", textAlign: "right", fontWeight: 500 }}>{fmtCur(parseFloat(r["T'BLE AMT"]))}</td>
                <td style={{ padding: "10px 12px", textAlign: "right", color: C.textMuted }}>{parseFloat(r["IGST"]) > 0 ? fmtCur(parseFloat(r["IGST"])) : "—"}</td>
                <td style={{ padding: "10px 12px", textAlign: "right", color: C.brandDark }}>{parseFloat(r["CGST"]) > 0 ? fmtCur(parseFloat(r["CGST"])) : "—"}</td>
                <td style={{ padding: "10px 12px", textAlign: "right", color: C.brandDark }}>{parseFloat(r["SGST"]) > 0 ? fmtCur(parseFloat(r["SGST"])) : "—"}</td>
                <td style={{ padding: "10px 12px", textAlign: "right" }}><RateBadge rate={parseFloat(r["Tax Rate"])} /></td>
              </tr>
            ))}
          </tbody>
          <tfoot>
            <tr style={{ background: C.bg, fontWeight: 600 }}>
              <td colSpan={4} style={{ padding: "10px 12px", fontSize: 12, color: C.textSub }}>Grand Total</td>
              <td style={{ padding: "10px 12px", textAlign: "right", fontSize: 12 }}>{fmt(totalQty)}</td>
              <td style={{ padding: "10px 12px", textAlign: "right", fontSize: 12 }}>{fmtCur(totalVal)}</td>
              <td style={{ padding: "10px 12px", textAlign: "right", fontSize: 12 }}>{fmtCur(totalTaxable)}</td>
              <td colSpan={4}></td>
            </tr>
          </tfoot>
        </table>
      </div>
    </div>
  );
}

// ─── Main App ─────────────────────────────────────────────────────────────────
const NAV = [
  { id: "overview", label: "Overview",  icon: "📊" },
  { id: "b2b",      label: "B2B Sales", icon: "🏢" },
  { id: "b2c",      label: "B2C Sales", icon: "🛒" },
  { id: "hsnb2b",   label: "HSN B2B",   icon: "🔖" },
  { id: "hsnb2c",   label: "HSN B2C",   icon: "🔖" },
];

export default function App() {
  const [summary, setSummary] = useState(null);
  const [activeTab, setActiveTab] = useState("overview");
  const [fileName, setFileName] = useState("");
  const [error, setError] = useState("");
  const [dragging, setDragging] = useState(false);
  const [loading, setLoading] = useState(false);

  const processFile = (file) => {
    if (!file) return;
    setError("");
    setFileName(file.name);
    setLoading(true);
    const reader = new FileReader();
    reader.onload = (e) => {
      setTimeout(() => {
        try {
          const data = new Uint8Array(e.target.result);
          const wb = XLSX.read(data, { type: "array" });
          const s = buildSummary(wb);
          setSummary(s);
          setActiveTab("overview");
        } catch {
          setError("Could not parse file. Please upload a valid GSTR-1 Excel file.");
          setSummary(null);
        } finally {
          setLoading(false);
        }
      }, 400);
    };
    reader.readAsArrayBuffer(file);
  };

  const onDrop = useCallback((e) => {
    e.preventDefault();
    setDragging(false);
    processFile(e.dataTransfer.files[0]);
  }, []);

  return (
    <div style={{ minHeight: "100vh", background: C.bg, fontFamily: "-apple-system, BlinkMacSystemFont, 'Segoe UI', sans-serif" }}>
      <style>{GLOBAL_STYLE}</style>

      {/* Header */}
      <div style={{ background: C.surface, borderBottom: `1px solid ${C.border}`, padding: "0 24px", position: "sticky", top: 0, zIndex: 100, boxShadow: "0 1px 3px rgba(0,0,0,0.06)" }}>
        <div style={{ maxWidth: 1100, margin: "0 auto", height: 56, display: "flex", alignItems: "center", justifyContent: "space-between" }}>
          <div style={{ display: "flex", alignItems: "center", gap: 10 }}>
            <div style={{ width: 32, height: 32, borderRadius: 8, background: C.brand, display: "flex", alignItems: "center", justifyContent: "center", fontSize: 16 }}>🧾</div>
            <div>
              <p style={{ fontSize: 14, fontWeight: 700, color: C.text, lineHeight: 1 }}>GSTR-1 Summary</p>
              <p style={{ fontSize: 11, color: C.textMuted, lineHeight: 1, marginTop: 2 }}>Sales return generator</p>
            </div>
          </div>
          {summary?.period && (
            <div style={{ background: C.brandLight, color: C.brandDark, fontSize: 12, fontWeight: 600, padding: "4px 12px", borderRadius: 20 }}>
              📅 {summary.period}
            </div>
          )}
          {summary && (
            <button onClick={() => { setSummary(null); setFileName(""); }} style={{
              padding: "6px 14px", borderRadius: 8, fontSize: 12, fontWeight: 500,
              background: C.bg, color: C.textSub, border: `1px solid ${C.border}`,
              display: "flex", alignItems: "center", gap: 5,
            }}>
              ↩ New file
            </button>
          )}
        </div>
      </div>

      <div style={{ maxWidth: 1100, margin: "0 auto", padding: "24px 24px 48px" }}>

        {/* Upload zone */}
        {!summary && !loading && (
          <div className="fade-up">
            <div style={{ textAlign: "center", padding: "40px 0 28px" }}>
              <p style={{ fontSize: 26, fontWeight: 700, color: C.text, marginBottom: 6 }}>Upload your sales file</p>
              <p style={{ fontSize: 14, color: C.textMuted }}>Supports GSTR-1 format Excel files with B2B, B2C, HSN sheets</p>
            </div>
            <div
              onDrop={onDrop}
              onDragOver={(e) => { e.preventDefault(); setDragging(true); }}
              onDragLeave={() => setDragging(false)}
              onClick={() => document.getElementById("gstr-file-input").click()}
              style={{
                border: `2px dashed ${dragging ? C.brand : C.borderStrong}`,
                borderRadius: 20, padding: "52px 40px", textAlign: "center",
                background: dragging ? C.brandLight : C.surface,
                cursor: "pointer", transition: "all 0.2s",
                boxShadow: dragging ? `0 0 0 4px ${C.brandLight}` : C.shadow,
                maxWidth: 520, margin: "0 auto",
              }}
            >
              <div style={{ fontSize: 48, marginBottom: 14 }}>{dragging ? "📂" : "📁"}</div>
              <p style={{ fontSize: 16, fontWeight: 600, color: C.text, marginBottom: 6 }}>
                {dragging ? "Release to upload" : "Drop your Excel file here"}
              </p>
              <p style={{ fontSize: 13, color: C.textMuted, marginBottom: 18 }}>or click to browse your files</p>
              <div style={{ display: "inline-block", background: C.brand, color: "#fff", padding: "9px 22px", borderRadius: 10, fontSize: 13, fontWeight: 600 }}>
                Choose .xlsx file
              </div>
              <input id="gstr-file-input" type="file" accept=".xlsx" style={{ display: "none" }} onChange={(e) => e.target.files[0] && processFile(e.target.files[0])} />
            </div>

            {error && (
              <div style={{ maxWidth: 520, margin: "16px auto 0", background: C.redBg, color: C.red, borderRadius: 10, padding: "12px 16px", fontSize: 13, border: `1px solid ${C.red}22` }}>
                ⚠️ {error}
              </div>
            )}
          </div>
        )}

        {/* Loading */}
        {loading && (
          <div style={{ textAlign: "center", padding: "80px 0" }}>
            <div style={{ width: 48, height: 48, border: `3px solid ${C.brandLight}`, borderTopColor: C.brand, borderRadius: "50%", animation: "spin 0.8s linear infinite", margin: "0 auto 16px" }} />
            <p style={{ fontSize: 14, color: C.textMuted }}>Parsing {fileName}…</p>
          </div>
        )}

        {/* Main content */}
        {summary && !loading && (
          <div>
            {/* Nav tabs */}
            <div style={{ display: "flex", gap: 4, marginBottom: 20, background: C.surface, padding: 4, borderRadius: 12, border: `1px solid ${C.border}`, width: "fit-content" }}>
              {NAV.map((t) => {
                const active = activeTab === t.id;
                return (
                  <button key={t.id} onClick={() => setActiveTab(t.id)} style={{
                    padding: "7px 16px", borderRadius: 9, fontSize: 13, fontWeight: active ? 600 : 400,
                    background: active ? C.brand : "transparent",
                    color: active ? "#fff" : C.textSub,
                    transition: "all 0.15s",
                    display: "flex", alignItems: "center", gap: 5,
                  }}>
                    <span>{t.icon}</span> {t.label}
                  </button>
                );
              })}
            </div>

            {/* Section panels */}
            {activeTab === "overview" && <OverviewSection summary={summary} />}
            {activeTab === "b2b" && summary.b2bSummary && <B2BSection data={summary.b2bSummary} />}
            {activeTab === "b2c" && summary.b2cSummary && <B2CSection data={summary.b2cSummary} />}
            {activeTab === "hsnb2b" && <HSNSection data={summary.hsnB2B} />}
            {activeTab === "hsnb2c" && <HSNSection data={summary.hsnB2C} />}
          </div>
        )}
      </div>
    </div>
  );
}