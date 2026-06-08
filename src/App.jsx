import { useState, useCallback } from "react";
import * as XLSX from "xlsx";

const fmt = (n) =>
  typeof n === "number" && !isNaN(n)
    ? n.toLocaleString("en-IN", { maximumFractionDigits: 2 })
    : n ?? "—";

const fmtCur = (n) =>
  typeof n === "number" && !isNaN(n)
    ? "₹" + n.toLocaleString("en-IN", { maximumFractionDigits: 2 })
    : "—";

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
  const find = (keyword) =>
    sheets.find((s) => s.toLowerCase().includes(keyword.toLowerCase()));

  const b2bName = find("b2b");
  const b2cName = find("b2c sales");
  const hsnB2BName = find("hsn") && sheets.find((s) => s.toLowerCase().includes("hsn") && s.toLowerCase().includes("b2b"));
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

const TAB_STYLE = (active) => ({
  padding: "8px 18px",
  border: "0.5px solid var(--color-border-tertiary)",
  borderBottom: active ? "2px solid #1D9E75" : "0.5px solid var(--color-border-tertiary)",
  background: active ? "var(--color-background-primary)" : "var(--color-background-secondary)",
  borderRadius: "var(--border-radius-md) var(--border-radius-md) 0 0",
  cursor: "pointer",
  fontSize: 13,
  fontWeight: active ? 500 : 400,
  color: active ? "#1D9E75" : "var(--color-text-secondary)",
  marginRight: 4,
  marginBottom: -1,
  transition: "all 0.15s",
});

const StatCard = ({ label, value, sub }) => (
  <div style={{
    background: "var(--color-background-secondary)",
    borderRadius: "var(--border-radius-md)",
    padding: "12px 16px",
    minWidth: 0,
  }}>
    <p style={{ fontSize: 12, color: "var(--color-text-secondary)", margin: "0 0 4px" }}>{label}</p>
    <p style={{ fontSize: 20, fontWeight: 500, margin: 0, color: "var(--color-text-primary)" }}>{value}</p>
    {sub && <p style={{ fontSize: 11, color: "var(--color-text-tertiary)", margin: "2px 0 0" }}>{sub}</p>}
  </div>
);

const Badge = ({ children, color }) => {
  const map = {
    green: { bg: "#E1F5EE", color: "#0F6E56" },
    blue: { bg: "#E6F1FB", color: "#185FA5" },
    amber: { bg: "#FAEEDA", color: "#854F0B" },
    gray: { bg: "#F1EFE8", color: "#5F5E5A" },
  };
  const c = map[color] || map.gray;
  return (
    <span style={{
      background: c.bg, color: c.color, fontSize: 11, fontWeight: 500,
      padding: "2px 8px", borderRadius: "var(--border-radius-md)", whiteSpace: "nowrap",
    }}>{children}</span>
  );
};

const RATE_COLORS = {
  0:  { bg: "#F1EFE8", text: "#5F5E5A", label: "0%" },
  5:  { bg: "#E1F5EE", text: "#0F6E56", label: "5%" },
  12: { bg: "#E6F1FB", text: "#185FA5", label: "12%" },
  18: { bg: "#FAEEDA", text: "#854F0B", label: "18%" },
  28: { bg: "#FCEBEB", text: "#A32D2D", label: "28%" },
};

function TaxRateBreakdown({ byRate }) {
  if (!byRate || byRate.length === 0) return null;
  return (
    <div style={{ marginTop: 14 }}>
      <p style={{ fontSize: 12, fontWeight: 500, color: "var(--color-text-secondary)", margin: "0 0 8px", textTransform: "uppercase", letterSpacing: "0.04em" }}>
        By tax rate
      </p>
      <table style={{ width: "100%", borderCollapse: "collapse", fontSize: 12 }}>
        <thead>
          <tr>
            {["Rate", "Taxable amt", "CGST", "SGST", "IGST", "Total tax"].map((h) => (
              <th key={h} style={{ padding: "5px 8px", textAlign: h === "Rate" ? "left" : "right", fontWeight: 500, color: "var(--color-text-tertiary)", borderBottom: "0.5px solid var(--color-border-tertiary)", whiteSpace: "nowrap" }}>{h}</th>
            ))}
          </tr>
        </thead>
        <tbody>
          {byRate.map((d) => {
            const rc = RATE_COLORS[d.rate] || RATE_COLORS[0];
            const totalTax = d.cgst + d.sgst + d.igst;
            return (
              <tr key={d.rate} style={{ borderBottom: "0.5px solid var(--color-border-tertiary)" }}>
                <td style={{ padding: "6px 8px" }}>
                  <span style={{ background: rc.bg, color: rc.text, fontWeight: 500, fontSize: 11, padding: "2px 8px", borderRadius: "var(--border-radius-md)" }}>{rc.label}</span>
                </td>
                <td style={{ padding: "6px 8px", textAlign: "right" }}>{fmtCur(d.taxableAmt)}</td>
                <td style={{ padding: "6px 8px", textAlign: "right" }}>{d.cgst > 0 ? fmtCur(d.cgst) : <span style={{ color: "var(--color-text-tertiary)" }}>—</span>}</td>
                <td style={{ padding: "6px 8px", textAlign: "right" }}>{d.sgst > 0 ? fmtCur(d.sgst) : <span style={{ color: "var(--color-text-tertiary)" }}>—</span>}</td>
                <td style={{ padding: "6px 8px", textAlign: "right" }}>{d.igst > 0 ? fmtCur(d.igst) : <span style={{ color: "var(--color-text-tertiary)" }}>—</span>}</td>
                <td style={{ padding: "6px 8px", textAlign: "right", fontWeight: 500 }}>{totalTax > 0 ? fmtCur(totalTax) : <span style={{ color: "var(--color-text-tertiary)" }}>—</span>}</td>
              </tr>
            );
          })}
        </tbody>
      </table>
    </div>
  );
}

function B2BSection({ data }) {
  const [page, setPage] = useState(0);
  const perPage = 10;
  const total = data.rows.length;
  const slice = data.rows.slice(page * perPage, (page + 1) * perPage);
  const pages = Math.ceil(total / perPage);

  return (
    <div>
      <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fit, minmax(140px, 1fr))", gap: 10, marginBottom: 20 }}>
        <StatCard label="Total invoices" value={fmt(data.invoices)} />
        <StatCard label="Recipients" value={fmt(data.recipients)} />
        <StatCard label="Invoice value" value={fmtCur(data.invoiceValue)} />
        <StatCard label="Taxable amount" value={fmtCur(data.taxableAmt)} />
        <StatCard label="CGST" value={fmtCur(data.cgst)} />
        <StatCard label="SGST" value={fmtCur(data.sgst)} />
      </div>
      <div style={{ overflowX: "auto" }}>
        <table style={{ width: "100%", borderCollapse: "collapse", fontSize: 12, tableLayout: "fixed" }}>
          <colgroup>
            <col style={{ width: 28 }} />
            <col style={{ width: 130 }} />
            <col style={{ width: 140 }} />
            <col style={{ width: 70 }} />
            <col style={{ width: 90 }} />
            <col style={{ width: 90 }} />
            <col style={{ width: 80 }} />
            <col style={{ width: 70 }} />
            <col style={{ width: 70 }} />
          </colgroup>
          <thead>
            <tr style={{ background: "var(--color-background-secondary)" }}>
              {["#", "Recipient", "GSTIN", "Invoice No", "Date", "Value", "Taxable Amt", "CGST", "SGST"].map((h) => (
                <th key={h} style={{ padding: "8px 6px", textAlign: "left", fontWeight: 500, color: "var(--color-text-secondary)", borderBottom: "0.5px solid var(--color-border-tertiary)", whiteSpace: "nowrap", overflow: "hidden", textOverflow: "ellipsis" }}>{h}</th>
              ))}
            </tr>
          </thead>
          <tbody>
            {slice.map((r, i) => (
              <tr key={i} style={{ borderBottom: "0.5px solid var(--color-border-tertiary)" }}>
                <td style={{ padding: "7px 6px", color: "var(--color-text-tertiary)" }}>{page * perPage + i + 1}</td>
                <td style={{ padding: "7px 6px", overflow: "hidden", textOverflow: "ellipsis", whiteSpace: "nowrap" }}>{r["RECEIPIENT"]}</td>
                <td style={{ padding: "7px 6px", fontFamily: "var(--font-mono)", fontSize: 11, overflow: "hidden", textOverflow: "ellipsis" }}>{r["GSTIN/UN OF RECEIPIENT"] || "—"}</td>
                <td style={{ padding: "7px 6px" }}>{r["INVOICE No"]}</td>
                <td style={{ padding: "7px 6px", color: "var(--color-text-secondary)" }}>{r["INVOICE DATE"]}</td>
                <td style={{ padding: "7px 6px", textAlign: "right" }}>{fmtCur(parseFloat(r["INVOICE VALUE"]))}</td>
                <td style={{ padding: "7px 6px", textAlign: "right" }}>{fmtCur(parseFloat(r["T'BLE AMT"]))}</td>
                <td style={{ padding: "7px 6px", textAlign: "right" }}>{fmtCur(parseFloat(r["CGST"]))}</td>
                <td style={{ padding: "7px 6px", textAlign: "right" }}>{fmtCur(parseFloat(r["SGST"]))}</td>
              </tr>
            ))}
          </tbody>
        </table>
      </div>
      {pages > 1 && (
        <div style={{ display: "flex", gap: 6, marginTop: 10, alignItems: "center", justifyContent: "flex-end" }}>
          <button onClick={() => setPage(Math.max(0, page - 1))} disabled={page === 0} style={{ padding: "4px 10px", fontSize: 12 }}>‹</button>
          <span style={{ fontSize: 12, color: "var(--color-text-secondary)" }}>{page + 1} / {pages}</span>
          <button onClick={() => setPage(Math.min(pages - 1, page + 1))} disabled={page === pages - 1} style={{ padding: "4px 10px", fontSize: 12 }}>›</button>
        </div>
      )}
    </div>
  );
}

function B2CSection({ data }) {
  const [page, setPage] = useState(0);
  const perPage = 10;
  const total = data.rows.length;
  const slice = data.rows.slice(page * perPage, (page + 1) * perPage);
  const pages = Math.ceil(total / perPage);

  return (
    <div>
      <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fit, minmax(140px, 1fr))", gap: 10, marginBottom: 20 }}>
        <StatCard label="Total line items" value={fmt(total)} />
        <StatCard label="Unique invoices" value={fmt(data.invoices)} />
        <StatCard label="Invoice value" value={fmtCur(data.invoiceValue)} />
        <StatCard label="Taxable amount" value={fmtCur(data.taxableAmt)} />
        <StatCard label="CGST" value={fmtCur(data.cgst)} />
        <StatCard label="SGST" value={fmtCur(data.sgst)} />
      </div>
      <div style={{ overflowX: "auto" }}>
        <table style={{ width: "100%", borderCollapse: "collapse", fontSize: 12, tableLayout: "fixed" }}>
          <colgroup>
            <col style={{ width: 28 }} />
            <col style={{ width: 130 }} />
            <col style={{ width: 70 }} />
            <col style={{ width: 90 }} />
            <col style={{ width: 90 }} />
            <col style={{ width: 90 }} />
            <col style={{ width: 70 }} />
            <col style={{ width: 70 }} />
          </colgroup>
          <thead>
            <tr style={{ background: "var(--color-background-secondary)" }}>
              {["#", "Recipient", "Invoice No", "Date", "Value", "Taxable Amt", "CGST", "SGST"].map((h) => (
                <th key={h} style={{ padding: "8px 6px", textAlign: "left", fontWeight: 500, color: "var(--color-text-secondary)", borderBottom: "0.5px solid var(--color-border-tertiary)", whiteSpace: "nowrap" }}>{h}</th>
              ))}
            </tr>
          </thead>
          <tbody>
            {slice.map((r, i) => (
              <tr key={i} style={{ borderBottom: "0.5px solid var(--color-border-tertiary)" }}>
                <td style={{ padding: "7px 6px", color: "var(--color-text-tertiary)" }}>{page * perPage + i + 1}</td>
                <td style={{ padding: "7px 6px", overflow: "hidden", textOverflow: "ellipsis", whiteSpace: "nowrap" }}>{r["RECEIPIENT"]}</td>
                <td style={{ padding: "7px 6px" }}>{r["INVOICE No"]}</td>
                <td style={{ padding: "7px 6px", color: "var(--color-text-secondary)" }}>{r["INVOICE DATE"]}</td>
                <td style={{ padding: "7px 6px", textAlign: "right" }}>{fmtCur(parseFloat(r["INVOICE VALUE"]))}</td>
                <td style={{ padding: "7px 6px", textAlign: "right" }}>{fmtCur(parseFloat(r["T'BLE AMT"]))}</td>
                <td style={{ padding: "7px 6px", textAlign: "right" }}>{fmtCur(parseFloat(r["CGST"]))}</td>
                <td style={{ padding: "7px 6px", textAlign: "right" }}>{fmtCur(parseFloat(r["SGST"]))}</td>
              </tr>
            ))}
          </tbody>
        </table>
      </div>
      {pages > 1 && (
        <div style={{ display: "flex", gap: 6, marginTop: 10, alignItems: "center", justifyContent: "flex-end" }}>
          <button onClick={() => setPage(Math.max(0, page - 1))} disabled={page === 0} style={{ padding: "4px 10px", fontSize: 12 }}>‹</button>
          <span style={{ fontSize: 12, color: "var(--color-text-secondary)" }}>{page + 1} / {pages}</span>
          <button onClick={() => setPage(Math.min(pages - 1, page + 1))} disabled={page === pages - 1} style={{ padding: "4px 10px", fontSize: 12 }}>›</button>
        </div>
      )}
    </div>
  );
}

function HSNSection({ data, label }) {
  if (!data || !data.rows.length) return <p style={{ color: "var(--color-text-tertiary)", fontSize: 13 }}>No HSN data found.</p>;

  const totalTaxable = sumField(data.rows, "T'BLE AMT");
  const totalQty = sumField(data.rows, "Tot.Qty");
  const totalVal = sumField(data.rows, "Tot Value");

  return (
    <div>
      <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fit, minmax(140px, 1fr))", gap: 10, marginBottom: 20 }}>
        <StatCard label="HSN codes" value={new Set(data.rows.map((r) => r["HSN CODE"])).size} />
        <StatCard label="Total quantity" value={fmt(totalQty) + " kg"} />
        <StatCard label="Total value" value={fmtCur(totalVal)} />
        <StatCard label="Taxable amount" value={fmtCur(totalTaxable)} />
      </div>
      <table style={{ width: "100%", borderCollapse: "collapse", fontSize: 12 }}>
        <thead>
          <tr style={{ background: "var(--color-background-secondary)" }}>
            {["#", "HSN Code", "Description", "UQC", "Total Qty", "Total Value", "Taxable Amt", "IGST", "CGST", "SGST", "Tax Rate"].map((h) => (
              <th key={h} style={{ padding: "8px 8px", textAlign: h === "#" ? "left" : "right", fontWeight: 500, color: "var(--color-text-secondary)", borderBottom: "0.5px solid var(--color-border-tertiary)", whiteSpace: "nowrap" }}>{h}</th>
            ))}
          </tr>
        </thead>
        <tbody>
          {data.rows.map((r, i) => (
            <tr key={i} style={{ borderBottom: "0.5px solid var(--color-border-tertiary)" }}>
              <td style={{ padding: "7px 8px", color: "var(--color-text-tertiary)" }}>{i + 1}</td>
              <td style={{ padding: "7px 8px", textAlign: "right", fontFamily: "var(--font-mono)", fontSize: 11 }}>{r["HSN CODE"]}</td>
              <td style={{ padding: "7px 8px", textAlign: "right" }}>{r["Description"]}</td>
              <td style={{ padding: "7px 8px", textAlign: "right", color: "var(--color-text-secondary)" }}>{r["UQC"]}</td>
              <td style={{ padding: "7px 8px", textAlign: "right" }}>{fmt(parseFloat(r["Tot.Qty"]))}</td>
              <td style={{ padding: "7px 8px", textAlign: "right" }}>{fmtCur(parseFloat(r["Tot Value"]))}</td>
              <td style={{ padding: "7px 8px", textAlign: "right" }}>{fmtCur(parseFloat(r["T'BLE AMT"]))}</td>
              <td style={{ padding: "7px 8px", textAlign: "right" }}>{fmtCur(parseFloat(r["IGST"]))}</td>
              <td style={{ padding: "7px 8px", textAlign: "right" }}>{fmtCur(parseFloat(r["CGST"]))}</td>
              <td style={{ padding: "7px 8px", textAlign: "right" }}>{fmtCur(parseFloat(r["SGST"]))}</td>
              <td style={{ padding: "7px 8px", textAlign: "right" }}>
                {parseFloat(r["Tax Rate"]) > 0
                  ? <Badge color="amber">{r["Tax Rate"]}%</Badge>
                  : <Badge color="gray">0%</Badge>}
              </td>
            </tr>
          ))}
        </tbody>
        <tfoot>
          <tr style={{ background: "var(--color-background-secondary)", fontWeight: 500 }}>
            <td colSpan={4} style={{ padding: "8px 8px", fontSize: 12 }}>Grand Total</td>
            <td style={{ padding: "8px 8px", textAlign: "right", fontSize: 12 }}>{fmt(totalQty)}</td>
            <td style={{ padding: "8px 8px", textAlign: "right", fontSize: 12 }}>{fmtCur(totalVal)}</td>
            <td style={{ padding: "8px 8px", textAlign: "right", fontSize: 12 }}>{fmtCur(totalTaxable)}</td>
            <td colSpan={4}></td>
          </tr>
        </tfoot>
      </table>
    </div>
  );
}

export default function App() {
  const [summary, setSummary] = useState(null);
  const [activeTab, setActiveTab] = useState("overview");
  const [fileName, setFileName] = useState("");
  const [error, setError] = useState("");
  const [dragging, setDragging] = useState(false);

  const processFile = (file) => {
    if (!file) return;
    setError("");
    setFileName(file.name);
    const reader = new FileReader();
    reader.onload = (e) => {
      try {
        const data = new Uint8Array(e.target.result);
        const wb = XLSX.read(data, { type: "array" });
        const s = buildSummary(wb);
        setSummary(s);
        setActiveTab("overview");
      } catch (err) {
        setError("Could not parse file. Please upload a valid GSTR-1 Excel file.");
        setSummary(null);
      }
    };
    reader.readAsArrayBuffer(file);
  };

  const onDrop = useCallback((e) => {
    e.preventDefault();
    setDragging(false);
    const file = e.dataTransfer.files[0];
    if (file) processFile(file);
  }, []);

  const onDragOver = (e) => { e.preventDefault(); setDragging(true); };
  const onDragLeave = () => setDragging(false);

  const tabs = [
    { id: "overview", label: "Overview" },
    { id: "b2b", label: "B2B Sales" },
    { id: "b2c", label: "B2C Sales" },
    { id: "hsnb2b", label: "HSN B2B" },
    { id: "hsnb2c", label: "HSN B2C" },
  ];

  return (
    <div style={{ padding: "1.5rem 0", fontFamily: "var(--font-sans)" }}>
      <h2 className="sr-only">GSTR-1 Summary Sheet Generator</h2>

      <div style={{ display: "flex", alignItems: "center", gap: 10, marginBottom: 6 }}>
        <i className="ti ti-file-invoice" style={{ fontSize: 22, color: "#1D9E75" }} aria-hidden="true"></i>
        <h2 style={{ margin: 0, fontSize: 18, fontWeight: 500 }}>GSTR-1 summary generator</h2>
        {summary && <Badge color="green">{summary.sheetNames.length} sheets loaded</Badge>}
      </div>
      <p style={{ fontSize: 13, color: "var(--color-text-secondary)", marginTop: 2, marginBottom: 18 }}>
        Upload your sales Excel file to generate a structured GSTR-1 summary
      </p>

      <div
        onDrop={onDrop}
        onDragOver={onDragOver}
        onDragLeave={onDragLeave}
        style={{
          border: `1.5px dashed ${dragging ? "#1D9E75" : "var(--color-border-secondary)"}`,
          borderRadius: "var(--border-radius-lg)",
          padding: "24px 20px",
          textAlign: "center",
          background: dragging ? "#E1F5EE" : "var(--color-background-secondary)",
          marginBottom: 20,
          transition: "all 0.15s",
          cursor: "pointer",
        }}
        onClick={() => document.getElementById("fileInput").click()}
      >
        <i className="ti ti-upload" style={{ fontSize: 28, color: dragging ? "#1D9E75" : "var(--color-text-tertiary)", display: "block", marginBottom: 8 }} aria-hidden="true"></i>
        <p style={{ margin: 0, fontSize: 13, color: "var(--color-text-primary)", fontWeight: 500 }}>
          {fileName ? fileName : "Drop your GSTR-1 Excel file here, or click to browse"}
        </p>
        <p style={{ margin: "4px 0 0", fontSize: 12, color: "var(--color-text-tertiary)" }}>.xlsx files only</p>
        <input
          id="fileInput"
          type="file"
          accept=".xlsx"
          style={{ display: "none" }}
          onChange={(e) => e.target.files[0] && processFile(e.target.files[0])}
        />
      </div>

      {error && (
        <div style={{ background: "var(--color-background-danger)", color: "var(--color-text-danger)", borderRadius: "var(--border-radius-md)", padding: "10px 14px", fontSize: 13, marginBottom: 16 }}>
          <i className="ti ti-alert-circle" style={{ marginRight: 6 }} aria-hidden="true"></i>{error}
        </div>
      )}

      {summary && (
        <div>
          {summary.period && (
            <div style={{ display: "flex", alignItems: "center", gap: 8, marginBottom: 16 }}>
              <i className="ti ti-calendar" style={{ fontSize: 15, color: "var(--color-text-secondary)" }} aria-hidden="true"></i>
              <span style={{ fontSize: 13, color: "var(--color-text-secondary)" }}>Period: <strong style={{ color: "var(--color-text-primary)" }}>{summary.period}</strong></span>
            </div>
          )}

          <div style={{ borderBottom: "0.5px solid var(--color-border-tertiary)", marginBottom: 0, display: "flex", flexWrap: "wrap" }}>
            {tabs.map((t) => (
              <button key={t.id} onClick={() => setActiveTab(t.id)} style={TAB_STYLE(activeTab === t.id)}>{t.label}</button>
            ))}
          </div>

          <div style={{ border: "0.5px solid var(--color-border-tertiary)", borderTop: "none", borderRadius: "0 0 var(--border-radius-lg) var(--border-radius-lg)", padding: "20px 16px", background: "var(--color-background-primary)" }}>

            {activeTab === "overview" && (
              <div>
                <p style={{ fontSize: 13, fontWeight: 500, color: "var(--color-text-secondary)", marginTop: 0, marginBottom: 14 }}>Consolidated summary</p>
                <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fit, minmax(150px, 1fr))", gap: 10, marginBottom: 24 }}>
                  <StatCard label="B2B taxable amount" value={fmtCur(summary.b2bSummary?.taxableAmt)} sub={`${summary.b2bSummary?.invoices} invoices`} />
                  <StatCard label="B2C taxable amount" value={fmtCur(summary.b2cSummary?.taxableAmt)} sub={`${summary.b2cSummary?.invoices} invoices`} />
                  <StatCard label="Total taxable" value={fmtCur((summary.b2bSummary?.taxableAmt || 0) + (summary.b2cSummary?.taxableAmt || 0))} />
                  <StatCard label="Total CGST" value={fmtCur((summary.b2bSummary?.cgst || 0) + (summary.b2cSummary?.cgst || 0))} />
                  <StatCard label="Total SGST" value={fmtCur((summary.b2bSummary?.sgst || 0) + (summary.b2cSummary?.sgst || 0))} />
                  <StatCard label="Total tax liability" value={fmtCur(((summary.b2bSummary?.cgst || 0) + (summary.b2bSummary?.sgst || 0) + (summary.b2cSummary?.cgst || 0) + (summary.b2cSummary?.sgst || 0) + (summary.b2cSummary?.igst || 0)))} />
                </div>

                <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 12 }}>
                  {[
                    { title: "B2B sales (bill of supply)", d: summary.b2bSummary },
                    { title: "B2C sales (unregistered)", d: summary.b2cSummary },
                  ].map(({ title, d }) => d && (
                    <div key={title} style={{ background: "var(--color-background-primary)", border: "0.5px solid var(--color-border-tertiary)", borderRadius: "var(--border-radius-lg)", padding: "14px 16px" }}>
                      <p style={{ margin: "0 0 12px", fontSize: 13, fontWeight: 500 }}>{title}</p>
                      {[
                        ["Invoice value", fmtCur(d.invoiceValue)],
                        ["Taxable amount", fmtCur(d.taxableAmt)],
                        ["CGST", fmtCur(d.cgst)],
                        ["SGST", fmtCur(d.sgst)],
                        ["IGST", fmtCur(d.igst)],
                      ].map(([k, v]) => (
                        <div key={k} style={{ display: "flex", justifyContent: "space-between", padding: "5px 0", borderBottom: "0.5px solid var(--color-border-tertiary)", fontSize: 13 }}>
                          <span style={{ color: "var(--color-text-secondary)" }}>{k}</span>
                          <span style={{ fontWeight: 500 }}>{v}</span>
                        </div>
                      ))}
                      <TaxRateBreakdown byRate={d.byRate} />
                    </div>
                  ))}
                </div>

                <div style={{ marginTop: 16, padding: "12px 14px", background: "var(--color-background-secondary)", borderRadius: "var(--border-radius-md)", fontSize: 12, color: "var(--color-text-secondary)" }}>
                  <i className="ti ti-info-circle" style={{ marginRight: 6 }} aria-hidden="true"></i>
                  Sheets detected: {summary.sheetNames.join(" · ")}
                </div>
              </div>
            )}

            {activeTab === "b2b" && summary.b2bSummary && <B2BSection data={summary.b2bSummary} />}
            {activeTab === "b2c" && summary.b2cSummary && <B2CSection data={summary.b2cSummary} />}
            {activeTab === "hsnb2b" && <HSNSection data={summary.hsnB2B} label="B2B HSN" />}
            {activeTab === "hsnb2c" && <HSNSection data={summary.hsnB2C} label="B2C HSN" />}
          </div>
        </div>
      )}
    </div>
  );
}