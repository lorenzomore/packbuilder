import { useState, useEffect, useRef, useMemo } from "react";
import * as XLSX from "xlsx";
import { jsPDF } from "jspdf";
import {
  Plus, Trash2, Upload, Search, X, Check, User,
  Pencil, Users, Package, FileSpreadsheet, Download, Sun, Snowflake, Copy, FileDown
} from "lucide-react";

const CATEGORIES = [
  "Clothing", "Footwear", "Shelter", "Sleep System", "Kitchen", "Water",
  "Navigation", "Electronics & Camera", "First Aid", "Sun & Skin Care",
  "Tools", "Food", "Miscellaneous"
];

const SEASONS = ["Summer", "Winter"];

const DEFAULT_GEAR = [];

function uid() {
  return Math.random().toString(36).slice(2, 10) + Date.now().toString(36);
}

function downloadBlob(filename, content, mime) {
  const blob = content instanceof Blob ? content : new Blob([content], { type: mime });
  const url = URL.createObjectURL(blob);
  const a = document.createElement("a");
  a.href = url;
  a.download = filename;
  document.body.appendChild(a);
  a.click();
  document.body.removeChild(a);
  setTimeout(() => URL.revokeObjectURL(url), 1000);
}

function xmlEscape(s) {
  return String(s).replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;").replace(/"/g, "&quot;");
}

function crc32(bytes) {
  let table = crc32.table;
  if (!table) {
    table = crc32.table = [];
    for (let n = 0; n < 256; n++) {
      let c = n;
      for (let k = 0; k < 8; k++) c = (c & 1) ? (0xEDB88320 ^ (c >>> 1)) : (c >>> 1);
      table[n] = c;
    }
  }
  let crc = 0 ^ (-1);
  for (let i = 0; i < bytes.length; i++) crc = (crc >>> 8) ^ table[(crc ^ bytes[i]) & 0xFF];
  return (crc ^ (-1)) >>> 0;
}

function buildZip(files) {
  const localParts = [];
  const centralParts = [];
  let offset = 0;
  const dosTime = 0;
  const dosDate = 22561;
  files.forEach(f => {
    const nameBytes = new TextEncoder().encode(f.name);
    const data = f.data;
    const crc = crc32(data);
    const size = data.length;
    const local = new Uint8Array(30 + nameBytes.length + size);
    const dv = new DataView(local.buffer);
    dv.setUint32(0, 0x04034b50, true);
    dv.setUint16(4, 20, true);
    dv.setUint16(6, 0, true);
    dv.setUint16(8, 0, true);
    dv.setUint16(10, dosTime, true);
    dv.setUint16(12, dosDate, true);
    dv.setUint32(14, crc, true);
    dv.setUint32(18, size, true);
    dv.setUint32(22, size, true);
    dv.setUint16(26, nameBytes.length, true);
    dv.setUint16(28, 0, true);
    local.set(nameBytes, 30);
    local.set(data, 30 + nameBytes.length);
    localParts.push(local);

    const central = new Uint8Array(46 + nameBytes.length);
    const cdv = new DataView(central.buffer);
    cdv.setUint32(0, 0x02014b50, true);
    cdv.setUint16(4, 20, true);
    cdv.setUint16(6, 20, true);
    cdv.setUint16(8, 0, true);
    cdv.setUint16(10, 0, true);
    cdv.setUint16(12, dosTime, true);
    cdv.setUint16(14, dosDate, true);
    cdv.setUint32(16, crc, true);
    cdv.setUint32(20, size, true);
    cdv.setUint32(24, size, true);
    cdv.setUint16(28, nameBytes.length, true);
    cdv.setUint16(30, 0, true);
    cdv.setUint16(32, 0, true);
    cdv.setUint16(34, 0, true);
    cdv.setUint16(36, 0, true);
    cdv.setUint32(38, 0, true);
    cdv.setUint32(42, offset, true);
    central.set(nameBytes, 46);
    centralParts.push(central);

    offset += local.length;
  });

  const centralSize = centralParts.reduce((s, p) => s + p.length, 0);
  const centralOffset = offset;
  const end = new Uint8Array(22);
  const edv = new DataView(end.buffer);
  edv.setUint32(0, 0x06054b50, true);
  edv.setUint16(4, 0, true);
  edv.setUint16(6, 0, true);
  edv.setUint16(8, files.length, true);
  edv.setUint16(10, files.length, true);
  edv.setUint32(12, centralSize, true);
  edv.setUint32(16, centralOffset, true);
  edv.setUint16(20, 0, true);

  return new Blob([...localParts, ...centralParts, end], {
    type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
  });
}

function buildTemplateXlsxBlob() {
  const catList = xmlEscape(CATEGORIES.join(","));
  const seasonList = xmlEscape(SEASONS.join(","));
  const contentTypes = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/><Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/></Types>`;
  const rootRels = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>`;
  const workbookRels = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/></Relationships>`;
  const workbook = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="Gear Template" sheetId="1" r:id="rId1"/></sheets></workbook>`;
  const styles = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><fonts count="1"><font><sz val="11"/><name val="Arial"/></font></fonts><fills count="1"><fill><patternFill patternType="none"/></fill></fills><borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders><cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs><cellXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/></cellXfs></styleSheet>`;
  const sheet = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><dimension ref="A1:D201"/><sheetViews><sheetView workbookViewId="0"/></sheetViews><cols><col min="1" max="1" width="28" customWidth="1"/><col min="2" max="2" width="22" customWidth="1"/><col min="3" max="3" width="14" customWidth="1"/><col min="4" max="4" width="40" customWidth="1"/></cols><sheetData><row r="1"><c r="A1" t="inlineStr"><is><t>Name</t></is></c><c r="B1" t="inlineStr"><is><t>Category</t></is></c><c r="C1" t="inlineStr"><is><t>Season</t></is></c><c r="D1" t="inlineStr"><is><t>Notes</t></is></c></row></sheetData><dataValidations count="2"><dataValidation type="list" allowBlank="1" showInputMessage="1" showErrorMessage="1" sqref="B2:B201"><formula1>&quot;${catList}&quot;</formula1></dataValidation><dataValidation type="list" allowBlank="1" showInputMessage="1" showErrorMessage="1" sqref="C2:C201"><formula1>&quot;${seasonList}&quot;</formula1></dataValidation></dataValidations></worksheet>`;

  const enc = (s) => new TextEncoder().encode(s);
  return buildZip([
    { name: "[Content_Types].xml", data: enc(contentTypes) },
    { name: "_rels/.rels", data: enc(rootRels) },
    { name: "xl/workbook.xml", data: enc(workbook) },
    { name: "xl/_rels/workbook.xml.rels", data: enc(workbookRels) },
    { name: "xl/styles.xml", data: enc(styles) },
    { name: "xl/worksheets/sheet1.xml", data: enc(sheet) }
  ]);
}

function mdEscape(s) {
  return String(s || "").replace(/\|/g, "\\|");
}

function categoryOf(it, byId) {
  if (it.custom) return "Custom";
  const g = byId[it.gearId];
  return g ? g.category : "Unknown";
}

function nameOf(it, byId) {
  if (it.custom) return it.name;
  const g = byId[it.gearId];
  return g ? g.name : "(removed item)";
}

function sortPackItems(items, byId) {
  return [...items].sort((a, b) => {
    const ca = categoryOf(a, byId);
    const cb = categoryOf(b, byId);
    return ca.localeCompare(cb) || nameOf(a, byId).localeCompare(nameOf(b, byId));
  });
}

function buildMarkdownExport(people, gearItems) {
  const byId = {};
  gearItems.forEach(g => { byId[g.id] = g; });
  const lines = ["# People and packs", ""];
  if (people.length === 0) lines.push("_No people added yet._");
  people.forEach(person => {
    lines.push(`## ${mdEscape(person.name)}`, "");
    if (person.packs.length === 0) {
      lines.push("_No packs yet._", "");
      return;
    }
    person.packs.forEach(pack => {
      lines.push(`### ${mdEscape(pack.name)}`, "");
      if (pack.items.length === 0) {
        lines.push("_No gear assigned._", "");
        return;
      }
      sortPackItems(pack.items, byId).forEach(it => {
        const box = it.packed ? "[x]" : "[ ]";
        const g = it.custom ? null : byId[it.gearId];
        const label = mdEscape(nameOf(it, byId));
        const cat = categoryOf(it, byId);
        const season = !it.custom && g && g.season ? `, ${g.season}` : "";
        const notes = it.notes ? ` — ${mdEscape(it.notes)}` : "";
        lines.push(`- ${box} ${label} (${cat}${season})${notes}`);
      });
      lines.push("");
    });
  });
  return lines.join("\n");
}

function buildPlainTextExport(people, gearItems) {
  const byId = {};
  gearItems.forEach(g => { byId[g.id] = g; });
  const lines = [];
  if (people.length === 0) lines.push("No people added yet.");
  people.forEach(person => {
    lines.push(`*${person.name}*`);
    if (person.packs.length === 0) {
      lines.push("No packs yet.", "");
      return;
    }
    person.packs.forEach(pack => {
      lines.push(`_${pack.name}_`);
      if (pack.items.length === 0) {
        lines.push("No gear assigned.");
      } else {
        sortPackItems(pack.items, byId).forEach(it => {
          const label = nameOf(it, byId);
          const cat = categoryOf(it, byId);
          const notes = it.notes ? ` — ${it.notes}` : "";
          const text = `${label} (${cat})${notes}`;
          lines.push(`- ${it.packed ? `~${text}~` : text}`);
        });
      }
      lines.push("");
    });
  });
  return lines.join("\n").trim();
}

function buildPdfBlob(people, gearItems) {
  const byId = {};
  gearItems.forEach(g => { byId[g.id] = g; });
  const doc = new jsPDF({ unit: "pt", format: "a4" });
  const marginX = 44;
  const pageHeight = doc.internal.pageSize.getHeight();
  const pageWidth = doc.internal.pageSize.getWidth();
  let y = 56;

  function ensureSpace(extra) {
    if (y + extra > pageHeight - 44) {
      doc.addPage();
      y = 56;
    }
  }

  doc.setFont("helvetica", "bold");
  doc.setFontSize(18);
  doc.text("PACKBUILDER", marginX, y);
  y += 16;
  doc.setDrawColor(0);
  doc.setLineWidth(1);
  doc.line(marginX, y, pageWidth - marginX, y);
  y += 26;

  if (people.length === 0) {
    doc.setFont("helvetica", "normal");
    doc.setFontSize(11);
    doc.text("No people added yet.", marginX, y);
  }

  people.forEach(person => {
    ensureSpace(30);
    doc.setFont("helvetica", "bold");
    doc.setFontSize(14);
    doc.text(person.name, marginX, y);
    y += 20;

    if (person.packs.length === 0) {
      doc.setFont("helvetica", "normal");
      doc.setFontSize(10);
      ensureSpace(16);
      doc.text("No packs yet.", marginX + 12, y);
      y += 20;
    }

    person.packs.forEach(pack => {
      ensureSpace(22);
      doc.setFont("helvetica", "bold");
      doc.setFontSize(11.5);
      doc.text(pack.name, marginX + 12, y);
      y += 17;

      if (pack.items.length === 0) {
        doc.setFont("helvetica", "normal");
        doc.setFontSize(10);
        ensureSpace(15);
        doc.text("No gear assigned.", marginX + 24, y);
        y += 17;
      } else {
        sortPackItems(pack.items, byId).forEach(it => {
          ensureSpace(15);
          const box = it.packed ? "[x]" : "[ ]";
          const label = nameOf(it, byId);
          const cat = categoryOf(it, byId);
          doc.setFont("helvetica", "normal");
          doc.setFontSize(10);
          let line = `${box} ${label}  (${cat})`;
          if (it.notes) line += `  — ${it.notes}`;
          doc.text(line, marginX + 24, y, { maxWidth: pageWidth - marginX * 2 - 24 });
          y += 15;
        });
      }
      y += 8;
    });
    y += 12;
  });

  return doc.output("blob");
}

async function loadKey(key, fallback) {
  try {
    const res = await window.storage.get(key, false);
    if (res && res.value) return JSON.parse(res.value);
    return fallback;
  } catch (e) {
    return fallback;
  }
}

async function saveKey(key, value) {
  try {
    await window.storage.set(key, JSON.stringify(value), false);
  } catch (e) {
    // best-effort
  }
}

function Tag({ label }) {
  return <span className="tag" title={label}>{label}</span>;
}

function SeasonBadge({ season }) {
  if (!season) return <span className="season-badge season-badge-neutral">— any season</span>;
  if (season === "Summer") return <span className="season-badge"><Sun size={12} /> Summer</span>;
  return <span className="season-badge"><Snowflake size={12} /> Winter</span>;
}

function CopyButton({ getText }) {
  const [copied, setCopied] = useState(false);

  async function handleCopy() {
    const text = getText();
    let ok = false;
    try {
      await navigator.clipboard.writeText(text);
      ok = true;
    } catch (e) {
      try {
        const ta = document.createElement("textarea");
        ta.value = text;
        ta.style.position = "fixed";
        ta.style.opacity = "0";
        document.body.appendChild(ta);
        ta.select();
        ok = document.execCommand("copy");
        document.body.removeChild(ta);
      } catch (e2) {
        ok = false;
      }
    }
    if (ok) {
      setCopied(true);
      setTimeout(() => setCopied(false), 2000);
    }
  }

  return (
    <button className="btn-ghost" onClick={handleCopy}>
      {copied ? <Check size={15} /> : <Copy size={15} />} {copied ? "Copied" : "Copy for chat"}
    </button>
  );
}

export default function PackbuilderApp() {
  const [tab, setTab] = useState("gear");
  const [gearItems, setGearItems] = useState([]);
  const [people, setPeople] = useState([]);
  const [loaded, setLoaded] = useState(false);
  const fileInputRef = useRef(null);

  useEffect(() => {
    (async () => {
      const gear = await loadKey("packbuilder-gear", null);
      const ppl = await loadKey("packbuilder-people", []);
      setGearItems(gear || DEFAULT_GEAR);
      setPeople(ppl || []);
      setLoaded(true);
      if (!gear) saveKey("packbuilder-gear", DEFAULT_GEAR);
    })();
  }, []);

  function persistGear(next) {
    setGearItems(next);
    saveKey("packbuilder-gear", next);
  }

  function persistPeople(next) {
    setPeople(next);
    saveKey("packbuilder-people", next);
  }

  if (!loaded) {
    return <div className="tw-root"><div className="loading-state">LOADING —</div></div>;
  }

  return (
    <div className="tw-root">
      <header className="tw-header">
        <div className="header-row">
          <span className="mono-label">PACKBUILDER</span>
          <span className="mono-label mono-dim">GEAR SYSTEM / V1</span>
        </div>
        <h1>Packbuilder</h1>
        <p className="tagline">Know what's in every pack.</p>
        <div className="tab-bar">
          <button className={tab === "gear" ? "tab active" : "tab"} onClick={() => setTab("gear")}>
            <span className="tab-index">01</span> Gear library
          </button>
          <button className={tab === "people" ? "tab active" : "tab"} onClick={() => setTab("people")}>
            <span className="tab-index">02</span> People &amp; packs
          </button>
        </div>
      </header>

      <main className="tw-main">
        {tab === "gear" ? (
          <GearLibrary gearItems={gearItems} onChange={persistGear} fileInputRef={fileInputRef} />
        ) : (
          <PeoplePacks gearItems={gearItems} people={people} onChange={persistPeople} />
        )}
      </main>
    </div>
  );
}

function GearLibrary({ gearItems, onChange, fileInputRef }) {
  const [search, setSearch] = useState("");
  const [catFilter, setCatFilter] = useState("All");
  const [adding, setAdding] = useState(false);
  const [editingId, setEditingId] = useState(null);
  const [form, setForm] = useState({ name: "", category: CATEGORIES[0], season: "", notes: "" });
  const [importMsg, setImportMsg] = useState("");
  const [confirmClear, setConfirmClear] = useState(false);

  const filtered = useMemo(() => {
    return gearItems
      .filter(g => catFilter === "All" || g.category === catFilter)
      .filter(g => g.name.toLowerCase().includes(search.toLowerCase()))
      .sort((a, b) => a.category.localeCompare(b.category) || a.name.localeCompare(b.name));
  }, [gearItems, search, catFilter]);

  function resetForm() {
    setForm({ name: "", category: CATEGORIES[0], season: "", notes: "" });
  }

  function startAdd() {
    resetForm();
    setEditingId(null);
    setAdding(true);
  }

  function startEdit(item) {
    setForm({ name: item.name, category: item.category, season: item.season || "", notes: item.notes || "" });
    setEditingId(item.id);
    setAdding(true);
  }

  function submitForm() {
    if (!form.name.trim()) return;
    if (editingId) {
      onChange(gearItems.map(g => g.id === editingId ? { ...g, ...form } : g));
    } else {
      onChange([...gearItems, { id: uid(), ...form }]);
    }
    setAdding(false);
    setEditingId(null);
    resetForm();
  }

  function deleteItem(id) {
    onChange(gearItems.filter(g => g.id !== id));
  }

  function clearAll() {
    onChange([]);
    setConfirmClear(false);
  }

  function handleFile(e) {
    const file = e.target.files && e.target.files[0];
    if (!file) return;
    const reader = new FileReader();
    reader.onload = (ev) => {
      try {
        const wb = XLSX.read(ev.target.result, { type: "array" });
        const sheet = wb.Sheets[wb.SheetNames[0]];
        const rows = XLSX.utils.sheet_to_json(sheet, { defval: "" });
        const imported = rows.map(row => {
          const keys = Object.keys(row);
          const get = (want) => {
            const k = keys.find(k => k.toLowerCase().trim() === want);
            return k ? row[k] : "";
          };
          const name = String(get("name") || get("item") || "").trim();
          const category = String(get("category") || "Miscellaneous").trim();
          const season = String(get("season") || "").trim();
          const notes = String(get("notes") || "");
          return name ? {
            id: uid(),
            name,
            category: CATEGORIES.includes(category) ? category : "Miscellaneous",
            season: SEASONS.includes(season) ? season : "",
            notes
          } : null;
        }).filter(Boolean);
        if (imported.length) {
          onChange([...gearItems, ...imported]);
          setImportMsg(`Added ${imported.length} item${imported.length === 1 ? "" : "s"} from ${file.name}.`);
        } else {
          setImportMsg("No rows recognized. Expect columns: Name, Category, Season, Notes.");
        }
      } catch (err) {
        setImportMsg("Could not read that file. Try a .csv or .xlsx export.");
      }
      e.target.value = "";
      setTimeout(() => setImportMsg(""), 6000);
    };
    reader.readAsArrayBuffer(file);
  }

  const grouped = useMemo(() => {
    const map = {};
    filtered.forEach(g => {
      if (!map[g.category]) map[g.category] = [];
      map[g.category].push(g);
    });
    return map;
  }, [filtered]);

  return (
    <section>
      <div className="toolbar">
        <div className="search-box">
          <Search size={15} />
          <input placeholder="Search gear" value={search} onChange={e => setSearch(e.target.value)} />
        </div>
        <div className="select-wrap">
          <select value={catFilter} onChange={e => setCatFilter(e.target.value)}>
            <option>All</option>
            {CATEGORIES.map(c => <option key={c}>{c}</option>)}
          </select>
        </div>
        <div className="spacer" />
        <button className="btn-ghost" onClick={() => downloadBlob("gear-template.xlsx", buildTemplateXlsxBlob())}>
          <Download size={15} /> Template
        </button>
        <button className="btn-ghost" onClick={() => fileInputRef.current.click()}>
          <Upload size={15} /> Upload
        </button>
        <input type="file" ref={fileInputRef} accept=".csv,.xlsx,.xls" style={{ display: "none" }} onChange={handleFile} />
        {confirmClear ? (
          <>
            <span className="confirm-text">Clear all gear?</span>
            <button className="btn-ghost" onClick={() => setConfirmClear(false)}>Cancel</button>
            <button className="btn-danger" onClick={clearAll}><Trash2 size={15} /> Confirm</button>
          </>
        ) : (
          <button className="btn-ghost" onClick={() => setConfirmClear(true)} disabled={gearItems.length === 0}>
            <Trash2 size={15} /> Clear all
          </button>
        )}
        <button className="btn-primary" onClick={startAdd}>
          <Plus size={15} /> Add item
        </button>
      </div>

      {importMsg && <div className="import-msg"><FileSpreadsheet size={14} /> {importMsg}</div>}

      {adding && (
        <div className="form-card">
          <div className="form-grid form-grid-4">
            <input placeholder="Item name" value={form.name} onChange={e => setForm({ ...form, name: e.target.value })} autoFocus />
            <select value={form.category} onChange={e => setForm({ ...form, category: e.target.value })}>
              {CATEGORIES.map(c => <option key={c}>{c}</option>)}
            </select>
            <select value={form.season} onChange={e => setForm({ ...form, season: e.target.value })}>
              <option value="">Any season</option>
              {SEASONS.map(s => <option key={s} value={s}>{s} only</option>)}
            </select>
            <input placeholder="Notes (optional)" value={form.notes} onChange={e => setForm({ ...form, notes: e.target.value })} />
          </div>
          <div className="form-actions">
            <button className="btn-ghost" onClick={() => { setAdding(false); setEditingId(null); }}><X size={15} /> Cancel</button>
            <button className="btn-primary" onClick={submitForm}><Check size={15} /> {editingId ? "Save changes" : "Add to library"}</button>
          </div>
        </div>
      )}

      {filtered.length === 0 ? (
        <div className="empty-state">
          <Package size={22} />
          <p>No gear yet. Add your first item, upload a spreadsheet, or download the template to fill in offline.</p>
        </div>
      ) : (
        <div className="gear-table">
          {Object.keys(grouped).sort().map(cat => (
            <div key={cat} className="cat-group">
              <div className="cat-heading">
                <Tag label={cat} />
                <span className="cat-count">{grouped[cat].length}</span>
              </div>
              {grouped[cat].map(item => (
                <div className="gear-row" key={item.id}>
                  <div className="gear-name">{item.name}</div>
                  <div><SeasonBadge season={item.season} /></div>
                  <div className="gear-notes">{item.notes}</div>
                  <div className="gear-actions">
                    <button className="icon-btn" onClick={() => startEdit(item)} aria-label="Edit"><Pencil size={14} /></button>
                    <button className="icon-btn" onClick={() => deleteItem(item.id)} aria-label="Delete"><Trash2 size={14} /></button>
                  </div>
                </div>
              ))}
            </div>
          ))}
        </div>
      )}

      <p className="hint">Spreadsheet columns recognized: Name, Category, Season (Summer/Winter, optional), Notes.</p>
    </section>
  );
}

function PeoplePacks({ gearItems, people, onChange }) {
  const [selectedPersonId, setSelectedPersonId] = useState(people[0]?.id ?? null);
  const [newPersonName, setNewPersonName] = useState("");
  const [addingPerson, setAddingPerson] = useState(false);

  useEffect(() => {
    if (!selectedPersonId && people.length) setSelectedPersonId(people[0].id);
    if (selectedPersonId && !people.find(p => p.id === selectedPersonId)) {
      setSelectedPersonId(people[0]?.id ?? null);
    }
  }, [people]);

  function addPerson() {
    if (!newPersonName.trim()) return;
    const p = { id: uid(), name: newPersonName.trim(), packs: [] };
    const next = [...people, p];
    onChange(next);
    setSelectedPersonId(p.id);
    setNewPersonName("");
    setAddingPerson(false);
  }

  function deletePerson(id) {
    onChange(people.filter(p => p.id !== id));
  }

  function updatePerson(id, updater) {
    onChange(people.map(p => p.id === id ? updater(p) : p));
  }

  const selected = people.find(p => p.id === selectedPersonId);

  return (
    <section>
      <div className="toolbar people-toolbar">
        <div className="spacer" />
        <CopyButton getText={() => buildPlainTextExport(people, gearItems)} />
        <button className="btn-ghost" onClick={() => downloadBlob("people-and-packs.pdf", buildPdfBlob(people, gearItems))}>
          <FileDown size={15} /> PDF export
        </button>
        <button className="btn-ghost" onClick={() => downloadBlob("people-and-packs.md", buildMarkdownExport(people, gearItems), "text/markdown")}>
          <Download size={15} /> Markdown export
        </button>
      </div>
      <div className="people-layout">
        <aside className="people-sidebar">
          <div className="sidebar-heading">People</div>
          {people.map(p => (
            <button
              key={p.id}
              className={"person-item" + (p.id === selectedPersonId ? " active" : "")}
              onClick={() => setSelectedPersonId(p.id)}
            >
              <span className="avatar">{p.name.slice(0, 2).toUpperCase()}</span>
              <span className="person-name">{p.name}</span>
              <span className="pack-count">{p.packs.length}</span>
            </button>
          ))}
          {addingPerson ? (
            <div className="add-person-form">
              <input
                autoFocus
                placeholder="Name"
                value={newPersonName}
                onChange={e => setNewPersonName(e.target.value)}
                onKeyDown={e => e.key === "Enter" && addPerson()}
              />
              <div className="form-actions">
                <button className="btn-ghost" onClick={() => setAddingPerson(false)}><X size={14} /></button>
                <button className="btn-primary" onClick={addPerson}><Check size={14} /></button>
              </div>
            </div>
          ) : (
            <button className="btn-ghost add-person-btn" onClick={() => setAddingPerson(true)}>
              <Plus size={15} /> Add person
            </button>
          )}
        </aside>

        <div className="people-main">
          {!selected ? (
            <div className="empty-state">
              <User size={22} />
              <p>No people yet. Add someone to start building their pack.</p>
            </div>
          ) : (
            <PersonPanel
              person={selected}
              gearItems={gearItems}
              onUpdate={(updater) => updatePerson(selected.id, updater)}
              onDelete={() => deletePerson(selected.id)}
            />
          )}
        </div>
      </div>
    </section>
  );
}

function PersonPanel({ person, gearItems, onUpdate, onDelete }) {
  function addPack() {
    onUpdate(p => ({
      ...p,
      packs: [...p.packs, { id: uid(), name: `Pack ${p.packs.length + 1}`, items: [] }]
    }));
  }

  function updatePack(packId, fn) {
    onUpdate(p => ({ ...p, packs: p.packs.map(pk => pk.id === packId ? fn(pk) : pk) }));
  }

  function deletePack(packId) {
    onUpdate(p => ({ ...p, packs: p.packs.filter(pk => pk.id !== packId) }));
  }

  return (
    <div>
      <div className="person-header">
        <h2>{person.name}</h2>
        <button className="btn-ghost" onClick={onDelete}><Trash2 size={14} /> Remove person</button>
      </div>
      {person.packs.length === 0 ? (
        <div className="empty-state">
          <Package size={20} />
          <p>No packs yet for {person.name}.</p>
        </div>
      ) : (
        person.packs.map(pack => (
          <PackCard
            key={pack.id}
            pack={pack}
            gearItems={gearItems}
            onChange={fn => updatePack(pack.id, fn)}
            onDelete={() => deletePack(pack.id)}
          />
        ))
      )}
      <button className="btn-primary add-pack-btn" onClick={addPack}><Plus size={15} /> Add pack</button>
    </div>
  );
}

function PackCard({ pack, gearItems, onChange, onDelete }) {
  const [pickerOpen, setPickerOpen] = useState(false);
  const [customOpen, setCustomOpen] = useState(false);
  const [customName, setCustomName] = useState("");
  const [nameEdit, setNameEdit] = useState(false);
  const [nameVal, setNameVal] = useState(pack.name);

  const itemsById = useMemo(() => {
    const m = {};
    gearItems.forEach(g => { m[g.id] = g; });
    return m;
  }, [gearItems]);

  function addGearBatch(gearIds) {
    const additions = gearIds
      .filter(gearId => !pack.items.find(it => it.gearId === gearId))
      .map(gearId => {
        const g = itemsById[gearId];
        return { id: uid(), gearId, packed: false, notes: (g && g.notes) || "" };
      });
    if (additions.length) {
      onChange(pk => ({ ...pk, items: [...pk.items, ...additions] }));
    }
    setPickerOpen(false);
  }

  function addCustomItem() {
    if (!customName.trim()) return;
    onChange(pk => ({ ...pk, items: [...pk.items, { id: uid(), custom: true, name: customName.trim(), packed: false, notes: "" }] }));
    setCustomName("");
    setCustomOpen(false);
  }

  function removeItem(itemId) {
    onChange(pk => ({ ...pk, items: pk.items.filter(it => it.id !== itemId) }));
  }

  function togglePacked(itemId) {
    onChange(pk => ({ ...pk, items: pk.items.map(it => it.id === itemId ? { ...it, packed: !it.packed } : it) }));
  }

  function setNotes(itemId, notes) {
    onChange(pk => ({ ...pk, items: pk.items.map(it => it.id === itemId ? { ...it, notes } : it) }));
  }

  function saveName() {
    onChange(pk => ({ ...pk, name: nameVal.trim() || pk.name }));
    setNameEdit(false);
  }

  return (
    <div className="pack-card">
      <div className="pack-card-top">
        <div className="pack-name-row">
          {nameEdit ? (
            <input
              autoFocus
              value={nameVal}
              onChange={e => setNameVal(e.target.value)}
              onBlur={saveName}
              onKeyDown={e => e.key === "Enter" && saveName()}
            />
          ) : (
            <h3 onClick={() => setNameEdit(true)}>{pack.name} <Pencil size={12} /></h3>
          )}
          <button className="icon-btn" onClick={onDelete} aria-label="Delete pack"><Trash2 size={14} /></button>
        </div>
      </div>

      {pack.items.length === 0 ? (
        <p className="pack-empty-hint">No gear assigned yet.</p>
      ) : (
        <div className="pack-items">
          {sortPackItems(pack.items, itemsById).map(it => {
            const g = it.custom ? null : itemsById[it.gearId];
            if (!it.custom && !g) return null;
            const name = it.custom ? it.name : g.name;
            const catLabel = it.custom ? "Custom" : g.category;
            return (
              <div className="pack-item-row" key={it.id}>
                <button className={"packed-check" + (it.packed ? " checked" : "")} onClick={() => togglePacked(it.id)} aria-label="Toggle packed">
                  {it.packed && <Check size={12} />}
                </button>
                <span className={"pack-item-name" + (it.packed ? " done" : "")}>{name}</span>
                <input
                  className="notes-input"
                  placeholder="Notes"
                  value={it.notes || ""}
                  onChange={e => setNotes(it.id, e.target.value)}
                />
                <Tag label={catLabel} />
                <button className="icon-btn" onClick={() => removeItem(it.id)} aria-label="Remove"><X size={14} /></button>
              </div>
            );
          })}
        </div>
      )}

      <div className="pack-add-row">
        {customOpen ? (
          <div className="custom-form">
            <input
              autoFocus
              placeholder="Custom item name"
              value={customName}
              onChange={e => setCustomName(e.target.value)}
              onKeyDown={e => e.key === "Enter" && addCustomItem()}
            />
            <button className="btn-ghost" onClick={() => { setCustomOpen(false); setCustomName(""); }}><X size={14} /></button>
            <button className="btn-primary" onClick={addCustomItem}><Check size={14} /></button>
          </div>
        ) : (
          <>
            <button className="btn-ghost add-gear-btn" onClick={() => setPickerOpen(true)}>
              <Plus size={14} /> Add gear
            </button>
            <button className="btn-ghost add-gear-btn" onClick={() => setCustomOpen(true)}>
              <Plus size={14} /> Add custom item
            </button>
          </>
        )}
      </div>

      {pickerOpen && (
        <GearPickerModal
          gearItems={gearItems}
          excludeIds={pack.items.filter(it => it.gearId).map(it => it.gearId)}
          onAdd={addGearBatch}
          onClose={() => setPickerOpen(false)}
        />
      )}
    </div>
  );
}

function GearPickerModal({ gearItems, excludeIds, onAdd, onClose }) {
  const [search, setSearch] = useState("");
  const [seasonFilter, setSeasonFilter] = useState("All");
  const [selected, setSelected] = useState(() => new Set());

  const excludeSet = useMemo(() => new Set(excludeIds), [excludeIds]);

  const available = useMemo(() => {
    return gearItems
      .filter(g => !excludeSet.has(g.id))
      .filter(g => g.name.toLowerCase().includes(search.toLowerCase()))
      .filter(g => {
        if (seasonFilter === "All") return true;
        return !g.season || g.season === seasonFilter;
      })
      .sort((a, b) => a.category.localeCompare(b.category) || a.name.localeCompare(b.name));
  }, [gearItems, excludeSet, search, seasonFilter]);

  const grouped = useMemo(() => {
    const map = {};
    available.forEach(g => {
      if (!map[g.category]) map[g.category] = [];
      map[g.category].push(g);
    });
    return map;
  }, [available]);

  function toggle(id) {
    setSelected(prev => {
      const next = new Set(prev);
      if (next.has(id)) next.delete(id); else next.add(id);
      return next;
    });
  }

  function selectAllVisible() {
    setSelected(prev => {
      const next = new Set(prev);
      available.forEach(g => next.add(g.id));
      return next;
    });
  }

  function confirmAdd() {
    onAdd(Array.from(selected));
  }

  return (
    <div className="modal-overlay" onClick={onClose}>
      <div className="modal-card" onClick={e => e.stopPropagation()}>
        <div className="modal-header">
          <h3>Add gear to pack</h3>
          <button className="icon-btn" onClick={onClose} aria-label="Close"><X size={16} /></button>
        </div>

        <div className="modal-controls">
          <div className="search-box">
            <Search size={14} />
            <input autoFocus placeholder="Search gear" value={search} onChange={e => setSearch(e.target.value)} />
          </div>
          <div className="season-switch">
            <button className={seasonFilter === "All" ? "season-toggle active" : "season-toggle"} onClick={() => setSeasonFilter("All")}>All</button>
            <button className={seasonFilter === "Summer" ? "season-toggle active" : "season-toggle"} onClick={() => setSeasonFilter("Summer")}><Sun size={13} /> Summer</button>
            <button className={seasonFilter === "Winter" ? "season-toggle active" : "season-toggle"} onClick={() => setSeasonFilter("Winter")}><Snowflake size={13} /> Winter</button>
          </div>
        </div>

        <div className="modal-list">
          {available.length === 0 ? (
            <div className="picker-empty">No matching items in the library.</div>
          ) : (
            Object.keys(grouped).sort().map(cat => (
              <div key={cat} className="modal-cat-group">
                <div className="modal-cat-heading">
                  <Tag label={cat} />
                </div>
                {grouped[cat].map(g => (
                  <label className="modal-row" key={g.id}>
                    <input
                      type="checkbox"
                      checked={selected.has(g.id)}
                      onChange={() => toggle(g.id)}
                    />
                    <span className="modal-row-name">{g.name}</span>
                    <SeasonBadge season={g.season} />
                  </label>
                ))}
              </div>
            ))
          )}
        </div>

        <div className="modal-footer">
          <button className="btn-ghost" onClick={selectAllVisible} disabled={available.length === 0}>Select all shown</button>
          <div className="spacer" />
          <button className="btn-ghost" onClick={onClose}>Cancel</button>
          <button className="btn-primary" onClick={confirmAdd} disabled={selected.size === 0}>
            <Check size={15} /> Add {selected.size > 0 ? selected.size : ""} item{selected.size === 1 ? "" : "s"}
          </button>
        </div>
      </div>
    </div>
  );
}
