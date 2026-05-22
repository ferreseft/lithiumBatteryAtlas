
import React, { useMemo, useState, useEffect } from "react";
import { createRoot } from "react-dom/client";
import { motion } from "framer-motion";
import * as XLSX from "xlsx";
import {
  Search, BatteryCharging, FileText, Microscope, Scale, AlertTriangle, Camera,
  Database, Info, Download, Ruler, ShieldAlert, Filter, BookOpen, ClipboardList,
  Upload, Image as ImageIcon, Columns3, TableProperties, X, Plus, Save, Trash2,
  RefreshCcw
} from "lucide-react";
import "./index.css";

const STORAGE_KEY = "liIonWeldAtlasEntriesV3";

const defaultAtlasEntries = [
  { id: "WP-001", category: "Power-tool pack", provenanceLayer: "Pack assembler / OEM family", cellFormat: "18650 cylindrical", method: "Resistance spot weld", pattern: "RSW_4_DOT_SQUARE", confidenceUse: "Strong for OEM-vs-rebuilt comparison when pack layout survives.", keyFeatures: ["Four-dot square pattern", "Nickel strip interconnect", "Repeatable terminal placement", "BMS-side orientation map"], cautions: "Do not attribute cell manufacturer from external nickel-strip welds alone." },
  { id: "WP-002", category: "Aftermarket/rebuilt pack", provenanceLayer: "Rebuilder / repair source", cellFormat: "18650 or 21700 cylindrical", method: "Resistance spot weld", pattern: "RSW_OFFSET_REWORK_SCARS", confidenceUse: "Strong for identifying prior installation or pack rebuild indicators.", keyFeatures: ["Old weld scars under new strip", "Offset second weld pattern", "Mixed cell date codes", "Hand-cut nickel strip"], cautions: "Old scars show prior use or rework, not necessarily fire causation." },
  { id: "WP-003", category: "E-bike / scooter pack", provenanceLayer: "Pack assembler / topology", cellFormat: "18650 or 21700 cylindrical", method: "Resistance spot weld", pattern: "RSW_STAGGERED", confidenceUse: "Moderate to strong when combined with full series-parallel layout.", keyFeatures: ["Staggered spot pairs", "Long nickel bus strips", "Parallel group mapping", "Thermistor and fuse routing"], cautions: "Many assemblers share similar staggered spot weld layouts." },
  { id: "WP-004", category: "EV module", provenanceLayer: "Module platform / automated line", cellFormat: "21700 cylindrical or prismatic", method: "Laser welding", pattern: "LASER_CIRCULAR_WOBBLE", confidenceUse: "High when compared to known module exemplars.", keyFeatures: ["Circular wobble path", "Repeatable coordinates", "Start-stop crater position", "Stamped busbar geometry"], cautions: "Country or factory attribution requires authenticated exemplars and records." },
  { id: "WP-005", category: "Pouch-cell module", provenanceLayer: "Cell maker or module assembler depending on weld location", cellFormat: "Pouch", method: "Ultrasonic metal weld", pattern: "ULTRASONIC_KNURL_RECTANGLE", confidenceUse: "High if sonotrode imprint is intact and compared to known exemplars.", keyFeatures: ["Rectangular knurled pad", "Tooth pitch", "Material thinning", "Tab stack alignment"], cautions: "Heat damage can erase or distort horn-imprint detail." },
  { id: "WP-006", category: "Battery module with wire bonds", provenanceLayer: "Module assembler / automation program", cellFormat: "Cylindrical or prismatic", method: "Ultrasonic wire or ribbon bonding", pattern: "ULTRASONIC_WIRE_BOND", confidenceUse: "High when bond map, wire diameter, and bond-foot geometry survive.", keyFeatures: ["Bond-foot shape", "Loop height", "Wire diameter", "Bond sequence map"], cautions: "Broken bonds may be damage from fire, impact, or disassembly." },
  { id: "WP-007", category: "Fire-altered cylindrical pack", provenanceLayer: "Indeterminate unless corroborated", cellFormat: "18650 or 21700 cylindrical", method: "Unknown or altered", pattern: "FIRE_ALTERED_OR_INDETERMINATE", confidenceUse: "Useful for exclusions only when major class features remain.", keyFeatures: ["Oxidized terminal", "Detached strip", "Melted insulation", "Partial weld scar"], cautions: "Do not over-attribute features created by heat exposure or suppression damage." },
  { id: "WP-008", category: "Portable power station", provenanceLayer: "Pack assembler / product family", cellFormat: "Prismatic or cylindrical", method: "Laser seam or resistance spot weld", pattern: "MIXED_PACK_LEVEL_JOINING", confidenceUse: "Moderate to high when combined with busbar, BMS, and enclosure evidence.", keyFeatures: ["Large current busbars", "Fuse links", "BMS harness routing", "Pack architecture"], cautions: "Multiple suppliers may build visually similar energy-storage packs." }
];

const glossary = [
  ["Class characteristic", "A feature shared by a group, such as weld method, strip width, cell format, or general spot count."],
  ["Subclass characteristic", "A feature produced by a shared tool, fixture, or short production run that may appear distinctive but is not unique."],
  ["Individual characteristic", "A random or progressive mark, such as unusual electrode damage, tool wear, or repeated asymmetry, that may support stronger association when validated."],
  ["Questioned evidence", "The unknown battery, pack, module, cell, weld, or component recovered from a case scene or claim file."],
  ["Known exemplar", "An authenticated comparison item with known manufacturer, model, source, and condition."],
  ["Provenance layer", "The source level being evaluated: cell manufacturer, pack assembler, module supplier, OEM, rebuilder, or device family."]
];

const workflowSteps = [
  ["1. Preserve and photograph before teardown", Camera, "Record the pack exterior, labels, enclosure, BMS position, cell layout, and visible welds before moving strips or cells."],
  ["2. Establish orientation and map cells", Ruler, "Assign row, column, polarity, side, and terminal orientation before logging weld measurements."],
  ["3. Classify the joining method", Microscope, "Identify resistance spot welding, laser welding, ultrasonic tab welding, wire bonding, soldering, or fire-altered/unknown status."],
  ["4. Measure repeatable features", Scale, "Capture spot count, diameter, pitch, strip width, heat-affected zone, indentation, weld path, tool imprint, and old weld scars."],
  ["5. Compare to known exemplars", Database, "Compare questioned evidence against authenticated exemplars using class, subclass, and possible individual characteristics."],
  ["6. Report with graded language", FileText, "Use terms such as inconsistent, limited similarity, consistent with, or strongly consistent with. Avoid unsupported absolute source claims."]
];

const weldTypes = [
  ["Resistance spot weld", "••    ••", "Common on cylindrical cell nickel strips. Useful for pack layout, assembler comparison, and rebuild detection."],
  ["Laser circular wobble", "◎  ◎  ◎", "Common on automated busbars. Weld path and start-stop artifacts can be highly useful with known exemplars."],
  ["Ultrasonic knurl pad", "▦▦▦", "Common on pouch/prismatic tab stacks. Horn or sonotrode imprint can preserve tool geometry."],
  ["Wire/ribbon bond", "⌒ ⌒ ⌒", "Common in some modules. Bond-foot geometry, loop shape, and map sequence are important."]
];

const fieldMap = {
  id: ["id", "entry id", "evidence_id", "atlas id", "reference id"],
  category: ["category", "device_category", "product category", "device category"],
  provenanceLayer: ["provenanceLayer", "provenance layer", "provenance_layer", "usually identifies"],
  cellFormat: ["cellFormat", "cell format", "cell_format"],
  method: ["method", "joining_method", "joining method", "weld method"],
  pattern: ["pattern", "pattern_type", "pattern type", "weld pattern"],
  confidenceUse: ["confidenceUse", "confidence use", "forensic value", "confidence_level", "confidence level"],
  keyFeatures: ["keyFeatures", "key features", "features", "toolmark_notes", "toolmark notes"],
  cautions: ["cautions", "caution", "notes", "limitations"]
};

function normalizeKey(key) { return String(key || "").trim().toLowerCase().replace(/\s+/g, " "); }
function findValue(row, candidates) {
  const entries = Object.entries(row);
  for (const candidate of candidates) {
    const found = entries.find(([key]) => normalizeKey(key) === normalizeKey(candidate));
    if (found && found[1] !== undefined && found[1] !== null && String(found[1]).trim() !== "") return found[1];
  }
  return "";
}
function rowToAtlasEntry(row, index) {
  const entry = {};
  Object.entries(fieldMap).forEach(([target, candidates]) => { entry[target] = findValue(row, candidates); });
  return {
    id: entry.id || `IMPORT-${String(index + 1).padStart(3, "0")}`,
    category: entry.category || "Imported atlas entry",
    provenanceLayer: entry.provenanceLayer || "Not specified",
    cellFormat: entry.cellFormat || "Not specified",
    method: entry.method || "Not specified",
    pattern: entry.pattern || "Not specified",
    confidenceUse: entry.confidenceUse || "Imported from spreadsheet. Review before relying on this entry.",
    keyFeatures: Array.isArray(entry.keyFeatures) ? entry.keyFeatures : String(entry.keyFeatures || "").split(/[;|,]/).map((x) => x.trim()).filter(Boolean),
    cautions: entry.cautions || "Review imported entry for completeness and source authentication."
  };
}
function cls(...items) { return items.filter(Boolean).join(" "); }
function Card({ children, className = "" }) { return <div className={cls("rounded-2xl border border-slate-200 bg-white shadow-sm", className)}>{children}</div>; }
function Button({ children, className = "", ...props }) { return <button className={cls("inline-flex items-center justify-center rounded-2xl bg-slate-950 px-4 py-3 text-sm font-semibold text-white shadow-sm transition hover:bg-slate-800 disabled:opacity-40", className)} {...props}>{children}</button>; }
function Badge({ children }) { return <span className="rounded-full bg-slate-100 px-3 py-1 text-xs font-medium text-slate-700">{children}</span>; }
function SectionHeader({ icon: Icon, title, subtitle }) {
  return <div className="mb-5 flex items-start gap-3"><div className="rounded-2xl bg-slate-900 p-3 text-white shadow-sm"><Icon className="h-5 w-5" /></div><div><h2 className="text-2xl font-semibold tracking-tight text-slate-950">{title}</h2><p className="mt-1 max-w-3xl text-sm leading-6 text-slate-600">{subtitle}</p></div></div>;
}

function nextWpId(entries) {
  const nums = entries.map(e => String(e.id || "").match(/^WP-(\d+)$/i)).filter(Boolean).map(m => Number(m[1]));
  const next = nums.length ? Math.max(...nums) + 1 : entries.length + 1;
  return `WP-${String(next).padStart(3, "0")}`;
}
function csvEscape(value) {
  const text = Array.isArray(value) ? value.join("; ") : String(value ?? "");
  if (/[",\n]/.test(text)) return `"${text.replace(/"/g, '""')}"`;
  return text;
}
function downloadAtlasCsv(entries) {
  const headers = ["id", "category", "provenance layer", "cell format", "method", "pattern", "confidence use", "key features", "cautions"];
  const rows = entries.map(e => [e.id, e.category, e.provenanceLayer, e.cellFormat, e.method, e.pattern, e.confidenceUse, e.keyFeatures, e.cautions]);
  const csv = [headers, ...rows].map(row => row.map(csvEscape).join(",")).join("\n");
  const blob = new Blob([csv], { type: "text/csv;charset=utf-8" });
  const url = URL.createObjectURL(blob);
  const a = document.createElement("a");
  a.href = url;
  a.download = "li-ion-weld-pattern-atlas-updated.csv";
  document.body.appendChild(a);
  a.click();
  a.remove();
  URL.revokeObjectURL(url);
}

function DataImport({ atlasEntries, setAtlasEntries }) {
  const [status, setStatus] = useState("Using saved or built-in atlas.");
  const [previewRows, setPreviewRows] = useState([]);
  async function importFile(file) {
    if (!file) return;
    const data = await file.arrayBuffer();
    const workbook = XLSX.read(data, { type: "array" });
    const preferredSheet = workbook.SheetNames.find((name) => /atlas|entry|reference/i.test(name)) || workbook.SheetNames[0];
    const sheet = workbook.Sheets[preferredSheet];
    const rows = XLSX.utils.sheet_to_json(sheet, { defval: "" });
    const imported = rows.map(rowToAtlasEntry).filter((entry) => entry.category || entry.method || entry.pattern);
    if (imported.length === 0) { setStatus("No usable rows found. Check that the spreadsheet has a header row."); return; }
    setAtlasEntries(imported);
    setPreviewRows(rows.slice(0, 5));
    setStatus(`Imported and saved ${imported.length} atlas entries from ${file.name}, sheet "${preferredSheet}".`);
  }
  return <div><SectionHeader icon={TableProperties} title="Atlas Data Import" subtitle="Upload the Excel workbook or CSV export of the atlas. The app maps common column names into searchable web entries." />
    <div className="grid gap-5 lg:grid-cols-[0.8fr_1.2fr]"><Card><div className="p-5"><label className="flex cursor-pointer flex-col items-center justify-center gap-3 rounded-2xl border border-dashed border-slate-300 bg-slate-50 px-4 py-10 text-center text-sm font-medium text-slate-600 hover:bg-slate-100"><Upload className="h-7 w-7" /><span>Upload .xlsx, .xls, or .csv atlas file</span><input type="file" accept=".xlsx,.xls,.csv" className="hidden" onChange={(event) => importFile(event.target.files?.[0])} /></label><div className="mt-4 rounded-xl bg-slate-100 p-3 text-sm leading-6 text-slate-700">{status}</div><div className="mt-4 flex flex-wrap gap-2"><Button onClick={() => { setAtlasEntries(defaultAtlasEntries); setStatus("Restored and saved built-in starter atlas."); setPreviewRows([]); }}>Restore starter data</Button><Button onClick={() => downloadAtlasCsv(atlasEntries)} className="bg-slate-700 hover:bg-slate-600"><Download className="mr-2 h-4 w-4" />Export CSV</Button></div></div></Card>
    <Card><div className="p-5"><h3 className="text-lg font-semibold text-slate-950">Loaded atlas summary</h3><div className="mt-4 grid gap-3 sm:grid-cols-3"><div className="rounded-2xl bg-slate-100 p-4"><div className="text-3xl font-bold">{atlasEntries.length}</div><div className="text-xs uppercase tracking-wide text-slate-500">entries</div></div><div className="rounded-2xl bg-slate-100 p-4"><div className="text-3xl font-bold">{new Set(atlasEntries.map((x) => x.method)).size}</div><div className="text-xs uppercase tracking-wide text-slate-500">methods</div></div><div className="rounded-2xl bg-slate-100 p-4"><div className="text-3xl font-bold">{new Set(atlasEntries.map((x) => x.cellFormat)).size}</div><div className="text-xs uppercase tracking-wide text-slate-500">cell formats</div></div></div><p className="mt-4 text-sm leading-6 text-slate-600">Recommended columns: id, category, provenance layer, cell format, method, pattern, key features, cautions, confidence use.</p>{previewRows.length > 0 && <div className="mt-4 max-h-72 overflow-auto rounded-xl border"><table className="w-full text-left text-xs"><thead className="bg-slate-100"><tr>{Object.keys(previewRows[0]).slice(0, 8).map((key) => <th key={key} className="whitespace-nowrap px-3 py-2 font-semibold">{key}</th>)}</tr></thead><tbody>{previewRows.map((row, index) => <tr key={index} className="border-t">{Object.keys(previewRows[0]).slice(0, 8).map((key) => <td key={key} className="max-w-48 truncate px-3 py-2">{String(row[key] || "")}</td>)}</tr>)}</tbody></table></div>}</div></Card></div></div>;
}

function AddEntry({ atlasEntries, setAtlasEntries }) {
  const blank = () => ({ id: nextWpId(atlasEntries), category: "", provenanceLayer: "", cellFormat: "", method: "Resistance spot weld", pattern: "", confidenceUse: "", keyFeaturesText: "", cautions: "" });
  const [form, setForm] = useState(blank());
  const [message, setMessage] = useState("");
  useEffect(() => { if (!form.id) setForm(blank()); }, [atlasEntries.length]);
  const update = (key, value) => setForm(cur => ({ ...cur, [key]: value }));
  function saveEntry() {
    if (!form.category.trim() || !form.method.trim() || !form.pattern.trim()) {
      setMessage("Category, method, and pattern are required.");
      return;
    }
    const entry = {
      id: form.id.trim() || nextWpId(atlasEntries),
      category: form.category.trim(),
      provenanceLayer: form.provenanceLayer.trim() || "Not specified",
      cellFormat: form.cellFormat.trim() || "Not specified",
      method: form.method.trim(),
      pattern: form.pattern.trim(),
      confidenceUse: form.confidenceUse.trim() || "Review against authenticated known exemplars before relying on this entry.",
      keyFeatures: form.keyFeaturesText.split(/[;\n]/).map(x => x.trim()).filter(Boolean),
      cautions: form.cautions.trim() || "Do not overstate source attribution without known exemplars and corroborating evidence."
    };
    setAtlasEntries(cur => {
      const withoutSame = cur.filter(e => e.id !== entry.id);
      return [...withoutSame, entry];
    });
    setMessage(`Saved ${entry.id}. It is now searchable in the Atlas tab and stored in this browser.`);
    setForm({ ...blank(), id: nextWpId([...atlasEntries, entry]) });
  }
  return <div><SectionHeader icon={Plus} title="Add New Atlas Entry" subtitle="Create permanent entries in this browser without editing code. Export CSV when you want to preserve or share the updated atlas." />
    <div className="grid gap-5 lg:grid-cols-[1fr_0.75fr]"><Card><div className="grid gap-4 p-5 md:grid-cols-2">
      {[ ["Entry ID", "id"], ["Category", "category"], ["Provenance layer", "provenanceLayer"], ["Cell format", "cellFormat"], ["Joining method", "method"], ["Pattern code", "pattern"] ].map(([label,key]) => <label key={key} className="text-sm font-medium text-slate-700">{label}<input value={form[key]} onChange={e => update(key, e.target.value)} className="mt-2 w-full rounded-xl border px-3 py-2 text-sm outline-none focus:ring-2 focus:ring-slate-300" /></label>)}
      <label className="md:col-span-2 text-sm font-medium text-slate-700">Confidence use / forensic value<textarea value={form.confidenceUse} onChange={e => update("confidenceUse", e.target.value)} rows={3} className="mt-2 w-full rounded-xl border px-3 py-2 text-sm outline-none focus:ring-2 focus:ring-slate-300" placeholder="Example: High when compared to known module exemplars with matching busbar geometry and weld path." /></label>
      <label className="md:col-span-2 text-sm font-medium text-slate-700">Key features, one per line or separated by semicolons<textarea value={form.keyFeaturesText} onChange={e => update("keyFeaturesText", e.target.value)} rows={4} className="mt-2 w-full rounded-xl border px-3 py-2 text-sm outline-none focus:ring-2 focus:ring-slate-300" placeholder={"Four-dot square pattern\nNickel strip width\nOld weld scars"} /></label>
      <label className="md:col-span-2 text-sm font-medium text-slate-700">Cautions<textarea value={form.cautions} onChange={e => update("cautions", e.target.value)} rows={3} className="mt-2 w-full rounded-xl border px-3 py-2 text-sm outline-none focus:ring-2 focus:ring-slate-300" placeholder="Do not attribute cell manufacturer from external pack-level welds alone." /></label>
      <div className="md:col-span-2 flex flex-wrap gap-2"><Button onClick={saveEntry}><Save className="mr-2 h-4 w-4" />Save entry</Button><Button className="bg-slate-700 hover:bg-slate-600" onClick={() => setForm(blank())}><RefreshCcw className="mr-2 h-4 w-4" />Reset form</Button><Button className="bg-slate-700 hover:bg-slate-600" onClick={() => downloadAtlasCsv(atlasEntries)}><Download className="mr-2 h-4 w-4" />Export CSV</Button></div>
      {message && <div className="md:col-span-2 rounded-xl bg-emerald-50 p-3 text-sm text-emerald-900">{message}</div>}
    </div></Card>
    <Card><div className="p-5"><h3 className="text-lg font-semibold text-slate-950">Recently saved entries</h3><div className="mt-4 space-y-3">{atlasEntries.slice(-5).reverse().map(entry => <div key={entry.id} className="rounded-xl border p-3"><div className="flex items-start justify-between gap-3"><div><div className="font-semibold text-slate-900">{entry.id} · {entry.category}</div><div className="mt-1 text-xs text-slate-500">{entry.method} · {entry.pattern}</div></div><button title="Delete entry" className="rounded-full p-2 text-slate-500 hover:bg-red-50 hover:text-red-700" onClick={() => setAtlasEntries(cur => cur.filter(e => e.id !== entry.id))}><Trash2 className="h-4 w-4" /></button></div></div>)}</div><div className="mt-4 rounded-xl bg-amber-50 p-3 text-sm leading-6 text-amber-900"><AlertTriangle className="mr-2 inline h-4 w-4" />Entries are saved in browser local storage. Use Export CSV for a backup or to move entries to another computer.</div></div></Card></div></div>;
}

function Atlas({ atlasEntries }) {
  const [query, setQuery] = useState("");
  const [method, setMethod] = useState("All");
  const methods = ["All", ...Array.from(new Set(atlasEntries.map((entry) => entry.method)))];
  const filtered = useMemo(() => {
    const q = query.trim().toLowerCase();
    return atlasEntries.filter((entry) => {
      const methodMatch = method === "All" || entry.method === method;
      const text = [entry.id, entry.category, entry.provenanceLayer, entry.cellFormat, entry.method, entry.pattern, entry.confidenceUse, ...(entry.keyFeatures || []), entry.cautions].join(" ").toLowerCase();
      return methodMatch && (!q || text.includes(q));
    });
  }, [query, method, atlasEntries]);
  return <div><SectionHeader icon={Database} title="Reference Atlas" subtitle="Search weld-pattern exemplars by cell format, joining method, provenance layer, product category, or forensic feature." />
    <div className="mb-5 grid gap-3 md:grid-cols-[1fr_260px]"><div className="flex items-center gap-2 rounded-2xl border bg-white px-4 py-3 shadow-sm"><Search className="h-4 w-4 text-slate-500" /><input className="w-full bg-transparent text-sm outline-none" placeholder="Search: rebuilt, laser, 21700, ultrasonic, old scars..." value={query} onChange={(event) => setQuery(event.target.value)} /></div><div className="flex items-center gap-2 rounded-2xl border bg-white px-4 py-3 shadow-sm"><Filter className="h-4 w-4 text-slate-500" /><select className="w-full bg-transparent text-sm outline-none" value={method} onChange={(event) => setMethod(event.target.value)}>{methods.map((item) => <option key={item}>{item}</option>)}</select></div></div>
    <div className="mb-4 flex justify-end"><Button className="bg-slate-700 hover:bg-slate-600" onClick={() => downloadAtlasCsv(atlasEntries)}><Download className="mr-2 h-4 w-4" />Export current atlas</Button></div>
    <div className="grid gap-4 lg:grid-cols-2">{filtered.map((entry) => <Card key={entry.id} className="overflow-hidden"><div className="p-5"><div className="mb-3 flex flex-wrap items-center gap-2"><Badge>{entry.id}</Badge><Badge>{entry.method}</Badge><Badge>{entry.cellFormat}</Badge></div><h3 className="text-lg font-semibold text-slate-950">{entry.category}</h3><p className="mt-1 text-sm text-slate-600"><strong>Provenance layer:</strong> {entry.provenanceLayer}</p><p className="mt-1 text-sm text-slate-600"><strong>Pattern:</strong> {entry.pattern}</p><p className="mt-3 text-sm leading-6 text-slate-700">{entry.confidenceUse}</p><div className="mt-4 flex flex-wrap gap-2">{(entry.keyFeatures || []).map((feature) => <Badge key={feature}>{feature}</Badge>)}</div><div className="mt-4 rounded-xl bg-amber-50 p-3 text-sm leading-6 text-amber-900"><AlertTriangle className="mr-2 inline h-4 w-4" />{entry.cautions}</div></div></Card>)}</div></div>;
}

function ImageUploadPanel({ title, images, setImages }) {
  const handleImages = (files) => { const mapped = Array.from(files || []).map((file) => ({ name: file.name, url: URL.createObjectURL(file) })); setImages((current) => [...current, ...mapped]); };
  return <Card><div className="p-5"><h3 className="flex items-center gap-2 text-lg font-semibold text-slate-950"><ImageIcon className="h-5 w-5" />{title}</h3><label className="mt-4 flex cursor-pointer flex-col items-center justify-center gap-3 rounded-2xl border border-dashed border-slate-300 bg-slate-50 px-4 py-8 text-center text-sm font-medium text-slate-600 hover:bg-slate-100"><Upload className="h-6 w-6" /><span>Upload one or more images</span><input type="file" multiple accept="image/*" className="hidden" onChange={(event) => handleImages(event.target.files)} /></label><div className="mt-4 grid gap-3">{images.length === 0 && <div className="rounded-xl bg-slate-100 p-4 text-sm text-slate-600">No images loaded.</div>}{images.map((image, index) => <div key={`${image.name}-${index}`} className="overflow-hidden rounded-xl border bg-white"><img src={image.url} alt={image.name} className="h-64 w-full object-contain bg-slate-950" /><div className="flex items-center justify-between gap-2 p-3 text-xs font-medium text-slate-600"><span className="truncate">{image.name}</span><button className="rounded-full p-1 hover:bg-slate-100" onClick={() => setImages((current) => current.filter((_, i) => i !== index))}><X className="h-4 w-4" /></button></div></div>)}</div></div></Card>;
}
function ImageCompare() {
  const [questioned, setQuestioned] = useState([]); const [known, setKnown] = useState([]); const [notes, setNotes] = useState("");
  return <div><SectionHeader icon={Columns3} title="Side-by-Side Image Comparison" subtitle="Upload questioned-evidence photos and known-exemplar photos for visual review. Images stay local in the browser unless you add backend storage later." /><div className="grid gap-5 lg:grid-cols-2"><ImageUploadPanel title="Questioned Evidence Images" images={questioned} setImages={setQuestioned} /><ImageUploadPanel title="Known Exemplar Images" images={known} setImages={setKnown} /></div><Card className="mt-5"><div className="p-5"><h3 className="text-lg font-semibold text-slate-950">Comparison notes</h3><textarea className="mt-3 w-full rounded-xl border px-3 py-2 text-sm outline-none focus:ring-2 focus:ring-slate-300" rows={6} value={notes} onChange={(event) => setNotes(event.target.value)} placeholder="Example: QE shows four-dot RSW square pattern with approx. similar pitch to KE-002, but strip width and BMS-side routing differ. Fire alteration obscures terminal C4." /><div className="mt-3 rounded-xl bg-amber-50 p-3 text-sm leading-6 text-amber-900">This panel is for examiner review only. It does not perform image matching, provenance attribution, or automated root-cause analysis.</div></div></Card></div>;
}
function Intake() {
  const [form, setForm] = useState({ evidenceId: "QE-001", deviceCategory: "Power-tool pack", condition: "Fire damaged", joiningMethod: "Resistance spot weld", cellFormat: "18650 cylindrical", packConfig: "5S2P", oldScars: "Unknown", notes: "" });
  const update = (key, value) => setForm((current) => ({ ...current, [key]: value }));
  return <div><SectionHeader icon={ClipboardList} title="Questioned Evidence Intake" subtitle="Enter the minimum case details before conducting weld-pattern comparison." /><div className="grid gap-5 lg:grid-cols-[1.2fr_0.8fr]"><Card><div className="grid gap-4 p-5 md:grid-cols-2">{[["Evidence ID", "evidenceId"], ["Device category", "deviceCategory"], ["Condition", "condition"], ["Joining method", "joiningMethod"], ["Cell format", "cellFormat"], ["Pack configuration", "packConfig"]].map(([label, key]) => <label key={key} className="text-sm font-medium text-slate-700">{label}<input value={form[key]} onChange={(event) => update(key, event.target.value)} className="mt-2 w-full rounded-xl border px-3 py-2 text-sm outline-none focus:ring-2 focus:ring-slate-300" /></label>)}<label className="text-sm font-medium text-slate-700">Old weld scars observed?<select value={form.oldScars} onChange={(event) => update("oldScars", event.target.value)} className="mt-2 w-full rounded-xl border px-3 py-2 text-sm outline-none focus:ring-2 focus:ring-slate-300"><option>Unknown</option><option>Yes</option><option>No</option><option>Indeterminate due to fire damage</option></select></label><label className="md:col-span-2 text-sm font-medium text-slate-700">Notes<textarea value={form.notes} onChange={(event) => update("notes", event.target.value)} rows={4} className="mt-2 w-full rounded-xl border px-3 py-2 text-sm outline-none focus:ring-2 focus:ring-slate-300" placeholder="Case observations, image references, measurement notes..." /></label></div></Card><Card className="bg-slate-950 text-white"><div className="p-5"><h3 className="text-lg font-semibold">Current Intake Summary</h3><div className="mt-4 space-y-3 text-sm text-slate-200">{Object.entries(form).filter(([key]) => key !== "notes").map(([key, value]) => <div key={key} className="flex justify-between gap-4 border-b border-white/10 pb-2"><span className="capitalize text-slate-400">{key.replace(/([A-Z])/g, " $1")}</span><span className="text-right font-medium">{value}</span></div>)}</div><div className="mt-5 rounded-xl bg-white/10 p-3 text-sm leading-6 text-slate-200">Preserve all images, scale references, and measurement notes before removing any busbars, nickel strips, or cell groups.</div></div></Card></div></div>;
}
function Comparison() {
  const [scores, setScores] = useState({ joiningMethod: 2, cellFormat: 2, layout: 1, weldPattern: 1, stripGeometry: 1, toolmarks: 0, fireDamage: -1 });
  const total = Object.values(scores).reduce((sum, value) => sum + Number(value), 0); let conclusion = "Indeterminate"; if (total <= -2) conclusion = "Inconsistent"; else if (total >= 7) conclusion = "Strongly consistent with"; else if (total >= 4) conclusion = "Consistent with"; else if (total >= 1) conclusion = "Limited similarity";
  const rows = [["joiningMethod", "Joining method"], ["cellFormat", "Cell format"], ["layout", "Pack / module layout"], ["weldPattern", "Weld count, pitch, and pattern"], ["stripGeometry", "Strip or busbar geometry"], ["toolmarks", "Potential individual toolmarks"], ["fireDamage", "Fire alteration penalty"]];
  return <div><SectionHeader icon={Scale} title="Comparison Scoring" subtitle="A structured scoring aid for internal triage. It does not replace examiner judgment, known-source exemplars, or formal validation." /><div className="grid gap-5 lg:grid-cols-[1fr_360px]"><Card><div className="p-5"><div className="grid gap-3">{rows.map(([key, label]) => <div key={key} className="grid gap-2 rounded-xl border p-3 md:grid-cols-[1fr_170px] md:items-center"><div><div className="text-sm font-semibold text-slate-800">{label}</div><div className="text-xs text-slate-500">Score from -2 to +2 based on questioned-to-known comparison.</div></div><select className="rounded-xl border px-3 py-2 text-sm outline-none focus:ring-2 focus:ring-slate-300" value={scores[key]} onChange={(event) => setScores((current) => ({ ...current, [key]: Number(event.target.value) }))}><option value={-2}>-2 incompatible</option><option value={-1}>-1 different/altered</option><option value={0}>0 not observed</option><option value={1}>+1 similar</option><option value={2}>+2 highly similar</option></select></div>)}</div></div></Card><Card><div className="p-5"><h3 className="text-lg font-semibold text-slate-950">Suggested conclusion</h3><div className="mt-5 rounded-3xl bg-slate-950 p-6 text-center text-white"><div className="text-5xl font-bold">{total}</div><div className="mt-2 text-sm uppercase tracking-wide text-slate-400">comparison score</div></div><div className="mt-5 rounded-2xl bg-slate-100 p-4 text-center text-xl font-semibold text-slate-950">{conclusion}</div><p className="mt-4 text-sm leading-6 text-slate-600">Report language should remain graded. Avoid claiming a specific country, factory, or unique source unless supported by authenticated exemplars and corroborating records.</p></div></Card></div></div>;
}
function Glossary() {
  return <div><SectionHeader icon={BookOpen} title="Definitions and Visual Guide" subtitle="Key definitions and simplified weld-pattern visuals for quick field reference. Replace these with real exemplar images as the reference library grows." /><div className="mb-6 grid gap-4 md:grid-cols-2 lg:grid-cols-4">{weldTypes.map(([name, visual, description]) => <Card key={name}><div className="p-5"><div className="mb-3 rounded-2xl bg-slate-100 p-5 text-center text-4xl font-bold tracking-widest text-slate-800">{visual}</div><h3 className="text-base font-semibold text-slate-950">{name}</h3><p className="mt-2 text-sm leading-6 text-slate-600">{description}</p></div></Card>)}</div><div className="grid gap-4 md:grid-cols-2">{glossary.map(([term, definition]) => <Card key={term}><div className="p-5"><h3 className="text-base font-semibold text-slate-950">{term}</h3><p className="mt-2 text-sm leading-6 text-slate-600">{definition}</p></div></Card>)}</div></div>;
}
function Workflow() {
  return <div><SectionHeader icon={FileText} title="SOP Workflow" subtitle="A condensed interactive version of the written procedure for using the weld-pattern atlas during evidence review." /><div className="grid gap-4 md:grid-cols-2 lg:grid-cols-3">{workflowSteps.map(([title, Icon, text], index) => <motion.div key={title} initial={{ opacity: 0, y: 12 }} animate={{ opacity: 1, y: 0 }} transition={{ delay: index * 0.04 }}><Card className="h-full"><div className="p-5"><div className="mb-4 flex h-11 w-11 items-center justify-center rounded-2xl bg-slate-900 text-white"><Icon className="h-5 w-5" /></div><h3 className="text-base font-semibold text-slate-950">{title}</h3><p className="mt-2 text-sm leading-6 text-slate-600">{text}</p></div></Card></motion.div>)}</div><div className="mt-6 grid gap-4 lg:grid-cols-2"><Card className="border-amber-200 bg-amber-50"><div className="p-5"><h3 className="flex items-center gap-2 text-base font-semibold text-amber-950"><ShieldAlert className="h-5 w-5" />Safety note</h3><p className="mt-2 text-sm leading-6 text-amber-900">Fire-damaged or mechanically damaged lithium-ion cells may retain energy and may reignite. Do not disassemble energized or unstable packs without appropriate controls.</p></div></Card><Card className="border-sky-200 bg-sky-50"><div className="p-5"><h3 className="flex items-center gap-2 text-base font-semibold text-sky-950"><Info className="h-5 w-5" />Attribution note</h3><p className="mt-2 text-sm leading-6 text-sky-900">Weld patterns are strongest when combined with pack layout, BMS design, cell markings, strip geometry, and known-source exemplars. Use caution with fire-altered features.</p></div></Card></div></div>;
}
function App() {
  const [activeTab, setActiveTab] = useState("atlas");
  const [atlasEntries, setAtlasEntriesRaw] = useState(() => {
    try { const saved = localStorage.getItem(STORAGE_KEY); return saved ? JSON.parse(saved) : defaultAtlasEntries; } catch { return defaultAtlasEntries; }
  });
  const setAtlasEntries = (updater) => {
    setAtlasEntriesRaw(cur => {
      const next = typeof updater === "function" ? updater(cur) : updater;
      localStorage.setItem(STORAGE_KEY, JSON.stringify(next));
      return next;
    });
  };
  const tabs = [["atlas", "Atlas", Database], ["add", "Add Entry", Plus], ["data", "Import/Export", TableProperties], ["intake", "Intake", ClipboardList], ["images", "Image Compare", Columns3], ["comparison", "Score", Scale], ["workflow", "SOP", FileText], ["glossary", "Glossary", BookOpen]];
  return <div className="min-h-screen bg-slate-50 text-slate-950"><header className="border-b bg-white/90 backdrop-blur"><div className="mx-auto flex max-w-7xl flex-col gap-5 px-5 py-6 lg:flex-row lg:items-center lg:justify-between"><div className="flex items-start gap-4"><div className="rounded-3xl bg-slate-950 p-4 text-white shadow-sm"><BatteryCharging className="h-7 w-7" /></div><div><h1 className="text-3xl font-bold tracking-tight">Li-ion Weld Pattern Reference Atlas</h1><p className="mt-1 max-w-3xl text-sm leading-6 text-slate-600">Interactive forensic aid for battery pack provenance, weld-pattern comparison, evidence intake, editable atlas entries, Excel import, and side-by-side image review.</p></div></div><Button onClick={() => window.print()}><Download className="mr-2 h-4 w-4" /> Print / save report</Button></div></header><main className="mx-auto max-w-7xl px-5 py-6"><div className="no-print mb-6 flex flex-wrap gap-2 rounded-3xl bg-white p-2 shadow-sm ring-1 ring-slate-200">{tabs.map(([id, label, Icon]) => <button key={id} onClick={() => setActiveTab(id)} className={cls("flex items-center gap-2 rounded-2xl px-4 py-3 text-sm font-medium transition", activeTab === id ? "bg-slate-950 text-white shadow-sm" : "text-slate-600 hover:bg-slate-100")}><Icon className="h-4 w-4" />{label}</button>)}</div><motion.div key={activeTab} initial={{ opacity: 0, y: 8 }} animate={{ opacity: 1, y: 0 }} transition={{ duration: 0.2 }}>{activeTab === "atlas" && <Atlas atlasEntries={atlasEntries} />}{activeTab === "add" && <AddEntry atlasEntries={atlasEntries} setAtlasEntries={setAtlasEntries} />}{activeTab === "data" && <DataImport atlasEntries={atlasEntries} setAtlasEntries={setAtlasEntries} />}{activeTab === "intake" && <Intake />}{activeTab === "images" && <ImageCompare />}{activeTab === "comparison" && <Comparison />}{activeTab === "workflow" && <Workflow />}{activeTab === "glossary" && <Glossary />}</motion.div></main><footer className="mx-auto max-w-7xl px-5 pb-8 pt-3 text-xs leading-6 text-slate-500">This prototype is a forensic workflow aid, not an automated source-identification decision engine. Entries saved in browser local storage should be exported to CSV for backup.</footer></div>;
}

createRoot(document.getElementById("root")).render(<App />);
