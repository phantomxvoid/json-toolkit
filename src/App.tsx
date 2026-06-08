import { useEffect, useMemo, useRef, useState } from "react";
import Editor from "@monaco-editor/react";
import JsonView from "@uiw/react-json-view";
import { vscodeTheme } from "@uiw/react-json-view/vscode";

import {
  Wand2,
  Minimize2,
  BadgeCheck,
  Copy as CopyIcon,
  Download,
  GitCompareArrows,
  Trash2,
  Moon,
  Sun,
  Plus,
  Minus,
  ArrowRightLeft,
  CheckCircle2,
  Braces,
  FileSpreadsheet,
  X,
  Search,
  Wrench,
  Table,
  Code2,
} from "lucide-react";

import { JSONPath } from "jsonpath-plus";
import { jsonrepair } from "jsonrepair";

import { diffJson, formatDiffAsText, type DiffEntry, type DiffResult } from "./utils/diff";
import { formatJson, minifyJson } from "./utils/formatter";
import { validateJson } from "./utils/validator";
import { downloadJson } from "./utils/downloader";
import { jsonToXml } from "./utils/xmlTransformer";
import { prepareExcelData, downloadExcel, type ExcelData } from "./utils/excelExporter";
import { jsonToCsv } from "./utils/csvExporter";
import { generateJsonSchema } from "./utils/schemaGenerator";

type Theme = "dark" | "light";

const getInitialTheme = (): Theme => {
  const storedTheme = localStorage.getItem("theme");
  if (storedTheme === "dark" || storedTheme === "light") return storedTheme;
  return window.matchMedia("(prefers-color-scheme: dark)").matches ? "dark" : "light";
};

// ─── Diff entry card ─────────────────────────────────────────────────────────

function DiffCard({ entry, darkMode }: { entry: DiffEntry; darkMode: boolean }) {
  const isAdded = entry.type === "added";
  const isRemoved = entry.type === "removed";
  const isChanged = entry.type === "changed";

  const accentBorder = isAdded
    ? "border-l-emerald-500"
    : isRemoved
    ? "border-l-rose-500"
    : "border-l-amber-500";

  const outerBorder = isAdded
    ? darkMode ? "border-emerald-800/60" : "border-emerald-200"
    : isRemoved
    ? darkMode ? "border-rose-800/60" : "border-rose-200"
    : darkMode ? "border-amber-800/60" : "border-amber-200";

  const headerBg = isAdded
    ? darkMode ? "bg-emerald-950/50" : "bg-emerald-50"
    : isRemoved
    ? darkMode ? "bg-rose-950/50" : "bg-rose-50"
    : darkMode ? "bg-amber-950/50" : "bg-amber-50";

  const divider = darkMode ? "border-slate-700/60" : "border-slate-200/80";

  const Icon = isAdded ? Plus : isRemoved ? Minus : ArrowRightLeft;
  const iconColor = isAdded
    ? "text-emerald-500"
    : isRemoved
    ? "text-rose-500"
    : "text-amber-500";

  const label = isAdded ? "Added" : isRemoved ? "Removed" : "Changed";

  return (
    <div className={`overflow-hidden rounded-xl border border-l-4 ${outerBorder} ${accentBorder}`}>
      {/* Header */}
      <div className={`flex items-center gap-2 px-3 py-2 border-b ${headerBg} ${divider}`}>
        <Icon size={13} className={`shrink-0 ${iconColor}`} />
        <span className={`text-xs font-semibold ${iconColor}`}>{label}</span>
        <span className={`text-xs font-mono truncate ml-1 ${darkMode ? "text-slate-300" : "text-slate-600"}`}>
          {entry.path}
        </span>
      </div>

      {/* Body */}
      {isChanged ? (
        <div className="grid grid-cols-1 sm:grid-cols-2">
          <div className={`px-3 py-2.5 border-b sm:border-b-0 sm:border-r ${divider} ${darkMode ? "bg-rose-950/30" : "bg-rose-50/70"}`}>
            <div className={`text-xs font-semibold mb-1.5 ${darkMode ? "text-rose-400" : "text-rose-600"}`}>Before</div>
            <code className={`block text-xs break-all whitespace-pre-wrap ${darkMode ? "text-rose-200" : "text-rose-900"}`}>
              {JSON.stringify(entry.before, null, 2)}
            </code>
          </div>
          <div className={`px-3 py-2.5 ${darkMode ? "bg-emerald-950/30" : "bg-emerald-50/70"}`}>
            <div className={`text-xs font-semibold mb-1.5 ${darkMode ? "text-emerald-400" : "text-emerald-600"}`}>After</div>
            <code className={`block text-xs break-all whitespace-pre-wrap ${darkMode ? "text-emerald-200" : "text-emerald-900"}`}>
              {JSON.stringify(entry.after, null, 2)}
            </code>
          </div>
        </div>
      ) : (
        <div className={`px-3 py-2.5 ${isAdded ? (darkMode ? "bg-emerald-950/20" : "bg-emerald-50/50") : (darkMode ? "bg-rose-950/20" : "bg-rose-50/50")}`}>
          <code className={`block text-xs break-all whitespace-pre-wrap ${isAdded ? (darkMode ? "text-emerald-200" : "text-emerald-900") : (darkMode ? "text-rose-200" : "text-rose-900")}`}>
            {JSON.stringify(isAdded ? entry.after : entry.before, null, 2)}
          </code>
        </div>
      )}
    </div>
  );
}

// ─── Excel preview modal ──────────────────────────────────────────────────────

const PREVIEW_LIMIT = 200;

function ExcelPreviewModal({
  data,
  darkMode,
  onDownload,
  onClose,
}: {
  data: ExcelData;
  darkMode: boolean;
  onDownload: () => void;
  onClose: () => void;
}) {
  const [activeTab, setActiveTab] = useState(0);

  useEffect(() => {
    const handler = (e: KeyboardEvent) => { if (e.key === "Escape") onClose(); };
    document.addEventListener("keydown", handler);
    return () => document.removeEventListener("keydown", handler);
  }, [onClose]);

  const sheet = data.sheets[activeTab] ?? data.sheets[0];
  const visibleRows = sheet.rows.slice(0, PREVIEW_LIMIT);
  const hasMore = sheet.rows.length > PREVIEW_LIMIT;
  const multiSheet = data.sheets.length > 1;
  const totalRows = data.sheets.reduce((sum, s) => sum + s.rows.length, 0);

  const cellBorder = darkMode ? "border-slate-800" : "border-slate-100";
  const headerBg   = darkMode ? "bg-slate-800"    : "bg-slate-50";
  const headerText = darkMode ? "text-slate-300"  : "text-slate-700";
  const rowHover   = darkMode ? "hover:bg-slate-800/50" : "hover:bg-slate-50";
  const mutedText  = darkMode ? "text-slate-500"  : "text-slate-400";
  const cellText   = darkMode ? "text-slate-300"  : "text-slate-700";

  return (
    <div
      className="fixed inset-0 z-50 flex items-center justify-center p-4"
      style={{ backgroundColor: "rgba(0,0,0,0.65)" }}
      onClick={onClose}
    >
      <div
        className={`flex w-full max-w-5xl flex-col rounded-2xl border shadow-2xl max-h-[92dvh] xl:max-h-[88dvh] ${darkMode ? "bg-slate-900 border-slate-700" : "bg-white border-slate-200"}`}
        onClick={(e) => e.stopPropagation()}
      >
        {/* Header */}
        <div className={`flex shrink-0 items-center justify-between border-b px-5 py-4 ${darkMode ? "border-slate-700" : "border-slate-200"}`}>
          <div>
            <h2 className={`text-sm font-bold ${darkMode ? "text-white" : "text-slate-900"}`}>Excel Preview</h2>
            <p className={`mt-0.5 text-xs ${mutedText}`}>
              {totalRows} row{totalRows !== 1 ? "s" : ""} &nbsp;·&nbsp; {data.sheets.length} sheet{data.sheets.length !== 1 ? "s" : ""}
            </p>
          </div>
          <button type="button" onClick={onClose}
            className={`rounded-lg p-1.5 transition ${darkMode ? "text-slate-400 hover:bg-slate-800 hover:text-white" : "text-slate-500 hover:bg-slate-100 hover:text-slate-900"}`}>
            <X size={17} />
          </button>
        </div>

        {/* Sheet tabs */}
        {multiSheet && (
          <div className={`flex shrink-0 gap-0.5 border-b px-4 pt-2 ${darkMode ? "border-slate-700" : "border-slate-200"}`}>
            {data.sheets.map((s, i) => (
              <button key={i} type="button" onClick={() => setActiveTab(i)}
                className={`-mb-px rounded-t-lg border-b-2 px-3 py-1.5 text-xs font-medium transition ${
                  activeTab === i
                    ? darkMode ? "border-emerald-400 text-emerald-400" : "border-emerald-600 text-emerald-700"
                    : darkMode ? "border-transparent text-slate-400 hover:text-slate-200" : "border-transparent text-slate-500 hover:text-slate-800"
                }`}>
                {s.name}
                <span className={`ml-1.5 rounded-full px-1.5 py-0.5 text-[10px] font-semibold ${darkMode ? "bg-slate-700 text-slate-400" : "bg-slate-100 text-slate-500"}`}>
                  {s.rows.length}
                </span>
              </button>
            ))}
          </div>
        )}

        {/* Sheet info bar */}
        <div className={`shrink-0 border-b px-5 py-1.5 text-xs ${mutedText} ${darkMode ? "border-slate-800" : "border-slate-100"}`}>
          {sheet.rows.length} row{sheet.rows.length !== 1 ? "s" : ""} &nbsp;·&nbsp; {sheet.headers.length} column{sheet.headers.length !== 1 ? "s" : ""}
          {hasMore && ` — showing first ${PREVIEW_LIMIT}`}
        </div>

        {/* Table */}
        <div className="min-h-0 flex-1 overflow-auto">
          <table className="w-full border-collapse text-xs">
            <thead className={`sticky top-0 z-10 ${headerBg}`}>
              <tr>
                <th className={`w-10 border-r px-3 py-2 text-left font-semibold ${cellBorder} ${mutedText}`}>#</th>
                {sheet.headers.map((h) => (
                  <th key={h} className={`whitespace-nowrap border-r px-3 py-2 text-left font-semibold ${cellBorder} ${headerText}`}>{h}</th>
                ))}
              </tr>
            </thead>
            <tbody>
              {visibleRows.map((row, i) => (
                <tr key={i} className={`border-t transition ${darkMode ? "border-slate-800" : "border-slate-100"} ${rowHover}`}>
                  <td className={`border-r px-3 py-1.5 font-mono ${cellBorder} ${mutedText}`}>{i + 1}</td>
                  {sheet.headers.map((h) => (
                    <td key={h} className={`max-w-[200px] truncate border-r px-3 py-1.5 ${cellBorder} ${cellText}`}>
                      {row[h] == null ? <span className="opacity-25">—</span> : String(row[h])}
                    </td>
                  ))}
                </tr>
              ))}
            </tbody>
          </table>
        </div>

        {/* Footer */}
        <div className={`flex shrink-0 items-center justify-end gap-2 border-t px-5 py-3 ${darkMode ? "border-slate-700" : "border-slate-200"}`}>
          <button type="button" onClick={onClose}
            className={`rounded-lg px-4 py-2 text-xs font-medium transition ${darkMode ? "bg-slate-800 text-slate-300 hover:bg-slate-700" : "bg-slate-100 text-slate-700 hover:bg-slate-200"}`}>
            Cancel
          </button>
          <button type="button" onClick={onDownload}
            className="flex items-center gap-1.5 rounded-lg bg-emerald-600 px-4 py-2 text-xs font-medium text-white transition hover:bg-emerald-500">
            <Download size={13} />Download .xlsx
          </button>
        </div>
      </div>
    </div>
  );
}

// ─── Main App ─────────────────────────────────────────────────────────────────

function App() {
  const [input, setInput] = useState("");
  const [compareInput, setCompareInput] = useState("");
  const [output, setOutput] = useState("");
  const [diffResult, setDiffResult] = useState<DiffResult | null>(null);
  const [status, setStatus] = useState("");
  const [toolMode, setToolMode] = useState<"single" | "diff" | "path">("single");
  const [viewMode, setViewMode] = useState<"raw" | "tree">("raw");
  const [outputLang, setOutputLang] = useState<"json" | "xml" | "csv">("json");
  const [excelPreview, setExcelPreview] = useState<ExcelData | null>(null);

  const [theme, setTheme] = useState<Theme>(getInitialTheme);
  const darkMode = theme === "dark";
  const statusTimerRef = useRef<ReturnType<typeof setTimeout> | null>(null);
  const clearTimerRef       = useRef<ReturnType<typeof setTimeout> | null>(null);
  const editorRef           = useRef<any>(null);
  const pasteDisposableRef  = useRef<{ dispose: () => void } | null>(null);
  const pathTimerRef        = useRef<ReturnType<typeof setTimeout> | null>(null);

  const [pathQuery,   setPathQuery]   = useState("");
  const [pathResults, setPathResults] = useState<Array<{ path: string; value: unknown }> | null>(null);
  const [pathError,   setPathError]   = useState("");
  const [clearPending, setClearPending] = useState(false);

  const inputStats = useMemo(() => {
    if (!input) return null;
    const lines = input.split("\n").length;
    const bytes = new Blob([input]).size;
    const size  = bytes < 1024 ? `${bytes} B` : `${(bytes / 1024).toFixed(1)} KB`;
    return `${lines} line${lines !== 1 ? "s" : ""} · ${size}`;
  }, [input]);

  useEffect(() => {
    return () => { pasteDisposableRef.current?.dispose(); };
  }, []);

  useEffect(() => {
    if (toolMode !== "path") return;
    if (pathTimerRef.current) clearTimeout(pathTimerRef.current);
    if (!pathQuery.trim() || !input.trim()) {
      setPathResults(null);
      setPathError("");
      return;
    }
    pathTimerRef.current = setTimeout(() => {
      try {
        const json = JSON.parse(input);
        const raw: any[] = JSONPath({ path: pathQuery, json, resultType: "all" });
        setPathResults(raw.map((r) => ({ path: r.path as string, value: r.value })));
        setPathError("");
      } catch (e) {
        setPathResults(null);
        setPathError((e as Error).message);
      }
    }, 300);
  }, [pathQuery, input, toolMode]);

  useEffect(() => {
    document.documentElement.dataset.theme = theme;
    document.documentElement.style.colorScheme = theme;
    localStorage.setItem("theme", theme);
  }, [theme]);

  const showStatus = (msg: string) => {
    setStatus(msg);
    if (statusTimerRef.current) clearTimeout(statusTimerRef.current);
    statusTimerRef.current = setTimeout(() => setStatus(""), 4000);
  };

  const handleFormat = () => {
    try {
      setOutput(formatJson(input));
      setOutputLang("json");
      showStatus("✅ Formatted");
    } catch (error) {
      showStatus(`❌ ${(error as Error).message}`);
    }
  };

  const handleMinify = () => {
    try {
      setOutput(minifyJson(input));
      setOutputLang("json");
      showStatus("✅ Minified");
    } catch (error) {
      showStatus(`❌ ${(error as Error).message}`);
    }
  };

  const handleJsonToXml = () => {
    try {
      setOutput(jsonToXml(input));
      setOutputLang("xml");
      setViewMode("raw");
      showStatus("✅ Converted to XML");
    } catch (error) {
      showStatus(`❌ ${(error as Error).message}`);
    }
  };

  const handleJsonToExcel = () => {
    try {
      setExcelPreview(prepareExcelData(input));
    } catch (error) {
      showStatus(`❌ ${(error as Error).message}`);
    }
  };

  const handleRepair = () => {
    try {
      const repaired = jsonrepair(input);
      setOutput(repaired);
      setOutputLang("json");
      showStatus("✅ Repaired");
    } catch (error) {
      showStatus(`❌ ${(error as Error).message}`);
    }
  };

  const handleJsonToCsv = () => {
    try {
      setOutput(jsonToCsv(input));
      setOutputLang("csv");
      setViewMode("raw");
      showStatus("✅ Converted to CSV");
    } catch (error) {
      showStatus(`❌ ${(error as Error).message}`);
    }
  };

  const handleJsonToSchema = () => {
    try {
      setOutput(generateJsonSchema(input));
      setOutputLang("json");
      showStatus("✅ Schema generated");
    } catch (error) {
      showStatus(`❌ ${(error as Error).message}`);
    }
  };

  const handleDownloadExcel = () => {
    if (!excelPreview) return;
    downloadExcel(excelPreview);
    setExcelPreview(null);
    showStatus("⬇️ Excel file downloaded");
  };

  const handleValidate = () => {
    const result = validateJson(input);
    showStatus(result.valid ? "✅ Valid JSON" : `❌ ${result.error}`);
  };

  const handleDiff = () => {
    try {
      setToolMode("diff");
      setViewMode("raw");
      const result = diffJson(input, compareInput);
      setDiffResult(result);
      setOutput(formatDiffAsText(result));
      showStatus(
        result.entries.length === 0
          ? "✅ No differences"
          : `✅ Found ${result.entries.length} difference${result.entries.length !== 1 ? "s" : ""}`
      );
    } catch (error) {
      setToolMode("diff");
      setDiffResult(null);
      showStatus(`❌ ${(error as Error).message}`);
    }
  };

  const handleCopy = async () => {
    if (!output) return;
    await navigator.clipboard.writeText(output);
    showStatus("📋 Copied");
  };

  const handleCopyPathResults = async () => {
    if (!pathResults?.length) return;
    await navigator.clipboard.writeText(JSON.stringify(pathResults.map((r) => r.value), null, 2));
    showStatus("📋 Copied results");
  };

  const handleDownload = () => {
    if (!output) return;
    const fileName = toolMode === "diff" ? "json-diff.txt" : outputLang === "xml" ? "output.xml" : outputLang === "csv" ? "output.csv" : "formatted.json";
    const mimeType = toolMode === "diff" ? "text/plain" : outputLang === "xml" ? "application/xml" : outputLang === "csv" ? "text/csv" : "application/json";
    downloadJson(output, fileName, mimeType);
    showStatus("⬇️ Download started");
  };

  const handleClear = () => {
    if (!input && !compareInput && !output && !pathQuery) return;
    if (!clearPending) {
      setClearPending(true);
      if (clearTimerRef.current) clearTimeout(clearTimerRef.current);
      clearTimerRef.current = setTimeout(() => setClearPending(false), 2500);
      return;
    }
    if (clearTimerRef.current) clearTimeout(clearTimerRef.current);
    setClearPending(false);
    setInput("");
    setCompareInput("");
    setOutput("");
    setDiffResult(null);
    setOutputLang("json");
    setExcelPreview(null);
    setStatus("");
    setPathQuery("");
    setPathResults(null);
    setPathError("");
  };

  const sep = (
    <div className={`hidden h-7 w-px shrink-0 sm:block ${darkMode ? "bg-slate-700" : "bg-slate-200"}`} />
  );

  const btnBase = "flex items-center gap-1.5 rounded-lg px-3 py-1.5 text-xs font-medium text-white transition";
  const btnDisabled = "disabled:cursor-not-allowed disabled:bg-slate-500 disabled:opacity-40";

  return (
    <div
      data-theme={theme}
      className="app-shell min-h-screen w-full transition-colors duration-300 xl:h-dvh xl:overflow-hidden"
    >
      <div className="w-full px-4 py-3 lg:px-6 xl:flex xl:h-full xl:min-h-0 xl:flex-col">

        {/* Header */}
        <div className="mb-3 flex shrink-0 items-center justify-between gap-4">
          <div className="min-w-0">
            <h1 className="text-2xl font-extrabold tracking-tight bg-gradient-to-r from-sky-400 via-cyan-300 to-blue-500 bg-clip-text text-transparent drop-shadow-[0_0_20px_rgba(56,189,248,0.4)] md:text-3xl">
              JSON Toolkit
            </h1>
            <p className={`mt-0.5 text-xs ${darkMode ? "text-slate-400" : "text-slate-500"}`}>
              Format, validate, diff, inspect and transform JSON instantly.
            </p>
          </div>

          <div className="flex shrink-0 items-center gap-2">
          <div
            role="group"
            aria-label="Theme mode"
            className={`flex shrink-0 overflow-hidden rounded-xl border p-1 transition ${darkMode ? "border-slate-700 bg-slate-900" : "border-slate-300 bg-white"}`}
          >
            <button type="button" onClick={() => setTheme("light")} aria-pressed={!darkMode} title="Light mode"
              className={`flex items-center gap-1.5 rounded-lg px-3 py-1.5 text-xs font-medium transition ${!darkMode ? "bg-slate-900 text-white" : "text-slate-400 hover:bg-slate-800 hover:text-slate-100"}`}>
              <Sun size={13} />Light
            </button>
            <button type="button" onClick={() => setTheme("dark")} aria-pressed={darkMode} title="Dark mode"
              className={`flex items-center gap-1.5 rounded-lg px-3 py-1.5 text-xs font-medium transition ${darkMode ? "bg-white text-slate-950" : "text-slate-600 hover:bg-slate-100 hover:text-slate-950"}`}>
              <Moon size={13} />Dark
            </button>
          </div>
          </div>
        </div>

        {/* Toolbar */}
        <div className={`z-10 mb-3 flex shrink-0 flex-wrap items-center gap-2 rounded-xl border p-2 shadow-sm backdrop-blur ${darkMode ? "border-slate-800 bg-slate-950/70" : "border-slate-200 bg-white/90"}`}>
          <div role="group" aria-label="Tool mode"
            className={`flex overflow-hidden rounded-lg border p-0.5 ${darkMode ? "border-slate-700 bg-slate-800" : "border-slate-300 bg-slate-100"}`}>
            <button type="button" onClick={() => {
                setToolMode("single");
                setOutput(""); setDiffResult(null); setOutputLang("json");
                setClearPending(false);
                if (clearTimerRef.current) clearTimeout(clearTimerRef.current);
              }} aria-pressed={toolMode === "single"}
              className={`rounded-md px-3 py-1.5 text-xs font-medium transition ${toolMode === "single" ? (darkMode ? "bg-white text-slate-950" : "bg-slate-900 text-white") : (darkMode ? "text-slate-400 hover:text-slate-100" : "text-slate-600 hover:text-slate-950")}`}>
              Tools
            </button>
            <button type="button" onClick={() => {
                setToolMode("diff");
                setViewMode("raw"); setOutput(""); setDiffResult(null);
                setClearPending(false);
                if (clearTimerRef.current) clearTimeout(clearTimerRef.current);
              }} aria-pressed={toolMode === "diff"}
              className={`rounded-md px-3 py-1.5 text-xs font-medium transition ${toolMode === "diff" ? (darkMode ? "bg-white text-slate-950" : "bg-slate-900 text-white") : (darkMode ? "text-slate-400 hover:text-slate-100" : "text-slate-600 hover:text-slate-950")}`}>
              Diff
            </button>
            <button type="button" onClick={() => {
                setToolMode("path");
                setOutput(""); setDiffResult(null);
                setClearPending(false);
                if (clearTimerRef.current) clearTimeout(clearTimerRef.current);
              }} aria-pressed={toolMode === "path"}
              className={`rounded-md px-3 py-1.5 text-xs font-medium transition ${toolMode === "path" ? (darkMode ? "bg-white text-slate-950" : "bg-slate-900 text-white") : (darkMode ? "text-slate-400 hover:text-slate-100" : "text-slate-600 hover:text-slate-950")}`}>
              Path
            </button>
          </div>

          {sep}

          <div className="flex flex-wrap gap-1.5">
            <button type="button" onClick={handleFormat} disabled={toolMode !== "single"}
              className={`${btnBase} ${btnDisabled} bg-blue-600 hover:bg-blue-500`}>
              <Wand2 size={14} />Format
            </button>
            <button type="button" onClick={handleMinify} disabled={toolMode !== "single"}
              className={`${btnBase} ${btnDisabled} bg-green-600 hover:bg-green-500`}>
              <Minimize2 size={14} />Minify
            </button>
            <button type="button" onClick={handleValidate} disabled={toolMode !== "single"}
              className={`${btnBase} ${btnDisabled} bg-purple-600 hover:bg-purple-500`}>
              <BadgeCheck size={14} />Validate
            </button>
            <button type="button" onClick={handleJsonToXml} disabled={toolMode !== "single"}
              className={`${btnBase} ${btnDisabled} bg-teal-600 hover:bg-teal-500`}>
              <Braces size={14} />To XML
            </button>
            <button type="button" onClick={handleJsonToExcel} disabled={toolMode !== "single"}
              className={`${btnBase} ${btnDisabled} bg-emerald-700 hover:bg-emerald-600`}>
              <FileSpreadsheet size={14} />To Excel
            </button>
            <button type="button" onClick={handleJsonToCsv} disabled={toolMode !== "single"}
              className={`${btnBase} ${btnDisabled} bg-orange-700 hover:bg-orange-600`}>
              <Table size={14} />To CSV
            </button>
            <button type="button" onClick={handleRepair} disabled={toolMode !== "single"}
              className={`${btnBase} ${btnDisabled} bg-rose-700 hover:bg-rose-600`}>
              <Wrench size={14} />Repair
            </button>
            <button type="button" onClick={handleJsonToSchema} disabled={toolMode !== "single"}
              className={`${btnBase} ${btnDisabled} bg-violet-600 hover:bg-violet-500`}>
              <Code2 size={14} />Schema
            </button>
          </div>

          {sep}

          <div className="flex gap-1.5">
            <button type="button" onClick={handleDiff}
              className={`${btnBase} ${toolMode === "diff" ? "bg-indigo-600 hover:bg-indigo-500" : "bg-slate-700 hover:bg-slate-600"}`}>
              <GitCompareArrows size={14} />Compare
            </button>
            <button type="button" onClick={handleClear}
              className={`${btnBase} transition-all ${clearPending ? "bg-orange-500 hover:bg-orange-400 ring-2 ring-orange-400/40" : "bg-red-600 hover:bg-red-500"}`}>
              <Trash2 size={14} />{clearPending ? "Confirm?" : "Clear"}
            </button>
          </div>

        </div>

        {/* Editor grid */}
        <div className={`grid w-full grid-cols-1 gap-4 xl:min-h-0 xl:flex-1 xl:overflow-hidden ${toolMode === "diff" ? "lg:grid-cols-3" : "lg:grid-cols-2"}`}>

          {/* Input */}
          <div className="flex min-w-0 flex-col xl:min-h-0">
            <div className="mb-2 flex shrink-0 items-center gap-2">
              <h2 className={`text-xs font-semibold uppercase tracking-wider ${darkMode ? "text-slate-500" : "text-slate-400"}`}>
                {toolMode === "diff" ? "Original JSON" : "Input JSON"}
              </h2>
              {inputStats && (
                <span className={`ml-auto rounded-md border px-2 py-0.5 font-mono text-xs tabular-nums ${darkMode ? "border-slate-700 bg-slate-800 text-slate-300" : "border-slate-200 bg-slate-100 text-slate-600"}`}>
                  {inputStats}
                </span>
              )}
            </div>
            <div className={`relative h-[48dvh] min-h-[320px] overflow-hidden rounded-xl border xl:h-auto xl:min-h-0 xl:flex-1 ${darkMode ? "border-slate-700" : "border-slate-300"}`}>
              {!input && (
                <div className="pointer-events-none absolute inset-0 z-10 flex flex-col items-center justify-center gap-2">
                  <Braces size={28} className={darkMode ? "text-slate-700" : "text-slate-300"} />
                  <p className={`text-sm ${darkMode ? "text-slate-600" : "text-slate-400"}`}>
                    Paste or type JSON here…
                  </p>
                </div>
              )}
              <Editor
                height="100%"
                defaultLanguage="json"
                theme={darkMode ? "vs-dark" : "light"}
                value={input}
                onChange={(v) => setInput(v || "")}
                onMount={(editor) => {
                  editorRef.current = editor;
                  pasteDisposableRef.current = editor.onDidPaste(() => {
                    setTimeout(() => {
                      const value = editor.getValue();
                      try {
                        setOutput(formatJson(value));
                        setOutputLang("json");
                        showStatus("✅ Auto-formatted");
                      } catch {
                        // not valid JSON yet — ignore
                      }
                    }, 150);
                  });
                }}
                options={{ minimap: { enabled: false }, fontSize: 14, automaticLayout: true, formatOnPaste: true, formatOnType: true, scrollBeyondLastLine: false }}
              />
            </div>
          </div>

          {/* Compare input (diff mode) */}
          {toolMode === "diff" && (
            <div className="flex min-w-0 flex-col xl:min-h-0">
              <h2 className={`mb-2 shrink-0 text-xs font-semibold uppercase tracking-wider ${darkMode ? "text-slate-500" : "text-slate-400"}`}>
                Compare JSON
              </h2>
              <div className={`h-[48dvh] min-h-[320px] overflow-hidden rounded-xl border xl:h-auto xl:min-h-0 xl:flex-1 ${darkMode ? "border-slate-700" : "border-slate-300"}`}>
                <Editor height="100%" defaultLanguage="json" theme={darkMode ? "vs-dark" : "light"} value={compareInput}
                  onChange={(v) => setCompareInput(v || "")}
                  options={{ minimap: { enabled: false }, fontSize: 14, automaticLayout: true, formatOnPaste: true, formatOnType: true, scrollBeyondLastLine: false }} />
              </div>
            </div>
          )}

          {/* Output */}
          <div className="flex min-w-0 flex-col xl:min-h-0">
            {/* Output header */}
            <div className="mb-2 flex shrink-0 items-center gap-2">
              <h2 className={`text-xs font-semibold uppercase tracking-wider ${darkMode ? "text-slate-500" : "text-slate-400"}`}>
                {toolMode === "diff" ? "Diff Result" : toolMode === "path" ? "Path Results" : "Output"}
              </h2>
              <div className="ml-auto flex items-center gap-1.5">
                {toolMode === "path" ? (
                  pathResults && pathResults.length > 0 && (
                    <>
                      <span className={`rounded-full px-2 py-0.5 text-xs font-semibold ${darkMode ? "bg-sky-900 text-sky-300" : "bg-sky-100 text-sky-700"}`}>
                        {pathResults.length} match{pathResults.length !== 1 ? "es" : ""}
                      </span>
                      <button type="button" onClick={handleCopyPathResults}
                        className={`${btnBase} bg-orange-600 hover:bg-orange-500`}>
                        <CopyIcon size={13} />Copy results
                      </button>
                    </>
                  )
                ) : (
                  <>
                    {toolMode === "single" && (
                      <div role="group" aria-label="View mode"
                        className={`flex overflow-hidden rounded-lg border p-0.5 ${darkMode ? "border-slate-700 bg-slate-800" : "border-slate-300 bg-slate-100"}`}>
                        <button type="button" onClick={() => setViewMode("raw")} aria-pressed={viewMode === "raw"}
                          className={`rounded-md px-2.5 py-1 text-xs font-medium transition ${viewMode === "raw" ? (darkMode ? "bg-white text-slate-950" : "bg-slate-900 text-white") : (darkMode ? "text-slate-400 hover:text-slate-100" : "text-slate-600 hover:text-slate-950")}`}>
                          Raw
                        </button>
                        <button type="button" onClick={() => setViewMode("tree")} aria-pressed={viewMode === "tree"}
                          disabled={outputLang === "xml" || outputLang === "csv"}
                          title={outputLang === "xml" ? "Tree view is not available for XML output" : outputLang === "csv" ? "Tree view is not available for CSV output" : undefined}
                          className={`rounded-md px-2.5 py-1 text-xs font-medium transition disabled:cursor-not-allowed disabled:opacity-30 ${viewMode === "tree" ? (darkMode ? "bg-white text-slate-950" : "bg-slate-900 text-white") : (darkMode ? "text-slate-400 hover:text-slate-100" : "text-slate-600 hover:text-slate-950")}`}>
                          Tree
                        </button>
                      </div>
                    )}
                    <button type="button" onClick={handleCopy} disabled={!output} title="Copy output"
                      className={`${btnBase} ${btnDisabled} bg-orange-600 hover:bg-orange-500`}>
                      <CopyIcon size={13} />Copy
                    </button>
                    <button type="button" onClick={handleDownload} disabled={!output} title="Download output"
                      className={`${btnBase} ${btnDisabled} bg-cyan-600 hover:bg-cyan-500`}>
                      <Download size={13} />Save
                    </button>
                  </>
                )}
              </div>
            </div>

            {/* Output body */}
            {toolMode === "path" ? (
              <div className="flex h-[48dvh] min-h-[320px] flex-col gap-2 xl:h-auto xl:min-h-0 xl:flex-1">
                {/* Query input */}
                <div className={`flex shrink-0 items-center gap-2 rounded-xl border px-3 py-2.5 ${darkMode ? "border-slate-700 bg-slate-900" : "border-slate-300 bg-white"}`}>
                  <Search size={14} className={`shrink-0 ${darkMode ? "text-slate-500" : "text-slate-400"}`} />
                  <input
                    type="text"
                    value={pathQuery}
                    onChange={(e) => setPathQuery(e.target.value)}
                    placeholder="$.store.book[*].title"
                    spellCheck={false}
                    className={`flex-1 bg-transparent font-mono text-sm outline-none ${darkMode ? "text-slate-100 placeholder-slate-600" : "text-slate-900 placeholder-slate-400"}`}
                  />
                  {pathQuery && (
                    <button type="button" onClick={() => setPathQuery("")}
                      className={`shrink-0 rounded p-0.5 transition ${darkMode ? "text-slate-500 hover:text-slate-300" : "text-slate-400 hover:text-slate-600"}`}>
                      <X size={13} />
                    </button>
                  )}
                </div>

                {/* Results area */}
                <div className={`flex-1 overflow-auto rounded-xl border p-3 ${darkMode ? "border-slate-700 bg-slate-950" : "border-slate-300 bg-slate-50"}`}>
                  {!pathQuery.trim() ? (
                    <div className="flex h-full flex-col items-center justify-center gap-2 py-8">
                      <Search size={32} className={darkMode ? "text-slate-700" : "text-slate-300"} />
                      <p className={`text-sm ${darkMode ? "text-slate-500" : "text-slate-400"}`}>Enter a JSONPath expression above.</p>
                      <p className={`font-mono text-xs ${darkMode ? "text-slate-600" : "text-slate-400"}`}>e.g. $.users[*].name &nbsp;·&nbsp; $..price &nbsp;·&nbsp; $.items[?(@.qty &gt; 0)]</p>
                    </div>
                  ) : pathError ? (
                    <div className={`rounded-xl border px-4 py-3 text-sm ${darkMode ? "border-rose-800 bg-rose-950/40 text-rose-300" : "border-rose-200 bg-rose-50 text-rose-700"}`}>
                      <span className="font-semibold">Error: </span>{pathError}
                    </div>
                  ) : !pathResults || pathResults.length === 0 ? (
                    <div className="flex h-full flex-col items-center justify-center gap-2 py-8">
                      <p className={`text-sm ${darkMode ? "text-slate-500" : "text-slate-400"}`}>No matches found.</p>
                    </div>
                  ) : (
                    <div className="flex flex-col gap-2">
                      {pathResults.map((r, i) => (
                        <div key={i} className={`overflow-hidden rounded-xl border ${darkMode ? "border-slate-700" : "border-slate-200"}`}>
                          <div className={`border-b px-3 py-1.5 font-mono text-xs ${darkMode ? "border-slate-700 bg-slate-900 text-sky-400" : "border-slate-200 bg-slate-50 text-sky-600"}`}>
                            {r.path}
                          </div>
                          <div className={`px-3 py-2 ${darkMode ? "bg-slate-950" : "bg-white"}`}>
                            <code className={`block break-all whitespace-pre-wrap text-xs ${darkMode ? "text-slate-300" : "text-slate-700"}`}>
                              {JSON.stringify(r.value, null, 2)}
                            </code>
                          </div>
                        </div>
                      ))}
                    </div>
                  )}
                </div>
              </div>
            ) : toolMode === "diff" ? (
              <div className={`h-[48dvh] min-h-[320px] overflow-auto rounded-xl border p-3 xl:h-auto xl:min-h-0 xl:flex-1 ${darkMode ? "border-slate-700 bg-slate-950" : "border-slate-300 bg-slate-50"}`}>
                {diffResult ? (
                  diffResult.entries.length === 0 ? (
                    /* Identical */
                    <div className="flex h-full flex-col items-center justify-center gap-2 py-8">
                      <CheckCircle2 size={36} className="text-emerald-500" />
                      <p className={`text-sm font-semibold ${darkMode ? "text-slate-200" : "text-slate-700"}`}>
                        No differences found
                      </p>
                      <p className="text-xs text-slate-500">Both JSON documents are identical.</p>
                    </div>
                  ) : (
                    <div className="flex flex-col gap-2">
                      {/* Summary pills */}
                      <div className={`mb-1 flex flex-wrap items-center gap-2 rounded-xl border px-3 py-2 ${darkMode ? "border-slate-800 bg-slate-900" : "border-slate-200 bg-white"}`}>
                        <span className={`text-xs font-semibold ${darkMode ? "text-slate-400" : "text-slate-500"}`}>
                          {diffResult.entries.length} difference{diffResult.entries.length !== 1 ? "s" : ""}
                        </span>
                        <div className="flex gap-2 ml-auto">
                          {diffResult.summary.added > 0 && (
                            <span className={`inline-flex items-center gap-1 rounded-full px-2.5 py-0.5 text-xs font-semibold ${darkMode ? "bg-emerald-950 text-emerald-300 border border-emerald-800" : "bg-emerald-100 text-emerald-700 border border-emerald-200"}`}>
                              <Plus size={10} />
                              {diffResult.summary.added} added
                            </span>
                          )}
                          {diffResult.summary.removed > 0 && (
                            <span className={`inline-flex items-center gap-1 rounded-full px-2.5 py-0.5 text-xs font-semibold ${darkMode ? "bg-rose-950 text-rose-300 border border-rose-800" : "bg-rose-100 text-rose-700 border border-rose-200"}`}>
                              <Minus size={10} />
                              {diffResult.summary.removed} removed
                            </span>
                          )}
                          {diffResult.summary.changed > 0 && (
                            <span className={`inline-flex items-center gap-1 rounded-full px-2.5 py-0.5 text-xs font-semibold ${darkMode ? "bg-amber-950 text-amber-300 border border-amber-800" : "bg-amber-100 text-amber-700 border border-amber-200"}`}>
                              <ArrowRightLeft size={10} />
                              {diffResult.summary.changed} changed
                            </span>
                          )}
                        </div>
                      </div>

                      {/* Entry cards */}
                      {diffResult.entries.map((entry, i) => (
                        <DiffCard key={i} entry={entry} darkMode={darkMode} />
                      ))}
                    </div>
                  )
                ) : (
                  /* Empty state */
                  <div className="flex h-full flex-col items-center justify-center gap-2 py-8">
                    <GitCompareArrows size={32} className={darkMode ? "text-slate-700" : "text-slate-300"} />
                    <p className="text-sm text-slate-500">Paste JSON into both panels and click Compare.</p>
                  </div>
                )}
              </div>
            ) : viewMode === "raw" ? (
              <div className={`relative h-[48dvh] min-h-[320px] overflow-hidden rounded-xl border xl:h-auto xl:min-h-0 xl:flex-1 ${darkMode ? "border-slate-700" : "border-slate-300"}`}>
                {!output && (
                  <div className="pointer-events-none absolute inset-0 z-10 flex flex-col items-center justify-center gap-2">
                    <Wand2 size={28} className={darkMode ? "text-slate-700" : "text-slate-300"} />
                    <p className={`text-sm ${darkMode ? "text-slate-600" : "text-slate-400"}`}>
                      Run Format, Minify or Validate to see output
                    </p>
                  </div>
                )}
                <Editor height="100%" language={outputLang === "csv" ? "plaintext" : outputLang} theme={darkMode ? "vs-dark" : "light"} value={output}
                  options={{ readOnly: true, minimap: { enabled: false }, fontSize: 14, automaticLayout: true, scrollBeyondLastLine: false }} />
              </div>
            ) : (
              <div className={`h-[48dvh] min-h-[320px] overflow-auto rounded-xl border p-4 xl:h-auto xl:min-h-0 xl:flex-1 ${darkMode ? "border-slate-700 bg-slate-900" : "border-slate-300 bg-white"}`}>
                {(() => {
                  if (!output) return <div className="text-sm text-slate-500">Format valid JSON to view tree structure.</div>;
                  try {
                    return <JsonView value={JSON.parse(output)} collapsed={1} displayDataTypes={false}
                      style={darkMode ? vscodeTheme : { backgroundColor: "#ffffff", color: "#111827" }} />;
                  } catch {
                    return <div className={`text-sm ${darkMode ? "text-rose-400" : "text-rose-600"}`}>Output is not valid JSON — switch to Raw view.</div>;
                  }
                })()}
              </div>
            )}
          </div>
        </div>

        {/* Footer */}
        <div className={`mt-3 shrink-0 border-t pt-2 text-center text-xs ${darkMode ? "border-slate-800 text-slate-600" : "border-slate-200 text-slate-400"}`}>
          JSON Toolkit • Build v1.0.0 &nbsp;|&nbsp; © 2026 All rights reserved
        </div>
      </div>

      {/* Excel preview modal */}
      {excelPreview && (
        <ExcelPreviewModal
          data={excelPreview}
          darkMode={darkMode}
          onDownload={handleDownloadExcel}
          onClose={() => setExcelPreview(null)}
        />
      )}

      {/* Corner toast */}
      <div
        className={`fixed bottom-5 right-5 z-50 max-w-xs rounded-xl border px-4 py-2.5 text-sm font-medium shadow-lg backdrop-blur-sm transition-all duration-300 ${
          status ? "translate-y-0 opacity-100" : "pointer-events-none translate-y-3 opacity-0"
        } ${
          status.startsWith("❌")
            ? darkMode ? "border-rose-700 bg-rose-950/90 text-rose-300" : "border-rose-200 bg-rose-50 text-rose-700"
            : status.startsWith("✅") || status.startsWith("📋") || status.startsWith("⬇️")
            ? darkMode ? "border-emerald-700 bg-emerald-950/90 text-emerald-300" : "border-emerald-200 bg-emerald-50 text-emerald-700"
            : darkMode ? "border-slate-700 bg-slate-900/90 text-slate-300" : "border-slate-200 bg-white text-slate-700"
        }`}
      >
        {status}
      </div>
    </div>
  );
}

export default App;
