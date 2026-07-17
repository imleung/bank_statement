"use client";

import { useEffect, useMemo, useState } from "react";
import type { ChangeEvent, DragEvent } from "react";
import * as XLSX from "xlsx";
import { languageOptions, messages } from "./i18n";
import type { CategoryKey, Language, MessageSet } from "./i18n";
import {
  buildMappingProfiles,
  createStatement,
  currencyOptions,
  defaultImportConfig,
  fieldOptions,
  inferField,
  normalizeStatements,
  normalizeText,
  parseMappingProfiles,
  parseStatementFile,
  transactionValue,
  validateTransactions,
} from "./statement";
import type { ImportConfig, IssueCode, MappingProfiles, MappingValue, StatementFile, Transaction } from "./statement";

const locales: Record<Language, string> = {
  vi: "vi-VN",
  en: "en-US",
  ja: "ja-JP",
  zh: "zh-CN",
  ko: "ko-KR",
};

const issueOrder: IssueCode[] = ["missingDate", "invalidDate", "noAmount", "bothSides", "negativeAmount", "balanceMismatch"];
const mappingStorageKey = "mach-tien-mapping-profiles";
const importConfigStorageKey = "mach-tien-import-config";

const demoFiles: StatementFile[] = [
  createStatement("SaoKe_VCB_Q2-2026.xlsx", "Vietcombank", [
    { "Ngày GD": "03/04/2026", "Số CT": "VCB-201", "Diễn giải": "Thu tiền bán hàng Công ty An Phú", "Ghi nợ": "", "Ghi có": 48_500_000, "Số dư": 128_500_000 },
    { "Ngày GD": "08/04/2026", "Số CT": "VCB-202", "Diễn giải": "Thanh toán tiền thuê văn phòng tháng 04", "Ghi nợ": 18_000_000, "Ghi có": "", "Số dư": 110_500_000 },
    { "Ngày GD": "22/04/2026", "Số CT": "VCB-203", "Diễn giải": "Nộp thuế GTGT quý I", "Ghi nợ": 12_600_000, "Ghi có": "", "Số dư": 97_900_000 },
    { "Ngày GD": "06/05/2026", "Số CT": "VCB-204", "Diễn giải": "Thu công nợ khách hàng Hải Đăng", "Ghi nợ": "", "Ghi có": 62_000_000, "Số dư": 159_900_000 },
    { "Ngày GD": "28/05/2026", "Số CT": "VCB-205", "Diễn giải": "Chi lương nhân viên tháng 05", "Ghi nợ": 46_000_000, "Ghi có": "", "Số dư": 113_900_000 },
  ]),
  createStatement("SaoKe_BIDV_Q2-2026.xlsx", "BIDV", [
    { Ngày: "11/04/2026", "Nội dung giao dịch": "Thu hợp đồng tư vấn Công ty Sao Việt", "Phát sinh Nợ": "", "Phát sinh Có": 35_000_000, "Số dư cuối": 65_000_000, "Người nhận / chuyển": "Công ty Sao Việt" },
    { Ngày: "18/04/2026", "Nội dung giao dịch": "Thanh toán NCC thiết bị văn phòng", "Phát sinh Nợ": 14_800_000, "Phát sinh Có": "", "Số dư cuối": 50_200_000, "Người nhận / chuyển": "NCC Minh Long" },
    { Ngày: "02/05/2026", "Nội dung giao dịch": "Lãi tiền gửi không kỳ hạn", "Phát sinh Nợ": "", "Phát sinh Có": 220_000, "Số dư cuối": 50_420_000, "Người nhận / chuyển": "BIDV" },
    { Ngày: "17/05/2026", "Nội dung giao dịch": "Phí quản lý tài khoản", "Phát sinh Nợ": 120_000, "Phát sinh Có": "", "Số dư cuối": 50_300_000, "Người nhận / chuyển": "BIDV" },
  ]),
];

interface FileError {
  name: string;
  message: string;
}

interface Filters {
  query: string;
  source: string;
  kind: "all" | Transaction["kind"];
  category: "all" | CategoryKey;
  issuesOnly: boolean;
}

type UndoItem =
  | { kind: "transaction"; transaction: Transaction; index: number }
  | { kind: "statement"; statement: StatementFile; index: number; transactions: Transaction[] }
  | { kind: "all"; statements: StatementFile[]; transactions: Transaction[] };

const initialFilters: Filters = {
  query: "",
  source: "all",
  kind: "all",
  category: "all",
  issuesOnly: false,
};

function formatMoney(value: number, language: Language) {
  return new Intl.NumberFormat(locales[language], { maximumFractionDigits: 2 }).format(value);
}

function formatDisplayDate(value: string, language: Language) {
  const match = value.match(/^(\d{4})-(\d{2})-(\d{2})$/);
  if (!match) return value || "—";
  const date = new Date(Date.UTC(Number(match[1]), Number(match[2]) - 1, Number(match[3])));
  return new Intl.DateTimeFormat(locales[language], { dateStyle: "medium", timeZone: "UTC" }).format(date);
}

function formatMonth(value: string, language: Language) {
  const match = value.match(/^(\d{4})-(\d{2})$/);
  if (!match) return value;
  const date = new Date(Date.UTC(Number(match[1]), Number(match[2]) - 1, 1));
  return new Intl.DateTimeFormat(locales[language], { month: "short", year: "2-digit", timeZone: "UTC" }).format(date);
}

function exportRows(transactions: Transaction[], copy: MessageSet) {
  const columns = copy.exportColumns;
  return transactions.map((transaction) => ({
    [columns.source]: transaction.source,
    [columns.date]: transaction.date,
    [columns.reference]: transaction.reference,
    [columns.description]: transaction.description,
    [columns.counterparty]: transaction.counterparty,
    [columns.debit]: transaction.debit,
    [columns.credit]: transaction.credit,
    [columns.balance]: transaction.balance,
    [columns.currency]: transaction.currency,
    [columns.kind]: transaction.kind === "Thu" ? copy.incomeKind : copy.expenseKind,
    [columns.category]: copy.categories[transaction.category],
  }));
}

function insertAt<T>(items: T[], item: T, index: number) {
  const next = [...items];
  next.splice(index, 0, item);
  return next;
}

function IssueBadges({ issues, copy }: { issues: IssueCode[]; copy: MessageSet }) {
  if (issues.length === 0) return null;
  return (
    <div className="issue-badges">
      {issues.map((issue) => <span key={issue}>{copy.review.issues[issue]}</span>)}
    </div>
  );
}

export default function Home() {
  const [language, setLanguage] = useState<Language>("vi");
  const [statements, setStatements] = useState<StatementFile[]>([]);
  const [transactions, setTransactions] = useState<Transaction[]>([]);
  const [templateFile, setTemplateFile] = useState<File | null>(null);
  const [importConfig, setImportConfig] = useState<ImportConfig>({ ...defaultImportConfig });
  const [mappingProfiles, setMappingProfiles] = useState<MappingProfiles>({});
  const [filters, setFilters] = useState<Filters>(initialFilters);
  const [duplicateCount, setDuplicateCount] = useState(0);
  const [fileErrors, setFileErrors] = useState<FileError[]>([]);
  const [undoItem, setUndoItem] = useState<UndoItem | null>(null);
  const [isDragging, setIsDragging] = useState(false);
  const [isReading, setIsReading] = useState(false);
  const [error, setError] = useState("");
  const copy = messages[language];

  useEffect(() => {
    const stored = window.localStorage.getItem("mach-tien-language");
    if (stored === "vi" || stored === "en" || stored === "ja" || stored === "zh" || stored === "ko") {
      setLanguage(stored);
    } else {
      const browserLanguage = window.navigator.language.toLowerCase();
      if (browserLanguage.startsWith("zh")) setLanguage("zh");
      else if (browserLanguage.startsWith("ko")) setLanguage("ko");
      else if (browserLanguage.startsWith("ja")) setLanguage("ja");
      else if (browserLanguage.startsWith("en")) setLanguage("en");
    }

    const storedProfiles = window.localStorage.getItem(mappingStorageKey);
    if (storedProfiles) {
      try {
        setMappingProfiles(parseMappingProfiles(storedProfiles));
      } catch {
        window.localStorage.removeItem(mappingStorageKey);
      }
    }

    const storedConfig = window.localStorage.getItem(importConfigStorageKey);
    if (storedConfig) {
      try {
        const parsed: unknown = JSON.parse(storedConfig);
        if (parsed && typeof parsed === "object" && !Array.isArray(parsed)) {
          const candidate = parsed as Partial<ImportConfig>;
          const dateFormat = candidate.dateFormat;
          const decimalSeparator = candidate.decimalSeparator;
          const currency = candidate.currency;
          if (
            (dateFormat === "auto" || dateFormat === "dmy" || dateFormat === "mdy" || dateFormat === "ymd")
            && (decimalSeparator === "auto" || decimalSeparator === "comma" || decimalSeparator === "dot")
            && currencyOptions.includes(currency as ImportConfig["currency"])
          ) {
            setImportConfig({ dateFormat, decimalSeparator, currency: currency as ImportConfig["currency"] });
          }
        }
      } catch {
        window.localStorage.removeItem(importConfigStorageKey);
      }
    }
  }, []);

  useEffect(() => {
    document.documentElement.lang = language;
    document.title = copy.metaTitle;
  }, [copy.metaTitle, language]);

  const validation = useMemo(() => validateTransactions(transactions), [transactions]);

  const analytics = useMemo(() => {
    const totalCredit = transactions.reduce((sum, item) => sum + item.credit, 0);
    const totalDebit = transactions.reduce((sum, item) => sum + item.debit, 0);
    const categories = new Map<CategoryKey, number>();
    const months = new Map<string, { credit: number; debit: number }>();
    const sources = new Map<string, { credit: number; debit: number }>();

    transactions.forEach((item) => {
      const month = /^\d{4}-\d{2}/.test(item.date) ? item.date.slice(0, 7) : "—";
      const monthly = months.get(month) ?? { credit: 0, debit: 0 };
      monthly.credit += item.credit;
      monthly.debit += item.debit;
      months.set(month, monthly);

      const source = sources.get(item.source) ?? { credit: 0, debit: 0 };
      source.credit += item.credit;
      source.debit += item.debit;
      sources.set(item.source, source);

      if (item.debit > 0) categories.set(item.category, (categories.get(item.category) ?? 0) + item.debit);
    });

    return {
      totalCredit,
      totalDebit,
      net: totalCredit - totalDebit,
      categories: [...categories.entries()].sort((a, b) => b[1] - a[1]),
      months: [...months.entries()].sort(([a], [b]) => a.localeCompare(b)),
      sources: [...sources.entries()],
    };
  }, [transactions]);

  const sourceOptions = useMemo(() => [...new Set(transactions.map((item) => item.source))].sort(), [transactions]);

  const filteredTransactions = useMemo(() => {
    const query = normalizeText(filters.query);
    return transactions.filter((transaction) => {
      if (filters.source !== "all" && transaction.source !== filters.source) return false;
      if (filters.kind !== "all" && transaction.kind !== filters.kind) return false;
      if (filters.category !== "all" && transaction.category !== filters.category) return false;
      if (filters.issuesOnly && !validation.issuesById.has(transaction.id)) return false;
      if (!query) return true;
      const haystack = normalizeText([transaction.source, transaction.date, transaction.reference, transaction.description, transaction.counterparty].join(" "));
      return haystack.includes(query);
    });
  }, [filters, transactions, validation.issuesById]);

  async function addFiles(files: File[]) {
    if (files.length === 0) return;
    setIsReading(true);
    setError("");

    const existing = new Set(statements.map((statement) => statement.fingerprint));
    const pending = new Set<string>();
    const accepted: File[] = [];
    const failures: FileError[] = [];

    files.forEach((file) => {
      const fingerprint = `${file.name}:${file.size}:${file.lastModified}`;
      if (existing.has(fingerprint) || pending.has(fingerprint)) {
        failures.push({ name: file.name, message: copy.errors.duplicateFile });
      } else {
        pending.add(fingerprint);
        accepted.push(file);
      }
    });

    const results = await Promise.allSettled(accepted.map((file) => parseStatementFile(file, copy, mappingProfiles)));
    const parsed: StatementFile[] = [];
    results.forEach((result, index) => {
      if (result.status === "fulfilled") parsed.push(result.value);
      else failures.push({ name: accepted[index].name, message: result.reason instanceof Error ? result.reason.message : copy.errors.readFile });
    });

    if (parsed.length > 0) {
      setStatements((current) => [...current, ...parsed]);
      setTransactions([]);
      setDuplicateCount(0);
    }
    setFileErrors(failures);
    setIsReading(false);
  }

  function handleFileInput(event: ChangeEvent<HTMLInputElement>) {
    void addFiles(Array.from(event.target.files ?? []));
    event.target.value = "";
  }

  function handleDrop(event: DragEvent<HTMLLabelElement>) {
    event.preventDefault();
    setIsDragging(false);
    void addFiles(Array.from(event.dataTransfer.files));
  }

  function updateImportConfig<Key extends keyof ImportConfig>(key: Key, value: ImportConfig[Key]) {
    const next = { ...importConfig, [key]: value };
    setImportConfig(next);
    window.localStorage.setItem(importConfigStorageKey, JSON.stringify(next));
    setTransactions([]);
  }

  function updateStatement(id: string, updater: (statement: StatementFile) => StatementFile) {
    setStatements((current) => current.map((statement) => (statement.id === id ? updater(statement) : statement)));
    setTransactions([]);
  }

  function removeStatement(id: string) {
    const index = statements.findIndex((statement) => statement.id === id);
    if (index < 0) return;
    setUndoItem({ kind: "statement", statement: statements[index], index, transactions });
    setStatements((current) => current.filter((statement) => statement.id !== id));
    setTransactions([]);
  }

  function clearAll() {
    if (statements.length === 0 && transactions.length === 0) return;
    setUndoItem({ kind: "all", statements, transactions });
    setStatements([]);
    setTransactions([]);
    setDuplicateCount(0);
    setFileErrors([]);
  }

  function restoreUndo() {
    if (!undoItem) return;
    if (undoItem.kind === "transaction") {
      setTransactions((current) => insertAt(current, undoItem.transaction, undoItem.index));
    } else if (undoItem.kind === "statement") {
      setStatements((current) => insertAt(current, undoItem.statement, undoItem.index));
      setTransactions(undoItem.transactions);
    } else {
      setStatements(undoItem.statements);
      setTransactions(undoItem.transactions);
    }
    setUndoItem(null);
  }

  function buildMasterTable() {
    const result = normalizeStatements(statements, importConfig);
    setTransactions(result.transactions);
    setDuplicateCount(result.duplicateCount);
    setFilters(initialFilters);
    setError(result.transactions.length === 0 ? copy.errors.noTransactions : "");

    const profiles = buildMappingProfiles(statements, mappingProfiles);
    setMappingProfiles(profiles);
    window.localStorage.setItem(mappingStorageKey, JSON.stringify(profiles));

    if (result.transactions.length > 0) {
      requestAnimationFrame(() => document.getElementById("insights")?.scrollIntoView({ behavior: "smooth", block: "start" }));
    }
  }

  function updateTransaction(id: string, patch: Partial<Transaction>) {
    setTransactions((current) => current.map((transaction) => {
      if (transaction.id !== id) return transaction;
      const next = { ...transaction, ...patch };
      if ("debit" in patch || "credit" in patch) {
        const previousKind = next.kind;
        next.kind = next.credit >= next.debit ? "Thu" : "Chi";
        if (next.kind !== previousKind && (next.category === "customerIncome" || next.category === "otherExpense")) {
          next.category = next.kind === "Thu" ? "customerIncome" : "otherExpense";
        }
      }
      if ("balance" in patch) next.balanceProvided = true;
      return next;
    }));
  }

  function removeTransaction(id: string) {
    const index = transactions.findIndex((transaction) => transaction.id === id);
    if (index < 0) return;
    setUndoItem({ kind: "transaction", transaction: transactions[index], index });
    setTransactions((current) => current.filter((transaction) => transaction.id !== id));
  }

  function exportWorkbook(compact: boolean) {
    const columns = copy.exportColumns;
    const rows = compact
      ? transactions.map((item) => ({
          [columns.compactDate]: item.date,
          [columns.description]: item.description,
          [columns.compactIncome]: item.credit,
          [columns.compactExpense]: item.debit,
          [columns.balance]: item.balance,
          [columns.currency]: item.currency,
          [columns.category]: copy.categories[item.category],
        }))
      : exportRows(transactions, copy);
    const workbook = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(workbook, XLSX.utils.json_to_sheet(rows), compact ? copy.export.sheetCompact : copy.export.sheetFull);
    XLSX.writeFile(workbook, compact ? copy.export.fileCompact : copy.export.fileFull);
  }

  async function exportToTemplate() {
    if (!templateFile) return;
    setError("");

    try {
      const workbook = XLSX.read(await templateFile.arrayBuffer(), { cellDates: true });
      const sheetName = workbook.SheetNames[0];
      const sheet = workbook.Sheets[sheetName];
      const grid = XLSX.utils.sheet_to_json<unknown[]>(sheet, { header: 1, defval: "" });
      const headerIndex = grid.slice(0, 20).findIndex((row) => row.filter((cell) => inferField(String(cell)) !== "ignore").length >= 2);

      if (headerIndex < 0) {
        XLSX.utils.book_append_sheet(workbook, XLSX.utils.json_to_sheet(exportRows(transactions, copy)), copy.export.sheetFallback);
      } else {
        const headers = grid[headerIndex].map((cell) => inferField(String(cell)));
        const rows = transactions.map((transaction) => headers.map((field) => transactionValue(transaction, field)));
        XLSX.utils.sheet_add_aoa(sheet, rows, { origin: { r: grid.length, c: 0 } });
      }

      XLSX.writeFile(workbook, `${copy.export.templatePrefix}${templateFile.name}`);
    } catch {
      setError(copy.errors.template);
    }
  }

  function changeLanguage(nextLanguage: Language) {
    setLanguage(nextLanguage);
    window.localStorage.setItem("mach-tien-language", nextLanguage);
  }

  const maxMonthly = Math.max(1, ...analytics.months.flatMap(([, value]) => [value.credit, value.debit]));
  const categoryKeys = Object.keys(copy.categories) as CategoryKey[];

  return (
    <main>
      <header className="site-header">
        <a className="brand" href="#top" aria-label={copy.backToTop} translate="no">
          <span className="brand-mark">M</span>
          <span><strong>Mạch Tiền</strong><small>{copy.brandTagline}</small></span>
        </a>
        <div className="header-actions">
          <div className="language-switcher" role="group" aria-label={copy.languageLabel}>
            {languageOptions.map((option) => (
              <button className={language === option.code ? "active" : ""} type="button" aria-pressed={language === option.code} lang={option.code} key={option.code} onClick={() => changeLanguage(option.code)}>
                {option.label}
              </button>
            ))}
          </div>
          <div className="privacy-pill"><span aria-hidden="true" /> {copy.privacy}</div>
        </div>
      </header>

      <section className="hero" id="top">
        <div className="hero-copy">
          <div className="eyebrow">{copy.hero.eyebrow}</div>
          <h1>{copy.hero.title} <em>{copy.hero.accent}</em></h1>
          <p>{copy.hero.body}</p>
          <div className="hero-badges">{copy.hero.badges.map((badge) => <span key={badge}>{badge}</span>)}</div>
        </div>
        <div className="hero-visual" aria-label={copy.hero.netCashFlow}>
          <div className="floating-note note-one">VCB · +{formatMoney(48_500_000, language)}</div>
          <div className="floating-note note-two">BIDV · −{formatMoney(14_800_000, language)}</div>
          <div className="statement-card">
            <div className="statement-top"><span>{copy.hero.quarter}</span><span>{copy.hero.transactions}</span></div>
            <div className="pulse-line" aria-hidden="true"><i /><i /><i /><i /><i /><i /></div>
            <div className="statement-total"><span>{copy.hero.netCashFlow}</span><strong>+{formatMoney(54_200_000, language)}</strong></div>
            <div className="mini-bars" aria-hidden="true"><i /><i /><i /><i /><i /></div>
          </div>
        </div>
      </section>

      <nav className="stepper" aria-label={copy.workflowLabel}>
        {copy.steps.map(([label, detail], index) => {
          const number = String(index + 1).padStart(2, "0");
          return (
            <div className={`step ${statements.length > 0 && index < 2 ? "active" : ""} ${transactions.length > 0 && index < 4 ? "active" : ""}`} key={number}>
              <span>{number}</span><div><strong>{label}</strong><small>{detail}</small></div>
            </div>
          );
        })}
      </nav>

      <section className="workspace-section">
        <div className="section-heading">
          <div><span className="section-number">01</span><h2>{copy.upload.title}</h2></div>
          <p>{copy.upload.description}</p>
        </div>

        <label className={`drop-zone ${isDragging ? "dragging" : ""}`} onDragEnter={() => setIsDragging(true)} onDragLeave={() => setIsDragging(false)} onDragOver={(event) => event.preventDefault()} onDrop={handleDrop}>
          <input name="statements" type="file" accept=".xlsx,.xls,.csv" multiple onChange={handleFileInput} aria-describedby="statement-formats" />
          <span className="upload-icon" aria-hidden="true">↥</span>
          <strong aria-live="polite">{isReading ? copy.upload.reading : copy.upload.ready}</strong>
          <small id="statement-formats">{copy.upload.formats}</small>
        </label>

        <details className="import-settings">
          <summary>
            <span><strong>{copy.importSettings.title}</strong><small>{copy.importSettings.hint}</small></span>
            <i aria-hidden="true">+</i>
          </summary>
          <div className="import-settings-grid">
            <label>
              <span>{copy.importSettings.dateFormat}</span>
              <select value={importConfig.dateFormat} onChange={(event) => updateImportConfig("dateFormat", event.target.value as ImportConfig["dateFormat"])}>
                <option value="auto">{copy.importSettings.auto}</option>
                <option value="dmy">DD/MM/YYYY</option>
                <option value="mdy">MM/DD/YYYY</option>
                <option value="ymd">YYYY/MM/DD</option>
              </select>
            </label>
            <label>
              <span>{copy.importSettings.decimalSeparator}</span>
              <select value={importConfig.decimalSeparator} onChange={(event) => updateImportConfig("decimalSeparator", event.target.value as ImportConfig["decimalSeparator"])}>
                <option value="auto">{copy.importSettings.auto}</option>
                <option value="comma">{copy.importSettings.comma}</option>
                <option value="dot">{copy.importSettings.dot}</option>
              </select>
            </label>
            <label>
              <span>{copy.importSettings.currency}</span>
              <select value={importConfig.currency} onChange={(event) => updateImportConfig("currency", event.target.value as ImportConfig["currency"])}>
                {currencyOptions.map((currency) => <option value={currency} key={currency}>{currency}</option>)}
              </select>
            </label>
          </div>
          <p>{copy.importSettings.note}</p>
        </details>

        <div className="demo-row">
          <button className="text-button" type="button" onClick={() => {
            setStatements(demoFiles.map((file) => ({ ...file, id: `${file.id}-${Math.random()}`, fingerprint: `${file.fingerprint}-${Math.random()}` })));
            setTransactions([]);
            setFileErrors([]);
            setError("");
          }}>
            {copy.upload.demo} <span>→</span>
          </button>
          <p>{copy.upload.demoHint}</p>
        </div>

        {fileErrors.length > 0 && (
          <div className="file-errors" role="status" aria-live="polite">
            <strong>{copy.errors.fileFailures}</strong>
            <ul>{fileErrors.map((item, index) => <li key={`${item.name}-${index}`}><b>{item.name}</b><span>{item.message}</span></li>)}</ul>
          </div>
        )}
        {error && <div className="error-message" role="alert">{error}</div>}
      </section>

      {statements.length > 0 && (
        <section className="mapping-section">
          <div className="section-heading">
            <div><span className="section-number">02</span><h2>{copy.mapping.title}</h2></div>
            <p>{copy.mapping.description} {copy.mapping.remembered}</p>
          </div>

          <div className="file-stack">
            {statements.map((statement) => (
              <article className="file-card" key={statement.id}>
                <div className="file-card-header">
                  <div className="file-identity"><span className="file-icon" aria-hidden="true">XL</span><div><strong>{statement.fileName}</strong><small>{statement.rows.length} {copy.mapping.rows}</small></div></div>
                  <button className="remove-button" type="button" onClick={() => removeStatement(statement.id)}>{copy.mapping.remove}</button>
                </div>
                <div className="source-control">
                  <label htmlFor={`source-${statement.id}`}>{copy.mapping.source}</label>
                  <input id={`source-${statement.id}`} name={`source-${statement.id}`} autoComplete="off" value={statement.source} onChange={(event) => updateStatement(statement.id, (current) => ({ ...current, source: event.target.value }))} />
                  <span>{statement.headers.filter((header) => statement.mapping[header] !== "ignore").length}/{statement.headers.length} {copy.mapping.mapped}</span>
                </div>
                <div className="mapping-table-wrap">
                  <table className="mapping-table">
                    <thead><tr>{copy.mapping.headers.map((header) => <th key={header}>{header}</th>)}</tr></thead>
                    <tbody>
                      {statement.headers.map((header) => (
                        <tr key={header}>
                          <td>{header}</td><td>{String(statement.rows[0]?.[header] ?? "—") || "—"}</td>
                          <td>
                            <select aria-label={`${copy.mapping.fieldLabel} ${header}`} value={statement.mapping[header]} onChange={(event) => updateStatement(statement.id, (current) => ({ ...current, mapping: { ...current.mapping, [header]: event.target.value as MappingValue } }))}>
                              {fieldOptions.map((option) => <option value={option} key={option}>{copy.fields[option]}</option>)}
                            </select>
                          </td>
                        </tr>
                      ))}
                    </tbody>
                  </table>
                </div>
              </article>
            ))}
          </div>

          <div className="mapping-actions">
            <button className="secondary-button" type="button" onClick={clearAll}>{copy.mapping.reset}</button>
            <button className="primary-button" type="button" onClick={buildMasterTable}>{copy.mapping.build} <span>→</span></button>
          </div>
        </section>
      )}

      {transactions.length > 0 && (
        <section className="insights-section" id="insights">
          <div className="section-heading light-heading">
            <div><span className="section-number">03</span><h2>{copy.insights.title}</h2></div>
            <p>{formatDisplayDate(transactions[0]?.date, language)} — {formatDisplayDate(transactions.at(-1)?.date ?? "", language)}</p>
          </div>

          <article className={`review-panel ${validation.issuesById.size === 0 ? "review-success" : ""}`}>
            <div className="review-heading">
              <div><span>{copy.review.eyebrow}</span><h3>{copy.review.title}</h3><p>{copy.review.description}</p></div>
              <strong>{importConfig.currency}</strong>
            </div>
            <div className="review-grid">
              <div><span>{copy.review.valid}</span><strong>{validation.validCount}/{transactions.length}</strong></div>
              <div><span>{copy.review.needsReview}</span><strong>{validation.issuesById.size}</strong></div>
              <div><span>{copy.review.duplicatesRemoved}</span><strong>{duplicateCount}</strong></div>
              <div><span>{copy.review.reconciled}</span><strong>{validation.reconciliationChecked}</strong><small>{validation.reconciliationMismatches} {copy.review.mismatches}</small></div>
            </div>
            {validation.issuesById.size === 0 ? (
              <p className="review-message">✓ {copy.review.allGood}</p>
            ) : (
              <div className="review-issues">
                {issueOrder.map((issue) => {
                  const count = validation.issueCounts.get(issue) ?? 0;
                  return count > 0 ? <button type="button" key={issue} onClick={() => setFilters((current) => ({ ...current, issuesOnly: true }))}><b>{count}</b>{copy.review.issues[issue]}</button> : null;
                })}
                {validation.reconciliationDifference > 0 && <span>{copy.review.difference}: {formatMoney(validation.reconciliationDifference, language)} {importConfig.currency}</span>}
              </div>
            )}
          </article>

          <div className="stat-grid">
            <div className="stat-card"><span>{copy.insights.transactionCount}</span><strong>{transactions.length}</strong><small>{statements.length} {copy.insights.statementSources}</small></div>
            <div className="stat-card positive"><span>{copy.insights.totalIncome}</span><strong>{formatMoney(analytics.totalCredit, language)}</strong><small>{copy.insights.incomeHint} · {importConfig.currency}</small></div>
            <div className="stat-card negative"><span>{copy.insights.totalExpense}</span><strong>{formatMoney(analytics.totalDebit, language)}</strong><small>{copy.insights.expenseHint} · {importConfig.currency}</small></div>
            <div className="stat-card net"><span>{copy.insights.net}</span><strong>{analytics.net >= 0 ? "+" : "−"}{formatMoney(Math.abs(analytics.net), language)}</strong><small>{analytics.net >= 0 ? copy.insights.surplus : copy.insights.deficit}</small></div>
          </div>

          <div className="insight-grid">
            <article className="chart-panel monthly-panel">
              <div className="panel-heading"><div><span>{copy.insights.trend}</span><h3>{copy.insights.monthly}</h3></div><div className="legend"><i className="credit-dot" /> {copy.insights.income} <i className="debit-dot" /> {copy.insights.expense}</div></div>
              <div className="bar-chart">
                {analytics.months.map(([month, value]) => (
                  <div className="bar-group" key={month}>
                    <div className="bars"><i className="credit-bar" style={{ height: `${Math.max(4, value.credit / maxMonthly * 100)}%` }} title={`${copy.insights.income} ${formatMoney(value.credit, language)}`} /><i className="debit-bar" style={{ height: `${Math.max(4, value.debit / maxMonthly * 100)}%` }} title={`${copy.insights.expense} ${formatMoney(value.debit, language)}`} /></div>
                    <strong>{formatMonth(month, language)}</strong><small>{formatMoney(value.credit - value.debit, language)}</small>
                  </div>
                ))}
              </div>
            </article>

            <article className="chart-panel category-panel">
              <div className="panel-heading"><div><span>{copy.insights.suggested}</span><h3>{copy.insights.spending}</h3></div></div>
              <div className="category-list">
                {analytics.categories.map(([category, value], index) => (
                  <div className="category-item" key={category}><div><span><i>{String(index + 1).padStart(2, "0")}</i>{copy.categories[category]}</span><strong>{formatMoney(value, language)}</strong></div><span className="category-track"><i style={{ width: `${analytics.totalDebit ? value / analytics.totalDebit * 100 : 0}%` }} /></span></div>
                ))}
              </div>
            </article>
          </div>

          <div className="source-strip">
            {analytics.sources.map(([source, value]) => <div key={source}><strong>{source}</strong><span><b>+{formatMoney(value.credit, language)}</b><em>−{formatMoney(value.debit, language)}</em></span></div>)}
          </div>

          <article className="master-panel">
            <div className="panel-heading"><div><span>{copy.insights.master}</span><h3>{transactions.length} {copy.insights.merged}</h3></div><small>{copy.actions.editHint}</small></div>

            <div className="transaction-filters">
              <label className="filter-search"><span>{copy.filters.search}</span><input name="transaction-search" type="search" autoComplete="off" placeholder={copy.filters.searchPlaceholder} value={filters.query} onChange={(event) => setFilters((current) => ({ ...current, query: event.target.value }))} /></label>
              <label><span>{copy.filters.source}</span><select value={filters.source} onChange={(event) => setFilters((current) => ({ ...current, source: event.target.value }))}><option value="all">{copy.filters.allSources}</option>{sourceOptions.map((source) => <option value={source} key={source}>{source}</option>)}</select></label>
              <label><span>{copy.filters.kind}</span><select value={filters.kind} onChange={(event) => setFilters((current) => ({ ...current, kind: event.target.value as Filters["kind"] }))}><option value="all">{copy.filters.allKinds}</option><option value="Thu">{copy.incomeKind}</option><option value="Chi">{copy.expenseKind}</option></select></label>
              <label><span>{copy.filters.category}</span><select value={filters.category} onChange={(event) => setFilters((current) => ({ ...current, category: event.target.value as Filters["category"] }))}><option value="all">{copy.filters.allCategories}</option>{categoryKeys.map((category) => <option value={category} key={category}>{copy.categories[category]}</option>)}</select></label>
              <label className="issues-filter"><input type="checkbox" checked={filters.issuesOnly} onChange={(event) => setFilters((current) => ({ ...current, issuesOnly: event.target.checked }))} /><span>{copy.filters.issuesOnly}</span></label>
              <button className="filter-reset" type="button" onClick={() => setFilters(initialFilters)}>{copy.filters.reset}</button>
              <p>{filteredTransactions.length}/{transactions.length} {copy.filters.results}</p>
            </div>

            {filteredTransactions.length === 0 ? <div className="empty-results">{copy.filters.empty}</div> : (
              <>
                <div className="master-table-wrap desktop-transactions">
                  <table className="master-table editable-table">
                    <thead><tr>{copy.insights.tableHeaders.map((header) => <th key={header}>{header}</th>)}<th>{copy.actions.action}</th></tr></thead>
                    <tbody>
                      {filteredTransactions.map((item) => {
                        const issues = validation.issuesById.get(item.id) ?? [];
                        return (
                          <tr className={issues.length > 0 ? "has-issues" : ""} key={item.id}>
                            <td><input aria-label={copy.insights.tableHeaders[0]} value={item.date} placeholder="YYYY-MM-DD" inputMode="numeric" onChange={(event) => updateTransaction(item.id, { date: event.target.value })} /></td>
                            <td><strong>{item.source}</strong><small>{item.reference || "—"}</small></td>
                            <td><textarea aria-label={copy.insights.tableHeaders[2]} rows={2} value={item.description} onChange={(event) => updateTransaction(item.id, { description: event.target.value })} /><IssueBadges issues={issues} copy={copy} /></td>
                            <td><input aria-label={copy.insights.tableHeaders[3]} value={item.counterparty} onChange={(event) => updateTransaction(item.id, { counterparty: event.target.value })} /></td>
                            <td><input aria-label={copy.insights.tableHeaders[4]} type="number" step="any" inputMode="decimal" value={item.debit} onChange={(event) => updateTransaction(item.id, { debit: Number(event.target.value) || 0 })} /></td>
                            <td><input aria-label={copy.insights.tableHeaders[5]} type="number" step="any" inputMode="decimal" value={item.credit} onChange={(event) => updateTransaction(item.id, { credit: Number(event.target.value) || 0 })} /></td>
                            <td><input aria-label={copy.insights.tableHeaders[6]} type="number" step="any" inputMode="decimal" value={item.balance} onChange={(event) => updateTransaction(item.id, { balance: Number(event.target.value) || 0 })} /><small>{item.currency}</small></td>
                            <td><select aria-label={copy.insights.tableHeaders[7]} value={item.category} onChange={(event) => updateTransaction(item.id, { category: event.target.value as CategoryKey })}>{categoryKeys.map((category) => <option value={category} key={category}>{copy.categories[category]}</option>)}</select></td>
                            <td><button className="row-delete" type="button" onClick={() => removeTransaction(item.id)}>{copy.actions.delete}</button></td>
                          </tr>
                        );
                      })}
                    </tbody>
                  </table>
                </div>

                <div className="transaction-cards">
                  {filteredTransactions.map((item) => {
                    const issues = validation.issuesById.get(item.id) ?? [];
                    return (
                      <article className={`transaction-card ${issues.length > 0 ? "has-issues" : ""}`} key={item.id}>
                        <div className="transaction-card-heading"><div><strong>{item.source}</strong><small>{item.reference || "—"}</small></div><span>{item.currency}</span></div>
                        <IssueBadges issues={issues} copy={copy} />
                        <label><span>{copy.insights.tableHeaders[0]}</span><input value={item.date} placeholder="YYYY-MM-DD" inputMode="numeric" onChange={(event) => updateTransaction(item.id, { date: event.target.value })} /></label>
                        <label><span>{copy.insights.tableHeaders[2]}</span><textarea rows={2} value={item.description} onChange={(event) => updateTransaction(item.id, { description: event.target.value })} /></label>
                        <label><span>{copy.insights.tableHeaders[3]}</span><input value={item.counterparty} onChange={(event) => updateTransaction(item.id, { counterparty: event.target.value })} /></label>
                        <div className="transaction-card-amounts">
                          <label><span>{copy.insights.tableHeaders[4]}</span><input type="number" step="any" inputMode="decimal" value={item.debit} onChange={(event) => updateTransaction(item.id, { debit: Number(event.target.value) || 0 })} /></label>
                          <label><span>{copy.insights.tableHeaders[5]}</span><input type="number" step="any" inputMode="decimal" value={item.credit} onChange={(event) => updateTransaction(item.id, { credit: Number(event.target.value) || 0 })} /></label>
                          <label><span>{copy.insights.tableHeaders[6]}</span><input type="number" step="any" inputMode="decimal" value={item.balance} onChange={(event) => updateTransaction(item.id, { balance: Number(event.target.value) || 0 })} /></label>
                        </div>
                        <label><span>{copy.insights.tableHeaders[7]}</span><select value={item.category} onChange={(event) => updateTransaction(item.id, { category: event.target.value as CategoryKey })}>{categoryKeys.map((category) => <option value={category} key={category}>{copy.categories[category]}</option>)}</select></label>
                        <button className="row-delete" type="button" onClick={() => removeTransaction(item.id)}>{copy.actions.delete}</button>
                      </article>
                    );
                  })}
                </div>
              </>
            )}
          </article>

          <div className="export-heading"><span className="section-number">04</span><div><h2>{copy.export.title}</h2><p>{copy.export.description}</p></div></div>
          <div className="export-grid">
            <article className="export-card featured-export">
              <span className="export-label">{copy.export.recommended}</span><div className="export-icon" aria-hidden="true">↳</div><h3>{copy.export.templateTitle}</h3><p>{copy.export.templateBody}</p>
              <label className="template-picker"><input name="excel-template" type="file" accept=".xlsx,.xls" onChange={(event) => setTemplateFile(event.target.files?.[0] ?? null)} /><span>{templateFile ? templateFile.name : copy.export.chooseTemplate}</span><b>{copy.export.chooseFile}</b></label>
              <button className="primary-button full-button" type="button" disabled={!templateFile} onClick={() => void exportToTemplate()}>{copy.export.fillDownload} <span>↓</span></button>
            </article>
            <article className="export-card"><div className="export-icon" aria-hidden="true">▦</div><h3>{copy.export.fullTitle}</h3><p>{copy.export.fullBody}</p><button className="outline-button" type="button" onClick={() => exportWorkbook(false)}>{copy.export.fullDownload} <span>↓</span></button></article>
            <article className="export-card"><div className="export-icon" aria-hidden="true">≡</div><h3>{copy.export.compactTitle}</h3><p>{copy.export.compactBody}</p><button className="outline-button" type="button" onClick={() => exportWorkbook(true)}>{copy.export.compactDownload} <span>↓</span></button></article>
          </div>
        </section>
      )}

      <footer>
        <a className="brand footer-brand" href="#top" aria-label={copy.backToTop} translate="no"><span className="brand-mark">M</span><span><strong>Mạch Tiền</strong><small>{copy.brandTagline}</small></span></a>
        <p>{copy.footerPrivacy}</p><span>{copy.footerNote}</span>
      </footer>

      {undoItem && <div className="undo-toast" role="status" aria-live="polite"><span>{copy.actions.removed}</span><button type="button" onClick={restoreUndo}>{copy.actions.undo}</button><button className="undo-close" type="button" aria-label={copy.actions.dismiss} onClick={() => setUndoItem(null)}>×</button></div>}
    </main>
  );
}
