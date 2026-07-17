"use client";

import { useEffect, useMemo, useState } from "react";
import type { ChangeEvent, DragEvent } from "react";
import * as XLSX from "xlsx";
import { languageOptions, messages } from "./i18n";
import type { CategoryKey, Language, MessageSet } from "./i18n";

type CanonicalKey =
  | "date"
  | "reference"
  | "description"
  | "debit"
  | "credit"
  | "amount"
  | "balance"
  | "counterparty";

type MappingValue = CanonicalKey | "ignore";

interface StatementFile {
  id: string;
  fileName: string;
  source: string;
  headers: string[];
  rows: Record<string, unknown>[];
  mapping: Record<string, MappingValue>;
}

interface Transaction {
  source: string;
  date: string;
  reference: string;
  description: string;
  counterparty: string;
  debit: number;
  credit: number;
  balance: number;
  kind: "Thu" | "Chi";
  category: CategoryKey;
}

const fieldOptions: MappingValue[] = ["ignore", "date", "reference", "description", "debit", "credit", "amount", "balance", "counterparty"];

const locales: Record<Language, string> = {
  vi: "vi-VN",
  en: "en-US",
  ja: "ja-JP",
  zh: "zh-CN",
  ko: "ko-KR",
};

const fieldAliases: Record<CanonicalKey, string[]> = {
  date: ["ngay giao dich", "ngay gd", "transaction date", "posting date", "ngay", "取引日", "日付", "交易日期", "记账日期", "日期", "거래일자", "거래일", "일자"],
  reference: ["so chung tu", "so ct", "ma giao dich", "transaction id", "reference", "ref", "取引番号", "参照番号", "交易编号", "参考号", "流水号", "凭证号", "거래번호", "참조번호", "전표번호"],
  description: ["noi dung giao dich", "dien giai", "description", "remark", "noi dung", "摘要", "取引内容", "交易说明", "交易内容", "备注", "적요", "거래내용", "내용"],
  debit: ["phat sinh no", "ghi no", "debit", "withdrawal", "tien ra", "出金", "支払", "借方", "支出", "付款金额", "출금", "지급", "차변"],
  credit: ["phat sinh co", "ghi co", "credit", "deposit", "tien vao", "入金", "預入", "贷方", "收入", "收款金额", "입금", "수입", "대변"],
  amount: ["so tien", "amount", "transaction amount", "gia tri", "金額", "取引金額", "金额", "交易金额", "금액", "거래금액"],
  balance: ["so du cuoi", "so du", "balance", "closing balance", "残高", "余额", "账户余额", "잔액"],
  counterparty: ["nguoi nhan", "nguoi chuyen", "doi tac", "beneficiary", "counterparty", "取引先", "受取人", "依頼人", "交易对方", "对方户名", "收款人", "付款人", "거래처", "수취인", "송금인"],
};

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

function normalizeText(value: string) {
  return value
    .normalize("NFD")
    .replace(/[\u0300-\u036f]/g, "")
    .toLowerCase()
    .replace(/đ/g, "d")
    .replace(/[^\p{L}\p{N}]+/gu, " ")
    .trim();
}

function inferField(header: string): MappingValue {
  const normalized = normalizeText(header);

  for (const [key, aliases] of Object.entries(fieldAliases) as [CanonicalKey, string[]][]) {
    if (aliases.some((alias) => {
      const normalizedAlias = normalizeText(alias);
      return normalized === normalizedAlias || normalized.includes(normalizedAlias);
    })) {
      return key;
    }
  }

  return "ignore";
}

function inferBankName(fileName: string, unknownBank: string) {
  const normalized = normalizeText(fileName);
  const banks: [string, string[]][] = [
    ["Vietcombank", ["vietcombank", "vcb"]],
    ["BIDV", ["bidv"]],
    ["VietinBank", ["vietinbank", "ctg"]],
    ["Techcombank", ["techcombank", "tcb"]],
    ["MBBank", ["mbbank", "mb bank"]],
    ["ACB", ["acb"]],
    ["Sacombank", ["sacombank", "stb"]],
    ["VPBank", ["vpbank", "vpb"]],
  ];

  return banks.find(([, aliases]) => aliases.some((alias) => normalized.includes(alias)))?.[0] ?? unknownBank;
}

function createStatement(fileName: string, source: string, rows: Record<string, unknown>[]): StatementFile {
  const headers = rows.length > 0 ? Object.keys(rows[0]) : [];
  return {
    id: `demo-${fileName}`,
    fileName,
    source,
    headers,
    rows,
    mapping: Object.fromEntries(headers.map((header) => [header, inferField(header)])),
  };
}

function uniqueHeaders(values: unknown[], columnLabel: string) {
  const counts = new Map<string, number>();

  return values.map((value, index) => {
    const base = String(value ?? "").trim() || `${columnLabel} ${index + 1}`;
    const count = counts.get(base) ?? 0;
    counts.set(base, count + 1);
    return count === 0 ? base : `${base} (${count + 1})`;
  });
}

async function parseStatementFile(file: File, copy: MessageSet): Promise<StatementFile> {
  const data = await file.arrayBuffer();
  const workbook = XLSX.read(data, { cellDates: true });
  const sheetName = workbook.SheetNames[0];

  if (!sheetName) {
    throw new Error(`${file.name} ${copy.errors.emptySheet}`);
  }

  const sheet = workbook.Sheets[sheetName];
  const grid = XLSX.utils.sheet_to_json<unknown[]>(sheet, { header: 1, defval: "", raw: true });
  const headerIndex = grid
    .slice(0, 15)
    .map((row, index) => ({
      index,
      populated: row.filter((cell) => String(cell).trim()).length,
      recognized: row.filter((cell) => inferField(String(cell)) !== "ignore").length,
    }))
    .filter((candidate) => candidate.populated >= 2)
    .sort((a, b) => (b.recognized * 10 + b.populated) - (a.recognized * 10 + a.populated))[0]?.index ?? -1;

  if (headerIndex < 0) {
    throw new Error(`${copy.errors.missingHeader} ${file.name}.`);
  }

  const headers = uniqueHeaders(grid[headerIndex], copy.mapping.column);
  const rows = grid
    .slice(headerIndex + 1)
    .filter((row) => row.some((cell) => String(cell).trim()))
    .map((row) => Object.fromEntries(headers.map((header, index) => [header, row[index] ?? ""])));

  if (rows.length === 0) {
    throw new Error(`${file.name} ${copy.errors.noRows}`);
  }

  return {
    id: `${file.name}-${file.lastModified}-${Math.random().toString(36).slice(2)}`,
    fileName: file.name,
    source: inferBankName(file.name, copy.unknownBank),
    headers,
    rows,
    mapping: Object.fromEntries(headers.map((header) => [header, inferField(header)])),
  };
}

function parseAmount(value: unknown) {
  if (typeof value === "number") return Number.isFinite(value) ? value : 0;
  if (value === null || value === undefined || value === "") return 0;

  const text = String(value).trim();
  const negative = /^\(.*\)$/.test(text) || text.startsWith("-");
  const digits = text.replace(/[^0-9]/g, "");
  const amount = digits ? Number(digits) : 0;
  return negative ? -amount : amount;
}

function parseDate(value: unknown) {
  if (value instanceof Date && !Number.isNaN(value.getTime())) {
    return value.toISOString().slice(0, 10);
  }

  if (typeof value === "number") {
    const parsed = XLSX.SSF.parse_date_code(value);
    if (parsed) {
      return `${parsed.y}-${String(parsed.m).padStart(2, "0")}-${String(parsed.d).padStart(2, "0")}`;
    }
  }

  const text = String(value ?? "").trim();
  const dayFirst = text.match(/^(\d{1,2})[/.\-](\d{1,2})[/.\-](\d{4})$/);
  if (dayFirst) {
    return `${dayFirst[3]}-${dayFirst[2].padStart(2, "0")}-${dayFirst[1].padStart(2, "0")}`;
  }

  const yearFirst = text.match(/^(\d{4})[/.\-](\d{1,2})[/.\-](\d{1,2})$/);
  if (yearFirst) {
    return `${yearFirst[1]}-${yearFirst[2].padStart(2, "0")}-${yearFirst[3].padStart(2, "0")}`;
  }

  return text;
}

function classifyTransaction(description: string, kind: Transaction["kind"]) {
  const text = normalizeText(description);
  const rules: [CategoryKey, string[]][] = [
    ["payroll", ["luong", "nhan vien", "bao hiem", "給与", "従業員", "社会保険", "工资", "员工", "社保", "급여", "직원", "보험"]],
    ["tax", ["thue", "ngan sach", "kho bac", "税", "国庫", "国库", "财政", "세금", "국고"]],
    ["bankFees", ["phi quan ly", "phi giao dich", "sms banking", "phi dich vu", "手数料", "口座管理", "手续费", "账户管理费", "短信服务费", "수수료", "계좌관리"]],
    ["suppliers", ["ncc", "nha cung cap", "thanh toan", "tien thue", "tien dien", "仕入", "支払", "家賃", "電気", "供应商", "付款", "房租", "电费", "采购", "공급업체", "결제", "임대료", "전기요금", "구매"]],
    ["finance", ["lai tien gui", "lai suat", "利息", "存款利息", "이자", "예금이자"]],
  ];

  return rules.find(([, words]) => words.some((word) => text.includes(normalizeText(word))))?.[0] ?? (kind === "Thu" ? "customerIncome" : "otherExpense");
}

function getMappedValue(statement: StatementFile, row: Record<string, unknown>, key: CanonicalKey) {
  const header = statement.headers.find((candidate) => statement.mapping[candidate] === key);
  return header ? row[header] : "";
}

function normalizeStatements(statements: StatementFile[]) {
  return statements
    .flatMap((statement) =>
      statement.rows.map((row) => {
        let debit = Math.abs(parseAmount(getMappedValue(statement, row, "debit")));
        let credit = Math.abs(parseAmount(getMappedValue(statement, row, "credit")));
        const signedAmount = parseAmount(getMappedValue(statement, row, "amount"));

        if (debit === 0 && credit === 0 && signedAmount !== 0) {
          if (signedAmount < 0) debit = Math.abs(signedAmount);
          else credit = signedAmount;
        }

        const kind: Transaction["kind"] = credit >= debit ? "Thu" : "Chi";
        const description = String(getMappedValue(statement, row, "description") ?? "").trim();

        return {
          source: statement.source,
          date: parseDate(getMappedValue(statement, row, "date")),
          reference: String(getMappedValue(statement, row, "reference") ?? "").trim(),
          description,
          counterparty: String(getMappedValue(statement, row, "counterparty") ?? "").trim(),
          debit,
          credit,
          balance: parseAmount(getMappedValue(statement, row, "balance")),
          kind,
          category: classifyTransaction(description, kind),
        } satisfies Transaction;
      }),
    )
    .filter((transaction) => transaction.date || transaction.description || transaction.debit || transaction.credit)
    .sort((a, b) => a.date.localeCompare(b.date));
}

function formatMoney(value: number, language: Language) {
  return new Intl.NumberFormat(locales[language]).format(value);
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
    [columns.kind]: transaction.kind === "Thu" ? copy.incomeKind : copy.expenseKind,
    [columns.category]: copy.categories[transaction.category],
  }));
}

function transactionValue(transaction: Transaction, field: MappingValue) {
  if (field === "ignore") return "";
  const values: Record<CanonicalKey, string | number> = {
    date: transaction.date,
    reference: transaction.reference,
    description: transaction.description,
    debit: transaction.debit,
    credit: transaction.credit,
    amount: transaction.credit - transaction.debit,
    balance: transaction.balance,
    counterparty: transaction.counterparty,
  };
  return values[field];
}

export default function Home() {
  const [language, setLanguage] = useState<Language>("vi");
  const [statements, setStatements] = useState<StatementFile[]>([]);
  const [transactions, setTransactions] = useState<Transaction[]>([]);
  const [templateFile, setTemplateFile] = useState<File | null>(null);
  const [isDragging, setIsDragging] = useState(false);
  const [isReading, setIsReading] = useState(false);
  const [error, setError] = useState("");
  const copy = messages[language];

  useEffect(() => {
    const stored = window.localStorage.getItem("mach-tien-language");
    if (stored === "vi" || stored === "en" || stored === "ja" || stored === "zh" || stored === "ko") {
      setLanguage(stored);
      return;
    }

    const browserLanguage = window.navigator.language.toLowerCase();
    if (browserLanguage.startsWith("zh")) setLanguage("zh");
    else if (browserLanguage.startsWith("ko")) setLanguage("ko");
    else if (browserLanguage.startsWith("ja")) setLanguage("ja");
    else if (browserLanguage.startsWith("en")) setLanguage("en");
  }, []);

  useEffect(() => {
    document.documentElement.lang = language;
    document.title = copy.metaTitle;
  }, [copy.metaTitle, language]);

  const analytics = useMemo(() => {
    const totalCredit = transactions.reduce((sum, item) => sum + item.credit, 0);
    const totalDebit = transactions.reduce((sum, item) => sum + item.debit, 0);
    const categories = new Map<CategoryKey, number>();
    const months = new Map<string, { credit: number; debit: number }>();
    const sources = new Map<string, { credit: number; debit: number }>();

    transactions.forEach((item) => {
      const month = item.date.slice(0, 7) || "—";
      const monthly = months.get(month) ?? { credit: 0, debit: 0 };
      monthly.credit += item.credit;
      monthly.debit += item.debit;
      months.set(month, monthly);

      const source = sources.get(item.source) ?? { credit: 0, debit: 0 };
      source.credit += item.credit;
      source.debit += item.debit;
      sources.set(item.source, source);

      if (item.debit > 0) {
        categories.set(item.category, (categories.get(item.category) ?? 0) + item.debit);
      }
    });

    return {
      totalCredit,
      totalDebit,
      net: totalCredit - totalDebit,
      closingBalance: transactions.at(-1)?.balance ?? 0,
      categories: [...categories.entries()].sort((a, b) => b[1] - a[1]),
      months: [...months.entries()].sort(([a], [b]) => a.localeCompare(b)),
      sources: [...sources.entries()],
    };
  }, [transactions]);

  async function addFiles(files: File[]) {
    if (files.length === 0) return;
    setIsReading(true);
    setError("");

    try {
      const parsed = await Promise.all(files.map((file) => parseStatementFile(file, copy)));
      setStatements((current) => [...current, ...parsed]);
      setTransactions([]);
    } catch (caught) {
      setError(caught instanceof Error ? caught.message : copy.errors.readFile);
    } finally {
      setIsReading(false);
    }
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

  function updateStatement(id: string, updater: (statement: StatementFile) => StatementFile) {
    setStatements((current) => current.map((statement) => (statement.id === id ? updater(statement) : statement)));
    setTransactions([]);
  }

  function buildMasterTable() {
    const normalized = normalizeStatements(statements);
    setTransactions(normalized);
    setError(normalized.length === 0 ? copy.errors.noTransactions : "");
    if (normalized.length > 0) {
      requestAnimationFrame(() => document.getElementById("insights")?.scrollIntoView({ behavior: "smooth", block: "start" }));
    }
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

  return (
    <main>
      <header className="site-header">
        <a className="brand" href="#top" aria-label={copy.backToTop}>
          <span className="brand-mark">M</span>
          <span>
            <strong>Mạch Tiền</strong>
            <small>{copy.brandTagline}</small>
          </span>
        </a>
        <div className="header-actions">
          <div className="language-switcher" role="group" aria-label={copy.languageLabel}>
            {languageOptions.map((option) => (
              <button
                className={language === option.code ? "active" : ""}
                type="button"
                aria-pressed={language === option.code}
                lang={option.code}
                key={option.code}
                onClick={() => changeLanguage(option.code)}
              >
                {option.label}
              </button>
            ))}
          </div>
          <div className="privacy-pill"><span /> {copy.privacy}</div>
        </div>
      </header>

      <section className="hero" id="top">
        <div className="hero-copy">
          <div className="eyebrow">{copy.hero.eyebrow}</div>
          <h1>{copy.hero.title} <em>{copy.hero.accent}</em></h1>
          <p>{copy.hero.body}</p>
          <div className="hero-badges">
            {copy.hero.badges.map((badge) => <span key={badge}>{badge}</span>)}
          </div>
        </div>
        <div className="hero-visual" aria-label={copy.hero.netCashFlow}>
          <div className="floating-note note-one">VCB · +{formatMoney(48_500_000, language)}</div>
          <div className="floating-note note-two">BIDV · −{formatMoney(14_800_000, language)}</div>
          <div className="statement-card">
            <div className="statement-top"><span>{copy.hero.quarter}</span><span>{copy.hero.transactions}</span></div>
            <div className="pulse-line"><i /><i /><i /><i /><i /><i /></div>
            <div className="statement-total">
              <span>{copy.hero.netCashFlow}</span>
              <strong>+{formatMoney(54_200_000, language)}</strong>
            </div>
            <div className="mini-bars"><i /><i /><i /><i /><i /></div>
          </div>
        </div>
      </section>

      <nav className="stepper" aria-label={copy.workflowLabel}>
        {copy.steps.map(([label, detail], index) => {
          const number = String(index + 1).padStart(2, "0");
          return (
          <div className={`step ${statements.length > 0 && index < 2 ? "active" : ""} ${transactions.length > 0 && index < 4 ? "active" : ""}`} key={number}>
            <span>{number}</span>
            <div><strong>{label}</strong><small>{detail}</small></div>
          </div>
          );
        })}
      </nav>

      <section className="workspace-section">
        <div className="section-heading">
          <div><span className="section-number">01</span><h2>{copy.upload.title}</h2></div>
          <p>{copy.upload.description}</p>
        </div>

        <label
          className={`drop-zone ${isDragging ? "dragging" : ""}`}
          onDragEnter={() => setIsDragging(true)}
          onDragLeave={() => setIsDragging(false)}
          onDragOver={(event) => event.preventDefault()}
          onDrop={handleDrop}
        >
          <input type="file" accept=".xlsx,.xls,.csv" multiple onChange={handleFileInput} />
          <span className="upload-icon">↥</span>
          <strong>{isReading ? copy.upload.reading : copy.upload.ready}</strong>
          <small>{copy.upload.formats}</small>
        </label>

        <div className="demo-row">
          <button
            className="text-button"
            type="button"
            onClick={() => {
              setStatements(demoFiles.map((file) => ({ ...file, id: `${file.id}-${Math.random()}` })));
              setTransactions([]);
              setError("");
            }}
          >
            {copy.upload.demo} <span>→</span>
          </button>
          <p>{copy.upload.demoHint}</p>
        </div>

        {error && <div className="error-message" role="alert">{error}</div>}
      </section>

      {statements.length > 0 && (
        <section className="mapping-section">
          <div className="section-heading">
            <div><span className="section-number">02</span><h2>{copy.mapping.title}</h2></div>
            <p>{copy.mapping.description}</p>
          </div>

          <div className="file-stack">
            {statements.map((statement) => (
              <article className="file-card" key={statement.id}>
                <div className="file-card-header">
                  <div className="file-identity">
                    <span className="file-icon">XL</span>
                    <div><strong>{statement.fileName}</strong><small>{statement.rows.length} {copy.mapping.rows}</small></div>
                  </div>
                  <button className="remove-button" type="button" onClick={() => setStatements((current) => current.filter((item) => item.id !== statement.id))}>{copy.mapping.remove}</button>
                </div>
                <div className="source-control">
                  <label htmlFor={`source-${statement.id}`}>{copy.mapping.source}</label>
                  <input
                    id={`source-${statement.id}`}
                    value={statement.source}
                    onChange={(event) => updateStatement(statement.id, (current) => ({ ...current, source: event.target.value }))}
                  />
                  <span>{statement.headers.filter((header) => statement.mapping[header] !== "ignore").length}/{statement.headers.length} {copy.mapping.mapped}</span>
                </div>
                <div className="mapping-table-wrap">
                  <table className="mapping-table">
                    <thead><tr>{copy.mapping.headers.map((header) => <th key={header}>{header}</th>)}</tr></thead>
                    <tbody>
                      {statement.headers.map((header) => (
                        <tr key={header}>
                          <td>{header}</td>
                          <td>{String(statement.rows[0]?.[header] ?? "—") || "—"}</td>
                          <td>
                            <select
                              aria-label={`${copy.mapping.fieldLabel} ${header}`}
                              value={statement.mapping[header]}
                              onChange={(event) => updateStatement(statement.id, (current) => ({
                                ...current,
                                mapping: { ...current.mapping, [header]: event.target.value as MappingValue },
                              }))}
                            >
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
            <button className="secondary-button" type="button" onClick={() => { setStatements([]); setTransactions([]); }}>{copy.mapping.reset}</button>
            <button className="primary-button" type="button" onClick={buildMasterTable}>{copy.mapping.build} <span>→</span></button>
          </div>
        </section>
      )}

      {transactions.length > 0 && (
        <section className="insights-section" id="insights">
          <div className="section-heading light-heading">
            <div><span className="section-number">03</span><h2>{copy.insights.title}</h2></div>
            <p>{transactions[0]?.date} — {transactions.at(-1)?.date}</p>
          </div>

          <div className="stat-grid">
            <div className="stat-card"><span>{copy.insights.transactionCount}</span><strong>{transactions.length}</strong><small>{statements.length} {copy.insights.statementSources}</small></div>
            <div className="stat-card positive"><span>{copy.insights.totalIncome}</span><strong>{formatMoney(analytics.totalCredit, language)}</strong><small>{copy.insights.incomeHint}</small></div>
            <div className="stat-card negative"><span>{copy.insights.totalExpense}</span><strong>{formatMoney(analytics.totalDebit, language)}</strong><small>{copy.insights.expenseHint}</small></div>
            <div className="stat-card net"><span>{copy.insights.net}</span><strong>{analytics.net >= 0 ? "+" : "−"}{formatMoney(Math.abs(analytics.net), language)}</strong><small>{analytics.net >= 0 ? copy.insights.surplus : copy.insights.deficit}</small></div>
          </div>

          <div className="insight-grid">
            <article className="chart-panel monthly-panel">
              <div className="panel-heading"><div><span>{copy.insights.trend}</span><h3>{copy.insights.monthly}</h3></div><div className="legend"><i className="credit-dot" /> {copy.insights.income} <i className="debit-dot" /> {copy.insights.expense}</div></div>
              <div className="bar-chart">
                {analytics.months.map(([month, value]) => (
                  <div className="bar-group" key={month}>
                    <div className="bars">
                      <i className="credit-bar" style={{ height: `${Math.max(4, value.credit / maxMonthly * 100)}%` }} title={`${copy.insights.income} ${formatMoney(value.credit, language)}`} />
                      <i className="debit-bar" style={{ height: `${Math.max(4, value.debit / maxMonthly * 100)}%` }} title={`${copy.insights.expense} ${formatMoney(value.debit, language)}`} />
                    </div>
                    <strong>{month.slice(5)}/{month.slice(2, 4)}</strong>
                    <small>{formatMoney(value.credit - value.debit, language)}</small>
                  </div>
                ))}
              </div>
            </article>

            <article className="chart-panel category-panel">
              <div className="panel-heading"><div><span>{copy.insights.suggested}</span><h3>{copy.insights.spending}</h3></div></div>
              <div className="category-list">
                {analytics.categories.map(([category, value], index) => (
                  <div className="category-item" key={category}>
                    <div><span><i>{String(index + 1).padStart(2, "0")}</i>{copy.categories[category]}</span><strong>{formatMoney(value, language)}</strong></div>
                    <span className="category-track"><i style={{ width: `${analytics.totalDebit ? value / analytics.totalDebit * 100 : 0}%` }} /></span>
                  </div>
                ))}
              </div>
            </article>
          </div>

          <div className="source-strip">
            {analytics.sources.map(([source, value]) => (
              <div key={source}><strong>{source}</strong><span><b>+{formatMoney(value.credit, language)}</b><em>−{formatMoney(value.debit, language)}</em></span></div>
            ))}
          </div>

          <article className="master-panel">
            <div className="panel-heading"><div><span>{copy.insights.master}</span><h3>{transactions.length} {copy.insights.merged}</h3></div><small>{copy.insights.scroll}</small></div>
            <div className="master-table-wrap">
              <table className="master-table">
                <thead><tr>{copy.insights.tableHeaders.map((header) => <th key={header}>{header}</th>)}</tr></thead>
                <tbody>
                  {transactions.map((item, index) => (
                    <tr key={`${item.source}-${item.date}-${item.reference}-${index}`}>
                      <td>{item.date}</td><td>{item.source}</td><td>{item.description || "—"}</td><td>{item.counterparty || "—"}</td>
                      <td className="money debit-text">{item.debit ? formatMoney(item.debit, language) : "—"}</td><td className="money credit-text">{item.credit ? formatMoney(item.credit, language) : "—"}</td>
                      <td className="money">{formatMoney(item.balance, language)}</td><td><span className={`category-tag ${item.kind === "Thu" ? "tag-credit" : ""}`}>{copy.categories[item.category]}</span></td>
                    </tr>
                  ))}
                </tbody>
              </table>
            </div>
          </article>

          <div className="export-heading">
            <span className="section-number">04</span>
            <div><h2>{copy.export.title}</h2><p>{copy.export.description}</p></div>
          </div>
          <div className="export-grid">
            <article className="export-card featured-export">
              <span className="export-label">{copy.export.recommended}</span>
              <div className="export-icon">↳</div>
              <h3>{copy.export.templateTitle}</h3>
              <p>{copy.export.templateBody}</p>
              <label className="template-picker">
                <input type="file" accept=".xlsx,.xls" onChange={(event) => setTemplateFile(event.target.files?.[0] ?? null)} />
                <span>{templateFile ? templateFile.name : copy.export.chooseTemplate}</span><b>{copy.export.chooseFile}</b>
              </label>
              <button className="primary-button full-button" type="button" disabled={!templateFile} onClick={() => void exportToTemplate()}>{copy.export.fillDownload} <span>↓</span></button>
            </article>
            <article className="export-card">
              <div className="export-icon">▦</div><h3>{copy.export.fullTitle}</h3>
              <p>{copy.export.fullBody}</p>
              <button className="outline-button" type="button" onClick={() => exportWorkbook(false)}>{copy.export.fullDownload} <span>↓</span></button>
            </article>
            <article className="export-card">
              <div className="export-icon">≡</div><h3>{copy.export.compactTitle}</h3>
              <p>{copy.export.compactBody}</p>
              <button className="outline-button" type="button" onClick={() => exportWorkbook(true)}>{copy.export.compactDownload} <span>↓</span></button>
            </article>
          </div>
        </section>
      )}

      <footer>
        <a className="brand footer-brand" href="#top" aria-label={copy.backToTop}><span className="brand-mark">M</span><span><strong>Mạch Tiền</strong><small>{copy.brandTagline}</small></span></a>
        <p>{copy.footerPrivacy}</p>
        <span>{copy.footerNote}</span>
      </footer>
    </main>
  );
}
