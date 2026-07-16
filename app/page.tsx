"use client";

import { useMemo, useState } from "react";
import type { ChangeEvent, DragEvent } from "react";
import * as XLSX from "xlsx";

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
  category: string;
}

interface FieldOption {
  key: MappingValue;
  label: string;
}

const fieldOptions: FieldOption[] = [
  { key: "ignore", label: "Bỏ qua cột này" },
  { key: "date", label: "Ngày giao dịch" },
  { key: "reference", label: "Số giao dịch / chứng từ" },
  { key: "description", label: "Diễn giải / nội dung" },
  { key: "debit", label: "Ghi nợ · tiền ra" },
  { key: "credit", label: "Ghi có · tiền vào" },
  { key: "amount", label: "Số tiền · một cột có dấu" },
  { key: "balance", label: "Số dư" },
  { key: "counterparty", label: "Đối tác" },
];

const fieldAliases: Record<CanonicalKey, string[]> = {
  date: ["ngay giao dich", "ngay gd", "transaction date", "posting date", "ngay"],
  reference: ["so chung tu", "so ct", "ma giao dich", "transaction id", "reference", "ref"],
  description: ["noi dung giao dich", "dien giai", "description", "remark", "noi dung"],
  debit: ["phat sinh no", "ghi no", "debit", "withdrawal", "tien ra"],
  credit: ["phat sinh co", "ghi co", "credit", "deposit", "tien vao"],
  amount: ["so tien", "amount", "transaction amount", "gia tri"],
  balance: ["so du cuoi", "so du", "balance", "closing balance"],
  counterparty: ["nguoi nhan", "nguoi chuyen", "doi tac", "beneficiary", "counterparty"],
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
    .replace(/[^a-zA-Z0-9]+/g, " ")
    .trim();
}

function inferField(header: string): MappingValue {
  const normalized = normalizeText(header);

  for (const [key, aliases] of Object.entries(fieldAliases) as [CanonicalKey, string[]][]) {
    if (aliases.some((alias) => normalized === alias || normalized.includes(alias))) {
      return key;
    }
  }

  return "ignore";
}

function inferBankName(fileName: string) {
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

  return banks.find(([, aliases]) => aliases.some((alias) => normalized.includes(alias)))?.[0] ?? "Ngân hàng chưa xác định";
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

function uniqueHeaders(values: unknown[]) {
  const counts = new Map<string, number>();

  return values.map((value, index) => {
    const base = String(value ?? "").trim() || `Cột ${index + 1}`;
    const count = counts.get(base) ?? 0;
    counts.set(base, count + 1);
    return count === 0 ? base : `${base} (${count + 1})`;
  });
}

async function parseStatementFile(file: File): Promise<StatementFile> {
  const data = await file.arrayBuffer();
  const workbook = XLSX.read(data, { cellDates: true });
  const sheetName = workbook.SheetNames[0];

  if (!sheetName) {
    throw new Error(`${file.name} không có trang dữ liệu.`);
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
    throw new Error(`Không tìm thấy dòng tiêu đề trong ${file.name}.`);
  }

  const headers = uniqueHeaders(grid[headerIndex]);
  const rows = grid
    .slice(headerIndex + 1)
    .filter((row) => row.some((cell) => String(cell).trim()))
    .map((row) => Object.fromEntries(headers.map((header, index) => [header, row[index] ?? ""])));

  if (rows.length === 0) {
    throw new Error(`${file.name} không có giao dịch bên dưới dòng tiêu đề.`);
  }

  return {
    id: `${file.name}-${file.lastModified}-${Math.random().toString(36).slice(2)}`,
    fileName: file.name,
    source: inferBankName(file.name),
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
  const rules: [string, string[]][] = [
    ["Lương & nhân sự", ["luong", "nhan vien", "bao hiem"]],
    ["Thuế & ngân sách", ["thue", "ngan sach", "kho bac"]],
    ["Phí ngân hàng", ["phi quan ly", "phi giao dich", "sms banking", "phi dich vu"]],
    ["Nhà cung cấp", ["ncc", "nha cung cap", "thanh toan", "tien thue", "tien dien"]],
    ["Lãi & tài chính", ["lai tien gui", "lai suat"]],
  ];

  return rules.find(([, words]) => words.some((word) => text.includes(word)))?.[0] ?? (kind === "Thu" ? "Thu khách hàng" : "Chi khác");
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

function formatMoney(value: number) {
  return new Intl.NumberFormat("vi-VN").format(value);
}

function exportRows(transactions: Transaction[]) {
  return transactions.map((transaction) => ({
    "Ngân hàng": transaction.source,
    "Ngày giao dịch": transaction.date,
    "Số giao dịch": transaction.reference,
    "Diễn giải": transaction.description,
    "Đối tác": transaction.counterparty,
    "Ghi nợ": transaction.debit,
    "Ghi có": transaction.credit,
    "Số dư": transaction.balance,
    "Loại": transaction.kind,
    "Nhãn": transaction.category,
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
  const [statements, setStatements] = useState<StatementFile[]>([]);
  const [transactions, setTransactions] = useState<Transaction[]>([]);
  const [templateFile, setTemplateFile] = useState<File | null>(null);
  const [isDragging, setIsDragging] = useState(false);
  const [isReading, setIsReading] = useState(false);
  const [error, setError] = useState("");

  const analytics = useMemo(() => {
    const totalCredit = transactions.reduce((sum, item) => sum + item.credit, 0);
    const totalDebit = transactions.reduce((sum, item) => sum + item.debit, 0);
    const categories = new Map<string, number>();
    const months = new Map<string, { credit: number; debit: number }>();
    const sources = new Map<string, { credit: number; debit: number }>();

    transactions.forEach((item) => {
      const month = item.date.slice(0, 7) || "Không rõ";
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
      const parsed = await Promise.all(files.map(parseStatementFile));
      setStatements((current) => [...current, ...parsed]);
      setTransactions([]);
    } catch (caught) {
      setError(caught instanceof Error ? caught.message : "Không thể đọc file sao kê này.");
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
    setError(normalized.length === 0 ? "Không tìm thấy giao dịch hợp lệ. Hãy kiểm tra lại phần khớp cột." : "");
    if (normalized.length > 0) {
      requestAnimationFrame(() => document.getElementById("insights")?.scrollIntoView({ behavior: "smooth", block: "start" }));
    }
  }

  function exportWorkbook(compact: boolean) {
    const rows = compact
      ? transactions.map((item) => ({
          Ngày: item.date,
          "Diễn giải": item.description,
          "Tiền thu": item.credit,
          "Tiền chi": item.debit,
          "Số dư": item.balance,
          "Nhãn": item.category,
        }))
      : exportRows(transactions);
    const workbook = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(workbook, XLSX.utils.json_to_sheet(rows), compact ? "Mẫu gọn" : "Bảng chuẩn");
    XLSX.writeFile(workbook, compact ? "sao-ke-mau-gon.xlsx" : "sao-ke-bang-chuan.xlsx");
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
        XLSX.utils.book_append_sheet(workbook, XLSX.utils.json_to_sheet(exportRows(transactions)), "Dữ liệu chuẩn");
      } else {
        const headers = grid[headerIndex].map((cell) => inferField(String(cell)));
        const rows = transactions.map((transaction) => headers.map((field) => transactionValue(transaction, field)));
        XLSX.utils.sheet_add_aoa(sheet, rows, { origin: { r: grid.length, c: 0 } });
      }

      XLSX.writeFile(workbook, `da-dien-${templateFile.name}`);
    } catch (caught) {
      setError(caught instanceof Error ? caught.message : "Không thể điền dữ liệu vào mẫu Excel này.");
    }
  }

  const maxMonthly = Math.max(1, ...analytics.months.flatMap(([, value]) => [value.credit, value.debit]));

  return (
    <main>
      <header className="site-header">
        <a className="brand" href="#top" aria-label="Mạch Tiền — về đầu trang">
          <span className="brand-mark">M</span>
          <span>
            <strong>Mạch Tiền</strong>
            <small>Bank statement studio</small>
          </span>
        </a>
        <div className="privacy-pill"><span /> Dữ liệu ở lại trên thiết bị</div>
      </header>

      <section className="hero" id="top">
        <div className="hero-copy">
          <div className="eyebrow">Sao kê vào · Bức tranh dòng tiền ra</div>
          <h1>Biến mọi sao kê thành <em>một bảng tiền rõ ràng.</em></h1>
          <p>
            Gom file từ nhiều ngân hàng, tự nhận diện cột, rà lại dòng tiền và xuất đúng cấu trúc kế toán bạn cần — ngay trong trình duyệt.
          </p>
          <div className="hero-badges">
            <span>Không tải dữ liệu lên máy chủ</span>
            <span>Excel · CSV · nhiều ngân hàng</span>
          </div>
        </div>
        <div className="hero-visual" aria-label="Minh họa luồng dữ liệu sao kê">
          <div className="floating-note note-one">VCB · +48,5 triệu</div>
          <div className="floating-note note-two">BIDV · −14,8 triệu</div>
          <div className="statement-card">
            <div className="statement-top"><span>Q2 / 2026</span><span>9 giao dịch</span></div>
            <div className="pulse-line"><i /><i /><i /><i /><i /><i /></div>
            <div className="statement-total">
              <span>Dòng tiền ròng</span>
              <strong>+54.200.000</strong>
            </div>
            <div className="mini-bars"><i /><i /><i /><i /><i /></div>
          </div>
        </div>
      </section>

      <nav className="stepper" aria-label="Quy trình xử lý">
        {[
          ["01", "Nạp sao kê", "Nhiều file, nhiều mẫu"],
          ["02", "Khớp cột", "Tự đoán, dễ chỉnh"],
          ["03", "Rà dòng tiền", "Tổng quan và phân loại"],
          ["04", "Xuất kết quả", "Bảng chuẩn hoặc mẫu riêng"],
        ].map(([number, label, detail], index) => (
          <div className={`step ${statements.length > 0 && index < 2 ? "active" : ""} ${transactions.length > 0 && index < 4 ? "active" : ""}`} key={number}>
            <span>{number}</span>
            <div><strong>{label}</strong><small>{detail}</small></div>
          </div>
        ))}
      </nav>

      <section className="workspace-section">
        <div className="section-heading">
          <div><span className="section-number">01</span><h2>Thả sao kê vào đây</h2></div>
          <p>Hệ thống đọc trang đầu tiên và tìm dòng tiêu đề trong 15 dòng đầu.</p>
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
          <strong>{isReading ? "Đang đọc sao kê…" : "Kéo file vào hoặc bấm để chọn"}</strong>
          <small>.xlsx · .xls · .csv · có thể chọn nhiều file</small>
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
            Dùng dữ liệu mẫu để xem trọn quy trình <span>→</span>
          </button>
          <p>2 ngân hàng · 9 giao dịch mẫu · không cần chuẩn bị file</p>
        </div>

        {error && <div className="error-message" role="alert">{error}</div>}
      </section>

      {statements.length > 0 && (
        <section className="mapping-section">
          <div className="section-heading">
            <div><span className="section-number">02</span><h2>Kiểm tra cách khớp cột</h2></div>
            <p>Chỉ cần sửa những cột hệ thống hiểu chưa đúng.</p>
          </div>

          <div className="file-stack">
            {statements.map((statement) => (
              <article className="file-card" key={statement.id}>
                <div className="file-card-header">
                  <div className="file-identity">
                    <span className="file-icon">XL</span>
                    <div><strong>{statement.fileName}</strong><small>{statement.rows.length} dòng dữ liệu</small></div>
                  </div>
                  <button className="remove-button" type="button" onClick={() => setStatements((current) => current.filter((item) => item.id !== statement.id))}>Bỏ file</button>
                </div>
                <div className="source-control">
                  <label htmlFor={`source-${statement.id}`}>Ngân hàng / nguồn</label>
                  <input
                    id={`source-${statement.id}`}
                    value={statement.source}
                    onChange={(event) => updateStatement(statement.id, (current) => ({ ...current, source: event.target.value }))}
                  />
                  <span>{statement.headers.filter((header) => statement.mapping[header] !== "ignore").length}/{statement.headers.length} cột đã khớp</span>
                </div>
                <div className="mapping-table-wrap">
                  <table className="mapping-table">
                    <thead><tr><th>Cột trong file</th><th>Giá trị mẫu</th><th>Khớp với</th></tr></thead>
                    <tbody>
                      {statement.headers.map((header) => (
                        <tr key={header}>
                          <td>{header}</td>
                          <td>{String(statement.rows[0]?.[header] ?? "—") || "—"}</td>
                          <td>
                            <select
                              aria-label={`Khớp cột ${header}`}
                              value={statement.mapping[header]}
                              onChange={(event) => updateStatement(statement.id, (current) => ({
                                ...current,
                                mapping: { ...current.mapping, [header]: event.target.value as MappingValue },
                              }))}
                            >
                              {fieldOptions.map((option) => <option value={option.key} key={option.key}>{option.label}</option>)}
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
            <button className="secondary-button" type="button" onClick={() => { setStatements([]); setTransactions([]); }}>Xóa và làm lại</button>
            <button className="primary-button" type="button" onClick={buildMasterTable}>Tạo bảng dòng tiền <span>→</span></button>
          </div>
        </section>
      )}

      {transactions.length > 0 && (
        <section className="insights-section" id="insights">
          <div className="section-heading light-heading">
            <div><span className="section-number">03</span><h2>Bức tranh dòng tiền</h2></div>
            <p>{transactions[0]?.date} — {transactions.at(-1)?.date}</p>
          </div>

          <div className="stat-grid">
            <div className="stat-card"><span>Số giao dịch</span><strong>{transactions.length}</strong><small>{statements.length} nguồn sao kê</small></div>
            <div className="stat-card positive"><span>Tổng thu</span><strong>{formatMoney(analytics.totalCredit)}</strong><small>Tiền vào tài khoản</small></div>
            <div className="stat-card negative"><span>Tổng chi</span><strong>{formatMoney(analytics.totalDebit)}</strong><small>Tiền ra khỏi tài khoản</small></div>
            <div className="stat-card net"><span>Dòng tiền ròng</span><strong>{analytics.net >= 0 ? "+" : "−"}{formatMoney(Math.abs(analytics.net))}</strong><small>{analytics.net >= 0 ? "Thặng dư trong kỳ" : "Thâm hụt trong kỳ"}</small></div>
          </div>

          <div className="insight-grid">
            <article className="chart-panel monthly-panel">
              <div className="panel-heading"><div><span>Xu hướng</span><h3>Thu / chi theo tháng</h3></div><div className="legend"><i className="credit-dot" /> Thu <i className="debit-dot" /> Chi</div></div>
              <div className="bar-chart">
                {analytics.months.map(([month, value]) => (
                  <div className="bar-group" key={month}>
                    <div className="bars">
                      <i className="credit-bar" style={{ height: `${Math.max(4, value.credit / maxMonthly * 100)}%` }} title={`Thu ${formatMoney(value.credit)}`} />
                      <i className="debit-bar" style={{ height: `${Math.max(4, value.debit / maxMonthly * 100)}%` }} title={`Chi ${formatMoney(value.debit)}`} />
                    </div>
                    <strong>{month.slice(5)}/{month.slice(2, 4)}</strong>
                    <small>{formatMoney(value.credit - value.debit)}</small>
                  </div>
                ))}
              </div>
            </article>

            <article className="chart-panel category-panel">
              <div className="panel-heading"><div><span>Phân loại gợi ý</span><h3>Tiền chi đi đâu?</h3></div></div>
              <div className="category-list">
                {analytics.categories.map(([category, value], index) => (
                  <div className="category-item" key={category}>
                    <div><span><i>{String(index + 1).padStart(2, "0")}</i>{category}</span><strong>{formatMoney(value)}</strong></div>
                    <span className="category-track"><i style={{ width: `${analytics.totalDebit ? value / analytics.totalDebit * 100 : 0}%` }} /></span>
                  </div>
                ))}
              </div>
            </article>
          </div>

          <div className="source-strip">
            {analytics.sources.map(([source, value]) => (
              <div key={source}><strong>{source}</strong><span><b>+{formatMoney(value.credit)}</b><em>−{formatMoney(value.debit)}</em></span></div>
            ))}
          </div>

          <article className="master-panel">
            <div className="panel-heading"><div><span>Bảng chuẩn</span><h3>{transactions.length} giao dịch đã hợp nhất</h3></div><small>Cuộn ngang để xem toàn bộ cột</small></div>
            <div className="master-table-wrap">
              <table className="master-table">
                <thead><tr><th>Ngày</th><th>Ngân hàng</th><th>Diễn giải</th><th>Đối tác</th><th>Ghi nợ</th><th>Ghi có</th><th>Số dư</th><th>Nhãn</th></tr></thead>
                <tbody>
                  {transactions.map((item, index) => (
                    <tr key={`${item.source}-${item.date}-${item.reference}-${index}`}>
                      <td>{item.date}</td><td>{item.source}</td><td>{item.description || "—"}</td><td>{item.counterparty || "—"}</td>
                      <td className="money debit-text">{item.debit ? formatMoney(item.debit) : "—"}</td><td className="money credit-text">{item.credit ? formatMoney(item.credit) : "—"}</td>
                      <td className="money">{formatMoney(item.balance)}</td><td><span className={`category-tag ${item.kind === "Thu" ? "tag-credit" : ""}`}>{item.category}</span></td>
                    </tr>
                  ))}
                </tbody>
              </table>
            </div>
          </article>

          <div className="export-heading">
            <span className="section-number">04</span>
            <div><h2>Đưa dữ liệu vào công việc</h2><p>Chọn cấu trúc phù hợp với quy trình kế toán hiện tại.</p></div>
          </div>
          <div className="export-grid">
            <article className="export-card featured-export">
              <span className="export-label">Khuyên dùng</span>
              <div className="export-icon">↳</div>
              <h3>Điền vào mẫu của bạn</h3>
              <p>Nạp file Excel đang dùng. Nếu nhận ra tiêu đề, hệ thống điền tiếp dữ liệu; nếu không, một trang “Dữ liệu chuẩn” sẽ được thêm vào.</p>
              <label className="template-picker">
                <input type="file" accept=".xlsx,.xls" onChange={(event) => setTemplateFile(event.target.files?.[0] ?? null)} />
                <span>{templateFile ? templateFile.name : "Chọn mẫu Excel"}</span><b>Chọn file</b>
              </label>
              <button className="primary-button full-button" type="button" disabled={!templateFile} onClick={() => void exportToTemplate()}>Điền và tải xuống <span>↓</span></button>
            </article>
            <article className="export-card">
              <div className="export-icon">▦</div><h3>Bảng chuẩn đầy đủ</h3>
              <p>Tất cả trường đã chuẩn hóa, phù hợp để kiểm tra, lọc, Pivot hoặc làm dữ liệu nguồn.</p>
              <button className="outline-button" type="button" onClick={() => exportWorkbook(false)}>Tải bảng chuẩn <span>↓</span></button>
            </article>
            <article className="export-card">
              <div className="export-icon">≡</div><h3>Mẫu hạch toán gọn</h3>
              <p>Ngày, diễn giải, thu, chi, số dư và nhãn — đủ để tiếp tục xử lý trong phần mềm kế toán.</p>
              <button className="outline-button" type="button" onClick={() => exportWorkbook(true)}>Tải mẫu gọn <span>↓</span></button>
            </article>
          </div>
        </section>
      )}

      <footer>
        <a className="brand footer-brand" href="#top"><span className="brand-mark">M</span><span><strong>Mạch Tiền</strong><small>Bank statement studio</small></span></a>
        <p>Dữ liệu sao kê được xử lý cục bộ trong trình duyệt và không được gửi đi.</p>
        <span>Thiết kế cho đội ngũ tài chính Việt Nam.</span>
      </footer>
    </main>
  );
}
