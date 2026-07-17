import * as XLSX from "xlsx";
import type { CategoryKey, MessageSet } from "./i18n";

export type CanonicalKey =
  | "date"
  | "reference"
  | "description"
  | "debit"
  | "credit"
  | "amount"
  | "balance"
  | "counterparty";

export type MappingValue = CanonicalKey | "ignore";
export type DateFormat = "auto" | "dmy" | "mdy" | "ymd";
export type DecimalSeparator = "auto" | "comma" | "dot";
export type CurrencyCode = "VND" | "USD" | "JPY" | "KRW" | "CNY";
export type IssueCode = "missingDate" | "invalidDate" | "noAmount" | "bothSides" | "negativeAmount" | "balanceMismatch";

export interface ImportConfig {
  dateFormat: DateFormat;
  decimalSeparator: DecimalSeparator;
  currency: CurrencyCode;
}

export interface StatementFile {
  id: string;
  fingerprint: string;
  fileName: string;
  source: string;
  headers: string[];
  rows: Record<string, unknown>[];
  mapping: Record<string, MappingValue>;
}

export interface Transaction {
  id: string;
  source: string;
  date: string;
  reference: string;
  description: string;
  counterparty: string;
  debit: number;
  credit: number;
  balance: number;
  balanceProvided: boolean;
  currency: CurrencyCode;
  kind: "Thu" | "Chi";
  category: CategoryKey;
}

export interface MappingProfile {
  source: string;
  mapping: Record<string, MappingValue>;
}

export type MappingProfiles = Record<string, MappingProfile>;

export interface ValidationSummary {
  issuesById: Map<string, IssueCode[]>;
  issueCounts: Map<IssueCode, number>;
  validCount: number;
  reconciliationChecked: number;
  reconciliationMismatches: number;
  reconciliationDifference: number;
}

export const defaultImportConfig: ImportConfig = {
  dateFormat: "auto",
  decimalSeparator: "auto",
  currency: "VND",
};

export const fieldOptions: MappingValue[] = ["ignore", "date", "reference", "description", "debit", "credit", "amount", "balance", "counterparty"];
export const currencyOptions: CurrencyCode[] = ["VND", "USD", "JPY", "KRW", "CNY"];

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

export function normalizeText(value: string) {
  return value
    .normalize("NFD")
    .replace(/[\u0300-\u036f]/g, "")
    .toLowerCase()
    .replace(/đ/g, "d")
    .replace(/[^\p{L}\p{N}]+/gu, " ")
    .trim();
}

export function inferField(header: string): MappingValue {
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

export function headerSignature(headers: string[]) {
  return headers.map(normalizeText).join("|");
}

export function createStatement(fileName: string, source: string, rows: Record<string, unknown>[]): StatementFile {
  const headers = rows.length > 0 ? Object.keys(rows[0]) : [];
  return {
    id: `demo-${fileName}`,
    fingerprint: `demo-${fileName}`,
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

export async function parseStatementFile(file: File, copy: MessageSet, profiles: MappingProfiles): Promise<StatementFile> {
  const data = await file.arrayBuffer();
  const workbook = XLSX.read(data, { cellDates: true });
  const sheetName = workbook.SheetNames[0];

  if (!sheetName) {
    throw new Error(`${file.name} ${copy.errors.emptySheet}`);
  }

  const sheet = workbook.Sheets[sheetName];
  const grid = XLSX.utils.sheet_to_json<unknown[]>(sheet, { header: 1, defval: "", raw: true });
  const headerIndex = grid
    .slice(0, 30)
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

  const profile = profiles[headerSignature(headers)];
  return {
    id: `${file.name}-${file.lastModified}-${Math.random().toString(36).slice(2)}`,
    fingerprint: `${file.name}:${file.size}:${file.lastModified}`,
    fileName: file.name,
    source: profile?.source || inferBankName(file.name, copy.unknownBank),
    headers,
    rows,
    mapping: Object.fromEntries(headers.map((header) => [header, profile?.mapping[header] ?? inferField(header)])),
  };
}

function detectDecimalSeparator(value: string, preference: DecimalSeparator): "." | "," | null {
  const requested = preference === "dot" ? "." : preference === "comma" ? "," : null;
  const lastDot = value.lastIndexOf(".");
  const lastComma = value.lastIndexOf(",");

  if (requested) {
    const decimalLength = value.length - value.lastIndexOf(requested) - 1;
    if (value.includes(requested === "." ? "," : ".")) return requested;
    return decimalLength > 0 && decimalLength <= 2 ? requested : null;
  }

  if (lastDot >= 0 && lastComma >= 0) return lastDot > lastComma ? "." : ",";
  const separator = lastDot >= 0 ? "." : lastComma >= 0 ? "," : null;
  if (!separator) return null;

  const parts = value.split(separator);
  return parts.length === 2 && parts[1].length > 0 && parts[1].length <= 2 ? separator : null;
}

export function parseAmount(value: unknown, config: ImportConfig) {
  if (typeof value === "number") return Number.isFinite(value) ? value : 0;
  if (value === null || value === undefined || value === "") return 0;

  const text = String(value).trim();
  const negative = /^\(.*\)$/.test(text) || /^[−-]/.test(text);
  const cleaned = text.replace(/[^0-9.,]/g, "");
  const decimalSeparator = detectDecimalSeparator(cleaned, config.decimalSeparator);

  if (!cleaned) return 0;
  if (!decimalSeparator) {
    const amount = Number(cleaned.replace(/[^0-9]/g, ""));
    return negative ? -amount : amount;
  }

  const decimalIndex = cleaned.lastIndexOf(decimalSeparator);
  const integer = cleaned.slice(0, decimalIndex).replace(/[^0-9]/g, "") || "0";
  const fraction = cleaned.slice(decimalIndex + 1).replace(/[^0-9]/g, "");
  const amount = Number(`${integer}.${fraction}`);
  return negative ? -amount : amount;
}

function validDate(year: number, month: number, day: number) {
  const date = new Date(Date.UTC(year, month - 1, day));
  return date.getUTCFullYear() === year && date.getUTCMonth() === month - 1 && date.getUTCDate() === day;
}

function toIsoDate(year: number, month: number, day: number) {
  if (!validDate(year, month, day)) return "";
  return `${year}-${String(month).padStart(2, "0")}-${String(day).padStart(2, "0")}`;
}

export function parseDate(value: unknown, config: ImportConfig) {
  if (value instanceof Date && !Number.isNaN(value.getTime())) {
    return value.toISOString().slice(0, 10);
  }

  if (typeof value === "number") {
    const parsed = XLSX.SSF.parse_date_code(value);
    if (parsed) return toIsoDate(parsed.y, parsed.m, parsed.d);
  }

  const text = String(value ?? "").trim();
  const yearFirst = text.match(/^(\d{4})[/.-](\d{1,2})[/.-](\d{1,2})$/);
  if (yearFirst) return toIsoDate(Number(yearFirst[1]), Number(yearFirst[2]), Number(yearFirst[3])) || text;

  const shortDate = text.match(/^(\d{1,2})[/.-](\d{1,2})[/.-](\d{4})$/);
  if (!shortDate || config.dateFormat === "ymd") return text;

  const first = Number(shortDate[1]);
  const second = Number(shortDate[2]);
  const year = Number(shortDate[3]);
  const format = config.dateFormat === "auto" ? (first > 12 ? "dmy" : second > 12 ? "mdy" : "dmy") : config.dateFormat;
  const month = format === "mdy" ? first : second;
  const day = format === "mdy" ? second : first;
  return toIsoDate(year, month, day) || text;
}

function classifyTransaction(description: string, kind: Transaction["kind"]): CategoryKey {
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

function hasValue(value: unknown) {
  return value !== null && value !== undefined && String(value).trim() !== "";
}

function duplicateKey(transaction: Transaction) {
  const identity = transaction.reference || normalizeText(transaction.description);
  return [transaction.source, transaction.currency, transaction.date, identity, transaction.debit, transaction.credit, transaction.balanceProvided ? transaction.balance : ""].join("|");
}

export function normalizeStatements(statements: StatementFile[], config: ImportConfig) {
  const seen = new Set<string>();
  const transactions: Transaction[] = [];
  let duplicateCount = 0;

  statements.forEach((statement) => {
    statement.rows.forEach((row, rowIndex) => {
      const debitRaw = getMappedValue(statement, row, "debit");
      const creditRaw = getMappedValue(statement, row, "credit");
      const amountRaw = getMappedValue(statement, row, "amount");
      const balanceRaw = getMappedValue(statement, row, "balance");
      let debit = Math.abs(parseAmount(debitRaw, config));
      let credit = Math.abs(parseAmount(creditRaw, config));
      const signedAmount = parseAmount(amountRaw, config);

      if (debit === 0 && credit === 0 && signedAmount !== 0) {
        if (signedAmount < 0) debit = Math.abs(signedAmount);
        else credit = signedAmount;
      }

      const kind: Transaction["kind"] = credit >= debit ? "Thu" : "Chi";
      const description = String(getMappedValue(statement, row, "description") ?? "").trim();
      const transaction: Transaction = {
        id: `${statement.id}:${rowIndex}`,
        source: statement.source,
        date: parseDate(getMappedValue(statement, row, "date"), config),
        reference: String(getMappedValue(statement, row, "reference") ?? "").trim(),
        description,
        counterparty: String(getMappedValue(statement, row, "counterparty") ?? "").trim(),
        debit,
        credit,
        balance: parseAmount(balanceRaw, config),
        balanceProvided: hasValue(balanceRaw),
        currency: config.currency,
        kind,
        category: classifyTransaction(description, kind),
      };

      if (!transaction.date && !transaction.description && !transaction.debit && !transaction.credit) return;
      const key = duplicateKey(transaction);
      if (seen.has(key)) {
        duplicateCount += 1;
        return;
      }
      seen.add(key);
      transactions.push(transaction);
    });
  });

  transactions.sort((a, b) => a.date.localeCompare(b.date));
  return { transactions, duplicateCount };
}

function isValidIsoDate(value: string) {
  const match = value.match(/^(\d{4})-(\d{2})-(\d{2})$/);
  return Boolean(match && validDate(Number(match[1]), Number(match[2]), Number(match[3])));
}

export function validateTransactions(transactions: Transaction[]): ValidationSummary {
  const issuesById = new Map<string, IssueCode[]>();
  const issueCounts = new Map<IssueCode, number>();
  const addIssue = (id: string, issue: IssueCode) => {
    const current = issuesById.get(id) ?? [];
    if (current.includes(issue)) return;
    issuesById.set(id, [...current, issue]);
    issueCounts.set(issue, (issueCounts.get(issue) ?? 0) + 1);
  };

  transactions.forEach((transaction) => {
    if (!transaction.date) addIssue(transaction.id, "missingDate");
    else if (!isValidIsoDate(transaction.date)) addIssue(transaction.id, "invalidDate");
    if (transaction.debit === 0 && transaction.credit === 0) addIssue(transaction.id, "noAmount");
    if (transaction.debit > 0 && transaction.credit > 0) addIssue(transaction.id, "bothSides");
    if (transaction.debit < 0 || transaction.credit < 0) addIssue(transaction.id, "negativeAmount");
  });

  const groups = new Map<string, Transaction[]>();
  transactions.forEach((transaction) => {
    const key = `${transaction.source}|${transaction.currency}`;
    groups.set(key, [...(groups.get(key) ?? []), transaction]);
  });

  let reconciliationChecked = 0;
  let reconciliationMismatches = 0;
  let reconciliationDifference = 0;
  groups.forEach((group) => {
    let runningBalance: number | undefined;
    group.forEach((transaction) => {
      if (runningBalance === undefined) {
        if (transaction.balanceProvided) runningBalance = transaction.balance;
        return;
      }

      const expected = runningBalance + transaction.credit - transaction.debit;
      if (!transaction.balanceProvided) {
        runningBalance = expected;
        return;
      }

      reconciliationChecked += 1;
      const difference = transaction.balance - expected;
      if (Math.abs(difference) > 0.01) {
        addIssue(transaction.id, "balanceMismatch");
        reconciliationMismatches += 1;
        reconciliationDifference += Math.abs(difference);
      }
      runningBalance = transaction.balance;
    });
  });

  return {
    issuesById,
    issueCounts,
    validCount: transactions.length - issuesById.size,
    reconciliationChecked,
    reconciliationMismatches,
    reconciliationDifference,
  };
}

export function buildMappingProfiles(statements: StatementFile[], current: MappingProfiles): MappingProfiles {
  const next = { ...current };
  statements.forEach((statement) => {
    next[headerSignature(statement.headers)] = {
      source: statement.source,
      mapping: { ...statement.mapping },
    };
  });
  return next;
}

export function parseMappingProfiles(value: string): MappingProfiles {
  const parsed: unknown = JSON.parse(value);
  if (!parsed || typeof parsed !== "object" || Array.isArray(parsed)) throw new Error("Invalid mapping profiles");

  const profiles: MappingProfiles = {};
  Object.entries(parsed).forEach(([signature, profile]) => {
    if (!profile || typeof profile !== "object" || Array.isArray(profile)) return;
    const source = "source" in profile && typeof profile.source === "string" ? profile.source : "";
    const rawMapping = "mapping" in profile && profile.mapping && typeof profile.mapping === "object" && !Array.isArray(profile.mapping) ? profile.mapping : {};
    const mapping: Record<string, MappingValue> = {};
    Object.entries(rawMapping).forEach(([header, field]) => {
      if (typeof field === "string" && fieldOptions.includes(field as MappingValue)) mapping[header] = field as MappingValue;
    });
    profiles[signature] = { source, mapping };
  });
  return profiles;
}

export function transactionValue(transaction: Transaction, field: MappingValue) {
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
