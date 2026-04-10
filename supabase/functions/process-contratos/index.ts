import { serve } from "https://deno.land/std@0.168.0/http/server.ts";
import { createClient } from "https://esm.sh/@supabase/supabase-js@2";
import * as XLSX from "xlsx";

const corsHeaders = {
  "Access-Control-Allow-Origin": "*",
  "Access-Control-Allow-Headers":
    "authorization, x-client-info, apikey, content-type, x-supabase-client-platform, x-supabase-client-platform-version, x-supabase-client-runtime, x-supabase-client-runtime-version",
};

// ============ HELPERS ============

function removeAccents(str: string): string {
  return str.normalize("NFD").replace(/[\u0300-\u036f]/g, "");
}

function norm(str: string): string {
  return removeAccents(str).toLowerCase().trim().replace(/[_\-\.\/]/g, " ").replace(/\s+/g, " ");
}

function extractDigits(val: string): string {
  return val.replace(/\D/g, "");
}

function isValidCPF(d: string): boolean {
  return d.length === 11 && !/^(\d)\1{10}$/.test(d);
}

function isValidCNPJ(d: string): boolean {
  return d.length === 14 && !/^(\d)\1{13}$/.test(d);
}

function detectDocumento(val: unknown): string {
  if (val == null) return "";
  const str = String(val).trim();
  const digits = extractDigits(str);
  if (isValidCPF(digits) || isValidCNPJ(digits)) return digits;
  if (/^\d{11,14}$/.test(str)) return str;
  return "";
}

function detectTipoPessoa(doc: string): "Física" | "Jurídica" | "Não definido" {
  if (doc.length === 11) return "Física";
  if (doc.length === 14) return "Jurídica";
  return "Não definido";
}

function findColumn(headers: string[], synonyms: string[], excludeTerms?: string[]): string | null {
  for (const syn of synonyms) {
    const nSyn = norm(syn);
    for (const h of headers) {
      const nH = norm(h);
      if (excludeTerms && excludeTerms.some(t => nH.includes(norm(t)))) continue;
      if (nH === nSyn || nH.includes(nSyn) || nSyn.includes(nH)) return h;
    }
  }
  return null;
}

function findAllColumns(headers: string[], synonyms: string[]): string[] {
  const result: string[] = [];
  for (const h of headers) {
    const nH = norm(h);
    if (synonyms.some(s => { const ns = norm(s); return nH === ns || nH.includes(ns) || ns.includes(nH); })) {
      if (!result.includes(h)) result.push(h);
    }
  }
  return result;
}

function normalizeCurrency(val: unknown): number {
  if (val == null) return NaN;
  if (typeof val === "number") return val;
  let str = String(val).replace(/[R$\s]/g, "").trim();
  if (str.includes(",") && str.includes(".")) {
    str = str.replace(/\./g, "").replace(",", ".");
  } else if (str.includes(",")) {
    str = str.replace(",", ".");
  }
  return parseFloat(str);
}

function formatCurrency(n: number): string {
  return isNaN(n) ? "" : n.toFixed(2);
}

function normalizeDate(val: unknown): string {
  if (val == null) return "";
  if (typeof val === "number") {
    // Small integers (1–31) represent day-of-month only — add current month/year.
    // XLSX serial date codes are typically in the tens of thousands for modern dates.
    if (Number.isInteger(val) && val >= 1 && val <= 31) {
      const now = new Date();
      return `${String(val).padStart(2, "0")}/${String(now.getMonth() + 1).padStart(2, "0")}/${now.getFullYear()}`;
    }
    try {
      const d = XLSX.SSF.parse_date_code(val);
      if (d) return `${String(d.d).padStart(2, "0")}/${String(d.m).padStart(2, "0")}/${d.y}`;
    } catch { /* fall through */ }
  }
  if (val instanceof Date) {
    return `${String(val.getDate()).padStart(2, "0")}/${String(val.getMonth() + 1).padStart(2, "0")}/${val.getFullYear()}`;
  }
  const str = String(val).trim();
  if (/^\d{2}\/\d{2}\/\d{4}$/.test(str)) return str;
  const iso = str.match(/^(\d{4})-(\d{2})-(\d{2})/);
  if (iso) return `${iso[3]}/${iso[2]}/${iso[1]}`;
  if (/^\d{1,2}$/.test(str)) {
    const now = new Date();
    return `${str.padStart(2, "0")}/${String(now.getMonth() + 1).padStart(2, "0")}/${now.getFullYear()}`;
  }
  return str;
}

function parseDate(dateStr: string): Date | null {
  const m = dateStr.match(/^(\d{2})\/(\d{2})\/(\d{4})$/);
  if (!m) return null;
  return new Date(parseInt(m[3]), parseInt(m[2]) - 1, parseInt(m[1]));
}

function formatDateDDMMYYYY(d: Date): string {
  return `${String(d.getDate()).padStart(2, "0")}/${String(d.getMonth() + 1).padStart(2, "0")}/${d.getFullYear()}`;
}

function lastDayOfMonth(year: number, month: number): Date {
  return new Date(year, month + 1, 0);
}

function firstDayOfMonth(year: number, month: number): Date {
  return new Date(year, month, 1);
}

function setCellValue(sheet: XLSX.WorkSheet, col: number, row: number, value: string | number): void {
  const cellRef = XLSX.utils.encode_cell({ c: col, r: row });
  const t = typeof value === "number" ? "n" : "s";
  if (!sheet[cellRef]) {
    sheet[cellRef] = { t, v: value };
  } else {
    sheet[cellRef].v = value;
    sheet[cellRef].t = t;
  }
}

function updateSheetRange(sheet: XLSX.WorkSheet, maxRow: number, maxCol: number): void {
  const range = XLSX.utils.decode_range(sheet["!ref"] || "A1");
  if (maxRow - 1 > range.e.r) range.e.r = maxRow - 1;
  if (maxCol - 1 > range.e.c) range.e.c = maxCol - 1;
  sheet["!ref"] = XLSX.utils.encode_range(range);
}

function captureRowStyles(
  sheet: XLSX.WorkSheet,
  protoRow: number,
  maxCol: number
): Map<number, Record<string, unknown>> {
  const styles = new Map<number, Record<string, unknown>>();
  for (let c = 0; c < maxCol; c++) {
    const ref = XLSX.utils.encode_cell({ c, r: protoRow });
    const cell = sheet[ref];
    if (!cell) continue;
    const snap: Record<string, unknown> = {};
    if (cell.s !== undefined) snap.s = JSON.parse(JSON.stringify(cell.s));
    if (cell.z !== undefined) snap.z = cell.z;
    styles.set(c, snap);
  }
  return styles;
}

function copyRowStyle(
  sheet: XLSX.WorkSheet,
  toRow: number,
  styles: Map<number, Record<string, unknown>>
): void {
  if (styles.size === 0) return;
  for (const [c, snap] of styles) {
    const ref = XLSX.utils.encode_cell({ c, r: toRow });
    if (!sheet[ref]) sheet[ref] = { t: "z" };
    if (snap.s !== undefined) sheet[ref].s = JSON.parse(JSON.stringify(snap.s));
    if (snap.z !== undefined) sheet[ref].z = snap.z;
  }
}

// Parse vigência range "MM/YYYY - MM/YYYY" or single "MM/YYYY"
function parseVigenciaRange(val: unknown): { inicio: string; fim: string } {
  const empty = { inicio: "", fim: "" };
  if (val == null) return empty;
  const str = String(val).trim();
  if (!str || str === "-") return empty;

  // Range: "05/2026 - 02/2027"
  const rangeMatch = str.match(/^(\d{2})\/(\d{4})\s*[-–]\s*(\d{2})\/(\d{4})/);
  if (rangeMatch) {
    const inicio = `01/${rangeMatch[1]}/${rangeMatch[2]}`;
    const endYear = parseInt(rangeMatch[4]);
    const endMonthIdx = parseInt(rangeMatch[3]); // 1-indexed
    const lastDay = new Date(endYear, endMonthIdx, 0).getDate();
    const fim = `${String(lastDay).padStart(2, "0")}/${rangeMatch[3]}/${rangeMatch[4]}`;
    return { inicio, fim };
  }
  // Single "MM/YYYY"
  const singleMatch = str.match(/^(\d{2})\/(\d{4})$/);
  if (singleMatch) {
    return { inicio: `01/${singleMatch[1]}/${singleMatch[2]}`, fim: "" };
  }
  return empty;
}

function parseDateForSort(dateStr: string): number {
  if (!dateStr) return 0;
  const m = dateStr.match(/^(\d{2})\/(\d{2})\/(\d{4})$/);
  if (m) return parseInt(m[3]) * 10000 + parseInt(m[2]) * 100 + parseInt(m[1]);
  return 0;
}

// ============ CONTRACT DETECTION ============

const CONTRACT_TERMS = ["recorrente", "mensalidade", "assinatura", "plano", "contrato", "vigencia", "renovacao", "periodo", "anual", "semestral", "trimestral"];
const EXCLUDE_TERMS = [
  "venda pontual", "venda simples", "venda unica", "venda avulsa",
  "parcelamento pontual", "cobranca avulsa", "cobranca pontual",
  "avulso", "pontual", "unica", "unico",
  "parcelamento", "parcela", "parcelado",
];

function isContract(description: string): boolean {
  const n = norm(description);
  // Skip rows that explicitly describe a one-time or punctual sale
  if (EXCLUDE_TERMS.some(t => n.includes(t))) return false;
  // Default: assume contract — the user already chose the "contratos" processor,
  // so every row is a contract unless it explicitly looks like a one-time sale.
  return true;
}

// ============ PAYMENT TYPE MAPPING ============

function mapTipoCobranca(val: string): string {
  if (!val || !val.trim()) return "Boleto";
  const n = norm(val);
  if (!n) return "Boleto";
  if (n.includes("debito em conta") || n.includes("debito conta")) return "Débito em Conta";
  if (n.includes("cartao de credito") || n.includes("credito") || n === "cc" || n === "cartao" || n === "cartão") return "Cartão de Crédito";
  if (n.includes("cartao de debito") || n.includes("debito") || n === "cd") return "Cartão de Débito";
  if (n.includes("pix")) return "Pix";
  if (n.includes("boleto")) return "Boleto";
  if (n.includes("dinheiro") || n.includes("cash")) return "Dinheiro";
  if (n.includes("ted")) return "TED";
  if (n.includes("doc")) return "DOC";
  if (n.includes("cheque")) return "Cheque";
  if (n.includes("cambio")) return "Câmbio";
  if (n === "outros") return "Outros";
  return "Boleto";
}

// ============ PERIODICITY ============

function mapPeriodicidade(val: string): string {
  if (!val) return "Mensal";
  const n = norm(val);
  if (n.includes("semanal")) return "Semanal";
  if (n.includes("quinzenal")) return "Quinzenal";
  if (n.includes("bimestral")) return "Bimestral";
  if (n.includes("trimestral")) return "Trimestral";
  if (n.includes("semestral")) return "Semestral";
  if (n.includes("anual")) return "Anual";
  if (n.includes("mensal")) return "Mensal";
  return "Mensal";
}

function countVendas(dataInicial: Date, dataFinal: Date, periodicidade: string): number {
  const diffMs = dataFinal.getTime() - dataInicial.getTime();
  if (diffMs <= 0) return 1;
  const diffDays = diffMs / (1000 * 60 * 60 * 24);
  switch (periodicidade) {
    case "Semanal": return Math.max(1, Math.ceil(diffDays / 7));
    case "Quinzenal": return Math.max(1, Math.ceil(diffDays / 15));
    case "Mensal": {
      const months = (dataFinal.getFullYear() - dataInicial.getFullYear()) * 12 + (dataFinal.getMonth() - dataInicial.getMonth());
      return Math.max(1, months + 1);
    }
    case "Bimestral": {
      const months = (dataFinal.getFullYear() - dataInicial.getFullYear()) * 12 + (dataFinal.getMonth() - dataInicial.getMonth());
      return Math.max(1, Math.ceil((months + 1) / 2));
    }
    case "Trimestral": {
      const months = (dataFinal.getFullYear() - dataInicial.getFullYear()) * 12 + (dataFinal.getMonth() - dataInicial.getMonth());
      return Math.max(1, Math.ceil((months + 1) / 3));
    }
    case "Semestral": {
      const months = (dataFinal.getFullYear() - dataInicial.getFullYear()) * 12 + (dataFinal.getMonth() - dataInicial.getMonth());
      return Math.max(1, Math.ceil((months + 1) / 6));
    }
    case "Anual": {
      const years = dataFinal.getFullYear() - dataInicial.getFullYear();
      return Math.max(1, years + 1);
    }
    default: {
      const months = (dataFinal.getFullYear() - dataInicial.getFullYear()) * 12 + (dataFinal.getMonth() - dataInicial.getMonth());
      return Math.max(1, months + 1);
    }
  }
}

// ============ MULTIPLE SERVICES DETECTION ============

interface ParsedService {
  description: string;
  value: number;
}

function parseMultipleServices(desc: string): ParsedService[] {
  const currencyPattern = /R?\$?\s*([\d]{1,3}(?:\.[\d]{3})*,[\d]{2}|[\d]+[.,][\d]{2})/g;
  const matches: { index: number; fullMatch: string; value: number }[] = [];
  let m: RegExpExecArray | null;
  while ((m = currencyPattern.exec(desc)) !== null) {
    const value = normalizeCurrency(m[1]);
    if (!isNaN(value) && value > 0) {
      matches.push({ index: m.index, fullMatch: m[0], value });
    }
  }
  if (matches.length < 2) return [];
  const services: ParsedService[] = [];
  for (let i = 0; i < matches.length; i++) {
    const start = i === 0 ? 0 : matches[i - 1].index + matches[i - 1].fullMatch.length;
    const textBefore = desc.substring(start, matches[i].index).trim().replace(/^[;|\/,\s]+/, "").replace(/[;|\/,\s]+$/, "");
    const name = textBefore || `Serviço ${i + 1}`;
    services.push({ description: name, value: matches[i].value });
  }
  return services;
}

// ============ RAW CONTRACT TYPE ============

// Validate that a value looks like a real identifier (not generic text like "Personalizado")
function isValidIdentifier(val: string): boolean {
  if (!val) return false;
  // Pure number or alphanumeric code -> valid
  if (/^\d+$/.test(val)) return true;
  if (/^[A-Za-z0-9\-_\/\.]{2,30}$/.test(val)) return true;
  // If it has spaces or is longer text, likely a description, not an ID
  return false;
}

interface RawContrato {
  rowIdx: number;
  documento: string;
  tipoPessoa: string;
  nome: string;
  dataInicio: string;
  dataFim: string;
  vencimento: string;
  valor: number;
  centroCusto: string;
  conta: string;
  tipoCobranca: string;
  observacoes: string;
  descricaoServico: string;
  planoServico: string;
  periodicidade: string;
  identificador: string;
  cep: string;
  numeroEndereco: string;
  email: string;
}

interface ConsolidatedContrato {
  documento: string;
  tipoPessoa: string;
  nome: string;
  dataCompetencia: string;
  dataInicial: string;
  dataFinal: string;
  primeiroVencimento: string;
  centroCusto: string;
  conta: string;
  tipoCobranca: string;
  observacoes: string;
  descricaoServico: string;
  codigoReferencia: string;
  valorUnitario: number;
  qtdVendas: number;
  periodicidade: string;
  primeiraVenda: string;
  multiServicos: boolean;
  parsedServices: ParsedService[];
  cep: string;
  numeroEndereco: string;
  email: string;
  warnings: string[];
}

// ============ MAIN HANDLER ============

serve(async (req) => {
  if (req.method === "OPTIONS") {
    return new Response(null, { headers: corsHeaders });
  }

  const supabase = createClient(
    Deno.env.get("SUPABASE_URL")!,
    Deno.env.get("SUPABASE_SERVICE_ROLE_KEY")!
  );

  try {
    const { upload_id } = await req.json();
    if (!upload_id) throw new Error("upload_id is required");

    console.log(`[process-contratos] Starting for upload: ${upload_id}`);

    // 1) Create run
    const { data: run, error: runError } = await supabase.from("runs").insert({
      upload_id,
      import_type: "contratos",
      status: "processing",
    }).select().single();
    if (runError || !run) throw new Error(`Failed to create run: ${runError?.message}`);
    const runId = run.id;
    console.log(`[process-contratos] Run created: ${runId}`);

    // 2) Fetch upload & download primary file
    const { data: upload, error: uploadErr } = await supabase.from("uploads").select("*").eq("id", upload_id).single();
    if (uploadErr || !upload) throw new Error(`Upload not found: ${uploadErr?.message}`);

    const { data: primaryFile, error: dlErr } = await supabase.storage.from("uploads").download(upload.storage_path);
    if (dlErr || !primaryFile) throw new Error(`Download primary failed: ${dlErr?.message}`);
    console.log(`[process-contratos] Primary file: ${upload.original_filename}`);

    // 3) Fetch template
    const { data: tmpl, error: tmplErr } = await supabase
      .from("import_templates")
      .select("*")
      .eq("import_type", "contratos")
      .eq("is_default", true)
      .single();
    if (tmplErr || !tmpl) throw new Error(`Template not found: ${tmplErr?.message}`);

    const { data: templateFile, error: tmplDlErr } = await supabase.storage.from("templates").download(tmpl.template_storage_path);
    if (tmplDlErr || !templateFile) throw new Error(`Download template failed: ${tmplDlErr?.message}`);
    console.log(`[process-contratos] Template downloaded: ${tmpl.template_storage_path}`);

    // 4) Read primary Excel
    const primaryBuffer = await primaryFile.arrayBuffer();
    const primaryWb = XLSX.read(new Uint8Array(primaryBuffer), { type: "array" });
    const primarySheet = primaryWb.Sheets[primaryWb.SheetNames[0]];
    const rawData: Record<string, unknown>[] = XLSX.utils.sheet_to_json(primarySheet);
    if (rawData.length === 0) throw new Error("Planilha primária vazia");
    console.log(`[process-contratos] Primary data: ${rawData.length} rows`);

    const headers = Object.keys(rawData[0]);
    console.log(`[process-contratos] Headers: ${headers.join(", ")}`);

    // 5) Map columns
    const colDoc = findColumn(headers, ["cpf", "cnpj", "cpf/cnpj", "cpf_cnpj", "documento", "doc"]);
    const colNome = findColumn(headers, ["cliente", "nome", "tomador", "sacado", "razao", "razao social", "razão social"]);
    const colDataInicio = findColumn(headers, ["inicio", "início", "start", "vigencia inicial", "vigência inicial", "data inicial", "data inicio", "data início"]);
    const colDataFim = findColumn(headers, ["fim", "final", "end", "vigencia final", "vigência final", "data final", "termino", "término"]);
    const colVencimento = findColumn(headers, ["venc", "vencimento", "due date", "dt vencimento", "data vencimento"]);
    const colValor = findColumn(headers, ["valor bruto", "valor", "total", "amount", "vlr", "valor total"], ["liquido", "líquido", "net", "desconto"]);
    const colCentroCusto = findColumn(headers, ["centro", "custo", "cc", "centro de custo", "centro custo"]);
    const colConta = findColumn(headers, ["conta", "carteira", "banco", "caixa"]);
    const colTipoCobranca = findColumn(headers, [
      "forma de pagamento", "tipo de pagamento", "tipo pagamento", "forma pagamento",
      "tipo de cobranca", "tipo de cobrança", "tipo cobranca", "tipo cobrança",
      "cobranca", "cobrança", "pagamento", "meio de pagamento", "meio pagamento",
      "meio", "forma", "modalidade"
    ]);
    const colObs = findColumn(headers, ["obs", "observacao", "observação", "comentario", "comentário", "historico", "histórico"]);
    const colDescricao = findColumn(headers, ["descricao", "descrição", "servico", "serviço", "produto", "historico", "histórico"]);
    const colId = findColumn(headers, ["identificador", "referencia", "referência", "ref", "titulo", "título",
      "nosso numero", "nosso número", "pedido", "fatura", "invoice", "numero contrato", "número contrato", "cod contrato", "código contrato"]);
    const colPlano = findColumn(headers, ["plano", "produto/servico", "produto/serviço", "produto servico", "produto serviço", "produto", "servico", "serviço"]);
    const colCep = findColumn(headers, ["cep", "zip", "codigo postal"]);
    const colNumeroEndereco = findColumn(headers, ["numero endereco", "número endereço", "num", "nro"]);
    const colEmail = findColumn(headers, ["email", "e-mail", "mail"]);
    const colPeriodicidade = findColumn(headers, ["periodicidade", "frequencia", "frequência", "recorrencia", "recorrência", "ciclo"]);
    // Column that may hold a range like "05/2026 - 02/2027"
    const colVigencia = findColumn(headers, ["vigencia do contrato", "vigência do contrato", "vigencia", "vigência"]);

    // Description columns for contract detection
    // Include "Contrato/Venda" so rows marked "Venda Pontual" are correctly excluded
    const descCols = findAllColumns(headers, ["descricao", "descrição", "categoria", "tipo", "historico", "histórico", "natureza", "servico", "serviço", "contrato/venda", "contrato venda"]);

    console.log(`[process-contratos] Mapping: doc=${colDoc}, nome=${colNome}, inicio=${colDataInicio}, fim=${colDataFim}, venc=${colVencimento}, val=${colValor}, desc=${colDescricao}, plano=${colPlano}, id=${colId}, cobranca=${colTipoCobranca}`);

    // 6) Extract raw contratos
    const rawContratos: RawContrato[] = [];
    const errors: { row: number; field: string; message: string; type: string }[] = [];

    for (let i = 0; i < rawData.length; i++) {
      const row = rawData[i];
      const rowNum = i + 2;

      // Skip ghost/empty rows (formatting residue from deleted rows)
      const isGhostRow = headers.every(h => {
        const v = row[h];
        return v == null || String(v).trim() === "";
      });
      if (isGhostRow) continue;

      // Check if this row is a contract
      let isContractRow = true;
      if (descCols.length > 0) {
        const combinedDesc = descCols.map(c => String(row[c] ?? "")).join(" ");
        if (combinedDesc.trim()) {
          isContractRow = isContract(combinedDesc);
        }
      }
      if (!isContractRow) continue;

      // Extract documento
      let documento = "";
      if (colDoc) documento = detectDocumento(row[colDoc]);
      if (!documento) {
        for (const key of headers) {
          const val = String(row[key] ?? "");
          const digits = extractDigits(val);
          if (isValidCPF(digits) || isValidCNPJ(digits)) { documento = digits; break; }
        }
      }

      // Extract nome
      let nome = "";
      if (colNome) nome = String(row[colNome] ?? "").trim();
      if (!nome) {
        for (const key of headers) {
          if (key === colDoc) continue;
          const val = String(row[key] ?? "").trim();
          if (val.length > 3 && /^[A-Za-zÀ-ú\s\.\-]+$/.test(val)) { nome = val; break; }
        }
      }

      if (!documento && !nome) {
        errors.push({ row: rowNum, field: "CPF/CNPJ + Nome", message: `Linha ${rowNum}: Documento e nome ausentes, registro ignorado`, type: "obrigatorio" });
        continue;
      }
      if (!nome) {
        errors.push({ row: rowNum, field: "Nome", message: `Linha ${rowNum}: Nome não encontrado, registro ignorado`, type: "obrigatorio" });
        continue;
      }

      // Extract valor
      const valorRaw = colValor ? row[colValor] : undefined;
      const valor = normalizeCurrency(valorRaw);
      if (isNaN(valor) || valor <= 0) {
        errors.push({ row: rowNum, field: "Valor", message: `Linha ${rowNum}: Valor ausente ou inválido (${valorRaw}), registro ignorado`, type: "obrigatorio" });
        continue;
      }

      // Dates — try dedicated columns first, then fall back to vigência range
      let dataInicio = colDataInicio ? normalizeDate(row[colDataInicio]) : "";
      let dataFim = colDataFim ? normalizeDate(row[colDataFim]) : "";
      if ((!dataInicio || !dataFim) && colVigencia) {
        const vig = parseVigenciaRange(row[colVigencia]);
        if (!dataInicio && vig.inicio) dataInicio = vig.inicio;
        if (!dataFim && vig.fim) dataFim = vig.fim;
      }
      const vencimento = colVencimento ? normalizeDate(row[colVencimento]) : "";

      // Other fields
      const centroCusto = colCentroCusto ? String(row[colCentroCusto] ?? "").trim() : "";
      const conta = colConta ? String(row[colConta] ?? "").trim() : "";
      const tipoCobrancaRaw = colTipoCobranca ? String(row[colTipoCobranca] ?? "").trim() : "";
      const observacoes = colObs ? String(row[colObs] ?? "").trim() : "";
      const descricaoServico = colDescricao ? String(row[colDescricao] ?? "").trim() : "";
      const planoServico = colPlano ? String(row[colPlano] ?? "").trim() : "";
      const identificadorRaw = colId ? String(row[colId] ?? "").trim() : "";
      // Validate identifier: must look like an ID (numeric, alphanumeric code, etc.), not generic text
      const identificador = isValidIdentifier(identificadorRaw) ? identificadorRaw : "";
      const cep = colCep ? String(row[colCep] ?? "").trim() : "";
      const numeroEndereco = colNumeroEndereco ? String(row[colNumeroEndereco] ?? "").trim() : "";
      const email = colEmail ? String(row[colEmail] ?? "").trim() : "";

      const periodicidadeRaw = colPeriodicidade ? String(row[colPeriodicidade] ?? "").trim() : "";

      rawContratos.push({
        rowIdx: rowNum, documento, tipoPessoa: detectTipoPessoa(documento), nome,
        dataInicio, dataFim, vencimento, valor, centroCusto, conta, tipoCobranca: tipoCobrancaRaw,
        observacoes, descricaoServico, planoServico, periodicidade: periodicidadeRaw,
        identificador, cep, numeroEndereco, email,
      });
    }

    console.log(`[process-contratos] Raw contratos: ${rawContratos.length}, errors so far: ${errors.length}`);

    // 7) Consolidate — key = doc + plano + descrição + vencDay + valor
    // Use BOTH planoServico and descricaoServico in the key so that contracts
    // with the same client/value/day but different plan names stay separate.
    const groups = new Map<string, RawContrato[]>();
    for (const c of rawContratos) {
      const docKey = extractDigits(c.documento);
      const planoKey = norm(c.planoServico || "");
      const descKey = norm(c.descricaoServico || "");
      // Combine both fields; if they are the same column they'll just repeat (harmless)
      const serviceKey = [planoKey, descKey].filter(Boolean).join("|") || "sem_servico";
      const vencDay = c.vencimento ? c.vencimento.substring(0, 2) : "";
      const key = `${docKey}|${serviceKey}|${vencDay}|${c.valor.toFixed(2)}`;
      if (!groups.has(key)) groups.set(key, []);
      groups.get(key)!.push(c);
    }

    // Log dedup stats
    const totalDeduped = rawContratos.length - groups.size;
    console.log(`[process-contratos] Dedup: ${rawContratos.length} raw -> ${groups.size} unique keys (${totalDeduped} collapsed)`);
    // Log sample keys (first 5)
    let keyIdx = 0;
    for (const [key, group] of groups) {
      if (keyIdx >= 5) break;
      console.log(`[process-contratos] Key sample ${keyIdx + 1}: "${key}" (${group.length} rows, nome=${group[0].nome})`);
      keyIdx++;
    }

    const now = new Date();
    const consolidated: ConsolidatedContrato[] = [];
    let refCounter = 1;

    for (const [, group] of groups) {
      group.sort((a, b) => parseDateForSort(a.vencimento) - parseDateForSort(b.vencimento));
      const first = group[0];
      const warnings: string[] = [];

      // Periodicidade — read from the field stored in RawContrato
      const periodicidadeRaw = first.periodicidade || "";
      const periodicidade = mapPeriodicidade(periodicidadeRaw);
      if (!periodicidadeRaw) {
        warnings.push("periodicidade_default: Periodicidade não encontrada, usando Mensal");
      }

      // --- Date logic ---
      // data_inicio_base
      let dataInicioBase: Date | null = null;
      if (first.dataInicio) dataInicioBase = parseDate(first.dataInicio);
      if (!dataInicioBase) dataInicioBase = firstDayOfMonth(now.getFullYear(), now.getMonth());

      // Data Competência: 1st day of month of dataInicioBase; if before current month, use current month
      let compDate = firstDayOfMonth(dataInicioBase.getFullYear(), dataInicioBase.getMonth());
      const currentMonthFirst = firstDayOfMonth(now.getFullYear(), now.getMonth());
      if (compDate < currentMonthFirst) compDate = currentMonthFirst;
      const dataCompetencia = formatDateDDMMYYYY(compDate);

      // Data Inicial = Data Competência
      const dataInicial = dataCompetencia;
      const dataInicialDate = compDate;

      // Data Final
      let dataFinalDate: Date | null = null;
      if (first.dataFim) dataFinalDate = parseDate(first.dataFim);
      if (dataFinalDate) {
        // Use last day of that month
        dataFinalDate = lastDayOfMonth(dataFinalDate.getFullYear(), dataFinalDate.getMonth());
      } else {
        // Default: 31/12 of the second subsequent year (e.g. if now is Apr/2026 → 31/12/2028)
        warnings.push("final_default: Data final não encontrada, usando 31/12 do 2º ano posterior");
        const year = now.getFullYear() + 2;
        dataFinalDate = new Date(year, 11, 31);
      }
      const dataFinal = formatDateDDMMYYYY(dataFinalDate);

      // Primeiro Vencimento
      let primeiroVencimento = "";
      if (first.vencimento) {
        primeiroVencimento = first.vencimento;
        // Validate: primeiro vencimento must not be before data inicial
        const pvParts = primeiroVencimento.match(/^(\d{2})\/(\d{2})\/(\d{4})$/);
        if (pvParts) {
          const pvDate = new Date(parseInt(pvParts[3]), parseInt(pvParts[2]) - 1, parseInt(pvParts[1]));
          if (pvDate.getFullYear() < dataInicialDate.getFullYear() ||
              (pvDate.getFullYear() === dataInicialDate.getFullYear() && pvDate.getMonth() < dataInicialDate.getMonth())) {
            const adjustedDay = pvDate.getDate();
            primeiroVencimento = `${String(adjustedDay).padStart(2, "0")}/${String(dataInicialDate.getMonth() + 1).padStart(2, "0")}/${dataInicialDate.getFullYear()}`;
            warnings.push(`primeiro_vencimento_ajustado: Data de vencimento era anterior à data inicial, ajustado para ${primeiroVencimento}`);
          }
        }
      } else {
        // Use day from dataInicioBase, month/year from dataInicial
        const day = dataInicioBase.getDate();
        primeiroVencimento = `${String(day).padStart(2, "0")}/${String(compDate.getMonth() + 1).padStart(2, "0")}/${compDate.getFullYear()}`;
      }

      // Qtd Vendas
      const qtdVendas = countVendas(dataInicialDate, dataFinalDate, periodicidade);

      // Payment type
      const tipoCobranca = mapTipoCobranca(first.tipoCobranca);
      if (first.tipoCobranca && tipoCobranca === "Outros") {
        warnings.push(`tipo_cobranca_desconhecido: "${first.tipoCobranca}" mapeado para "Outros"`);
      }

      // Reference code: use valid identifier from input, else generate sequential number
      const codigoReferencia = first.identificador || String(refCounter);
      refCounter++;

      // Multiple services
      const services = parseMultipleServices(first.descricaoServico);
      const multiServicos = services.length >= 2;

      if (multiServicos) {
        const serviceTotal = services.reduce((s, sv) => s + sv.value, 0);
        if (serviceTotal > 0 && Math.abs(serviceTotal - first.valor) > 0.01) {
          // Proportionally adjust service values to match the total
          const ratio = first.valor / serviceTotal;
          let runningSum = 0;
          for (let si = 0; si < services.length; si++) {
            if (si === services.length - 1) {
              services[si].value = Math.round((first.valor - runningSum) * 100) / 100;
            } else {
              services[si].value = Math.round(services[si].value * ratio * 100) / 100;
              runningSum += services[si].value;
            }
          }
          warnings.push(`multiplos_servicos_ajustado: soma dos serviços (${serviceTotal.toFixed(2)}) ajustada proporcionalmente para bater com valor total (${first.valor.toFixed(2)})`);
        }
      }

      consolidated.push({
        documento: first.documento,
        tipoPessoa: first.tipoPessoa,
        nome: first.nome,
        dataCompetencia,
        dataInicial,
        dataFinal,
        primeiroVencimento,
        centroCusto: first.centroCusto,
        conta: first.conta || "Caixinha",
        tipoCobranca,
        observacoes: first.observacoes,
        descricaoServico: first.descricaoServico,
        codigoReferencia,
        valorUnitario: first.valor,
        qtdVendas,
        periodicidade,
        primeiraVenda: primeiroVencimento,
        multiServicos,
        parsedServices: multiServicos ? services : [],
        cep: first.cep,
        numeroEndereco: first.numeroEndereco,
        email: first.email,
        warnings,
      });

      for (const w of warnings) {
        errors.push({ row: first.rowIdx, field: "aviso", message: w, type: "warning" });
      }
    }

    const ignoredRows = rawData.length - rawContratos.length - errors.filter(e => e.type === "obrigatorio").length;
    console.log(`[process-contratos] Summary: ${rawData.length} input rows, ${rawContratos.length} classified as contract, ${groups.size} unique after dedup (${totalDeduped} collapsed), ${consolidated.length} exported, ${errors.filter(e => e.type === "obrigatorio").length} errors, ${ignoredRows} ignored (not contract)`);
    console.log(`[process-contratos] Consolidated contratos: ${consolidated.length}`);

    // 8) Read template and fill
    const templateBuffer = await templateFile.arrayBuffer();
    const templateWb = XLSX.read(new Uint8Array(templateBuffer), { type: "array", cellStyles: true });

    // --- Fill Contratos sheet (index 1: "Contratos") ---
    const contratosSheet = templateWb.Sheets["Contratos"];
    if (!contratosSheet) throw new Error("Aba 'Contratos' não encontrada no template");

    const DATA_START_ROW = 3; // 0-indexed (headers at row 2, data starts row 3)

    // Column indices (0-indexed) from template inspection
    const COL = {
      TIPO_TOMADOR: 0,       // A
      CPF_CNPJ: 1,           // B
      RAZAO_SOCIAL: 2,       // C
      CEP: 3,                // D
      NUMERO_ENDERECO: 4,    // E
      EMAIL: 5,              // F
      ID_INTEGRACAO: 6,      // G
      DATA_COMPETENCIA: 7,   // H
      DATA_INICIAL: 8,       // I
      DATA_FINAL: 9,         // J
      DATA_ASSINATURA: 10,   // K
      PRIMEIRO_VENCIMENTO: 11, // L
      CATEGORIA: 12,         // M
      DESTINO_OPERACAO: 13,  // N
      CONTA: 14,             // O
      CENTRO_CUSTO: 15,      // P
      TIPO_COBRANCA: 16,     // Q
      CONTATO: 17,           // R
      CODIGO_REF: 18,        // S
      ORDEM_COMPRA: 19,      // T
      ORDEM_SERVICO: 20,     // U
      VENDEDOR: 21,          // V
      OBSERVACOES: 22,       // W
      INFO_NFSE: 23,         // X
      MULTIPLOS_SERVICOS: 24, // Y
      NOME_SERVICO: 25,      // Z
      DESCRICAO: 26,         // AA (index 26)
      QUANTIDADE: 27,        // AB (index 27)
      VALOR_UNITARIO: 28,    // AC (index 28)
      // AD(29)=Total Bruto, AE(30)=Acréscimos, AF(31)=Descontos, AG(32)=Total Líquido
      QTD_VENDAS: 33,        // AH (index 33)
      PERIODICIDADE: 34,     // AI (index 34)
      PRIMEIRA_VENDA: 35,    // AJ (index 35)
    };

    const contratosProtoStyles = captureRowStyles(contratosSheet, DATA_START_ROW, 36);

    for (let i = 0; i < consolidated.length; i++) {
      const c = consolidated[i];
      const row = DATA_START_ROW + i;

      copyRowStyle(contratosSheet, row, contratosProtoStyles);
      setCellValue(contratosSheet, COL.TIPO_TOMADOR, row, c.tipoPessoa);
      if (c.documento) setCellValue(contratosSheet, COL.CPF_CNPJ, row, c.documento);
      setCellValue(contratosSheet, COL.RAZAO_SOCIAL, row, c.nome);
      if (c.cep) setCellValue(contratosSheet, COL.CEP, row, c.cep);
      if (c.numeroEndereco) setCellValue(contratosSheet, COL.NUMERO_ENDERECO, row, c.numeroEndereco);
      if (c.email) setCellValue(contratosSheet, COL.EMAIL, row, c.email);
      setCellValue(contratosSheet, COL.DATA_COMPETENCIA, row, c.dataCompetencia);
      setCellValue(contratosSheet, COL.DATA_INICIAL, row, c.dataInicial);
      setCellValue(contratosSheet, COL.DATA_FINAL, row, c.dataFinal);
      setCellValue(contratosSheet, COL.PRIMEIRO_VENCIMENTO, row, c.primeiroVencimento);
      setCellValue(contratosSheet, COL.CATEGORIA, row, "01.01.01");
      setCellValue(contratosSheet, COL.CONTA, row, c.conta);
      if (c.centroCusto) setCellValue(contratosSheet, COL.CENTRO_CUSTO, row, c.centroCusto);
      setCellValue(contratosSheet, COL.TIPO_COBRANCA, row, c.tipoCobranca);
      setCellValue(contratosSheet, COL.CODIGO_REF, row, c.codigoReferencia);
      if (c.observacoes) setCellValue(contratosSheet, COL.OBSERVACOES, row, c.observacoes);

      if (c.multiServicos) {
        setCellValue(contratosSheet, COL.MULTIPLOS_SERVICOS, row, "Sim");
      } else {
        setCellValue(contratosSheet, COL.NOME_SERVICO, row, "Prestação de Serviço");
        if (c.descricaoServico) setCellValue(contratosSheet, COL.DESCRICAO, row, c.descricaoServico);
        setCellValue(contratosSheet, COL.QUANTIDADE, row, 1);
        setCellValue(contratosSheet, COL.VALOR_UNITARIO, row, c.valorUnitario);
      }

      setCellValue(contratosSheet, COL.QTD_VENDAS, row, c.qtdVendas);
      setCellValue(contratosSheet, COL.PERIODICIDADE, row, c.periodicidade);
      setCellValue(contratosSheet, COL.PRIMEIRA_VENDA, row, c.primeiraVenda);
    }

    // Set range to exactly the data written (no extra empty rows)
    const lastDataRow = DATA_START_ROW + consolidated.length - 1;
    const range = XLSX.utils.decode_range(contratosSheet["!ref"] || "A1");
    range.e.r = lastDataRow;
    contratosSheet["!ref"] = XLSX.utils.encode_range(range);

    // --- Fill Serviços sheet ---
    const servicosSheet = templateWb.Sheets["Serviços"];
    if (servicosSheet) {
      const SVC_COL = {
        CODIGO_REF: 0,
        NOME_SERVICO: 1,
        DESCRICAO: 2,
        QUANTIDADE: 3,
        VALOR_UNITARIO: 4,
      };
      let svcRow = 1; // Data starts at row 2 (0-indexed row 1)

      const svcProtoStyles = captureRowStyles(servicosSheet, 1, 9);

      for (const c of consolidated) {
        if (!c.multiServicos) continue;
        for (const svc of c.parsedServices) {
          copyRowStyle(servicosSheet, svcRow, svcProtoStyles);
          setCellValue(servicosSheet, SVC_COL.CODIGO_REF, svcRow, c.codigoReferencia);
          setCellValue(servicosSheet, SVC_COL.NOME_SERVICO, svcRow, "Prestação de Serviço");
          setCellValue(servicosSheet, SVC_COL.DESCRICAO, svcRow, svc.description);
          setCellValue(servicosSheet, SVC_COL.QUANTIDADE, svcRow, 1);
          setCellValue(servicosSheet, SVC_COL.VALOR_UNITARIO, svcRow, svc.value);
          svcRow++;
        }
      }

      if (svcRow > 1) {
        updateSheetRange(servicosSheet, svcRow, 9);
      }
    }

    // 9) Generate preview
    const previewRows: Record<string, string | number>[] = [];
    for (let i = 0; i < Math.min(consolidated.length, 20); i++) {
      const c = consolidated[i];
      previewRows.push({
        "Tipo Tomador": c.tipoPessoa,
        "CPF/CNPJ": c.documento,
        "Razão Social": c.nome,
        "Competência": c.dataCompetencia,
        "Data Inicial": c.dataInicial,
        "Data Final": c.dataFinal,
        "1º Venc.": c.primeiroVencimento,
        "Conta": c.conta,
        "Cobrança": c.tipoCobranca,
        "Cód. Ref.": c.codigoReferencia,
        "Multi Svc?": c.multiServicos ? "Sim" : "",
        "Valor": formatCurrency(c.valorUnitario),
        "Qtd Vendas": c.qtdVendas,
        "Periodicidade": c.periodicidade,
      });
    }

    // 10) Write output XLSX
    const outputBuffer = XLSX.write(templateWb, { type: "array", bookType: "xlsx", cellStyles: true });
    const outputPath = `contratos/${runId}.xlsx`;

    const { error: uploadOutputErr } = await supabase.storage
      .from("outputs")
      .upload(outputPath, new Uint8Array(outputBuffer), {
        contentType: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        upsert: true,
      });
    if (uploadOutputErr) throw new Error(`Upload output failed: ${uploadOutputErr.message}`);

    console.log(`[process-contratos] Output saved: ${outputPath}`);

    // 11) Update run
    await supabase.from("runs").update({
      status: "done",
      preview_json: previewRows,
      error_report_json: errors.length > 0 ? errors : null,
      output_storage_path: outputPath,
      output_xlsx_path: outputPath,
    }).eq("id", runId);

    console.log(`[process-contratos] Run ${runId} completed. ${consolidated.length} contratos, ${errors.length} errors`);

    return new Response(JSON.stringify({
      success: true,
      run_id: runId,
      total_contratos: consolidated.length,
      total_errors: errors.length,
      output_path: outputPath,
    }), {
      headers: { ...corsHeaders, "Content-Type": "application/json" },
    });
  } catch (e) {
    console.error("[process-contratos] Error:", e);
    const message = e instanceof Error ? e.message : "Unknown error";

    try {
      const body = await req.clone().json();
      if (body.upload_id) {
        const { data: runs } = await supabase.from("runs").select("id").eq("upload_id", body.upload_id).eq("status", "processing").limit(1);
        if (runs && runs[0]) {
          await supabase.from("runs").update({
            status: "error",
            error_report_json: [{ row: 0, field: "system", message, type: "system" }],
          }).eq("id", runs[0].id);
        }
      }
    } catch { /* ignore */ }

    return new Response(JSON.stringify({ error: message }), {
      status: 500,
      headers: { ...corsHeaders, "Content-Type": "application/json" },
    });
  }
});
