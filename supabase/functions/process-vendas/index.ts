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
  return removeAccents(str).toLowerCase().trim().replace(/[_\-\./]/g, " ").replace(/\s+/g, " ");
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

function cellToString(val: unknown): string {
  if (val == null) return "";
  if (typeof val === "number") {
    // Prevent scientific notation, preserve digits
    if (Number.isInteger(val)) return val.toFixed(0);
    return String(val);
  }
  return String(val).trim();
}

function detectDocumento(val: unknown): string {
  if (val == null) return "";
  const str = cellToString(val);
  const digits = extractDigits(str);
  if (isValidCPF(digits) || isValidCNPJ(digits)) return digits;
  if (/^\d{11,14}$/.test(digits)) return digits;
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
      // For very short synonyms (<=3 chars), require exact or word-boundary match
      if (nSyn.length <= 3) {
        const words = nH.split(" ");
        if (nH === nSyn || words.includes(nSyn)) return h;
      } else {
        if (nH === nSyn || nH.includes(nSyn) || nSyn.includes(nH)) return h;
      }
    }
  }
  return null;
}

function findAllColumns(headers: string[], synonyms: string[]): string[] {
  const result: string[] = [];
  for (const h of headers) {
    const nH = norm(h);
    if (synonyms.some(s => {
      const ns = norm(s);
      if (ns.length <= 3) {
        const words = nH.split(" ");
        return nH === ns || words.includes(ns);
      }
      return nH === ns || nH.includes(ns) || ns.includes(nH);
    })) {
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

// ============ CLASSIFICATION: VENDA vs CONTRATO ============

const CONTRACT_TERMS = [
  "vigencia", "contrato", "renovacao", "renovação", "mensalidade", "assinatura",
  "plano", "recorrente", "anual", "semestral", "trimestral", "12 meses", "24 meses",
];

const SALE_TERMS = [
  "venda", "venda pontual", "venda simples", "parcelamento pontual",
  "servico avulso", "avulso", "pontual", "job", "horas", "implantacao", "implantação",
];

// Columns that indicate the input has period/end dates (contract evidence)
const PERIOD_COL_TERMS = [
  "data final", "data fim", "fim", "termino", "término", "vigencia final",
  "vigência final", "comp final", "comp. final", "ate", "até", "periodo",
  "período", "data final contrato",
];

function hasContractPeriodColumns(headers: string[]): boolean {
  for (const h of headers) {
    const nH = norm(h);
    if (PERIOD_COL_TERMS.some(t => nH.includes(t))) return true;
  }
  return false;
}

function classifyRow(description: string, hasPeriodCols: boolean, rowHasPeriodValue: boolean): "venda" | "contrato" {
  const n = norm(description);

  // 1) If description has contract terms -> contrato
  if (CONTRACT_TERMS.some(t => n.includes(t))) return "contrato";

  // 2) If description has sale terms -> venda
  if (SALE_TERMS.some(t => n.includes(t))) return "venda";

  // 3) If the input has period columns and this row has a value in them -> contrato
  if (hasPeriodCols && rowHasPeriodValue) return "contrato";

  // 4) Default: if period columns exist at all, route to contrato to be safe
  if (hasPeriodCols) return "contrato";

  // 5) No signals at all: assume venda
  return "venda";
}

// ============ PAYMENT TYPE MAPPING ============

function mapTipoCobranca(val: string): string {
  if (!val || !val.trim()) return "Boleto";
  const n = norm(val);
  if (n.includes("debito em conta") || n.includes("debito conta")) return "Débito em Conta";
  if (n.includes("cartao de credito") || n.includes("credito") || n === "cc") return "Cartão de Crédito";
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

// ============ REFERENCE CODE VALIDATION ============

const INVALID_REF_TERMS = [
  "mensal", "semanal", "quinzenal", "bimestral", "trimestral", "semestral", "anual",
  "recorrente", "personalizado", "padrao", "padrão", "contrato", "sim", "nao", "não",
  "ativo", "inativo", "cancelado",
];

function isValidIdentifier(val: string): boolean {
  if (!val || !val.trim()) return false;
  const n = norm(val);
  if (INVALID_REF_TERMS.includes(n)) return false;
  // Must have at least one alphanumeric char
  if (!/[a-z0-9]/i.test(val)) return false;
  return true;
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

// ============ INSTALLMENT GROUPING ============

interface RawVenda {
  rowIdx: number;
  documento: string;
  tipoPessoa: string;
  nome: string;
  vencimento: string;
  competencia: string;
  valor: number;
  centroCusto: string;
  conta: string;
  tipoCobranca: string;
  observacoes: string;
  descricaoServico: string;
  identificadorVenda: string;
  cep: string;
  numeroEndereco: string;
  email: string;
  installmentKey: string;
}

function buildInstallmentKey(doc: string, desc: string, idVenda: string, valor: number): string {
  const docN = extractDigits(doc);
  const descN = norm(desc).replace(/[^a-z0-9 ]/g, "").replace(/\s+/g, " ").trim();
  const idN = norm(idVenda).replace(/[^a-z0-9 ]/g, "").replace(/\s+/g, " ").trim();
  const valN = isNaN(valor) ? "" : valor.toFixed(2);
  return `${docN}|${descN}|${idN}|${valN}`;
}

interface ConsolidatedVenda {
  documento: string;
  tipoPessoa: string;
  nome: string;
  competencia: string;
  centroCusto: string;
  conta: string;
  tipoCobranca: string;
  observacoes: string;
  descricaoServico: string;
  codigoReferencia: string;
  parcelas: number;
  valorTotal: number;
  periodicidade: string;
  primeiroVencimento: string;
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

    console.log(`[process-vendas] Starting for upload: ${upload_id}`);

    // 1) Create run
    const { data: run, error: runError } = await supabase.from("runs").insert({
      upload_id,
      import_type: "vendas",
      status: "processing",
    }).select().single();
    if (runError || !run) throw new Error(`Failed to create run: ${runError?.message}`);
    const runId = run.id;
    console.log(`[process-vendas] Run created: ${runId}`);

    // 2) Fetch upload & download primary file
    const { data: upload, error: uploadErr } = await supabase.from("uploads").select("*").eq("id", upload_id).single();
    if (uploadErr || !upload) throw new Error(`Upload not found: ${uploadErr?.message}`);

    const { data: primaryFile, error: dlErr } = await supabase.storage.from("uploads").download(upload.storage_path);
    if (dlErr || !primaryFile) throw new Error(`Download primary failed: ${dlErr?.message}`);
    console.log(`[process-vendas] Primary file: ${upload.original_filename}`);

    // 3) Fetch template
    const { data: tmpl, error: tmplErr } = await supabase
      .from("import_templates")
      .select("*")
      .eq("import_type", "vendas")
      .eq("is_default", true)
      .single();
    if (tmplErr || !tmpl) throw new Error(`Template not found: ${tmplErr?.message}`);

    const { data: templateFile, error: tmplDlErr } = await supabase.storage.from("templates").download(tmpl.template_storage_path);
    if (tmplDlErr || !templateFile) throw new Error(`Download template failed: ${tmplDlErr?.message}`);
    console.log(`[process-vendas] Template downloaded: ${tmpl.template_storage_path}`);

    // 4) Read primary Excel
    const primaryBuffer = await primaryFile.arrayBuffer();
    const primaryWb = XLSX.read(new Uint8Array(primaryBuffer), { type: "array" });
    const primarySheet = primaryWb.Sheets[primaryWb.SheetNames[0]];
    const rawData: Record<string, unknown>[] = XLSX.utils.sheet_to_json(primarySheet);
    if (rawData.length === 0) throw new Error("Planilha primária vazia");
    console.log(`[process-vendas] Primary data: ${rawData.length} rows`);

    const headers = Object.keys(rawData[0]);
    console.log(`[process-vendas] Headers: ${headers.join(", ")}`);

    // 5) Detect if input has period/end columns (contract evidence)
    const hasPeriodCols = hasContractPeriodColumns(headers);
    const periodColNames = headers.filter(h => PERIOD_COL_TERMS.some(t => norm(h).includes(t)));
    console.log(`[process-vendas] Has period columns: ${hasPeriodCols}${hasPeriodCols ? ` (${periodColNames.join(", ")})` : ""}`);

    // 6) Map columns from primary
    const colDoc = findColumn(headers, ["cpf", "cnpj", "cpf/cnpj", "cpf_cnpj", "documento", "doc"]);
    const colNome = findColumn(headers, ["cliente", "nome", "tomador", "sacado", "razao", "razao social", "razão social", "nome fantasia"]);
    const colVencimento = findColumn(headers, ["venc", "vencimento", "due date", "dt vencimento", "data vencimento", "dia do vencimento"]);
    const colCompetencia = findColumn(headers, ["compet", "competencia", "competência", "comp. inicial", "comp inicial"]);
    const colValor = findColumn(headers, ["valor bruto", "valor assinatura", "valor", "total", "amount", "vlr", "valor total"], ["liquido", "líquido", "net", "desconto"]);
    const colCentroCusto = findColumn(headers, ["centro", "custo", "centro de custo", "centro custo"]);
    const colConta = findColumn(headers, ["conta", "carteira", "banco", "caixa"]);
    const colTipoCobranca = findColumn(headers, [
      "forma de pagamento", "tipo de pagamento", "tipo pagamento", "forma pagamento",
      "tipo de cobranca", "tipo de cobrança", "tipo cobranca", "tipo cobrança",
      "cobranca", "cobrança", "pagamento", "meio de pagamento", "meio pagamento",
      "meio", "forma", "modalidade"
    ]);
    const colObs = findColumn(headers, ["obs", "observacao", "observação", "comentario", "comentário", "historico", "histórico", "informacao adicional", "informação adicional"]);
    const colDescricao = findColumn(headers, ["descricao", "descrição", "servico", "serviço", "produto/servico", "produto/serviço", "produto", "discriminacao", "discriminação"]);
    const colCep = findColumn(headers, ["cep", "zip", "codigo postal"]);
    const colNumeroEndereco = findColumn(headers, ["numero endereco", "número endereço", "num", "nro"]);
    const colEmail = findColumn(headers, ["email", "e-mail", "mail"]);

    // Identifier columns — exclude generic terms; use longer synonyms to avoid false matches
    const idSynonyms = [
      "numero venda", "num venda", "identificador", "referencia", "referência",
      "titulo", "título", "nosso numero", "nosso número", "pedido", "fatura",
      "invoice", "numero pedido", "número pedido", "codigo", "código",
    ];
    const colIdVenda = findColumn(headers, idSynonyms);

    // Description columns for classification
    const descCols = findAllColumns(headers, ["descricao", "descrição", "categoria", "tipo contrato", "historico", "histórico", "natureza", "servico", "serviço", "modelo", "produto/servico", "produto/serviço"]);

    console.log(`[process-vendas] Mapping: doc=${colDoc}, nome=${colNome}, venc=${colVencimento}, comp=${colCompetencia}, val=${colValor}, desc=${colDescricao}, id=${colIdVenda}`);

    // 7) Extract and classify rows
    const rawVendas: RawVenda[] = [];
    const errors: { row: number; field: string; message: string; type: string }[] = [];
    let routedToContrato = 0;
    const contratoReasons: string[] = [];

    for (let i = 0; i < rawData.length; i++) {
      const row = rawData[i];
      const rowNum = i + 2;

      // Build combined description for classification
      let combinedDesc = "";
      if (descCols.length > 0) {
        combinedDesc = descCols.map(c => cellToString(row[c])).join(" ");
      }

      // Check if this row has values in period columns
      const rowHasPeriodValue = periodColNames.some(col => {
        const v = cellToString(row[col]);
        return v.length > 0;
      });

      // Classify
      const classification = classifyRow(combinedDesc, hasPeriodCols, rowHasPeriodValue);
      if (classification === "contrato") {
        routedToContrato++;
        if (contratoReasons.length < 5) {
          contratoReasons.push(`Row ${rowNum}: "${combinedDesc.substring(0, 80)}" (period_col_value=${rowHasPeriodValue})`);
        }
        continue;
      }

      // Extract documento
      let documento = "";
      if (colDoc) documento = detectDocumento(row[colDoc]);
      if (!documento) {
        for (const key of headers) {
          const val = cellToString(row[key]);
          const digits = extractDigits(val);
          if (isValidCPF(digits) || isValidCNPJ(digits)) { documento = digits; break; }
        }
      }

      // Extract nome
      let nome = "";
      if (colNome) nome = cellToString(row[colNome]);
      if (!nome) {
        for (const key of headers) {
          if (key === colDoc) continue;
          const val = cellToString(row[key]);
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

      // Extract dates
      const vencimentoRaw = colVencimento ? normalizeDate(row[colVencimento]) : "";
      const competenciaRaw = colCompetencia ? normalizeDate(row[colCompetencia]) : "";
      const competencia = competenciaRaw || vencimentoRaw;
      const vencimento = vencimentoRaw;

      if (!competencia) {
        errors.push({ row: rowNum, field: "Data", message: `Linha ${rowNum}: Nenhuma data de competência ou vencimento encontrada`, type: "data" });
      }

      // Other fields
      const centroCusto = colCentroCusto ? cellToString(row[colCentroCusto]) : "";
      const conta = colConta ? cellToString(row[colConta]) : "";
      const tipoCobrancaRaw = colTipoCobranca ? cellToString(row[colTipoCobranca]) : "";
      const observacoes = colObs ? cellToString(row[colObs]) : "";
      const descricaoServico = colDescricao ? cellToString(row[colDescricao]) : "";
      const identificadorVenda = colIdVenda ? cellToString(row[colIdVenda]) : "";
      const cep = colCep ? cellToString(row[colCep]) : "";
      const numeroEndereco = colNumeroEndereco ? cellToString(row[colNumeroEndereco]) : "";
      const email = colEmail ? cellToString(row[colEmail]) : "";

      const installmentKey = buildInstallmentKey(documento, descricaoServico, identificadorVenda, valor);

      rawVendas.push({
        rowIdx: rowNum,
        documento,
        tipoPessoa: detectTipoPessoa(documento),
        nome,
        vencimento,
        competencia,
        valor,
        centroCusto,
        conta,
        tipoCobranca: tipoCobrancaRaw,
        observacoes,
        descricaoServico,
        identificadorVenda,
        cep,
        numeroEndereco,
        email,
        installmentKey,
      });
    }

    console.log(`[process-vendas] Classification: ${rawVendas.length} vendas, ${routedToContrato} routed to contrato`);
    if (contratoReasons.length > 0) {
      console.log(`[process-vendas] Contrato routing samples: ${JSON.stringify(contratoReasons)}`);
    }
    console.log(`[process-vendas] Raw vendas: ${rawVendas.length}, errors so far: ${errors.length}`);

    // 8) Group by installment key to detect parcelas
    const groups = new Map<string, RawVenda[]>();
    for (const v of rawVendas) {
      const key = v.installmentKey;
      if (!groups.has(key)) groups.set(key, []);
      groups.get(key)!.push(v);
    }

    const consolidated: ConsolidatedVenda[] = [];
    let refCounter = 1;
    let dedupCount = 0;

    for (const [_key, group] of groups) {
      group.sort((a, b) => parseDateForSort(a.vencimento) - parseDateForSort(b.vencimento));

      const first = group[0];
      const warnings: string[] = [];

      const uniqueDates = new Set(group.map(g => g.vencimento).filter(Boolean));
      const isInstallment = group.length >= 2 && uniqueDates.size >= 2;

      if (isInstallment) {
        dedupCount += group.length - 1;
        if (!first.identificadorVenda) {
          warnings.push(`parcelamento_sem_id: ${group.length} parcelas detectadas sem identificador (linhas ${group.map(g => g.rowIdx).join(",")})`);
        }
      }

      const parcelas = isInstallment ? group.length : 1;
      const valorTotal = isInstallment
        ? group.reduce((sum, g) => sum + g.valor, 0)
        : first.valor;

      // Multiple services detection
      const services = parseMultipleServices(first.descricaoServico);
      const multiServicos = services.length >= 2;

      if (multiServicos) {
        const serviceTotal = services.reduce((s, sv) => s + sv.value, 0);
        if (serviceTotal > 0 && Math.abs(serviceTotal - valorTotal) > 0.01) {
          const ratio = valorTotal / serviceTotal;
          let runningSum = 0;
          for (let si = 0; si < services.length; si++) {
            if (si === services.length - 1) {
              services[si].value = Math.round((valorTotal - runningSum) * 100) / 100;
            } else {
              services[si].value = Math.round(services[si].value * ratio * 100) / 100;
              runningSum += services[si].value;
            }
          }
          warnings.push(`multiplos_servicos_ajustado: soma ajustada para ${valorTotal.toFixed(2)}`);
        }
      }

      // Payment type - default Boleto
      const tipoCobranca = mapTipoCobranca(first.tipoCobranca);

      // Reference code - validate before using
      let codigoReferencia: string;
      if (isValidIdentifier(first.identificadorVenda)) {
        codigoReferencia = first.identificadorVenda;
      } else {
        codigoReferencia = String(refCounter);
        refCounter++;
      }

      consolidated.push({
        documento: first.documento,
        tipoPessoa: first.tipoPessoa,
        nome: first.nome,
        competencia: first.competencia,
        centroCusto: first.centroCusto,
        conta: first.conta || "Caixinha",
        tipoCobranca,
        observacoes: first.observacoes,
        descricaoServico: first.descricaoServico,
        codigoReferencia,
        parcelas,
        valorTotal,
        periodicidade: "Mensal",
        primeiroVencimento: first.vencimento || first.competencia,
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

    console.log(`[process-vendas] Consolidated vendas: ${consolidated.length}, deduped installments: ${dedupCount}`);

    // 9) Read template and fill
    const templateBuffer = await templateFile.arrayBuffer();
    const templateWb = XLSX.read(new Uint8Array(templateBuffer), { type: "array" });

    const dadosSheet = templateWb.Sheets[templateWb.SheetNames[0]];
    if (!dadosSheet) throw new Error("Aba 'Dados' não encontrada no template");

    const DATA_START_ROW = 3; // 0-indexed

    const dadosProtoStyles = captureRowStyles(dadosSheet, DATA_START_ROW, 30);

    const COL = {
      TIPO_TOMADOR: 0,
      CPF_CNPJ: 1,
      RAZAO_SOCIAL: 2,
      CEP: 3,
      NUMERO_ENDERECO: 4,
      EMAIL: 5,
      ID_INTEGRACAO: 6,
      DATA_COMPETENCIA: 7,
      CATEGORIA: 8,
      CENTRO_CUSTO: 9,
      CONTA: 10,
      TIPO_COBRANCA: 11,
      CODIGO_REF: 12,
      ORDEM_COMPRA: 13,
      ORDEM_SERVICO: 14,
      VENDEDOR: 15,
      OBSERVACOES: 16,
      INFO_NFSE: 17,
      MULTIPLOS_SERVICOS: 18,
      NOME_SERVICO: 19,
      DESCRICAO: 20,
      QUANTIDADE: 21,
      VALOR_UNITARIO: 22,
      PARCELAS: 27,
      PERIODICIDADE: 28,
      PRIMEIRO_VENCIMENTO: 29,
    };

    for (let i = 0; i < consolidated.length; i++) {
      const v = consolidated[i];
      const row = DATA_START_ROW + i;
      copyRowStyle(dadosSheet, row, dadosProtoStyles);

      setCellValue(dadosSheet, COL.TIPO_TOMADOR, row, v.tipoPessoa);
      if (v.documento) setCellValue(dadosSheet, COL.CPF_CNPJ, row, v.documento);
      setCellValue(dadosSheet, COL.RAZAO_SOCIAL, row, v.nome);
      if (v.cep) setCellValue(dadosSheet, COL.CEP, row, v.cep);
      if (v.numeroEndereco) setCellValue(dadosSheet, COL.NUMERO_ENDERECO, row, v.numeroEndereco);
      if (v.email) setCellValue(dadosSheet, COL.EMAIL, row, v.email);
      setCellValue(dadosSheet, COL.DATA_COMPETENCIA, row, v.competencia);
      setCellValue(dadosSheet, COL.CATEGORIA, row, "01.01.02");
      if (v.centroCusto) setCellValue(dadosSheet, COL.CENTRO_CUSTO, row, v.centroCusto);
      setCellValue(dadosSheet, COL.CONTA, row, v.conta);
      setCellValue(dadosSheet, COL.TIPO_COBRANCA, row, v.tipoCobranca);
      setCellValue(dadosSheet, COL.CODIGO_REF, row, v.codigoReferencia);
      if (v.observacoes) setCellValue(dadosSheet, COL.OBSERVACOES, row, v.observacoes);

      if (v.multiServicos) {
        setCellValue(dadosSheet, COL.MULTIPLOS_SERVICOS, row, "Sim");
      } else {
        setCellValue(dadosSheet, COL.NOME_SERVICO, row, "Prestação de Serviço");
        if (v.descricaoServico) setCellValue(dadosSheet, COL.DESCRICAO, row, v.descricaoServico);
        setCellValue(dadosSheet, COL.QUANTIDADE, row, 1);
        setCellValue(dadosSheet, COL.VALOR_UNITARIO, row, v.valorTotal);
      }

      setCellValue(dadosSheet, COL.PARCELAS, row, v.parcelas);
      setCellValue(dadosSheet, COL.PERIODICIDADE, row, v.periodicidade);
      if (v.primeiroVencimento) setCellValue(dadosSheet, COL.PRIMEIRO_VENCIMENTO, row, v.primeiroVencimento);
    }

    // Trim sheet range to exact data (no extra empty rows)
    const lastDataRow = DATA_START_ROW + consolidated.length;
    const templateRange = XLSX.utils.decode_range(dadosSheet["!ref"] || "A1");
    templateRange.e.r = Math.max(templateRange.e.r, lastDataRow - 1);
    dadosSheet["!ref"] = XLSX.utils.encode_range(templateRange);

    // --- Fill Serviços sheet ---
    const servicosSheetName = templateWb.SheetNames[1];
    if (servicosSheetName) {
      const servicosSheet = templateWb.Sheets[servicosSheetName];
      if (servicosSheet) {
        const SVC_COL = {
          CODIGO_REF: 0,
          NOME_SERVICO: 1,
          DESCRICAO: 2,
          QUANTIDADE: 3,
          VALOR_UNITARIO: 4,
        };
        let svcRow = 1;

        const svcProtoStyles = captureRowStyles(servicosSheet, svcRow, 5);

        for (const v of consolidated) {
          if (!v.multiServicos) continue;
          for (const svc of v.parsedServices) {
            copyRowStyle(servicosSheet, svcRow, svcProtoStyles);
            setCellValue(servicosSheet, SVC_COL.CODIGO_REF, svcRow, v.codigoReferencia);
            setCellValue(servicosSheet, SVC_COL.NOME_SERVICO, svcRow, "Prestação de Serviço");
            setCellValue(servicosSheet, SVC_COL.DESCRICAO, svcRow, svc.description);
            setCellValue(servicosSheet, SVC_COL.QUANTIDADE, svcRow, 1);
            setCellValue(servicosSheet, SVC_COL.VALOR_UNITARIO, svcRow, svc.value);
            svcRow++;
          }
        }

        if (svcRow > 1) {
          updateSheetRange(servicosSheet, svcRow, 5);
        }
      }
    }

    // 10) Generate preview
    const previewRows: Record<string, string | number>[] = [];
    for (let i = 0; i < Math.min(consolidated.length, 20); i++) {
      const v = consolidated[i];
      previewRows.push({
        "Tipo Tomador": v.tipoPessoa,
        "CPF/CNPJ": v.documento,
        "Razão Social": v.nome,
        "Competência": v.competencia,
        "Conta": v.conta,
        "Cobrança": v.tipoCobranca,
        "Cód. Ref.": v.codigoReferencia,
        "Multi Svc?": v.multiServicos ? "Sim" : "",
        "Parcelas": v.parcelas,
        "Valor": formatCurrency(v.valorTotal),
        "1º Venc.": v.primeiroVencimento,
      });
    }

    // 11) Write output XLSX
    const outputBuffer = XLSX.write(templateWb, { type: "array", bookType: "xlsx" });
    const outputPath = `vendas/${runId}.xlsx`;

    const { error: uploadOutputErr } = await supabase.storage
      .from("outputs")
      .upload(outputPath, new Uint8Array(outputBuffer), {
        contentType: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        upsert: true,
      });
    if (uploadOutputErr) throw new Error(`Upload output failed: ${uploadOutputErr.message}`);

    console.log(`[process-vendas] Output saved: ${outputPath}`);

    // 12) Update run
    await supabase.from("runs").update({
      status: "done",
      preview_json: previewRows,
      error_report_json: errors.length > 0 ? errors : null,
      output_storage_path: outputPath,
      output_xlsx_path: outputPath,
    }).eq("id", runId);

    console.log(`[process-vendas] Run ${runId} completed. ${consolidated.length} vendas exported, ${routedToContrato} routed to contrato, ${errors.length} errors/warnings`);

    return new Response(JSON.stringify({
      success: true,
      run_id: runId,
      total_vendas: consolidated.length,
      total_routed_contrato: routedToContrato,
      total_errors: errors.length,
      output_path: outputPath,
    }), {
      headers: { ...corsHeaders, "Content-Type": "application/json" },
    });
  } catch (e) {
    console.error("[process-vendas] Error:", e);
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

function parseDateForSort(dateStr: string): number {
  if (!dateStr) return 0;
  const m = dateStr.match(/^(\d{2})\/(\d{2})\/(\d{4})$/);
  if (m) return parseInt(m[3]) * 10000 + parseInt(m[2]) * 100 + parseInt(m[1]);
  return 0;
}
