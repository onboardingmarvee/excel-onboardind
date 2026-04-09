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
  return removeAccents(str).toLowerCase().trim().replace(/[_\-.\/]/g, " ").replace(/\s+/g, " ");
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
  // Try text-based due date parsing
  const textParsed = parseDueDateFromText(str, new Date());
  if (textParsed) return textParsed;
  return str;
}

/**
 * Parse textual due date descriptions like "dia 5 do mês", "último dia de cada mês", etc.
 * Returns dd/mm/yyyy string or null.
 */
function parseDueDateFromText(raw: string, baseDate: Date): string | null {
  if (!raw) return null;
  const cleaned = removeAccents(raw).toLowerCase().trim();

  // Already a date dd/mm/yyyy
  if (/^\d{2}\/\d{2}\/\d{4}$/.test(cleaned)) return cleaned;

  const year = baseDate.getFullYear();
  const month = baseDate.getMonth(); // 0-based

  // "ultimo dia de cada mes" / "último dia do mês"
  if (/ultim[oa]\s+dia/.test(cleaned)) {
    const lastDay = new Date(year, month + 1, 0).getDate();
    return `${String(lastDay).padStart(2, "0")}/${String(month + 1).padStart(2, "0")}/${year}`;
  }

  // "dia 5 do mes", "dia 5 de cada mes", "dia 5", "05 de cada mes", "03 de cada mes"
  const dayMatch = cleaned.match(/(?:dia\s+)?(\d{1,2})(?:\s+(?:do|de\s+cada)\s+mes)?/);
  if (dayMatch) {
    let day = parseInt(dayMatch[1]);
    if (day >= 1 && day <= 31) {
      // Clamp to last day of month
      const lastDay = new Date(year, month + 1, 0).getDate();
      if (day > lastDay) day = lastDay;
      return `${String(day).padStart(2, "0")}/${String(month + 1).padStart(2, "0")}/${year}`;
    }
  }

  return null;
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

// ============ EXPENSE DETECTION ============

const EXPENSE_TERMS = [
  "aluguel", "condominio", "condomínio", "energia", "luz", "internet", "telefone",
  "software", "assinatura", "pro-labore", "pró-labore", "pro labore", "salario", "salário",
  "folha", "imposto", "taxa", "tarifa", "honorarios", "honorários",
  "prestador", "fornecedor", "compra", "fatura", "boleto", "pix",
  "despesa", "saida", "saída", "pagamento", "contas a pagar",
  "remuneracao", "remuneração", "distribuicao", "distribuição", "lucros",
  "material", "manutencao", "manutenção", "seguro", "agua", "água",
  "combustivel", "combustível", "transporte", "frete", "cartorio", "cartório",
  "contabilidade", "advocacia", "consultoria"
];

const REVENUE_TERMS = [
  "venda", "recebimento", "receita", "pagamento cliente", "fatura recebida",
  "contas a receber", "entrada", "credito", "crédito"
];

// Government/institutional names that contain revenue-like words but are expense payees
const EXPENSE_INSTITUTIONS = [
  "receita federal", "secretaria da receita", "receita estadual", "receita municipal",
  "ministerio da fazenda", "sefaz", "prefeitura", "governo",
  "darf", "das", "gps", "gare", "guia de recolhimento"
];

function isExpense(row: Record<string, unknown>, headers: string[], tipoCols: string[]): "expense" | "revenue" | "unknown" {
  // Check tipo/natureza column
  for (const col of tipoCols) {
    const val = norm(String(row[col] ?? ""));
    if (val.includes("despesa") || val.includes("saida") || val.includes("pagamento") || val.includes("contas a pagar")) return "expense";
    if (val.includes("receita") || val.includes("entrada") || val.includes("recebimento") || val.includes("contas a receber")) {
      // Before classifying as revenue, check if it's actually an institution name
      if (EXPENSE_INSTITUTIONS.some(t => val.includes(norm(t)))) return "expense";
      return "revenue";
    }
  }

  // Check description
  const allText = headers.map(h => String(row[h] ?? "")).join(" ");
  const n = norm(allText);

  // Check expense institutions BEFORE revenue terms (overrides "receita" match)
  if (EXPENSE_INSTITUTIONS.some(t => n.includes(norm(t)))) return "expense";
  if (REVENUE_TERMS.some(t => n.includes(norm(t)))) return "revenue";
  if (EXPENSE_TERMS.some(t => n.includes(norm(t)))) return "expense";

  // Check for negative values (might indicate expense in some formats)
  for (const h of headers) {
    const val = row[h];
    if (typeof val === "number" && val < 0) return "expense";
  }

  return "unknown";
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
  // Default per spec: "Caso não informe, preencher como padrão Boleto"
  return "Boleto";
}

// ============ PIX KEY TYPE DETECTION ============

function detectPixKeyType(key: string): string {
  if (!key) return "";
  const digits = extractDigits(key);
  if (digits.length === 11 && isValidCPF(digits)) return "CPF";
  if (digits.length === 14 && isValidCNPJ(digits)) return "CNPJ";
  if (/^[a-z0-9._%+-]+@[a-z0-9.-]+\.[a-z]{2,}$/i.test(key.trim())) return "E-mail";
  if (/^\+?55?\d{10,11}$/.test(digits) || (digits.length >= 10 && digits.length <= 13)) return "Telefone";
  return "Aleatória";
}

// ============ INSTALLMENT DETECTION ============

function detectInstallment(desc: string): { current: number; total: number } | null {
  const n = norm(desc);
  // "parcela 1/10", "1/10", "1 de 10", "parc 4/7"
  const m = n.match(/(?:parcela|parc\.?|parcel\.?)?\s*(\d{1,3})\s*(?:\/|de)\s*(\d{1,3})/);
  if (m) {
    const current = parseInt(m[1]);
    const total = parseInt(m[2]);
    if (current > 0 && total > 0 && current <= total) return { current, total };
  }
  return null;
}

// ============ RECURRENCE DETECTION ============

const RECURRENCE_TERMS = [
  "mensal", "recorrente", "assinatura", "mensalidade", "aluguel",
  "condominio", "condomínio", "energia", "luz", "internet",
  "pro-labore", "pró-labore", "pro labore", "salario", "salário",
  "folha", "remuneracao", "remuneração", "distribuicao de lucros", "distribuição de lucros",
  // CLT / trabalhista
  "clt", "funcionario", "funcionário", "empregado", "trabalhista",
  "ferias", "férias", "13o", "decimo terceiro", "décimo terceiro",
  "fgts", "inss", "rescisao", "rescisão",
  // PJ / prestação de serviço
  "pj", "pessoa juridica", "pessoa jurídica",
  "prestacao de servico", "prestação de serviço",
  // Veículo / financiamento
  "financiamento", "leasing", "veiculo", "veículo",
  "carro", "automovel", "automóvel", "parcela carro", "parcela veiculo",
  // Impostos
  "iss", "irpf", "simples nacional", "pis", "cofins", "csll", "irpj",
  "darf", "das", "gps", "gare", "imposto", "tributo"
];

const PERIODICITY_MAP: [string, string][] = [
  ["semanal", "Semanal"],
  ["quinzenal", "Quinzenal"],
  ["bimestral", "Bimestral"],
  ["trimestral", "Trimestral"],
  ["semestral", "Semestral"],
  ["anual", "Anual"],
  ["mensal", "Mensal"],
];

function detectPeriodicity(desc: string): string | null {
  const n = norm(desc);
  for (const [term, label] of PERIODICITY_MAP) {
    if (n.includes(term)) return label;
  }
  return null;
}

function computeRepetitions(periodicidade: string): number {
  const now = new Date();
  const currentMonth = now.getMonth(); // 0-based
  const currentYear = now.getFullYear();
  const endYear = currentMonth >= 6 ? currentYear + 1 : currentYear; // past July -> next year
  const endDate = new Date(endYear, 11, 31); // Dec 31

  const months = (endDate.getFullYear() - now.getFullYear()) * 12 + (endDate.getMonth() - now.getMonth());

  switch (periodicidade) {
    case "Semanal": return Math.floor((months * 30) / 7);
    case "Quinzenal": return months * 2;
    case "Mensal": return months;
    case "Bimestral": return Math.floor(months / 2);
    case "Trimestral": return Math.floor(months / 3);
    case "Semestral": return Math.floor(months / 6);
    case "Anual": return Math.floor(months / 12);
    default: return months;
  }
}

// ============ TAX DEFAULTS ============

interface TaxDefault {
  terms: string[];         // detection terms
  diaVencimento: number | null;  // default due day (null = compute last business day)
  conta: string;           // default account
  periodicidade: string;   // "Mensal" or "Trimestral"
  mesesTrimestre?: number[]; // for quarterly taxes: months (1-indexed)
}

const TAX_DEFAULTS: TaxDefault[] = [
  { terms: ["fgts"], diaVencimento: 20, conta: "Caixinha", periodicidade: "Mensal" },
  { terms: ["iss", "imposto sobre servico", "imposto sobre serviço"], diaVencimento: 10, conta: "Caixinha", periodicidade: "Mensal" },
  { terms: ["inss", "gps"], diaVencimento: 20, conta: "Caixinha", periodicidade: "Mensal" },
  { terms: ["irpf", "imposto de renda pessoa fisica", "imposto de renda pessoa física"], diaVencimento: 20, conta: "Caixinha", periodicidade: "Mensal" },
  { terms: ["simples nacional", "simples", "das simples"], diaVencimento: 20, conta: "Caixinha", periodicidade: "Mensal" },
  { terms: ["pis"], diaVencimento: 25, conta: "Caixinha", periodicidade: "Mensal" },
  { terms: ["cofins"], diaVencimento: 25, conta: "Caixinha", periodicidade: "Mensal" },
  { terms: ["csll"], diaVencimento: null, conta: "Caixinha", periodicidade: "Trimestral", mesesTrimestre: [1, 4, 7, 10] },
  { terms: ["irpj", "imposto de renda pessoa juridica", "imposto de renda pessoa jurídica"], diaVencimento: null, conta: "Caixinha", periodicidade: "Trimestral", mesesTrimestre: [1, 4, 7, 10] },
];

function detectTaxDefault(descBase: string, nome: string): TaxDefault | null {
  const n = norm(descBase) + " " + norm(nome);
  for (const tax of TAX_DEFAULTS) {
    if (tax.terms.some(t => n.includes(norm(t)))) return tax;
  }
  return null;
}

/** Get last business day of a month (skip Sat/Sun) */
function lastBusinessDay(year: number, month: number): number {
  // month is 1-indexed
  const lastDay = new Date(year, month, 0).getDate(); // last calendar day
  const d = new Date(year, month - 1, lastDay);
  while (d.getDay() === 0 || d.getDay() === 6) {
    d.setDate(d.getDate() - 1);
  }
  return d.getDate();
}

function computeQuarterlyRepetitions(mesesTrimestre: number[]): number {
  const now = new Date();
  const currentMonth = now.getMonth() + 1; // 1-indexed
  const currentYear = now.getFullYear();
  const endYear = currentMonth > 7 ? currentYear + 1 : currentYear;

  let count = 0;
  for (let y = currentYear; y <= endYear; y++) {
    for (const m of mesesTrimestre) {
      const ref = new Date(y, m - 1, 1);
      if (ref > now && ref <= new Date(endYear, 11, 31)) count++;
    }
  }
  return Math.max(count - 1, 0); // subtract 1 because first occurrence is the current one
}

// ============ DESCRIPTION CLEANING FOR CATEGORY ============

function cleanDescriptionForCategory(descBase: string, descFinal: string): string {
  if (descFinal === "Financeiro por assinatura") return descFinal;
  let text = descBase;
  // Remove prefixes
  text = text.replace(/^PROVISÃO\s*-\s*/i, "").replace(/^Cartão de Crédito\s*-\s*/i, "");
  // Remove noise words
  const noiseWords = ["provisao", "provisão", "cartao", "cartão", "credito", "crédito"];
  let cleaned = norm(text);
  for (const w of noiseWords) {
    cleaned = cleaned.replace(new RegExp(`\\b${norm(w)}\\b`, "g"), "");
  }
  return cleaned.replace(/\s+/g, " ").trim() || descBase;
}

// ============ MAIN HANDLER ============

serve(async (req) => {
  if (req.method === "OPTIONS") {
    return new Response(null, { headers: corsHeaders });
  }

  const supabaseUrl = Deno.env.get("SUPABASE_URL")!;
  const serviceKey = Deno.env.get("SUPABASE_SERVICE_ROLE_KEY")!;
  const supabase = createClient(supabaseUrl, serviceKey);

  try {
    const { upload_id } = await req.json();
    if (!upload_id) throw new Error("upload_id is required");

    console.log(`[process-movimentacoes] Starting for upload: ${upload_id}`);

    // 1) Create run
    const { data: run, error: runError } = await supabase.from("runs").insert({
      upload_id,
      import_type: "movimentacoes",
      status: "processing",
    }).select().single();
    if (runError || !run) throw new Error(`Failed to create run: ${runError?.message}`);
    const runId = run.id;
    console.log(`[process-movimentacoes] Run created: ${runId}`);

    // 2) Fetch upload & download primary file
    const { data: upload, error: uploadErr } = await supabase.from("uploads").select("*").eq("id", upload_id).single();
    if (uploadErr || !upload) throw new Error(`Upload not found: ${uploadErr?.message}`);

    const { data: primaryFile, error: dlErr } = await supabase.storage.from("uploads").download(upload.storage_path);
    if (dlErr || !primaryFile) throw new Error(`Download primary failed: ${dlErr?.message}`);
    console.log(`[process-movimentacoes] Primary file: ${upload.original_filename}`);

    // 3) Fetch template
    const { data: tmpl, error: tmplErr } = await supabase
      .from("import_templates")
      .select("*")
      .eq("import_type", "movimentacoes")
      .eq("is_default", true)
      .single();
    if (tmplErr || !tmpl) throw new Error(`Template not found: ${tmplErr?.message}`);

    const { data: templateFile, error: tmplDlErr } = await supabase.storage.from("templates").download(tmpl.template_storage_path);
    if (tmplDlErr || !templateFile) throw new Error(`Download template failed: ${tmplDlErr?.message}`);
    console.log(`[process-movimentacoes] Template downloaded: ${tmpl.template_storage_path}`);

    // 4) Read primary Excel
    const primaryBuffer = await primaryFile.arrayBuffer();
    const primaryWb = XLSX.read(new Uint8Array(primaryBuffer), { type: "array" });
    const primarySheet = primaryWb.Sheets[primaryWb.SheetNames[0]];
    const rawData: Record<string, unknown>[] = XLSX.utils.sheet_to_json(primarySheet);
    if (rawData.length === 0) throw new Error("Planilha primária vazia");
    console.log(`[process-movimentacoes] Primary data: ${rawData.length} rows`);

    const headers = Object.keys(rawData[0]);
    console.log(`[process-movimentacoes] Headers: ${headers.join(", ")}`);

    // 5) Map columns
    const colDoc = findColumn(headers, ["cpf", "cnpj", "cpf/cnpj", "cpf_cnpj", "documento", "doc"]);
    const colNome = findColumn(headers, ["fornecedor", "prestador", "colaborador", "favorecido", "nome", "beneficiario", "beneficiário", "razao", "razao social", "razão social", "cliente"]);
    const colDescricao = findColumn(headers, ["descricao", "descrição", "historico", "histórico", "memo", "detalhe", "servico", "serviço", "produto", "categoria texto"]);
    const colTipoCobranca = findColumn(headers, [
      "forma de pagamento", "tipo de pagamento", "tipo pagamento", "forma pagamento",
      "tipo de cobranca", "tipo de cobrança", "tipo cobranca", "tipo cobrança",
      "cobranca", "cobrança", "pagamento", "meio de pagamento", "meio pagamento",
      "meio", "forma", "modalidade"
    ]);
    const colVencimento = findColumn(headers, ["venc", "vencimento", "due date", "dt vencimento", "data vencimento", "dia do vencimento", "dia vencimento", "dia venc"]);
    const colValor = findColumn(headers, ["valor bruto", "valor", "total", "amount", "vlr", "valor total"], ["liquido", "líquido", "net", "desconto"]);
    const colConta = findColumn(headers, ["conta", "carteira", "banco", "caixa"]);
    const colObs = findColumn(headers, ["obs", "observacao", "observação", "comentario", "comentário"]);
    const colChavePix = findColumn(headers, ["pix", "chave", "qr", "copia e cola", "copiaecola", "chave pix"]);
    const colCentroCusto = findColumn(headers, ["centro", "custo", "cc", "centro de custo", "centro custo"]);
    const colCep = findColumn(headers, ["cep", "zip", "codigo postal"]);
    const colNumero = findColumn(headers, ["numero endereco", "número endereço", "num", "nro"]);
    const colEmail = findColumn(headers, ["email", "e-mail", "mail"]);

    // Tipo/natureza columns for expense detection
    const tipoCols: string[] = [];
    for (const h of headers) {
      const nH = norm(h);
      if (["tipo", "natureza", "entrada saida", "entrada/saida", "classificacao", "classificação"].some(s => nH.includes(s))) {
        tipoCols.push(h);
      }
    }

    console.log(`[process-movimentacoes] Mapping: doc=${colDoc}, nome=${colNome}, desc=${colDescricao}, val=${colValor}, venc=${colVencimento}, tipo_cols=${tipoCols.join(",")}`);

    // 6) Extract despesas
    interface RawDespesa {
      rowIdx: number;
      documento: string;
      tipoPessoa: string;
      nome: string;
      descricaoBase: string;
      tipoCobrancaRaw: string;
      vencimento: string;
      valor: number;
      conta: string;
      observacoes: string;
      chavePix: string;
      centroCusto: string;
      cep: string;
      numero: string;
      email: string;
    }

    const rawDespesas: RawDespesa[] = [];
    const errors: { row: number; field: string; message: string; type: string }[] = [];

    for (let i = 0; i < rawData.length; i++) {
      const row = rawData[i];
      const rowNum = i + 2;

      // Expense filter — this sheet is for expenses only, so:
      // • "revenue" rows are always skipped
      // • "unknown" rows are treated as expenses by default (the user is
      //   explicitly using the movimentacoes processor for their despesas file)
      const classification = isExpense(row, headers, tipoCols);
      if (classification === "revenue") continue;
      if (classification === "unknown") {
        // Log as informational warning but do NOT skip — process as expense
        errors.push({ row: rowNum, field: "classificação", message: `Linha ${rowNum}: Classificação ambígua, processado como despesa por padrão`, type: "linha_presumida_despesa" });
      }

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

      if (!nome) {
        errors.push({ row: rowNum, field: "Razão Social", message: `Linha ${rowNum}: Nome/Razão Social ausente, registro ignorado`, type: "obrigatorio" });
        continue;
      }

      if (!documento) {
        errors.push({ row: rowNum, field: "CPF/CNPJ", message: `Linha ${rowNum}: Documento não encontrado`, type: "aviso" });
      }

      // Extract valor
      const valorRaw = colValor ? row[colValor] : undefined;
      let valor = normalizeCurrency(valorRaw);
      if (isNaN(valor) || valor === 0) {
        valor = 0.01;
        errors.push({ row: rowNum, field: "Valor", message: `Linha ${rowNum}: Valor ausente ou inválido (${valorRaw}), usando 0,01 como padrão`, type: "aviso" });
      }
      valor = Math.abs(valor); // Ensure positive

      // Extract vencimento (no hard skip — defer resolution to processing phase)
      const vencRaw = colVencimento ? row[colVencimento] : undefined;
      const vencimento = colVencimento ? normalizeDate(vencRaw) : "";
      const vencIsValidDate = /^\d{2}\/\d{2}\/\d{4}$/.test(vencimento);
      if (!vencIsValidDate) {
        const rawStr = vencRaw != null ? String(vencRaw).trim() : "(vazio)";
        errors.push({ row: rowNum, field: "Vencimento", message: `Linha ${rowNum}: Vencimento não encontrado no input (raw: "${rawStr}"), será inferido se possível`, type: "aviso_vencimento" });
      }

      // Other fields
      const descricaoBase = colDescricao ? String(row[colDescricao] ?? "").trim() : "";
      const tipoCobrancaRaw = colTipoCobranca ? String(row[colTipoCobranca] ?? "").trim() : "";
      const conta = colConta ? String(row[colConta] ?? "").trim() : "";
      const observacoes = colObs ? String(row[colObs] ?? "").trim() : "";
      const chavePix = colChavePix ? String(row[colChavePix] ?? "").trim() : "";
      const centroCusto = colCentroCusto ? String(row[colCentroCusto] ?? "").trim() : "";
      const cep = colCep ? String(row[colCep] ?? "").trim() : "";
      const numero = colNumero ? String(row[colNumero] ?? "").trim() : "";
      const email = colEmail ? String(row[colEmail] ?? "").trim() : "";

      rawDespesas.push({
        rowIdx: rowNum, documento, tipoPessoa: detectTipoPessoa(documento),
        nome, descricaoBase, tipoCobrancaRaw, vencimento, valor,
        conta, observacoes, chavePix, centroCusto, cep, numero, email,
      });
    }

    // Logging: extraction stats
    const vencFromInput = rawDespesas.filter(d => /^\d{2}\/\d{2}\/\d{4}$/.test(d.vencimento)).length;
    const vencMissing = rawDespesas.length - vencFromInput;
    console.log(`[process-movimentacoes] Raw despesas: ${rawDespesas.length}, venc from input: ${vencFromInput}, venc missing (will infer later): ${vencMissing}, errors so far: ${errors.length}`);
    // Sample: log first 10 raw vencimento values from input
    const sampleSize = Math.min(rawData.length, 10);
    for (let s = 0; s < sampleSize; s++) {
      const rawVenc = colVencimento ? rawData[s][colVencimento] : undefined;
      const parsed = colVencimento ? normalizeDate(rawVenc) : "";
      console.log(`[process-movimentacoes] Sample[${s}] venc raw="${rawVenc}" -> parsed="${parsed}"`);
    }

    // 7) Process each despesa and build output rows
    interface OutputRow {
      tipoPessoa: string;
      documento: string;
      nome: string;
      cep: string;
      numero: string;
      email: string;
      tipoDocumento: string;
      descricao: string;
      nFatura: number;
      dataCompetencia: string;
      categoria: string;
      centroCusto: string;
      conta: string;
      tipoCobranca: string;
      tipoChavePix: string;
      chavePix: string;
      codigoReferencia: string;
      observacoes: string;
      qtdRepeticoes: number | string;
      periodicidade: string;
      vencimento: string;
      valor: number;
    }

    const outputRows: OutputRow[] = [];
    let faturaCounter = 1;

    // First day of current month
    const now = new Date();
    const dataCompetencia = `01/${String(now.getMonth() + 1).padStart(2, "0")}/${now.getFullYear()}`;

    for (const d of rawDespesas) {
      const descBase = d.descricaoBase || "Despesa";

      // Tipo cobrança
      const tipoCobranca = mapTipoCobranca(d.tipoCobrancaRaw);
      if (d.tipoCobrancaRaw && tipoCobranca === "Outros") {
        errors.push({ row: d.rowIdx, field: "Tipo Cobrança", message: `"${d.tipoCobrancaRaw}" mapeado para "Outros"`, type: "warning" });
      }

      // Description rules
      let descricaoFinal: string;
      const nomeNorm = norm(d.nome);
      const descNorm = norm(descBase);

      if (nomeNorm.includes("marvee") || descNorm.includes("marvee")) {
        descricaoFinal = "Financeiro por assinatura";
      } else if (
        tipoCobranca === "Cartão de Crédito" ||
        descNorm.includes("visa") || descNorm.includes("master") || descNorm.includes("amex") ||
        descNorm.includes("cartao") || descNorm.includes("credito")
      ) {
        descricaoFinal = "Cartão de Crédito - " + descBase;
      } else {
        descricaoFinal = "PROVISÃO - " + descBase;
      }

      // Category suggestion — combine fornecedor name + descrição for richer signal
      const textoParaCategoria = cleanDescriptionForCategory(
        `${d.nome} ${descBase}`.trim(),
        descricaoFinal
      );

      let categoriaCode = "02.01.01"; // default expense category
      try {
        const suggestResp = await fetch(`${supabaseUrl}/functions/v1/suggest-financial-category`, {
          method: "POST",
          headers: {
            "Content-Type": "application/json",
            "Authorization": `Bearer ${serviceKey}`,
          },
          body: JSON.stringify({ description: textoParaCategoria }),
        });
        if (suggestResp.ok) {
          const suggestion = await suggestResp.json();
          if (suggestion.code) {
            categoriaCode = suggestion.code;
            if (suggestion.score < 0.25) {
              errors.push({
                row: d.rowIdx, field: "Categoria", type: "categoria_suspeita",
                message: `Score baixo (${suggestion.score}) para "${textoParaCategoria}". Top 3: ${(suggestion.top || []).slice(0, 3).map((t: any) => `${t.code}:${t.name}(${t.score})`).join(", ")}`,
              });
            }
          }
        }
      } catch (e) {
        console.error(`[process-movimentacoes] Category suggestion failed for row ${d.rowIdx}:`, e);
      }

      // PIX key
      let tipoChavePix = "";
      let chavePix = "";
      if (tipoCobranca === "Pix" && d.chavePix) {
        tipoChavePix = detectPixKeyType(d.chavePix);
        chavePix = d.chavePix;
      } else if (tipoCobranca === "Pix" && !d.chavePix) {
        errors.push({ row: d.rowIdx, field: "Chave Pix", message: `Tipo Pix mas chave não identificada`, type: "pix_sem_chave" });
      }

      // Tax defaults detection
      const taxDefault = detectTaxDefault(descBase, d.nome);

      // Installment detection
      const installment = detectInstallment(descBase);
      let qtdRepeticoes: number | string = "";
      let periodicidade = "";

      if (installment) {
        qtdRepeticoes = installment.total - installment.current;
        periodicidade = "Mensal";
      } else if (taxDefault) {
        // Tax-specific recurrence
        periodicidade = taxDefault.periodicidade === "Trimestral" ? "Trimestral" : "Mensal";
        if (taxDefault.periodicidade === "Trimestral" && taxDefault.mesesTrimestre) {
          qtdRepeticoes = computeQuarterlyRepetitions(taxDefault.mesesTrimestre);
        } else {
          qtdRepeticoes = computeRepetitions(periodicidade);
        }
      } else {
        // General recurrence detection
        const isRecurrent = RECURRENCE_TERMS.some(t => norm(descBase).includes(norm(t)) || nomeNorm.includes(norm(t)));
        if (isRecurrent) {
          periodicidade = detectPeriodicity(descBase) || "Mensal";
          qtdRepeticoes = computeRepetitions(periodicidade);
        }
      }

      // Vencimento resolution pipeline: input -> tax default -> recurrence -> fallback
      let vencimentoFinal = /^\d{2}\/\d{2}\/\d{4}$/.test(d.vencimento) ? d.vencimento : "";
      let vencInferenceMethod = vencimentoFinal ? "input" : "";

      const y = now.getFullYear();
      const m = now.getMonth() + 1;

      // Step 1: Tax default day
      if (!vencimentoFinal && taxDefault) {
        if (taxDefault.diaVencimento) {
          vencimentoFinal = `${String(taxDefault.diaVencimento).padStart(2, "0")}/${String(m).padStart(2, "0")}/${y}`;
        } else {
          const lbd = lastBusinessDay(y, m);
          vencimentoFinal = `${String(lbd).padStart(2, "0")}/${String(m).padStart(2, "0")}/${y}`;
        }
        vencInferenceMethod = "tax_default";
        errors.push({ row: d.rowIdx, field: "Vencimento", message: `Vencimento inferido via imposto: ${vencimentoFinal}`, type: "venc_inferido" });
      }

      // Step 2: Recurrence — use day 1 of next month
      if (!vencimentoFinal && (periodicidade || qtdRepeticoes)) {
        const nextM = m === 12 ? 1 : m + 1;
        const nextY = m === 12 ? y + 1 : y;
        vencimentoFinal = `01/${String(nextM).padStart(2, "0")}/${nextY}`;
        vencInferenceMethod = "recurrence";
        errors.push({ row: d.rowIdx, field: "Vencimento", message: `Vencimento inferido via recorrência: ${vencimentoFinal}`, type: "venc_inferido" });
      }

      // Step 3: Generic fallback — 1st of next month as provisioned date
      if (!vencimentoFinal) {
        const nextM = m === 12 ? 1 : m + 1;
        const nextY = m === 12 ? y + 1 : y;
        vencimentoFinal = `01/${String(nextM).padStart(2, "0")}/${nextY}`;
        vencInferenceMethod = "fallback";
        errors.push({ row: d.rowIdx, field: "Vencimento", message: `Vencimento não encontrado, usando provisão: ${vencimentoFinal}`, type: "venc_fallback" });
      }

      // Conta — apply tax default if missing
      const conta = d.conta || (taxDefault?.conta) || "Caixinha";

      // Código referência
      const codigoReferencia = `${runId.slice(0, 8)}-${String(faturaCounter).padStart(4, "0")}`;

      outputRows.push({
        tipoPessoa: d.tipoPessoa,
        documento: d.documento,
        nome: d.nome,
        cep: d.cep,
        numero: d.numero,
        email: d.email,
        tipoDocumento: "Contas a Pagar",
        descricao: descricaoFinal,
        nFatura: faturaCounter,
        dataCompetencia,
        categoria: categoriaCode,
        centroCusto: d.centroCusto,
        conta,
        tipoCobranca,
        tipoChavePix,
        chavePix,
        codigoReferencia,
        observacoes: d.observacoes,
        qtdRepeticoes,
        periodicidade,
        vencimento: vencimentoFinal,
        valor: d.valor,
      });

      faturaCounter++;
    }

    // Log vencimento inference summary
    const vencStats = { input: 0, tax_default: 0, recurrence: 0, fallback: 0 };
    for (const o of outputRows) {
      // Count based on error types logged
      const inferErr = errors.find(e => e.row === 0 && false); // placeholder
    }
    // Count from errors array
    vencStats.input = outputRows.length - errors.filter(e => e.type === "venc_inferido" || e.type === "venc_fallback").length;
    vencStats.tax_default = errors.filter(e => e.type === "venc_inferido" && e.message.includes("imposto")).length;
    vencStats.recurrence = errors.filter(e => e.type === "venc_inferido" && e.message.includes("recorrência")).length;
    vencStats.fallback = errors.filter(e => e.type === "venc_fallback").length;
    console.log(`[process-movimentacoes] Vencimento resolution: from_input=${vencStats.input}, tax_default=${vencStats.tax_default}, recurrence=${vencStats.recurrence}, fallback=${vencStats.fallback}`);

    console.log(`[process-movimentacoes] Output rows: ${outputRows.length}`);

    // 8) Fill template
    const templateBuffer = await templateFile.arrayBuffer();
    const templateWb = XLSX.read(new Uint8Array(templateBuffer), { type: "array" });

    const dadosSheet = templateWb.Sheets["Dados"];
    if (!dadosSheet) throw new Error("Aba 'Dados' não encontrada no template");

    // Headers at row 3 (0-indexed row 2), data starts at row 4 (0-indexed row 3)
    const DATA_START_ROW = 3;

    // Column mapping (0-indexed) based on template analysis
    const COL = {
      TIPO_PESSOA: 0,         // A
      CPF_CNPJ: 1,            // B
      RAZAO_SOCIAL: 2,        // C
      CEP: 3,                 // D
      NUMERO_ENDERECO: 4,     // E
      EMAIL: 5,               // F
      ID_INTEGRACAO: 6,       // G
      TIPO_DOCUMENTO: 7,      // H
      DESCRICAO: 8,           // I
      N_FATURA: 9,            // J
      DATA_COMPETENCIA: 10,   // K
      CATEGORIA: 11,          // L
      CENTRO_CUSTO: 12,       // M
      CONTA: 13,              // N
      TIPO_COBRANCA: 14,      // O
      TIPO_CHAVE_PIX: 15,     // P
      CHAVE_PIX: 16,          // Q
      CODIGO_REF: 17,         // R
      ORDEM_COMPRA: 18,       // S
      ORDEM_SERVICO: 19,      // T
      OBSERVACOES: 20,        // U
      QTD_REPETICOES: 21,     // V
      PERIODICIDADE: 22,      // W
      MANTER_COMPETENCIA: 23, // X
      VENCIMENTO: 24,         // Y
      VALOR: 25,              // Z
    };

    for (let i = 0; i < outputRows.length; i++) {
      const o = outputRows[i];
      const row = DATA_START_ROW + i;

      setCellValue(dadosSheet, COL.TIPO_PESSOA, row, o.tipoPessoa);
      if (o.documento) setCellValue(dadosSheet, COL.CPF_CNPJ, row, o.documento);
      setCellValue(dadosSheet, COL.RAZAO_SOCIAL, row, o.nome);
      if (o.cep) setCellValue(dadosSheet, COL.CEP, row, o.cep);
      if (o.numero) setCellValue(dadosSheet, COL.NUMERO_ENDERECO, row, o.numero);
      if (o.email) setCellValue(dadosSheet, COL.EMAIL, row, o.email);
      setCellValue(dadosSheet, COL.TIPO_DOCUMENTO, row, o.tipoDocumento);
      setCellValue(dadosSheet, COL.DESCRICAO, row, o.descricao);
      setCellValue(dadosSheet, COL.N_FATURA, row, o.nFatura);
      setCellValue(dadosSheet, COL.DATA_COMPETENCIA, row, o.dataCompetencia);
      setCellValue(dadosSheet, COL.CATEGORIA, row, o.categoria);
      if (o.centroCusto) setCellValue(dadosSheet, COL.CENTRO_CUSTO, row, o.centroCusto);
      setCellValue(dadosSheet, COL.CONTA, row, o.conta);
      setCellValue(dadosSheet, COL.TIPO_COBRANCA, row, o.tipoCobranca);
      if (o.tipoChavePix) setCellValue(dadosSheet, COL.TIPO_CHAVE_PIX, row, o.tipoChavePix);
      if (o.chavePix) setCellValue(dadosSheet, COL.CHAVE_PIX, row, o.chavePix);
      setCellValue(dadosSheet, COL.CODIGO_REF, row, o.codigoReferencia);
      if (o.observacoes) setCellValue(dadosSheet, COL.OBSERVACOES, row, o.observacoes);
      if (typeof o.qtdRepeticoes === "number" && o.qtdRepeticoes > 0) {
        setCellValue(dadosSheet, COL.QTD_REPETICOES, row, o.qtdRepeticoes);
      }
      if (o.periodicidade) setCellValue(dadosSheet, COL.PERIODICIDADE, row, o.periodicidade);
      setCellValue(dadosSheet, COL.VENCIMENTO, row, o.vencimento);
      setCellValue(dadosSheet, COL.VALOR, row, o.valor);
    }

    updateSheetRange(dadosSheet, DATA_START_ROW + outputRows.length, 36);

    // 9) Generate preview
    const previewRows: Record<string, string | number>[] = [];
    for (let i = 0; i < Math.min(outputRows.length, 20); i++) {
      const o = outputRows[i];
      previewRows.push({
        "Tipo Pessoa": o.tipoPessoa,
        "CPF/CNPJ": o.documento,
        "Razão Social": o.nome,
        "Descrição": o.descricao,
        "Categoria": o.categoria,
        "Conta": o.conta,
        "Cobrança": o.tipoCobranca,
        "Vencimento": o.vencimento,
        "Valor": formatCurrency(o.valor),
        "Repetições": o.qtdRepeticoes,
        "Periodicidade": o.periodicidade,
      });
    }

    // 10) Write output
    const outputBuffer = XLSX.write(templateWb, { type: "array", bookType: "xlsx" });
    const outputPath = `movimentacoes/${runId}.xlsx`;

    const { error: uploadOutputErr } = await supabase.storage
      .from("outputs")
      .upload(outputPath, new Uint8Array(outputBuffer), {
        contentType: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        upsert: true,
      });
    if (uploadOutputErr) throw new Error(`Upload output failed: ${uploadOutputErr.message}`);

    console.log(`[process-movimentacoes] Output saved: ${outputPath}`);

    // 11) Update run
    await supabase.from("runs").update({
      status: "done",
      preview_json: previewRows,
      error_report_json: errors.length > 0 ? errors : null,
      output_storage_path: outputPath,
      output_xlsx_path: outputPath,
    }).eq("id", runId);

    console.log(`[process-movimentacoes] Run ${runId} completed. ${outputRows.length} despesas, ${errors.length} errors`);

    return new Response(JSON.stringify({
      success: true,
      run_id: runId,
      total_despesas: outputRows.length,
      total_errors: errors.length,
      output_path: outputPath,
    }), {
      headers: { ...corsHeaders, "Content-Type": "application/json" },
    });
  } catch (e) {
    console.error("[process-movimentacoes] Error:", e);
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
