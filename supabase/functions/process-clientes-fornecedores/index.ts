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

function normalize(str: string): string {
  return removeAccents(str).toLowerCase().trim().replace(/[_\-\.\/]/g, " ").replace(/\s+/g, " ");
}

function cellToString(val: unknown): string {
  if (val == null) return "";
  if (typeof val === "number") {
    // Avoid scientific notation; preserve full digits
    return val.toFixed(0).replace(/\.0+$/, "");
  }
  return String(val).trim();
}

function extractDigits(val: string): string {
  return val.replace(/\D/g, "");
}

// ============ CPF/CNPJ VALIDATION WITH DV ============

function validateCPFDV(digits: string): boolean {
  if (digits.length !== 11) return false;
  if (/^(\d)\1{10}$/.test(digits)) return false;
  // First DV
  let sum = 0;
  for (let i = 0; i < 9; i++) sum += parseInt(digits[i]) * (10 - i);
  let rem = (sum * 10) % 11;
  if (rem === 10) rem = 0;
  if (rem !== parseInt(digits[9])) return false;
  // Second DV
  sum = 0;
  for (let i = 0; i < 10; i++) sum += parseInt(digits[i]) * (11 - i);
  rem = (sum * 10) % 11;
  if (rem === 10) rem = 0;
  return rem === parseInt(digits[10]);
}

function validateCNPJDV(digits: string): boolean {
  if (digits.length !== 14) return false;
  if (/^(\d)\1{13}$/.test(digits)) return false;
  const w1 = [5,4,3,2,9,8,7,6,5,4,3,2];
  let sum = 0;
  for (let i = 0; i < 12; i++) sum += parseInt(digits[i]) * w1[i];
  let rem = sum % 11;
  const dv1 = rem < 2 ? 0 : 11 - rem;
  if (dv1 !== parseInt(digits[12])) return false;
  const w2 = [6,5,4,3,2,9,8,7,6,5,4,3,2];
  sum = 0;
  for (let i = 0; i < 13; i++) sum += parseInt(digits[i]) * w2[i];
  rem = sum % 11;
  const dv2 = rem < 2 ? 0 : 11 - rem;
  return dv2 === parseInt(digits[13]);
}

interface DocResult {
  raw: string;
  digits: string;
  type: "Física" | "Jurídica" | "Não definido";
  dvValid: boolean;
}

function extractDocument(row: Record<string, unknown>, headers: string[], colDoc: string | null): DocResult {
  const result: DocResult = { raw: "", digits: "", type: "Não definido", dvValid: false };

  // Try the document column first
  const sources: string[] = [];
  if (colDoc) sources.push(cellToString(row[colDoc]));
  // Fallback: scan all cells for CPF/CNPJ patterns
  for (const h of headers) {
    if (h === colDoc) continue;
    sources.push(cellToString(row[h]));
  }

  for (const raw of sources) {
    if (!raw) continue;
    const digits = extractDigits(raw);
    // Pad CPF with leading zeros if needed (e.g. 9 digits from numeric cell)
    let padded = digits;
    if (padded.length >= 9 && padded.length <= 10) padded = padded.padStart(11, "0");
    if (padded.length >= 12 && padded.length <= 13) padded = padded.padStart(14, "0");

    if (padded.length === 11) {
      result.raw = raw;
      result.digits = padded;
      result.type = "Física";
      result.dvValid = validateCPFDV(padded);
      return result;
    }
    if (padded.length === 14) {
      result.raw = raw;
      result.digits = padded;
      result.type = "Jurídica";
      result.dvValid = validateCNPJDV(padded);
      return result;
    }
  }

  return result;
}

// ============ EMAIL EXTRACTION ============

function extractEmails(raw: string): string[] {
  if (!raw) return [];
  let text = raw.trim().toLowerCase();

  // Step 1: Replace common separators with space
  text = text.replace(/[;,|\\/\n\t]+/g, " ");

  // Step 2: Split concatenated emails at .com.br / .com boundaries.
  // Per spec all client emails end with .com.br or .com, so when one of those
  // TLDs is immediately followed by an alphanumeric character we know a new
  // token (possibly another email) is starting — insert a space.
  text = text.replace(/\.(com\.br|com)(?=[a-z0-9])/g, ".$1 ");

  // Step 3: Extract emails — only .com.br and .com TLDs (per spec)
  const re = /[a-z0-9._%+-]+@[a-z0-9.-]+\.(?:com\.br|com)\b/g;
  const results: string[] = [];
  let match: RegExpExecArray | null;
  while ((match = re.exec(text)) !== null) {
    const email = match[0];
    if (!results.includes(email)) results.push(email);
  }
  return results;
}

// ============ PHONE EXTRACTION ============

function extractPhones(raw: string): string[] {
  if (!raw) return [];
  const digitsOnly = String(raw).replace(/\D/g, " ");
  const sequences = digitsOnly.split(/\s+/).filter(Boolean);
  const results: string[] = [];
  for (const seq of sequences) {
    if (seq.length >= 10 && seq.length <= 13) {
      let phone = seq;
      if (phone.startsWith("55") && phone.length >= 12) {
        phone = phone.slice(2);
      }
      if (phone.length >= 10 && phone.length <= 11 && !results.includes(phone)) {
        results.push(phone);
      }
    }
  }
  return results;
}

function normalizePhone(val: unknown): string {
  if (val == null) return "";
  return extractDigits(cellToString(val));
}

// ============ COLUMN DETECTION ============

function findColumn(headers: string[], synonyms: string[]): string | null {
  for (const syn of synonyms) {
    const nSyn = normalize(syn);
    for (const h of headers) {
      const nH = normalize(h);
      if (nH === nSyn || nH.includes(nSyn) || nSyn.includes(nH)) {
        return h;
      }
    }
  }
  return null;
}

function findMultipleColumns(headers: string[], synonyms: string[]): string[] {
  const cols: string[] = [];
  for (const h of headers) {
    const nH = normalize(h);
    if (synonyms.some(s => {
      const ns = normalize(s);
      return nH.includes(ns) || ns.includes(nH);
    })) {
      if (!cols.includes(h)) cols.push(h);
    }
  }
  return cols;
}

function normalizeCEP(val: unknown): string {
  if (val == null) return "";
  const digits = extractDigits(cellToString(val));
  if (digits.length === 8) return digits;
  return "";
}

// ============ ADDRESS PARSING ============

interface ParsedAddress {
  cep: string;
  numero: string;
  complemento: string;
}

function parseAddress(raw: string): ParsedAddress {
  if (!raw || !raw.trim()) return { cep: "", numero: "", complemento: "" };

  let text = raw.trim();
  let cep = "";
  let numero = "";

  // Extract CEP: "CEP 74.343-340", "74343-340", "74343340"
  const cepRe = /(?:cep\s*[:\-]?\s*)?(\d{2}\.?\d{3}[\-.]?\d{3})/i;
  const cepMatch = text.match(cepRe);
  if (cepMatch) {
    const cepDigits = extractDigits(cepMatch[1]);
    if (cepDigits.length === 8) {
      cep = cepDigits;
      text = text.replace(cepMatch[0], " ").trim();
    }
  }

  // Extract Número: "nº 10", "numero 111", "número 03", "n° 111", "no 111"
  const numRe = /(?:n[ºo°]?|numero|número)\s*[:\-]?\s*(\d{1,6})/i;
  const numMatch = text.match(numRe);
  if (numMatch) {
    numero = numMatch[1];
    text = text.replace(numMatch[0], " ").trim();
  } else {
    // Fallback: first number after comma "Rua X, 123"
    const commaNumRe = /,\s*(\d{1,6})\b/;
    const commaMatch = text.match(commaNumRe);
    if (commaMatch) {
      numero = commaMatch[1];
      text = text.replace(commaMatch[0], ",").trim();
    }
  }

  // Complemento: remaining text minus CEP/Numero, cleaned up
  let complemento = text
    .replace(/\s{2,}/g, " ")
    .replace(/^[,;\s]+|[,;\s]+$/g, "")
    .trim();

  // Don't return the entire address as complemento if nothing was extracted
  if (complemento === raw.trim() && !cep && !numero) {
    // Nothing was extracted, keep full address as complemento for reference
  }

  return { cep, numero, complemento };
}

// ============ DATA EXTRACTION ============

interface ClienteRecord {
  documento: string;
  tipoPessoa: "Física" | "Jurídica" | "Não definido";
  dvValid: boolean;
  nome: string;
  cep: string;
  numero: string;
  complemento: string;
  telefones: string[];
  emails: string[];
  emitirNF: boolean;
}

interface ExtractError {
  documento: string;
  tipo: string;
  campo: string;
  mensagem: string;
}

function extractClientes(rawData: Record<string, unknown>[]): {
  clientes: ClienteRecord[];
  errors: ExtractError[];
  diagnostics: string[];
} {
  if (rawData.length === 0) return { clientes: [], errors: [], diagnostics: ["No rows in primary file"] };

  const headers = Object.keys(rawData[0]);
  const diagnostics: string[] = [];
  diagnostics.push(`Headers: ${headers.join(" | ")}`);

  const colDoc = findColumn(headers, ["cpf", "cnpj", "cpf/cnpj", "cpf_cnpj", "documento", "doc", "cpf cnpj"]);
  const colNome = findColumn(headers, ["razao social", "razao", "nome", "name", "empresa", "company", "cliente", "fornecedor", "razão social", "nome fantasia"]);
  const colCep = findColumn(headers, ["cep", "zip", "codigo postal", "código postal"]);
  const colNumero = findColumn(headers, ["numero", "número", "num", "nro", "number"]);
  const colComplemento = findColumn(headers, ["complemento", "compl"]);
  const colEndereco = findColumn(headers, ["endereco", "endereço", "logradouro", "rua", "address"]);

  const phoneSynonyms = ["telefone", "tel", "fone", "phone", "celular", "cel", "mobile", "whatsapp"];
  const phoneCols = findMultipleColumns(headers, phoneSynonyms);

  const emailSynonyms = ["email", "e-mail", "mail"];
  const emailCols = findMultipleColumns(headers, emailSynonyms);

  // Contact columns that may contain BOTH phones and emails mixed
  const contactSynonyms = ["enviar cobranca via", "enviar cobrança via", "cobranca", "cobrança", "contato"];
  const contactCols = findMultipleColumns(headers, contactSynonyms);

  const colDDD = findColumn(headers, ["ddd", "codigo area", "código área"]);
  const colEmitirNF = findColumn(headers, ["emitir nf", "emitir nf?", "emitir nota fiscal", "emitir nota", "nfse", "nf?"]);

  diagnostics.push(`Mapped: doc=${colDoc}, nome=${colNome}, cep=${colCep}, endereco=${colEndereco}, phones=[${phoneCols.join(",")}], emails=[${emailCols.join(",")}], contacts=[${contactCols.join(",")}], ddd=${colDDD}, emitirNF=${colEmitirNF}`);

  const byDoc = new Map<string, ClienteRecord>();
  const errors: ExtractError[] = [];
  let loggedCount = 0;

  for (let i = 0; i < rawData.length; i++) {
    const row = rawData[i];
    const lineNum = i + 2; // Excel line (header=1 or header row)

    // Extract document
    const doc = extractDocument(row, headers, colDoc);

    // Get nome
    let nome = "";
    if (colNome) nome = cellToString(row[colNome]);
    if (!nome) {
      for (const key of headers) {
        if (key === colDoc) continue;
        const val = cellToString(row[key]);
        if (val.length > 3 && /^[A-Za-zÀ-ú\s\.\-]+$/.test(val)) {
          nome = val;
          break;
        }
      }
    }

    if (!nome) {
      errors.push({ documento: doc.digits || `linha_${lineNum}`, tipo: "obrigatorio", campo: "Nome", mensagem: `Linha ${lineNum}: Nome não encontrado, registro ignorado` });
      continue;
    }

    // DV validation warning
    if (doc.digits && !doc.dvValid) {
      errors.push({ documento: doc.digits, tipo: "dv_invalido", campo: "CPF/CNPJ", mensagem: `Linha ${lineNum}: DV inválido para ${doc.type} ${doc.digits}. Inserido como "Não definido".` });
      doc.type = "Não definido";
    }

    // Phones
    const phones: string[] = [];
    const ddd = colDDD ? normalizePhone(row[colDDD]) : "";
    for (const pc of phoneCols) {
      const cellVal = cellToString(row[pc]);
      const extracted = extractPhones(cellVal);
      if (extracted.length > 0) {
        for (const p of extracted) { if (!phones.includes(p)) phones.push(p); }
      } else {
        const simple = normalizePhone(row[pc]);
        if (simple && simple.length >= 8) {
          const full = (ddd && simple.length <= 9) ? ddd + simple : simple;
          if (!phones.includes(full)) phones.push(full);
        }
      }
    }

    // Emails
    const emails: string[] = [];
    for (const ec of emailCols) {
      const cellVal = cellToString(row[ec]);
      const extracted = extractEmails(cellVal);
      for (const e of extracted) { if (!emails.includes(e)) emails.push(e); }
    }

    // Mixed contact columns (may contain both phones and emails separated by / ; , etc.)
    for (const cc of contactCols) {
      const cellVal = cellToString(row[cc]);
      if (!cellVal) continue;
      // Split by common separators first
      const parts = cellVal.split(/[\/;,|\\\n\t]+/);
      for (const part of parts) {
        const trimmed = part.trim();
        // Try emails
        const foundEmails = extractEmails(trimmed);
        for (const e of foundEmails) { if (!emails.includes(e)) emails.push(e); }
        // Try phones
        const foundPhones = extractPhones(trimmed);
        for (const p of foundPhones) { if (!phones.includes(p)) phones.push(p); }
        // If no email/phone found but looks like a phone (8+ digits)
        if (foundEmails.length === 0 && foundPhones.length === 0) {
          const simple = extractDigits(trimmed);
          if (simple.length >= 8 && simple.length <= 13) {
            let phone = simple;
            if (phone.startsWith("55") && phone.length >= 12) phone = phone.slice(2);
            if (phone.length >= 10 && phone.length <= 11 && !phones.includes(phone)) {
              phones.push(phone);
            }
          }
        }
      }
    }

    // Emitir NF
    let emitirNF = false;
    if (colEmitirNF) {
      const nfVal = cellToString(row[colEmitirNF]).toLowerCase().trim();
      emitirNF = ["sim", "s", "yes", "y", "1", "x", "true"].includes(nfVal);
    }

    // CEP, Numero, Complemento — from dedicated columns or parsed from Endereço
    let cep = colCep ? normalizeCEP(row[colCep]) : "";
    let numero = colNumero ? extractDigits(cellToString(row[colNumero])) : "";
    let complemento = colComplemento ? cellToString(row[colComplemento]) : "";

    // If dedicated columns didn't yield results, try parsing from Endereço column
    if (colEndereco && (!cep || !numero)) {
      const addrRaw = cellToString(row[colEndereco]);
      if (addrRaw) {
        const parsed = parseAddress(addrRaw);
        if (!cep && parsed.cep) cep = parsed.cep;
        if (!numero && parsed.numero) numero = parsed.numero;
        if (!complemento && parsed.complemento) complemento = parsed.complemento;
      }
    }

    // Diagnostic log (first 10)
    if (loggedCount < 10) {
      diagnostics.push(`[${lineNum}] nome="${nome}" doc="${doc.digits}" type=${doc.type} dv=${doc.dvValid} phones=${phones.length} emails=${emails.length} cep=${cep} num=${numero} compl="${complemento.substring(0, 30)}"`);
      loggedCount++;
    }

    // Consolidate by documento
    const key = doc.digits || `__nodoc_${i}`;
    if (doc.digits && byDoc.has(doc.digits)) {
      const existing = byDoc.get(doc.digits)!;
      if (nome.length > existing.nome.length) existing.nome = nome;
      for (const p of phones) { if (!existing.telefones.includes(p)) existing.telefones.push(p); }
      for (const e of emails) { if (!existing.emails.includes(e)) existing.emails.push(e); }
      if (!existing.cep && cep) existing.cep = cep;
      if (!existing.numero && numero) existing.numero = numero;
      if (!existing.complemento && complemento) existing.complemento = complemento;
      // If any row says to issue NF, keep true
      if (emitirNF) existing.emitirNF = true;
    } else {
      if (!doc.digits) {
        errors.push({ documento: `linha_${lineNum}`, tipo: "documento", campo: "CPF/CNPJ", mensagem: `Linha ${lineNum}: Documento (CPF/CNPJ) não identificado. Inserido como "Não definido".` });
      }
      byDoc.set(key, {
        documento: doc.digits,
        tipoPessoa: doc.type,
        dvValid: doc.dvValid,
        nome, cep, numero, complemento,
        telefones: phones,
        emails,
        emitirNF,
      });
    }
  }

  return { clientes: Array.from(byDoc.values()), errors, diagnostics };
}

// ============ FILL TEMPLATE ============

function setCellValue(sheet: XLSX.WorkSheet, col: number, row: number, value: string): void {
  const cellRef = XLSX.utils.encode_cell({ c: col, r: row });
  sheet[cellRef] = { t: "s", v: value };
}

function updateSheetRange(sheet: XLSX.WorkSheet, maxRow: number, maxCol: number): void {
  const range = XLSX.utils.decode_range(sheet["!ref"] || "A1");
  if (maxRow > range.e.r) range.e.r = maxRow - 1;
  if (maxCol - 1 > range.e.c) range.e.c = maxCol - 1;
  sheet["!ref"] = XLSX.utils.encode_range(range);
}

function fillTemplate(
  workbook: XLSX.WorkBook,
  clientes: ClienteRecord[]
): { dadosCount: number; contatosCount: number } {
  // --- ABA DADOS (sheet index 0) ---
  const dadosSheet = workbook.Sheets[workbook.SheetNames[0]];
  if (!dadosSheet) throw new Error("Aba 'Dados' não encontrada no template");

  // Headers row 3 (0-indexed row 2), data starts row 4 (0-indexed row 3)
  const DATA_START_ROW = 3;
  let dadosCount = 0;

  for (let i = 0; i < clientes.length; i++) {
    const c = clientes[i];
    const row = DATA_START_ROW + i;

    setCellValue(dadosSheet, 0, row, c.tipoPessoa);                          // A: Tipo de Pessoa
    setCellValue(dadosSheet, 1, row, c.documento);                            // B: CPF/CNPJ (always string)
    setCellValue(dadosSheet, 2, row, c.tipoPessoa === "Jurídica" ? "Sim" : "Não"); // C: Consultar RF
    setCellValue(dadosSheet, 3, row, c.nome);                                 // D: Razão Social
    setCellValue(dadosSheet, 4, row, c.nome);                                 // E: Nome Fantasia
    setCellValue(dadosSheet, 5, row, "Somente Cliente");                      // F: Cliente/Fornecedor
    setCellValue(dadosSheet, 6, row, "Ativo");                                // G: Status

    if (c.tipoPessoa === "Física") {
      if (c.cep) setCellValue(dadosSheet, 7, row, c.cep);
      if (c.numero) setCellValue(dadosSheet, 8, row, c.numero);
      if (c.complemento) setCellValue(dadosSheet, 9, row, c.complemento);
    }
    dadosCount++;
  }

  updateSheetRange(dadosSheet, DATA_START_ROW + clientes.length, 24);

  // --- ABA CONTATOS (sheet index 1) ---
  const contatosSheetName = workbook.SheetNames[1];
  if (!contatosSheetName) {
    console.log("[fill] No Contatos sheet found, skipping");
    return { dadosCount, contatosCount: 0 };
  }
  const contatosSheet = workbook.Sheets[contatosSheetName];
  if (!contatosSheet) return { dadosCount, contatosCount: 0 };

  // Headers row 1 (0-indexed row 0), data starts row 2 (0-indexed row 1)
  const CONTATOS_START_ROW = 1;
  let contatoRow = CONTATOS_START_ROW;
  let contatosCount = 0;

  for (const c of clientes) {
    if (c.telefones.length === 0 && c.emails.length === 0) continue;

    const maxLines = Math.max(c.telefones.length, c.emails.length, 1);
    for (let j = 0; j < maxLines; j++) {
      setCellValue(contatosSheet, 0, contatoRow, c.documento);     // A: CPF/CNPJ
      setCellValue(contatosSheet, 1, contatoRow, "Financeiro");    // B: Contato
      setCellValue(contatosSheet, 2, contatoRow, "Financeiro");    // C: Cargo
      // D: Data Nascimento — skip
      if (j < c.telefones.length) {
        setCellValue(contatosSheet, 4, contatoRow, c.telefones[j]); // E: Celular
      }
      if (j < c.emails.length) {
        setCellValue(contatosSheet, 5, contatoRow, c.emails[j]);   // F: E-mail
      }
      setCellValue(contatosSheet, 6, contatoRow, "Sim");                          // G: Enviar Boleto
      setCellValue(contatosSheet, 7, contatoRow, c.emitirNF ? "Sim" : "Não");   // H: Enviar NFSe
      contatoRow++;
      contatosCount++;
    }
  }

  updateSheetRange(contatosSheet, contatoRow, 8);
  return { dadosCount, contatosCount };
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

    console.log(`[process] Starting for upload: ${upload_id}`);

    // 1) Create run
    const { data: run, error: runError } = await supabase.from("runs").insert({
      upload_id,
      import_type: "clientes_fornecedores",
      status: "processing",
    }).select().single();
    if (runError || !run) throw new Error(`Failed to create run: ${runError?.message}`);
    const runId = run.id;
    console.log(`[process] Run created: ${runId}`);

    // 2) Fetch upload and download primary file
    const { data: upload, error: uploadErr } = await supabase
      .from("uploads").select("*").eq("id", upload_id).single();
    if (uploadErr || !upload) throw new Error(`Upload not found: ${uploadErr?.message}`);

    const filename = upload.original_filename || "";
    const ext = filename.split(".").pop()?.toLowerCase() || "xlsx";
    console.log(`[process] Primary file: ${filename} (ext: .${ext})`);

    const { data: primaryFile, error: dlErr } = await supabase.storage
      .from("uploads").download(upload.storage_path);
    if (dlErr || !primaryFile) throw new Error(`Download primary file failed: ${dlErr?.message}`);

    // 3) Fetch template
    const { data: tmpl, error: tmplErr } = await supabase
      .from("import_templates").select("*")
      .eq("import_type", "clientes_fornecedores").eq("is_default", true).single();
    if (tmplErr || !tmpl) throw new Error(`Template not found: ${tmplErr?.message}`);

    const { data: templateFile, error: tmplDlErr } = await supabase.storage
      .from("templates").download(tmpl.template_storage_path);
    if (tmplDlErr || !templateFile) throw new Error(`Download template failed: ${tmplDlErr?.message}`);
    console.log(`[process] Template: ${tmpl.template_storage_path}`);

    // 4) Read primary Excel (supports .xls and .xlsx via SheetJS)
    const primaryBuffer = await primaryFile.arrayBuffer();
    const primaryWb = XLSX.read(new Uint8Array(primaryBuffer), { type: "array" });
    const primarySheet = primaryWb.Sheets[primaryWb.SheetNames[0]];
    const rawData: Record<string, unknown>[] = XLSX.utils.sheet_to_json(primarySheet, { raw: false, defval: "" });

    if (rawData.length === 0) throw new Error("Planilha primária vazia");
    console.log(`[process] Primary data: ${rawData.length} rows, sheets: ${primaryWb.SheetNames.join(", ")}`);

    // 5) Extract and consolidate clientes
    const { clientes, errors, diagnostics } = extractClientes(rawData);
    for (const d of diagnostics) console.log(`[diag] ${d}`);
    console.log(`[process] Extracted ${clientes.length} clientes, ${errors.length} errors`);

    // 6) Read template workbook and fill
    const templateBuffer = await templateFile.arrayBuffer();
    const templateWb = XLSX.read(new Uint8Array(templateBuffer), { type: "array" });

    const { dadosCount, contatosCount } = fillTemplate(templateWb, clientes);
    console.log(`[process] Filled: ${dadosCount} rows in Dados, ${contatosCount} rows in Contatos`);

    // 7) Preview
    const previewRows: Record<string, string>[] = [];
    for (let i = 0; i < Math.min(clientes.length, 20); i++) {
      const c = clientes[i];
      previewRows.push({
        "Tipo de Pessoa": c.tipoPessoa,
        "CPF/CNPJ": c.documento,
        "DV Válido": c.dvValid ? "Sim" : "Não",
        "Razão Social": c.nome,
        "Telefones": c.telefones.join(", "),
        "Emails": c.emails.join(", "),
      });
    }

    // 8) Write output
    const outputBuffer = XLSX.write(templateWb, { type: "array", bookType: "xlsx" });
    const outputPath = `clientes_fornecedores/${runId}.xlsx`;

    const { error: uploadOutputErr } = await supabase.storage
      .from("outputs")
      .upload(outputPath, new Uint8Array(outputBuffer), {
        contentType: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        upsert: true,
      });
    if (uploadOutputErr) throw new Error(`Upload output failed: ${uploadOutputErr.message}`);

    // 9) Update run
    await supabase.from("runs").update({
      status: "done",
      preview_json: previewRows,
      error_report_json: errors,
      output_storage_path: outputPath,
      output_xlsx_path: outputPath,
    }).eq("id", runId);

    console.log(`[process] Run ${runId} completed. Dados=${dadosCount}, Contatos=${contatosCount}, Errors=${errors.length}`);

    return new Response(JSON.stringify({
      success: true,
      run_id: runId,
      total_clientes: clientes.length,
      total_dados: dadosCount,
      total_contatos: contatosCount,
      total_errors: errors.length,
      output_path: outputPath,
    }), {
      headers: { ...corsHeaders, "Content-Type": "application/json" },
    });
  } catch (e) {
    console.error("[process] Error:", e);
    const message = e instanceof Error ? e.message : "Unknown error";
    try {
      const body = await req.clone().json();
      if (body.upload_id) {
        const { data: runs } = await supabase.from("runs").select("id")
          .eq("upload_id", body.upload_id).eq("status", "processing").limit(1);
        if (runs && runs[0]) {
          await supabase.from("runs").update({
            status: "error",
            error_report_json: [{ documento: "system", tipo: "system", campo: "system", mensagem: message }],
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
