import { serve } from "https://deno.land/std@0.168.0/http/server.ts";
import { createClient } from "https://esm.sh/@supabase/supabase-js@2";
import * as XLSX from "xlsx";

const corsHeaders = {
  "Access-Control-Allow-Origin": "*",
  "Access-Control-Allow-Headers":
    "authorization, x-client-info, apikey, content-type, x-supabase-client-platform, x-supabase-client-platform-version, x-supabase-client-runtime, x-supabase-client-runtime-version",
};

// ============ TEMPLATE PROFILES (mirrors frontend) ============

interface TemplateColumn {
  name: string;
  required: boolean;
  type: "text" | "number" | "date" | "phone" | "currency" | "email" | "cpf_cnpj" | "list" | "integer";
  aliases: string[];
}

const TEMPLATE_PROFILES: Record<string, TemplateColumn[]> = {
  clientes_fornecedores: [
    { name: "Tipo de Pessoa", required: true, type: "list", aliases: ["tipo pessoa", "tipo de pessoa", "person type", "tipo"] },
    { name: "CPF/CNPJ", required: false, type: "cpf_cnpj", aliases: ["cpf", "cnpj", "cpf/cnpj", "cpf_cnpj", "documento", "doc"] },
    { name: "Consultar Receita Federal?", required: false, type: "list", aliases: ["consultar receita", "receita federal", "consulta rf"] },
    { name: "Razão Social", required: true, type: "text", aliases: ["razao social", "razao", "nome", "name", "empresa", "company", "razão social"] },
    { name: "Nome Fantasia", required: false, type: "text", aliases: ["nome fantasia", "fantasia", "trade name"] },
    { name: "Cliente/Fornecedor", required: true, type: "list", aliases: ["cliente fornecedor", "cliente/fornecedor", "tipo cadastro"] },
    { name: "Status", required: false, type: "list", aliases: ["status", "situacao", "estado"] },
    { name: "CEP", required: false, type: "text", aliases: ["cep", "zip", "codigo postal"] },
    { name: "Número", required: false, type: "integer", aliases: ["numero", "num", "nro", "number"] },
    { name: "Complemento", required: false, type: "text", aliases: ["complemento", "compl"] },
    { name: "DDD", required: false, type: "integer", aliases: ["ddd", "codigo area"] },
    { name: "Telefone", required: false, type: "text", aliases: ["telefone", "tel", "fone", "phone"] },
    { name: "E-mail", required: false, type: "email", aliases: ["email", "e-mail", "mail"] },
    { name: "Tipo Entidade", required: false, type: "list", aliases: ["tipo entidade", "entidade", "entity type"] },
    { name: "Regime Tributário", required: false, type: "list", aliases: ["regime tributario", "regime tributário", "tax regime"] },
    { name: "Modalidade de Lucro", required: false, type: "list", aliases: ["modalidade lucro", "modalidade de lucro", "profit mode"] },
    { name: "Regime Especial", required: false, type: "list", aliases: ["regime especial", "special regime"] },
    { name: "Inscrição Municipal", required: false, type: "text", aliases: ["inscricao municipal", "inscrição municipal", "im"] },
    { name: "Perfil de Tributação", required: false, type: "list", aliases: ["perfil tributacao", "perfil de tributação", "tax profile"] },
    { name: "CNAE Principal (Código)", required: false, type: "text", aliases: ["cnae", "cnae principal", "codigo cnae"] },
    { name: "Data Abertura/Nascimento", required: false, type: "date", aliases: ["data abertura", "data nascimento", "abertura", "nascimento", "birthday"] },
    { name: "Tipo Chave Pix", required: false, type: "list", aliases: ["tipo chave pix", "tipo pix", "pix type"] },
    { name: "Chave Pix", required: false, type: "text", aliases: ["chave pix", "pix", "pix key"] },
    { name: "Id Integração", required: false, type: "text", aliases: ["id integracao", "id integração", "integration id"] },
  ],
  movimentacoes: [
    { name: "Data", required: true, type: "date", aliases: ["data", "dt", "date", "data movimento"] },
    { name: "Tipo", required: true, type: "text", aliases: ["tipo", "type", "natureza"] },
    { name: "Descricao", required: true, type: "text", aliases: ["descricao", "desc", "historico", "description"] },
    { name: "Valor", required: true, type: "currency", aliases: ["valor", "value", "amount", "vlr"] },
    { name: "ContaDebito", required: false, type: "text", aliases: ["conta debito", "debito", "debit"] },
    { name: "ContaCredito", required: false, type: "text", aliases: ["conta credito", "credito", "credit"] },
    { name: "CentroCusto", required: false, type: "text", aliases: ["centro de custo", "centro custo", "cc"] },
    { name: "Documento", required: false, type: "text", aliases: ["documento", "doc", "nf", "nota fiscal"] },
    { name: "Observacao", required: false, type: "text", aliases: ["observacao", "obs", "nota"] },
  ],
  vendas: [
    { name: "NumeroVenda", required: true, type: "text", aliases: ["numero venda", "num venda", "nv", "pedido", "order"] },
    { name: "DataVenda", required: true, type: "date", aliases: ["data venda", "data", "dt venda", "date"] },
    { name: "Cliente", required: true, type: "text", aliases: ["cliente", "customer", "comprador"] },
    { name: "CPFCNPJCliente", required: false, type: "cpf_cnpj", aliases: ["cpf cliente", "cnpj cliente", "cpf/cnpj", "documento"] },
    { name: "Produto", required: true, type: "text", aliases: ["produto", "item", "product", "descricao"] },
    { name: "Quantidade", required: true, type: "number", aliases: ["quantidade", "qtd", "qty", "quant"] },
    { name: "ValorUnitario", required: true, type: "currency", aliases: ["valor unitario", "preco", "unit", "price"] },
    { name: "ValorTotal", required: true, type: "currency", aliases: ["valor total", "total", "subtotal"] },
    { name: "Desconto", required: false, type: "currency", aliases: ["desconto", "discount", "desc"] },
    { name: "FormaPagamento", required: false, type: "text", aliases: ["forma pagamento", "pagamento", "payment"] },
    { name: "Observacao", required: false, type: "text", aliases: ["observacao", "obs", "nota"] },
  ],
  contratos: [
    { name: "NumeroContrato", required: true, type: "text", aliases: ["numero contrato", "num contrato", "contrato", "contract"] },
    { name: "Cliente", required: true, type: "text", aliases: ["cliente", "contratante", "customer"] },
    { name: "CPFCNPJCliente", required: false, type: "cpf_cnpj", aliases: ["cpf", "cnpj", "cpf/cnpj", "documento"] },
    { name: "DataInicio", required: true, type: "date", aliases: ["data inicio", "inicio", "start", "dt inicio"] },
    { name: "DataFim", required: false, type: "date", aliases: ["data fim", "fim", "end", "dt fim", "vencimento"] },
    { name: "ValorMensal", required: true, type: "currency", aliases: ["valor mensal", "mensalidade", "monthly"] },
    { name: "ValorTotal", required: false, type: "currency", aliases: ["valor total", "total"] },
    { name: "Descricao", required: true, type: "text", aliases: ["descricao", "objeto", "description"] },
    { name: "Status", required: false, type: "text", aliases: ["status", "situacao", "estado"] },
    { name: "Observacao", required: false, type: "text", aliases: ["observacao", "obs", "nota"] },
  ],
};

// ============ HELPERS ============

function removeAccents(str: string): string {
  return str.normalize("NFD").replace(/[\u0300-\u036f]/g, "");
}

function normalize(str: string): string {
  return removeAccents(str).toLowerCase().trim().replace(/[_\-.]/g, " ");
}

function mapColumns(sourceHeaders: string[], template: TemplateColumn[]): Record<string, string | null> {
  const mapping: Record<string, string | null> = {};
  const normalizedSource = sourceHeaders.map((h) => ({ original: h, normalized: normalize(h) }));

  for (const col of template) {
    const normalizedName = normalize(col.name);
    let matched: string | null = null;

    // Try exact match first
    for (const src of normalizedSource) {
      if (src.normalized === normalizedName) {
        matched = src.original;
        break;
      }
    }

    // Try aliases
    if (!matched) {
      for (const alias of col.aliases) {
        const na = normalize(alias);
        for (const src of normalizedSource) {
          if (src.normalized === na || src.normalized.includes(na) || na.includes(src.normalized)) {
            matched = src.original;
            break;
          }
        }
        if (matched) break;
      }
    }

    mapping[col.name] = matched;
  }

  return mapping;
}

function normalizePhone(val: string): string {
  return val.replace(/\D/g, "");
}

function normalizeCurrency(val: string): string {
  // Remove R$, spaces, dots (thousands), keep comma as decimal
  let cleaned = val.replace(/[R$\s]/g, "").trim();
  // If has dot and comma: 1.234,56 -> 1234.56
  if (cleaned.includes(",") && cleaned.includes(".")) {
    cleaned = cleaned.replace(/\./g, "").replace(",", ".");
  } else if (cleaned.includes(",")) {
    cleaned = cleaned.replace(",", ".");
  }
  const num = parseFloat(cleaned);
  return isNaN(num) ? "" : num.toFixed(2);
}

function normalizeDate(val: unknown): string {
  if (val == null) return "";
  // If it's a JS serial date number from Excel
  if (typeof val === "number") {
    try {
      const date = XLSX.SSF.parse_date_code(val);
      if (date) {
        const y = String(date.y).padStart(4, "0");
        const m = String(date.m).padStart(2, "0");
        const d = String(date.d).padStart(2, "0");
        return `${d}/${m}/${y}`;
      }
    } catch { /* fall through */ }
  }
  const str = String(val).trim();
  // Try DD/MM/YYYY already
  if (/^\d{2}\/\d{2}\/\d{4}$/.test(str)) return str;
  // Try YYYY-MM-DD
  const iso = str.match(/^(\d{4})-(\d{2})-(\d{2})/);
  if (iso) return `${iso[3]}/${iso[2]}/${iso[1]}`;
  // Try MM/DD/YYYY
  const us = str.match(/^(\d{2})\/(\d{2})\/(\d{4})$/);
  if (us) return `${us[2]}/${us[1]}/${us[3]}`;
  return str;
}

function normalizeValue(val: unknown, type: string): string {
  if (val == null || val === "") return "";
  const str = String(val).trim();
  switch (type) {
    case "phone":
      return normalizePhone(str);
    case "currency":
      return normalizeCurrency(str);
    case "date":
      return normalizeDate(val);
    case "number": {
      const n = parseFloat(str.replace(",", "."));
      return isNaN(n) ? str : String(n);
    }
    case "integer": {
      const digits = str.replace(/\D/g, "");
      return digits || str;
    }
    case "cpf_cnpj":
      return str.replace(/\D/g, "");
    case "email":
      return str.toLowerCase().trim();
    case "list":
    default:
      return str;
  }
}

function validateValue(val: string, col: TemplateColumn, rowIdx: number): string | null {
  if (col.required && !val) {
    return `Campo obrigatório '${col.name}' vazio`;
  }
  if (!val) return null;

  switch (col.type) {
    case "email":
      if (!val.includes("@")) return `Email inválido '${val}'`;
      break;
    case "number":
    case "currency":
      if (isNaN(parseFloat(val))) return `Valor numérico inválido '${val}'`;
      break;
    case "date":
      if (!/^\d{2}\/\d{2}\/\d{4}$/.test(val) && val.length > 0)
        return `Formato de data inválido '${val}' (esperado DD/MM/AAAA)`;
      break;
    case "cpf_cnpj":
      if (val.length !== 11 && val.length !== 14)
        return `CPF/CNPJ com tamanho inválido (${val.length} dígitos)`;
      break;
  }
  return null;
}

// ============ AI TRANSFORM RULES ============

interface TransformRules {
  date_format?: string;
  decimal_separator?: string;
  strip_masks?: boolean;
  defaults?: Record<string, string>;
  uppercase_fields?: string[];
}

async function getTransformRules(instructions: string): Promise<TransformRules> {
  const LOVABLE_API_KEY = Deno.env.get("LOVABLE_API_KEY");
  if (!LOVABLE_API_KEY) {
    console.log("No LOVABLE_API_KEY, skipping AI transform");
    return {};
  }

  try {
    const response = await fetch("https://ai.gateway.lovable.dev/v1/chat/completions", {
      method: "POST",
      headers: {
        Authorization: `Bearer ${LOVABLE_API_KEY}`,
        "Content-Type": "application/json",
      },
      body: JSON.stringify({
        model: "google/gemini-3-flash-preview",
        messages: [
          {
            role: "system",
            content: `You extract transform rules from user instructions. Return ONLY a JSON object with these optional keys:
- date_format: target date format string (e.g. "DD/MM/YYYY")
- decimal_separator: "." or ","
- strip_masks: boolean, whether to remove formatting masks
- defaults: object mapping field names to default values
- uppercase_fields: array of field names to uppercase
No other keys allowed. If instruction is unclear, return {}.`,
          },
          { role: "user", content: instructions },
        ],
        tools: [
          {
            type: "function",
            function: {
              name: "set_transform_rules",
              description: "Set transformation rules",
              parameters: {
                type: "object",
                properties: {
                  date_format: { type: "string" },
                  decimal_separator: { type: "string", enum: [".", ","] },
                  strip_masks: { type: "boolean" },
                  defaults: {
                    type: "object",
                    additionalProperties: { type: "string" },
                  },
                  uppercase_fields: {
                    type: "array",
                    items: { type: "string" },
                  },
                },
                additionalProperties: false,
              },
            },
          },
        ],
        tool_choice: { type: "function", function: { name: "set_transform_rules" } },
      }),
    });

    if (!response.ok) {
      console.error("AI gateway error:", response.status);
      return {};
    }

    const data = await response.json();
    const toolCall = data.choices?.[0]?.message?.tool_calls?.[0];
    if (toolCall?.function?.arguments) {
      return JSON.parse(toolCall.function.arguments) as TransformRules;
    }
  } catch (e) {
    console.error("AI transform error:", e);
  }
  return {};
}

function applyTransformRules(row: Record<string, string>, rules: TransformRules): Record<string, string> {
  const result = { ...row };

  if (rules.defaults) {
    for (const [field, defaultVal] of Object.entries(rules.defaults)) {
      if (field in result && !result[field]) {
        result[field] = defaultVal;
      }
    }
  }

  if (rules.uppercase_fields) {
    for (const field of rules.uppercase_fields) {
      if (field in result && result[field]) {
        result[field] = result[field].toUpperCase();
      }
    }
  }

  return result;
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
    const { run_id } = await req.json();
    if (!run_id) throw new Error("run_id is required");

    console.log(`[process-excel] Starting run: ${run_id}`);

    // Fetch run
    const { data: run, error: runError } = await supabase
      .from("runs")
      .select("*, uploads(*)")
      .eq("id", run_id)
      .single();
    if (runError || !run) throw new Error(`Run not found: ${runError?.message}`);

    // Update status
    await supabase.from("runs").update({ status: "processing" }).eq("id", run_id);
    await supabase.from("uploads").update({ status: "processing" }).eq("id", run.upload_id);

    const upload = run.uploads as { storage_path: string; original_filename: string };
    const template = TEMPLATE_PROFILES[run.import_type];
    if (!template) throw new Error(`Unknown import type: ${run.import_type}`);

    console.log(`[process-excel] Template: ${run.import_type}, file: ${upload.original_filename}`);

    // Download file from storage
    const { data: fileData, error: dlError } = await supabase.storage
      .from("uploads")
      .download(upload.storage_path);
    if (dlError || !fileData) throw new Error(`Download failed: ${dlError?.message}`);

    // Read Excel
    const arrayBuffer = await fileData.arrayBuffer();
    const workbook = XLSX.read(new Uint8Array(arrayBuffer), { type: "array" });
    const sheetName = workbook.SheetNames[0];
    const sheet = workbook.Sheets[sheetName];
    const rawData: Record<string, unknown>[] = XLSX.utils.sheet_to_json(sheet);

    if (rawData.length === 0) throw new Error("Planilha vazia");

    const sourceHeaders = Object.keys(rawData[0]);
    console.log(`[process-excel] Source columns: ${sourceHeaders.join(", ")}`);

    // Map columns
    const columnMapping = mapColumns(sourceHeaders, template);
    console.log(`[process-excel] Column mapping:`, JSON.stringify(columnMapping));

    // Get AI transform rules if instructions provided
    let transformRules: TransformRules = {};
    if (run.transform_instructions) {
      console.log(`[process-excel] Getting AI transform rules...`);
      transformRules = await getTransformRules(run.transform_instructions);
      console.log(`[process-excel] Transform rules:`, JSON.stringify(transformRules));
    }

    // Process rows
    const outputRows: Record<string, string>[] = [];
    const errors: { row: number; field: string; message: string }[] = [];

    for (let i = 0; i < rawData.length; i++) {
      const sourceRow = rawData[i];
      const outputRow: Record<string, string> = {};

      for (const col of template) {
        const sourceCol = columnMapping[col.name];
        const rawVal = sourceCol ? sourceRow[sourceCol] : undefined;
        let val = normalizeValue(rawVal, col.type);

        outputRow[col.name] = val;

        // Validate
        const err = validateValue(val, col, i + 2); // +2 for 1-based + header row
        if (err) {
          errors.push({ row: i + 2, field: col.name, message: err });
        }
      }

      // Apply transform rules
      const finalRow = applyTransformRules(outputRow, transformRules);
      outputRows.push(finalRow);
    }

    console.log(`[process-excel] Processed ${outputRows.length} rows, ${errors.length} errors`);

    // Generate preview (first 20 rows)
    const preview = outputRows.slice(0, 20);

    // Generate output XLSX
    const templateHeaders = template.map((c) => c.name);
    const outputWs = XLSX.utils.json_to_sheet(outputRows, { header: templateHeaders });
    const outputWb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(outputWb, outputWs, "Dados");
    const xlsxBuffer = XLSX.write(outputWb, { type: "array", bookType: "xlsx" });

    // Generate CSV
    const csvContent = XLSX.utils.sheet_to_csv(outputWs);
    const csvEncoder = new TextEncoder();
    const csvBuffer = csvEncoder.encode(csvContent);

    // Upload outputs
    const basePath = `${run_id}`;
    const xlsxPath = `${basePath}/output.xlsx`;
    const csvPath = `${basePath}/output.csv`;

    await supabase.storage.from("outputs").upload(xlsxPath, new Uint8Array(xlsxBuffer), {
      contentType: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
      upsert: true,
    });

    await supabase.storage.from("outputs").upload(csvPath, csvBuffer, {
      contentType: "text/csv",
      upsert: true,
    });

    // Update run
    await supabase.from("runs").update({
      status: "done",
      preview_json: preview,
      error_report_json: errors,
      output_xlsx_path: xlsxPath,
      output_csv_path: csvPath,
    }).eq("id", run_id);

    await supabase.from("uploads").update({ status: "done" }).eq("id", run.upload_id);

    console.log(`[process-excel] Done! Output saved.`);

    return new Response(JSON.stringify({ success: true, errors: errors.length }), {
      headers: { ...corsHeaders, "Content-Type": "application/json" },
    });
  } catch (e) {
    console.error("[process-excel] Error:", e);
    const message = e instanceof Error ? e.message : "Unknown error";

    // Try to update run status
    try {
      const { run_id } = await req.clone().json();
      if (run_id) {
        await supabase.from("runs").update({
          status: "error",
          error_report_json: [{ row: 0, field: "system", message }],
        }).eq("id", run_id);
      }
    } catch { /* ignore */ }

    return new Response(JSON.stringify({ error: message }), {
      status: 500,
      headers: { ...corsHeaders, "Content-Type": "application/json" },
    });
  }
});
