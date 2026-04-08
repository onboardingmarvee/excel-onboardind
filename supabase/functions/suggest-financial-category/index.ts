import { serve } from "https://deno.land/std@0.168.0/http/server.ts";
import { createClient } from "https://esm.sh/@supabase/supabase-js@2";

const corsHeaders = {
  "Access-Control-Allow-Origin": "*",
  "Access-Control-Allow-Headers": "authorization, x-client-info, apikey, content-type",
};

const STOPWORDS = new Set(["de","da","do","das","dos","para","com","sem","e","a","o","as","os","em","no","na","nos","nas","por","um","uma","uns","umas","ao","aos","que","se","ou"]);

function norm(text: string): string {
  return text.toLowerCase().normalize("NFD").replace(/[\u0300-\u036f]/g, "").replace(/[^a-z0-9\s]/g, " ").replace(/\s+/g, " ").trim();
}

function toTokens(text: string): string[] {
  const n = norm(text);
  return [...new Set(n.split(" ").filter(t => t.length >= 3 && !STOPWORDS.has(t)))];
}

function cleanCategoryQueryText(desc: string): string {
  let c = desc;
  c = c.replace(/^(PROVISÃO\s*-\s*|Cartão de Crédito\s*-\s*)/i, "");
  c = c.replace(/\b(provisao|provisão|cartao|cartão|credito|crédito)\b/gi, "");
  return c.replace(/\s+/g, " ").trim();
}

// --- ETAPA A: Overrides determinísticos ---

const MARKETING_SIGNALS = ["marketing","ads","google","meta","facebook","instagram","vendas","comissao","comissão","crm","leads"];
const ADM_SIGNALS = ["administrativo","adm","financeiro","contabil","contábil","rh","recursos humanos"];

function hasAny(text: string, terms: string[]): boolean {
  const n = norm(text);
  return terms.some(t => n.includes(norm(t)));
}

function hasAnyToken(tokens: string[], terms: string[]): boolean {
  const normTerms = terms.map(t => norm(t));
  return tokens.some(tk => normTerms.some(nt => tk.includes(nt) || nt.includes(tk)));
}

function matchesAny(raw: string, cleaned: string, tokens: string[], terms: string[]): boolean {
  return hasAny(raw, terms) || hasAny(cleaned, terms) || hasAnyToken(tokens, terms);
}

function detectContext(raw: string, cleaned: string, tokens: string[]): "marketing" | "adm" | "default" {
  if (matchesAny(raw, cleaned, tokens, MARKETING_SIGNALS)) return "marketing";
  if (matchesAny(raw, cleaned, tokens, ADM_SIGNALS)) return "adm";
  return "default";
}

interface Override { code: string; name: string; reason: string }

function ruleBasedCategoryOverride(raw: string, cleaned: string, tokens: string[]): Override | null {
  // 1) Financeiro por assinatura -> Assessoria Financeira
  if (hasAny(raw, ["financeiro por assinatura"]) || hasAny(cleaned, ["financeiro por assinatura"])) {
    return { code: "02.06.01", name: "Assessoria Financeira", reason: "override_financeiro_assinatura" };
  }

  // 2) Remuneração / folha / colaborador
  const remuTerms = ["pro-labore","pro labore","salario","salário","folha","clt","pj","pagamento colaborador","remuneracao","remuneração","honorarios","honorários"];
  if (matchesAny(raw, cleaned, tokens, remuTerms)) {
    const ctx = detectContext(raw, cleaned, tokens);
    if (ctx === "marketing") return { code: "02.03.04", name: "Salários e Remunerações Marketing e Vendas", reason: "override_remuneracao" };
    if (ctx === "adm") return { code: "02.03.02", name: "Salários e Remunerações Administrativo", reason: "override_remuneracao" };
    return { code: "02.03.03", name: "Salários e Remunerações Operacional", reason: "override_remuneracao" };
  }

  // 3) Ferramentas / software
  const ferrTerms = ["ferramenta","software","licenca","licença","assinatura","saas","plataforma","sistema","app","aplicativo"];
  if (matchesAny(raw, cleaned, tokens, ferrTerms)) {
    const ctx = detectContext(raw, cleaned, tokens);
    if (ctx === "marketing") return { code: "02.05.02", name: "Ferramentas, Softwares e Sistemas - Comercial e Marketing", reason: "override_ferramentas" };
    if (ctx === "adm") return { code: "02.06.11", name: "Ferramentas, Softwares e Sistemas - Administrativo/Financeiro", reason: "override_ferramentas" };
    return { code: "02.02.06", name: "Ferramentas, Softwares e Sistemas - Prestação de Serviço", reason: "override_ferramentas" };
  }

  // 4) Alvará
  if (matchesAny(raw, cleaned, tokens, ["alvara","alvará"])) {
    return { code: "02.06.18", name: "Alvarás", reason: "override_alvara" };
  }

  // 5) Veículo / carro
  if (matchesAny(raw, cleaned, tokens, ["carro","veiculo","veículo","automovel","automóvel","parcela carro","financiamento carro"])) {
    return { code: "04.02.04", name: "Compra de Veículos", reason: "override_veiculo" };
  }

  return null;
}

// --- Main handler ---

serve(async (req) => {
  if (req.method === "OPTIONS") {
    return new Response(null, { headers: corsHeaders });
  }

  try {
    const { description } = await req.json();
    if (!description || typeof description !== "string") {
      return new Response(JSON.stringify({ error: "description is required" }), {
        status: 400, headers: { ...corsHeaders, "Content-Type": "application/json" },
      });
    }

    const raw = description;
    const cleaned = cleanCategoryQueryText(raw);
    const tokens = toTokens(cleaned);

    // ETAPA A: Override
    const override = ruleBasedCategoryOverride(raw, cleaned, tokens);
    if (override) {
      return new Response(JSON.stringify({
        code: override.code,
        name: override.name,
        score: 1,
        top: [{ code: override.code, name: override.name, score: 1 }],
        strategy: "override",
        reason: override.reason,
        cleaned_description: cleaned,
      }), { headers: { ...corsHeaders, "Content-Type": "application/json" } });
    }

    // ETAPA B: Similaridade (fallback)
    if (tokens.length === 0) {
      return new Response(JSON.stringify({
        code: null, name: null, score: 0, top: [],
        strategy: "similarity", reason: null, cleaned_description: cleaned,
      }), { headers: { ...corsHeaders, "Content-Type": "application/json" } });
    }

    const supabaseUrl = Deno.env.get("SUPABASE_URL")!;
    const serviceKey = Deno.env.get("SUPABASE_SERVICE_ROLE_KEY")!;
    const sb = createClient(supabaseUrl, serviceKey);

    const { data: candidates, error } = await sb
      .from("financial_categories")
      .select("code, name, full_label, tokens")
      .overlaps("tokens", tokens);

    if (error) throw error;

    const codeMatch = raw.match(/\d{2}\.\d{2}\.\d{2}/);

    const scored = (candidates || []).map((c: any) => {
      const catTokens: string[] = c.tokens;
      const intersection = tokens.filter(t => catTokens.includes(t)).length;
      let score = (2 * intersection) / (catTokens.length + tokens.length);
      if (codeMatch && c.code === codeMatch[0]) score += 0.3;
      return { code: c.code, name: c.name, score: Math.round(score * 1000) / 1000 };
    });

    scored.sort((a: any, b: any) => b.score - a.score);
    const top = scored.slice(0, 5);
    const best = top[0] || { code: null, name: null, score: 0 };

    // ETAPA C: Guardrails
    const warning = best.score > 0 && best.score < 0.25 ? "categoria_suspeita" : undefined;

    return new Response(JSON.stringify({
      code: best.code,
      name: best.name,
      score: best.score,
      top,
      strategy: "similarity",
      reason: warning || null,
      cleaned_description: cleaned,
    }), { headers: { ...corsHeaders, "Content-Type": "application/json" } });
  } catch (err) {
    console.error("Suggest error:", err);
    return new Response(JSON.stringify({ error: err.message }), {
      status: 500, headers: { ...corsHeaders, "Content-Type": "application/json" },
    });
  }
});
