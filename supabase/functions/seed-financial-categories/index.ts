import { serve } from "https://deno.land/std@0.168.0/http/server.ts";
import { createClient } from "https://esm.sh/@supabase/supabase-js@2";
import * as XLSX from "https://esm.sh/xlsx@0.18.5";

const corsHeaders = {
  "Access-Control-Allow-Origin": "*",
  "Access-Control-Allow-Headers": "authorization, x-client-info, apikey, content-type",
};

const STOPWORDS = new Set(["de","da","do","das","dos","para","com","sem","e","a","o","as","os","em","no","na","nos","nas","por","um","uma","uns","umas","ao","aos","que","se","ou"]);

function normalizeTokens(text: string): string[] {
  const n = text.toLowerCase().normalize("NFD").replace(/[\u0300-\u036f]/g, "").replace(/[^a-z0-9\s]/g, " ").replace(/\s+/g, " ").trim();
  return [...new Set(n.split(" ").filter(t => t.length >= 3 && !STOPWORDS.has(t)))];
}

serve(async (req) => {
  if (req.method === "OPTIONS") return new Response(null, { headers: corsHeaders });

  try {
    const sb = createClient(Deno.env.get("SUPABASE_URL")!, Deno.env.get("SUPABASE_SERVICE_ROLE_KEY")!);

    // Try multiple filename conventions for the categories file
    const candidateNames = [
      "categorias_financeiras_2026-03-20.xlsx",
      "Categorias - 20-03-2026.xlsx",
      "categorias.xlsx",
      "Categorias.xlsx",
    ];
    let fileData: Blob | null = null;
    for (const name of candidateNames) {
      const { data, error } = await sb.storage.from("templates").download(name);
      if (!error && data) { fileData = data; break; }
    }
    if (!fileData) throw new Error(`Arquivo de categorias não encontrado no bucket 'templates'. Tentativas: ${candidateNames.join(", ")}`);

    const wb = XLSX.read(new Uint8Array(await fileData.arrayBuffer()), { type: "array" });
    const rows: any[] = XLSX.utils.sheet_to_json(wb.Sheets[wb.SheetNames[0]], { header: 1 });

    const LINE_RE = /^(\d{2}\.\d{2}\.\d{2})\s*-\s*(.+)$/;
    const records: any[] = [];

    for (const row of rows) {
      const cell = String(row[0] ?? "").trim();
      if (!cell) continue;
      const m = cell.match(LINE_RE);
      if (!m) continue;
      const code = m[1], name = m[2].trim();
      records.push({
        code, name, full_label: `${code} - ${name}`,
        tokens: normalizeTokens(name),
        updated_at: new Date().toISOString(),
      });
    }

    // Batch upsert
    const { error: upsertError } = await sb.from("financial_categories").upsert(records, { onConflict: "code" });
    if (upsertError) throw upsertError;

    const ignored = rows.length - records.length;
    console.log(`Seed: ${records.length} upserted, ${ignored} ignored`);

    return new Response(JSON.stringify({ success: true, upserted: records.length, ignored }), {
      headers: { ...corsHeaders, "Content-Type": "application/json" },
    });
  } catch (err) {
    console.error("Seed error:", err);
    return new Response(JSON.stringify({ success: false, error: (err as Error).message }), {
      status: 500, headers: { ...corsHeaders, "Content-Type": "application/json" },
    });
  }
});
