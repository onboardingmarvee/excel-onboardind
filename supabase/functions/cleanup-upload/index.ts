import { serve } from "https://deno.land/std@0.168.0/http/server.ts";
import { createClient } from "https://esm.sh/@supabase/supabase-js@2";

const corsHeaders = {
  "Access-Control-Allow-Origin": "*",
  "Access-Control-Allow-Headers":
    "authorization, x-client-info, apikey, content-type, x-supabase-client-platform, x-supabase-client-platform-version, x-supabase-client-runtime, x-supabase-client-runtime-version",
};

async function cleanupSingleUpload(supabase: any, uid: string) {
  const { data: upload, error: uploadErr } = await supabase
    .from("uploads")
    .select("*")
    .eq("id", uid)
    .single();

  if (uploadErr || !upload) {
    console.warn(`[cleanup] Upload not found: ${uid}`);
    return;
  }

  const storagePath: string = upload.storage_path;
  if (storagePath.startsWith("templates/") || storagePath.includes("/templates/")) {
    console.warn(`[cleanup] Skipping template path: ${storagePath}`);
    return;
  }

  // Find related runs
  const { data: runs } = await supabase
    .from("runs")
    .select("id, output_storage_path, output_xlsx_path, output_csv_path")
    .eq("upload_id", uid);

  // Delete output files from 'outputs' bucket
  if (runs && runs.length > 0) {
    const outputPaths: string[] = [];
    for (const run of runs) {
      if (run.output_storage_path) outputPaths.push(run.output_storage_path);
      if (run.output_xlsx_path && run.output_xlsx_path !== run.output_storage_path) {
        outputPaths.push(run.output_xlsx_path);
      }
      if (run.output_csv_path) outputPaths.push(run.output_csv_path);
    }

    const safePaths = outputPaths.filter(
      (p) => !p.startsWith("templates/") && !p.includes("/templates/")
    );
    if (safePaths.length > 0) {
      const { error: delOutputErr } = await supabase.storage
        .from("outputs")
        .remove(safePaths);
      if (delOutputErr) console.warn(`[cleanup] Failed to delete outputs: ${delOutputErr.message}`);
      else console.log(`[cleanup] Deleted ${safePaths.length} output file(s)`);
    }

    const runIds = runs.map((r: any) => r.id);
    const { error: delRunsErr } = await supabase.from("runs").delete().in("id", runIds);
    if (delRunsErr) console.warn(`[cleanup] Failed to delete runs: ${delRunsErr.message}`);
    else console.log(`[cleanup] Deleted ${runIds.length} run(s)`);
  }

  // Delete primary file
  const { error: delFileErr } = await supabase.storage.from("uploads").remove([storagePath]);
  if (delFileErr) console.warn(`[cleanup] Failed to delete primary file: ${delFileErr.message}`);
  else console.log(`[cleanup] Deleted primary file: ${storagePath}`);

  // Delete upload record
  const { error: delUploadErr } = await supabase.from("uploads").delete().eq("id", uid);
  if (delUploadErr) console.warn(`[cleanup] Failed to delete upload record: ${delUploadErr.message}`);
  else console.log(`[cleanup] Deleted upload record: ${uid}`);
}

serve(async (req) => {
  if (req.method === "OPTIONS") {
    return new Response(null, { headers: corsHeaders });
  }

  const supabase = createClient(
    Deno.env.get("SUPABASE_URL")!,
    Deno.env.get("SUPABASE_SERVICE_ROLE_KEY")!
  );

  try {
    const { upload_id, clean_all } = await req.json();

    let uploadIds: string[] = [];

    if (clean_all) {
      const { data: allUploads, error: allErr } = await supabase
        .from("uploads")
        .select("id, storage_path");
      if (allErr) throw new Error(`Failed to fetch uploads: ${allErr.message}`);
      if (!allUploads || allUploads.length === 0) {
        return new Response(
          JSON.stringify({ success: true, message: "Nenhum upload para limpar" }),
          { headers: { ...corsHeaders, "Content-Type": "application/json" } }
        );
      }
      const safe = allUploads.filter(
        (u: any) => !u.storage_path.startsWith("templates/") && !u.storage_path.includes("/templates/")
      );
      uploadIds = safe.map((u: any) => u.id);
      console.log(`[cleanup] clean_all: found ${uploadIds.length} upload(s)`);
    } else if (upload_id) {
      uploadIds = [upload_id];
    } else {
      throw new Error("upload_id or clean_all is required");
    }

    for (const uid of uploadIds) {
      await cleanupSingleUpload(supabase, uid);
    }

    console.log(`[cleanup] Cleanup complete. Processed ${uploadIds.length} upload(s).`);

    return new Response(
      JSON.stringify({ success: true, cleaned: uploadIds.length }),
      { headers: { ...corsHeaders, "Content-Type": "application/json" } }
    );
  } catch (e) {
    console.error("[cleanup] Error:", e);
    const message = e instanceof Error ? e.message : "Unknown error";
    return new Response(
      JSON.stringify({ success: false, error: message }),
      { status: 500, headers: { ...corsHeaders, "Content-Type": "application/json" } }
    );
  }
});
