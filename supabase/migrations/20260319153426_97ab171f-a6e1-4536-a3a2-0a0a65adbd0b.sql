
-- Enums
CREATE TYPE public.upload_status AS ENUM ('uploaded', 'processing', 'done', 'error');
CREATE TYPE public.run_status AS ENUM ('queued', 'processing', 'done', 'error');
CREATE TYPE public.import_type AS ENUM ('clientes_fornecedores', 'movimentacoes', 'vendas', 'contratos');

-- Uploads table
CREATE TABLE public.uploads (
  id UUID NOT NULL DEFAULT gen_random_uuid() PRIMARY KEY,
  original_filename TEXT NOT NULL,
  storage_path TEXT NOT NULL,
  status upload_status NOT NULL DEFAULT 'uploaded',
  error_message TEXT,
  created_at TIMESTAMP WITH TIME ZONE NOT NULL DEFAULT now()
);

ALTER TABLE public.uploads ENABLE ROW LEVEL SECURITY;

CREATE POLICY "Anyone can view uploads" ON public.uploads FOR SELECT USING (true);
CREATE POLICY "Anyone can insert uploads" ON public.uploads FOR INSERT WITH CHECK (true);
CREATE POLICY "Anyone can update uploads" ON public.uploads FOR UPDATE USING (true);

-- Runs table
CREATE TABLE public.runs (
  id UUID NOT NULL DEFAULT gen_random_uuid() PRIMARY KEY,
  upload_id UUID NOT NULL REFERENCES public.uploads(id) ON DELETE CASCADE,
  import_type import_type NOT NULL,
  transform_instructions TEXT,
  status run_status NOT NULL DEFAULT 'queued',
  preview_json JSONB,
  error_report_json JSONB,
  output_xlsx_path TEXT,
  output_csv_path TEXT,
  created_at TIMESTAMP WITH TIME ZONE NOT NULL DEFAULT now()
);

ALTER TABLE public.runs ENABLE ROW LEVEL SECURITY;

CREATE POLICY "Anyone can view runs" ON public.runs FOR SELECT USING (true);
CREATE POLICY "Anyone can insert runs" ON public.runs FOR INSERT WITH CHECK (true);
CREATE POLICY "Anyone can update runs" ON public.runs FOR UPDATE USING (true);

-- Storage buckets
INSERT INTO storage.buckets (id, name, public) VALUES ('uploads', 'uploads', false);
INSERT INTO storage.buckets (id, name, public) VALUES ('outputs', 'outputs', true);

-- Storage policies
CREATE POLICY "Anyone can upload files" ON storage.objects FOR INSERT WITH CHECK (bucket_id = 'uploads');
CREATE POLICY "Anyone can read uploads" ON storage.objects FOR SELECT USING (bucket_id = 'uploads');
CREATE POLICY "Anyone can read outputs" ON storage.objects FOR SELECT USING (bucket_id = 'outputs');
CREATE POLICY "Service can write outputs" ON storage.objects FOR INSERT WITH CHECK (bucket_id = 'outputs');
CREATE POLICY "Service can update outputs" ON storage.objects FOR UPDATE USING (bucket_id = 'outputs');
