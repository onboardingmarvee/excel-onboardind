-- Create templates bucket
INSERT INTO storage.buckets (id, name, public) VALUES ('templates', 'templates', true)
ON CONFLICT (id) DO NOTHING;

-- Allow public read on templates bucket
CREATE POLICY "Public read templates" ON storage.objects FOR SELECT TO public USING (bucket_id = 'templates');
-- Allow service role to upload templates
CREATE POLICY "Service upload templates" ON storage.objects FOR INSERT TO public WITH CHECK (bucket_id = 'templates');

-- Create import_templates table
CREATE TABLE IF NOT EXISTS public.import_templates (
  id uuid PRIMARY KEY DEFAULT gen_random_uuid(),
  import_type text UNIQUE NOT NULL,
  name text NOT NULL,
  template_storage_path text NOT NULL,
  is_default boolean DEFAULT true,
  version integer DEFAULT 1,
  created_at timestamptz DEFAULT now() NOT NULL
);

ALTER TABLE public.import_templates ENABLE ROW LEVEL SECURITY;

CREATE POLICY "Anyone can view import_templates" ON public.import_templates FOR SELECT TO public USING (true);
CREATE POLICY "Anyone can insert import_templates" ON public.import_templates FOR INSERT TO public WITH CHECK (true);

-- Add output_storage_path to runs if not exists
ALTER TABLE public.runs ADD COLUMN IF NOT EXISTS output_storage_path text;