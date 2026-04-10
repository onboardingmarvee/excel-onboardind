-- Add UPDATE policy to import_templates (needed to update template paths)
CREATE POLICY "Anyone can update import_templates"
  ON public.import_templates FOR UPDATE TO public USING (true) WITH CHECK (true);

-- Add UPDATE and DELETE policies to storage.objects for templates bucket
CREATE POLICY "Public update templates"
  ON storage.objects FOR UPDATE TO public USING (bucket_id = 'templates') WITH CHECK (bucket_id = 'templates');

CREATE POLICY "Public delete templates"
  ON storage.objects FOR DELETE TO public USING (bucket_id = 'templates');

-- Update template_storage_path to point to the new uploaded files
UPDATE public.import_templates
SET template_storage_path = 'contratos_v2.xlsx', version = 2
WHERE import_type = 'contratos';

UPDATE public.import_templates
SET template_storage_path = 'movimentacoes_v2.xlsx', version = 2
WHERE import_type = 'movimentacoes';

UPDATE public.import_templates
SET template_storage_path = 'vendas_v3.xlsx', version = 3
WHERE import_type = 'vendas';

UPDATE public.import_templates
SET template_storage_path = 'clientes_fornecedores/Template-Clientes-Fornecedores_v2.xlsx', version = 2
WHERE import_type = 'clientes_fornecedores';
