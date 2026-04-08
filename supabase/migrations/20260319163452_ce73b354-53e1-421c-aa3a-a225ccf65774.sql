-- Allow deleting runs and uploads (needed for cleanup)
CREATE POLICY "Anyone can delete runs" ON public.runs FOR DELETE TO public USING (true);
CREATE POLICY "Anyone can delete uploads" ON public.uploads FOR DELETE TO public USING (true);

-- Allow deleting from storage buckets uploads and outputs
CREATE POLICY "Allow delete uploads storage" ON storage.objects FOR DELETE TO public USING (bucket_id = 'uploads');
CREATE POLICY "Allow delete outputs storage" ON storage.objects FOR DELETE TO public USING (bucket_id = 'outputs');