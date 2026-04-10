ALTER TABLE public.runs ADD COLUMN IF NOT EXISTS config_json jsonb;
