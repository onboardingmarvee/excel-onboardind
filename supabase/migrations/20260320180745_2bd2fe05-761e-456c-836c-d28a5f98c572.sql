CREATE TABLE public.financial_categories (
  id uuid PRIMARY KEY DEFAULT gen_random_uuid(),
  code text UNIQUE NOT NULL,
  name text NOT NULL,
  full_label text NOT NULL,
  tokens text[] NOT NULL DEFAULT '{}',
  created_at timestamptz NOT NULL DEFAULT now(),
  updated_at timestamptz NOT NULL DEFAULT now()
);

CREATE INDEX idx_financial_categories_tokens ON public.financial_categories USING GIN (tokens);

ALTER TABLE public.financial_categories ENABLE ROW LEVEL SECURITY;

CREATE POLICY "Anyone can view financial_categories" ON public.financial_categories FOR SELECT TO public USING (true);
CREATE POLICY "Anyone can insert financial_categories" ON public.financial_categories FOR INSERT TO public WITH CHECK (true);
CREATE POLICY "Anyone can update financial_categories" ON public.financial_categories FOR UPDATE TO public USING (true);
CREATE POLICY "Anyone can delete financial_categories" ON public.financial_categories FOR DELETE TO public USING (true);