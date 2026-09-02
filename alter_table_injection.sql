-- Rodar no SQL Editor do projeto Supabase "Valores" (sydamnqagkdmczmgkvso)
-- Colunas novas pra detecção de prompt injection no campo Observação/Link,
-- gravadas automaticamente pelo index.html ao salvar (não altera o fluxo
-- de preenchimento) — usadas pelo notificar_injection_cadastro.py pra
-- avisar a Controladoria quando aparecer algo suspeito.

alter table public.cadastros_deloitte add column if not exists injection_score integer default 0;
alter table public.cadastros_deloitte add column if not exists injection_nivel text default 'LIMPO';
alter table public.cadastros_deloitte add column if not exists injection_categorias text;

NOTIFY pgrst, 'reload schema';
