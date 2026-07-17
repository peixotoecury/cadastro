-- Rodar no SQL Editor do projeto Supabase "Valores" (sydamnqagkdmczmgkvso)

create table public.cadastros_deloitte (
  id uuid primary key default gen_random_uuid(),
  created_at timestamptz not null default now(),
  nome_reclamada text not null,
  numero_contrato text,
  numero_processo text not null,
  nome_reclamante text not null,
  cargo_reclamante text not null,
  grupo_processos text not null,
  sigla_adv text not null,
  sigla_assistente text not null,
  caso_estrategico text not null,
  data_admissao date not null,
  contrato_ativo text not null,
  data_desligamento date,
  urgente text not null,
  qual_prazo text,
  data_audiencia date,
  data_prazo date,
  observacao text,
  link_caso text not null
);

alter table public.cadastros_deloitte enable row level security;

create policy "cadastros_deloitte_anon_all" on public.cadastros_deloitte
  for all using (true) with check (true);

alter publication supabase_realtime add table public.cadastros_deloitte;
