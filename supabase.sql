create table if not exists public.empresas_atendimento (
  cnpj text primary key,
  razao_social text default '',
  nome_fantasia text default '',
  situacao text default '',
  cnae text default '',
  logradouro text default '',
  numero text default '',
  complemento text default '',
  bairro text default '',
  cep text default '',
  municipio text default '',
  uf text default '',
  telefone text default '',
  email text default '',
  status_visita smallint not null default 0 check (status_visita in (0,1,2)),
  primeira_visita_em timestampt,
  segunda_visita_em timestampt,
  observacoes text default '',
  criado_em timestamptz not null default now(),
  atualizado_em timestamptz not null default now()
);

create or replace function public.set_atualizado_em()
returns trigger language plpgsql as $$
begin new.atualizado_em = now(); return new; end; $$;

drop trigger if exists trg_empresas_atendimento_atualizado on public.empresas_atendimento;
create trigger trg_empresas_atendimento_atualizado
before update on public.empresas_atendimento
for each row execute function public.set_atualizado_em();

create index if not exists idx_empresas_atendimento_status on public.empresas_atendimento(status_visita);
create index if not exists idx_empresas_atendimento_bairro on public.empresas_atendimento(bairro);

-- O app usa a chave secreta/service_role apenas no backend Flask.
-- Não coloque essa chave no HTML/JavaScript do navegador.
alter table public.empresas_atendimento enable row level security;
