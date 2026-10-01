# prosense_atdm — Supabase

Painel Flask responsivo para importar CNPJs, consultar dados cadastrais, editar contatos e controlar visitas.

## Banco Supabase
1. No Supabase, abra **SQL Editor** e execute `supabase.sql`.
2. No provedor onde o Flask ficará hospedado, crie as variáveis:
   - `SUPABASE_URL`
   - `SUPABASE_KEY` (chave secreta/service_role, somente no backend)
   - `SUPABASE_TABLE=empresas_atendimento` (opcional)
3. Nunca coloque `SUPABASE_KEY` no HTML, JavaScript público ou GitHub.

Pode usar o mesmo projeto Supabase do FS Bolões: esta aplicação usa somente a tabela `empresas_atendimento` e não altera as tabelas do bolão.

## Rodar
`pip install -r requirements.txt`
`python app.py`

## Importação
Cole CNPJs por linha, vírgula, ponto-e-vírgula ou espaço. O backend consulta a BrasilAPI e faz upsert no Supabase.

## Status
- 0 = pendente (sem cor)
- 1 = primeira visita (verde)
- 2 = segunda visita (azul)

As datas da primeira e segunda visita são gravadas automaticamente na primeira vez que cada estágio é marcado.
