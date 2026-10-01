from flask import Flask, render_template, request, jsonify, Response
import os, requests, re, csv, io, time
from datetime import datetime, timezone

app = Flask(__name__)
BRASIL_API = 'https://brasilapi.com.br/api/cnpj/v1/'
SUPABASE_URL = os.getenv('SUPABASE_URL', '').rstrip('/')
SUPABASE_KEY = os.getenv('SUPABASE_KEY', '')  # use Service Role/secret only on server
TABLE = os.getenv('SUPABASE_TABLE', 'empresas_atendimento')


def sb_headers(prefer=None):
    if not SUPABASE_URL or not SUPABASE_KEY:
        raise RuntimeError('Configure SUPABASE_URL e SUPABASE_KEY no provedor.')
    h = {'apikey': SUPABASE_KEY, 'Authorization': f'Bearer {SUPABASE_KEY}', 'Content-Type': 'application/json'}
    if prefer: h['Prefer'] = prefer
    return h

def sb_url(): return f'{SUPABASE_URL}/rest/v1/{TABLE}'
def digits(s): return re.sub(r'\D','',s or '')
def fmt(c):
    c=digits(c); return f'{c[:2]}.{c[2:5]}.{c[5:8]}/{c[8:12]}-{c[12:14]}' if len(c)==14 else c

def fetch_cnpj(cnpj):
    r=requests.get(BRASIL_API+digits(cnpj), timeout=25)
    if r.status_code != 200: raise Exception(f'Consulta cadastral HTTP {r.status_code}')
    j=r.json()
    return dict(
      cnpj=fmt(j.get('cnpj') or cnpj), razao_social=j.get('razao_social',''), nome_fantasia=j.get('nome_fantasia',''),
      situacao=j.get('descricao_situacao_cadastral',''), cnae=f"{j.get('cnae_fiscal','')} - {j.get('cnae_fiscal_descricao','')}",
      logradouro=j.get('logradouro',''), numero=j.get('numero',''), complemento=j.get('complemento',''), bairro=j.get('bairro',''),
      cep=str(j.get('cep') or ''), municipio=j.get('municipio',''), uf=j.get('uf',''), telefone=j.get('ddd_telefone_1',''), email=j.get('email',''))

def get_one(cnpj):
    r=requests.get(sb_url(), headers=sb_headers(), params={'cnpj':f'eq.{fmt(cnpj)}','select':'*','limit':'1'}, timeout=25)
    r.raise_for_status(); data=r.json(); return data[0] if data else None

def upsert(d):
    # Supabase/PostgREST upsert preserves visit fields because they are omitted from payload.
    r=requests.post(sb_url(), headers=sb_headers('resolution=merge-duplicates,return=minimal'), params={'on_conflict':'cnpj'}, json=d, timeout=25)
    if not r.ok: raise Exception(f'Supabase {r.status_code}: {r.text[:300]}')

@app.route('/')
def home(): return render_template('index.html')

@app.route('/health')
def health():
    return jsonify({'ok':True,'database':'supabase','configured':bool(SUPABASE_URL and SUPABASE_KEY),'table':TABLE})

@app.route('/api/empresas')
def empresas():
    try:
        r=requests.get(sb_url(), headers=sb_headers(), params={'select':'*','order':'nome_fantasia.asc.nullslast,razao_social.asc'}, timeout=25)
        r.raise_for_status(); return jsonify(r.json())
    except Exception as e: return jsonify({'erro':str(e)}),500

@app.route('/api/importar', methods=['POST'])
def importar():
    raw=(request.json or {}).get('cnpjs',''); vals=[]
    for x in re.split(r'[,;\n\r\t ]+',raw):
        d=digits(x)
        if len(d)==14 and d not in vals: vals.append(d)
    ok=[]; erros=[]
    for i,x in enumerate(vals):
        try: d=fetch_cnpj(x); upsert(d); ok.append(d['cnpj'])
        except Exception as e: erros.append({'cnpj':fmt(x),'erro':str(e)})
        if i < len(vals)-1: time.sleep(.15)
    return jsonify({'ok':ok,'erros':erros})

@app.route('/api/empresa/<path:cnpj>', methods=['PUT'])
def editar(cnpj):
    data=request.json or {}
    allowed=['razao_social','nome_fantasia','situacao','cnae','logradouro','numero','complemento','bairro','cep','municipio','uf','telefone','email','status_visita','dia_mes_sse','observacoes']
    patch={k:data[k] for k in allowed if k in data}
    status = patch.get('status_visita')
    now=datetime.now(timezone.utc).isoformat()
    if status is not None:
        try: status=int(status); patch['status_visita']=status
        except: return jsonify({'erro':'status_visita inválido'}),400
        old=get_one(cnpj)
        if not old: return jsonify({'erro':'CNPJ não encontrado'}),404
        if status == 1 and not old.get('primeira_visita_em'): patch['primeira_visita_em']=now
        if status == 2:
            if not old.get('primeira_visita_em'): patch['primeira_visita_em']=now
            if not old.get('segunda_visita_em'): patch['segunda_visita_em']=now
    if not patch: return jsonify({'ok':True})
    r=requests.patch(sb_url(), headers=sb_headers('return=minimal'), params={'cnpj':f'eq.{fmt(cnpj)}'}, json=patch, timeout=25)
    if not r.ok: return jsonify({'erro':r.text}),500
    return jsonify({'ok':True})

@app.route('/api/empresa/<path:cnpj>', methods=['DELETE'])
def apagar(cnpj):
    r=requests.delete(sb_url(), headers=sb_headers('return=minimal'), params={'cnpj':f'eq.{fmt(cnpj)}'}, timeout=25)
    if not r.ok: return jsonify({'erro':r.text}),500
    return jsonify({'ok':True})

@app.route('/exportar.csv')
def exportar():
    r=requests.get(sb_url(), headers=sb_headers(), params={'select':'*','order':'nome_fantasia.asc.nullslast,razao_social.asc'}, timeout=25)
    r.raise_for_status(); rows=r.json(); out=io.StringIO(); w=csv.writer(out,delimiter=';')
    if rows:
        keys=list(rows[0].keys()); w.writerow(keys)
        for row in rows: w.writerow([row.get(k,'') for k in keys])
    return Response('\ufeff'+out.getvalue(), mimetype='text/csv', headers={'Content-Disposition':'attachment; filename=prosense_atdm.csv'})

if __name__=='__main__':
    app.run(host='0.0.0.0', port=int(os.getenv('PORT','5000')), debug=os.getenv('FLASK_DEBUG')=='1')
