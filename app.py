from flask import Flask, render_template, request, jsonify, Response
import os, requests, re, csv, io, time, unicodedata, uuid, mimetypes
from datetime import datetime, timezone

app = Flask(__name__)
BRASIL_API = 'https://brasilapi.com.br/api/cnpj/v1/'
BRASIL_CEP = 'https://brasilapi.com.br/api/cep/v2/'
SUPABASE_URL = os.getenv('SUPABASE_URL', '').rstrip('/')
SUPABASE_KEY = os.getenv('SUPABASE_KEY', '')
TABLE = os.getenv('SUPABASE_TABLE', 'empresas_atendimento')
LIST_TABLE = os.getenv('SUPABASE_LIST_TABLE', 'listas_atendimento')
LINK_TABLE = os.getenv('SUPABASE_LINK_TABLE', 'empresas_listas')

AREA_TABLE = os.getenv('SUPABASE_AREA_TABLE', 'portfolio_areas')
TIPO_TABLE = os.getenv('SUPABASE_TIPO_TABLE', 'portfolio_tipos')
SOLUCAO_TABLE = os.getenv('SUPABASE_SOLUCAO_TABLE', 'portfolio_solucoes')
DIAG_TABLE = os.getenv('SUPABASE_DIAG_TABLE', 'diagnosticos_empresa')
PROSP_TABLE = os.getenv('SUPABASE_PROSP_TABLE', 'prospeccao_fs')
FS_CONFIG_TABLE = os.getenv('SUPABASE_FS_CONFIG_TABLE', 'configuracoes_fs')

def sb_headers(prefer=None):
    if not SUPABASE_URL or not SUPABASE_KEY:
        raise RuntimeError('Configure SUPABASE_URL e SUPABASE_KEY no provedor.')
    h={'apikey':SUPABASE_KEY,'Authorization':f'Bearer {SUPABASE_KEY}','Content-Type':'application/json'}
    if prefer: h['Prefer']=prefer
    return h

def url(table): return f'{SUPABASE_URL}/rest/v1/{table}'
def digits(s): return re.sub(r'\D','',s or '')
def fmt(c):
    c=digits(c); return f'{c[:2]}.{c[2:5]}.{c[5:8]}/{c[8:12]}-{c[12:14]}' if len(c)==14 else c

def fetch_cnpj(cnpj):
    r=requests.get(BRASIL_API+digits(cnpj), timeout=25)
    if r.status_code != 200: raise Exception(f'Consulta cadastral HTTP {r.status_code}')
    j=r.json()
    return dict(cnpj=fmt(j.get('cnpj') or cnpj), razao_social=j.get('razao_social',''), nome_fantasia=j.get('nome_fantasia',''),
      situacao=j.get('descricao_situacao_cadastral',''), cnae=f"{j.get('cnae_fiscal','')} - {j.get('cnae_fiscal_descricao','')}",
      logradouro=j.get('logradouro',''), numero=j.get('numero',''), complemento=j.get('complemento',''), bairro=j.get('bairro',''),
      cep=str(j.get('cep') or ''), municipio=j.get('municipio',''), uf=j.get('uf',''), telefone=j.get('ddd_telefone_1',''), email=j.get('email',''))

def get_one(cnpj):
    r=requests.get(url(TABLE),headers=sb_headers(),params={'cnpj':f'eq.{fmt(cnpj)}','select':'*','limit':'1'},timeout=25);r.raise_for_status();d=r.json();return d[0] if d else None

def upsert_empresa(d):
    r=requests.post(url(TABLE),headers=sb_headers('resolution=merge-duplicates,return=minimal'),params={'on_conflict':'cnpj'},json=d,timeout=25)
    if not r.ok: raise Exception(f'Supabase {r.status_code}: {r.text[:300]}')

def get_or_create_list(nome):
    nome=(nome or '').strip()
    if not nome: raise ValueError('Informe um nome para a lista.')
    r=requests.get(url(LIST_TABLE),headers=sb_headers(),params={'nome':f'eq.{nome}','select':'*','limit':'1'},timeout=25);r.raise_for_status();d=r.json()
    if d:return d[0]
    r=requests.post(url(LIST_TABLE),headers=sb_headers('return=representation'),json={'nome':nome},timeout=25);r.raise_for_status();return r.json()[0]

def link_empresa(lista_id,cnpj):
    payload={'lista_id':lista_id,'cnpj':fmt(cnpj)}
    r=requests.post(url(LINK_TABLE),headers=sb_headers('resolution=ignore-duplicates,return=minimal'),params={'on_conflict':'lista_id,cnpj'},json=payload,timeout=25)
    if not r.ok: raise Exception(f'Erro ao vincular à lista: {r.text[:200]}')

def _valid_coord(lat, lon):
    try:
        lat=float(lat); lon=float(lon)
        return (-34 <= lat <= 6 and -74 <= lon <= -28), lat, lon
    except Exception:
        return False, None, None

def _nominatim(q):
    if not q: return None, None
    try:
        r=requests.get('https://nominatim.openstreetmap.org/search', params={
            'q':q, 'format':'jsonv2', 'limit':1, 'countrycodes':'br', 'addressdetails':1
        }, headers={'User-Agent':'prosense_atdm/1.1 (field-service geocoder)'}, timeout=20)
        if r.ok and r.json():
            ok,lat,lon=_valid_coord(r.json()[0].get('lat'),r.json()[0].get('lon'))
            if ok:return lat,lon
    except Exception: pass
    return None,None

def geocode_empresa(e):
    ok,lat,lon=_valid_coord(e.get('latitude'),e.get('longitude'))
    if ok:return lat,lon,'cache'
    tentativas=[]
    log=(e.get('logradouro') or '').strip(); num=(e.get('numero') or '').strip(); bairro=(e.get('bairro') or '').strip()
    mun=(e.get('municipio') or 'Teresina').strip(); uf=(e.get('uf') or 'PI').strip(); cep=digits(e.get('cep'))
    if log:
        tentativas.append(', '.join(x for x in [log,num,bairro,mun,uf,'Brasil'] if x))
        tentativas.append(', '.join(x for x in [log,bairro,mun,uf,'Brasil'] if x))
    if len(cep)==8:
        # Primeiro tenta o CEP no BrasilAPI; só aceita coordenadas plausíveis no Brasil.
        try:
            r=requests.get(BRASIL_CEP+cep,timeout=15)
            if r.ok:
                c=r.json().get('location',{}).get('coordinates',{}) or {}
                ok,la,lo=_valid_coord(c.get('latitude'),c.get('longitude'))
                if ok: lat,lon=la,lo
        except Exception: pass
        if lat is None: tentativas.append(f'{cep}, {mun}, {uf}, Brasil')
    usado=''
    if lat is None:
        vistos=set()
        for q in tentativas:
            if not q or q in vistos:continue
            vistos.add(q); usado=q
            lat,lon=_nominatim(q)
            if lat is not None:break
            time.sleep(.15)
    if lat is not None:
        try:requests.patch(url(TABLE),headers=sb_headers('return=minimal'),params={'cnpj':f"eq.{e['cnpj']}"},json={'latitude':lat,'longitude':lon},timeout=15)
        except Exception:pass
        return lat,lon,usado or 'CEP'
    return None,None,(tentativas[0] if tentativas else 'Endereço insuficiente')

@app.route('/')
def home(): return render_template('index.html')
@app.route('/health')
def health(): return jsonify({'ok':True,'database':'supabase','configured':bool(SUPABASE_URL and SUPABASE_KEY),'table':TABLE})

@app.route('/api/listas')
def listas():
    try:
        r=requests.get(url(LIST_TABLE),headers=sb_headers(),params={'select':'*','order':'criado_em.desc'},timeout=25);r.raise_for_status();ls=r.json()
        lr=requests.get(url(LINK_TABLE),headers=sb_headers(),params={'select':'lista_id'},timeout=25);lr.raise_for_status();links=lr.json()
        counts={}
        for x in links: counts[x['lista_id']]=counts.get(x['lista_id'],0)+1
        for x in ls:x['total']=counts.get(x['id'],0)
        return jsonify(ls)
    except Exception as e:return jsonify({'erro':str(e)}),500

@app.route('/api/lista/<int:lista_id>',methods=['PUT'])
def renomear_lista(lista_id):
    nome=((request.json or {}).get('nome') or '').strip()
    if not nome:return jsonify({'erro':'Informe o novo nome da lista.'}),400
    r=requests.patch(url(LIST_TABLE),headers=sb_headers('return=representation'),params={'id':f'eq.{lista_id}'},json={'nome':nome},timeout=25)
    if not r.ok:return jsonify({'erro':r.text}),500
    return jsonify({'ok':True,'lista':(r.json()[0] if r.json() else {'id':lista_id,'nome':nome})})

@app.route('/api/empresas')
def empresas():
    try:
        lista_id=request.args.get('lista_id')
        if lista_id:
            lr=requests.get(url(LINK_TABLE),headers=sb_headers(),params={'lista_id':f'eq.{lista_id}','select':'cnpj'},timeout=25);lr.raise_for_status();cnpjs=[x['cnpj'] for x in lr.json()]
            if not cnpjs:return jsonify([])
            r=requests.get(url(TABLE),headers=sb_headers(),params={'cnpj':'in.('+','.join('"'+c+'"' for c in cnpjs)+')','select':'*','order':'nome_fantasia.asc.nullslast,razao_social.asc'},timeout=25)
        else:r=requests.get(url(TABLE),headers=sb_headers(),params={'select':'*','order':'nome_fantasia.asc.nullslast,razao_social.asc'},timeout=25)
        r.raise_for_status(); rows=r.json()
        # Marca empresas que já possuem pelo menos um diagnóstico salvo.
        # Se a tabela de diagnósticos estiver temporariamente indisponível, a lista continua funcionando.
        try:
            dr=requests.get(url(DIAG_TABLE),headers=sb_headers(),params={'select':'cnpj'},timeout=20)
            diag_cnpjs={x.get('cnpj') for x in dr.json()} if dr.ok else set()
        except Exception:
            diag_cnpjs=set()
        for x in rows:x['diagnosticado']=x.get('cnpj') in diag_cnpjs
        # Situação comercial FS, independente do diagnóstico Sebrae.
        try:
            pr=requests.get(url(PROSP_TABLE),headers=sb_headers(),params={'select':'cnpj,status,ultimo_contato,proximo_contato,nao_contatar'},timeout=20)
            pm={x.get('cnpj'):x for x in (pr.json() if pr.ok else [])}
        except Exception: pm={}
        for x in rows:
            px=pm.get(x.get('cnpj'),{})
            x['prospeccao_status']=px.get('status') or 'nao_abordado'
            x['prospeccao_ultimo_contato']=px.get('ultimo_contato')
            x['prospeccao_proximo_contato']=px.get('proximo_contato')
            x['nao_contatar']=bool(px.get('nao_contatar'))
        return jsonify(rows)
    except Exception as e:return jsonify({'erro':str(e)}),500

@app.route('/api/importar',methods=['POST'])
def importar():
    body=request.json or {}; raw=body.get('cnpjs',''); nome=body.get('nome_lista',''); vals=[]
    try:lista=get_or_create_list(nome)
    except Exception as e:return jsonify({'erro':str(e)}),400
    for x in re.split(r'[,;\n\r\t ]+',raw):
        d=digits(x)
        if len(d)==14 and d not in vals:vals.append(d)
    ok=[];erros=[]
    for i,x in enumerate(vals):
        try:d=fetch_cnpj(x);upsert_empresa(d);link_empresa(lista['id'],d['cnpj']);ok.append(d['cnpj'])
        except Exception as e:erros.append({'cnpj':fmt(x),'erro':str(e)})
        if i<len(vals)-1:time.sleep(.15)
    return jsonify({'ok':ok,'erros':erros,'lista':lista})

@app.route('/api/empresa/<path:cnpj>',methods=['PUT'])
def editar(cnpj):
    data=request.json or {};allowed=['razao_social','nome_fantasia','situacao','cnae','logradouro','numero','complemento','bairro','cep','municipio','uf','telefone','email','status_visita','dia_mes_sse','observacoes']
    patch={k:data[k] for k in allowed if k in data};status=patch.get('status_visita');now=datetime.now(timezone.utc).isoformat()
    if status is not None:
        try:status=int(status);patch['status_visita']=status
        except:return jsonify({'erro':'status_visita inválido'}),400
        old=get_one(cnpj)
        if not old:return jsonify({'erro':'CNPJ não encontrado'}),404
        if status==1 and not old.get('primeira_visita_em'):patch['primeira_visita_em']=now
        if status==2:
            if not old.get('primeira_visita_em'):patch['primeira_visita_em']=now
            if not old.get('segunda_visita_em'):patch['segunda_visita_em']=now
        if any(k in patch for k in ['logradouro','numero','bairro','cep','municipio','uf']):patch.update({'latitude':None,'longitude':None})
    r=requests.patch(url(TABLE),headers=sb_headers('return=minimal'),params={'cnpj':f'eq.{fmt(cnpj)}'},json=patch,timeout=25)
    if not r.ok:return jsonify({'erro':r.text}),500
    return jsonify({'ok':True})

@app.route('/api/empresa/<path:cnpj>/foto',methods=['POST'])
def foto_empresa(cnpj):
    f=request.files.get('foto')
    if not f or not f.filename:return jsonify({'erro':'Selecione uma foto.'}),400
    mime=(f.mimetype or '').lower()
    if not mime.startswith('image/'):return jsonify({'erro':'O arquivo precisa ser uma imagem.'}),400
    data=f.read()
    if len(data)>8*1024*1024:return jsonify({'erro':'A foto deve ter no máximo 8 MB.'}),400
    ext=mimetypes.guess_extension(mime) or os.path.splitext(f.filename)[1] or '.jpg'
    if ext=='.jpe':ext='.jpg'
    path=f"{digits(cnpj)}-{uuid.uuid4().hex[:10]}{ext}"
    storage=f'{SUPABASE_URL}/storage/v1/object/fachadas/{path}'
    h={'apikey':SUPABASE_KEY,'Authorization':f'Bearer {SUPABASE_KEY}','Content-Type':mime,'x-upsert':'true'}
    r=requests.post(storage,headers=h,data=data,timeout=45)
    if not r.ok:return jsonify({'erro':f'Falha ao enviar foto: {r.text[:250]}'}),500
    foto_url=f'{SUPABASE_URL}/storage/v1/object/public/fachadas/{path}'
    pr=requests.patch(url(TABLE),headers=sb_headers('return=minimal'),params={'cnpj':f'eq.{fmt(cnpj)}'},json={'foto_url':foto_url},timeout=25)
    if not pr.ok:return jsonify({'erro':pr.text}),500
    return jsonify({'ok':True,'foto_url':foto_url})

@app.route('/api/empresa/<path:cnpj>',methods=['DELETE'])
def apagar(cnpj):
    r=requests.delete(url(TABLE),headers=sb_headers('return=minimal'),params={'cnpj':f'eq.{fmt(cnpj)}'},timeout=25)
    if not r.ok:return jsonify({'erro':r.text}),500
    return jsonify({'ok':True})

@app.route('/api/mapa',methods=['POST'])
def mapa():
    cnpjs=(request.json or {}).get('cnpjs',[])[:150];pontos=[];nao=[]
    for i,c in enumerate(cnpjs):
        e=get_one(c)
        if not e:
            nao.append({'cnpj':fmt(c),'nome_fantasia':'','motivo':'Registro não encontrado','endereco_tentado':''});continue
        lat,lon,tentado=geocode_empresa(e)
        base={'cnpj':e['cnpj'],'nome_fantasia':e.get('nome_fantasia') or e.get('razao_social') or '',
              'dia_mes_sse':e.get('dia_mes_sse') or '','status_visita':e.get('status_visita',0),
              'bairro':e.get('bairro') or '','logradouro':e.get('logradouro') or '','numero':e.get('numero') or '',
              'cep':e.get('cep') or '','municipio':e.get('municipio') or '','uf':e.get('uf') or '',
              'foto_url':e.get('foto_url') or ''}
        if lat is not None and lon is not None:
            base.update({'latitude':lat,'longitude':lon});pontos.append(base)
        else:
            base.update({'motivo':'Endereço não localizado','endereco_tentado':tentado});nao.append(base)
        if i<len(cnpjs)-1 and (e.get('latitude') is None or e.get('longitude') is None):time.sleep(.20)
    return jsonify({'pontos':pontos,'nao_localizados':nao,'solicitados':len(cnpjs),'localizados':len(pontos),'pendentes_localizacao':len(nao)})


@app.route('/api/mapa/posicao',methods=['PUT'])
def atualizar_posicao_mapa():
    body=request.json or {}; cnpjs=body.get('cnpjs') or []
    try: lat=float(body.get('latitude')); lon=float(body.get('longitude'))
    except Exception:return jsonify({'erro':'Coordenadas inválidas.'}),400
    ok,lat,lon=_valid_coord(lat,lon)
    if not ok:return jsonify({'erro':'Coordenadas fora da área válida.'}),400
    atualizados=[]; erros=[]
    for c in cnpjs[:150]:
        c=fmt(c)
        try:
            r=requests.patch(url(TABLE),headers=sb_headers('return=minimal'),params={'cnpj':f'eq.{c}'},json={'latitude':lat,'longitude':lon},timeout=20)
            if r.ok: atualizados.append(c)
            else: erros.append({'cnpj':c,'erro':r.text[:160]})
        except Exception as e: erros.append({'cnpj':c,'erro':str(e)})
    return jsonify({'ok':True,'atualizados':atualizados,'erros':erros,'latitude':lat,'longitude':lon})


@app.route('/api/fs/config',methods=['GET','PUT'])
def fs_config():
    if request.method=='GET':
        r=requests.get(url(FS_CONFIG_TABLE),headers=sb_headers(),params={'id':'eq.1','select':'*','limit':'1'},timeout=20)
        if not r.ok:return jsonify({'erro':r.text}),500
        d=r.json()
        return jsonify(d[0] if d else {})
    b=request.json or {}
    allowed=['link_triagem','mensagem_padrao','slogan','whatsapp_escritorio']
    payload={'id':1,**{k:b.get(k,'') for k in allowed if k in b}}
    r=requests.post(url(FS_CONFIG_TABLE),headers=sb_headers('resolution=merge-duplicates,return=representation'),params={'on_conflict':'id'},json=payload,timeout=20)
    if not r.ok:return jsonify({'erro':r.text}),500
    return jsonify(r.json()[0] if r.json() else payload)

@app.route('/api/prospeccao/<path:cnpj>',methods=['GET','PUT'])
def prospeccao(cnpj):
    c=fmt(cnpj)
    if request.method=='GET':
        r=requests.get(url(PROSP_TABLE),headers=sb_headers(),params={'cnpj':f'eq.{c}','select':'*','limit':'1'},timeout=20)
        if not r.ok:return jsonify({'erro':r.text}),500
        d=r.json(); return jsonify(d[0] if d else {'cnpj':c,'status':'nao_abordado'})
    b=request.json or {}
    allowed=['status','interesse','ultimo_contato','proximo_contato','observacoes','nao_contatar','consentiu_whatsapp']
    payload={'cnpj':c,**{k:b.get(k) for k in allowed if k in b}}
    r=requests.post(url(PROSP_TABLE),headers=sb_headers('resolution=merge-duplicates,return=representation'),params={'on_conflict':'cnpj'},json=payload,timeout=20)
    if not r.ok:return jsonify({'erro':r.text}),500
    return jsonify(r.json()[0] if r.json() else payload)

@app.route('/api/prospeccao/marcar-envio/<path:cnpj>',methods=['POST'])
def marcar_envio(cnpj):
    c=fmt(cnpj); now=datetime.now(timezone.utc).isoformat()
    payload={'cnpj':c,'status':'apresentacao_enviada','ultimo_contato':now}
    r=requests.post(url(PROSP_TABLE),headers=sb_headers('resolution=merge-duplicates,return=representation'),params={'on_conflict':'cnpj'},json=payload,timeout=20)
    if not r.ok:return jsonify({'erro':r.text}),500
    return jsonify({'ok':True,'ultimo_contato':now})

@app.route('/api/portfolio/catalogo')
def portfolio_catalogo():
    try:
        a=requests.get(url(AREA_TABLE),headers=sb_headers(),params={'select':'*','order':'nome.asc'},timeout=20);a.raise_for_status()
        t=requests.get(url(TIPO_TABLE),headers=sb_headers(),params={'select':'*','order':'nome.asc'},timeout=20);t.raise_for_status()
        s=requests.get(url(SOLUCAO_TABLE),headers=sb_headers(),params={'select':'*','order':'nome.asc'},timeout=20);s.raise_for_status()
        return jsonify({'areas':a.json(),'tipos':t.json(),'solucoes':s.json()})
    except Exception as e:return jsonify({'erro':str(e)}),500

@app.route('/api/portfolio/<kind>',methods=['POST'])
def portfolio_categoria_criar(kind):
    table=AREA_TABLE if kind=='area' else TIPO_TABLE if kind=='tipo' else None
    if not table:return jsonify({'erro':'Categoria inválida'}),404
    nome=((request.json or {}).get('nome') or '').strip()
    if not nome:return jsonify({'erro':'Informe o nome.'}),400
    r=requests.post(url(table),headers=sb_headers('resolution=ignore-duplicates,return=representation'),params={'on_conflict':'nome'},json={'nome':nome},timeout=20)
    if not r.ok:return jsonify({'erro':r.text}),500
    return jsonify({'ok':True,'item':(r.json()[0] if r.json() else {'nome':nome})})

@app.route('/api/portfolio/<kind>/<int:item_id>',methods=['DELETE'])
def portfolio_categoria_apagar(kind,item_id):
    table=AREA_TABLE if kind=='area' else TIPO_TABLE if kind=='tipo' else None
    if not table:return jsonify({'erro':'Categoria inválida'}),404
    r=requests.delete(url(table),headers=sb_headers('return=minimal'),params={'id':f'eq.{item_id}'},timeout=20)
    if not r.ok:return jsonify({'erro':'Não foi possível excluir. Verifique se há soluções usando esta categoria.'}),409
    return jsonify({'ok':True})

@app.route('/api/portfolio/solucao',methods=['POST'])
def portfolio_solucao_criar():
    b=request.json or {}
    try: area_id=int(b.get('area_id')); tipo_id=int(b.get('tipo_id'))
    except Exception:return jsonify({'erro':'Selecione área e tipo.'}),400
    nome=(b.get('nome') or '').strip().upper(); link=(b.get('link') or '').strip()
    if not nome:return jsonify({'erro':'Informe o nome da solução.'}),400
    r=requests.post(url(SOLUCAO_TABLE),headers=sb_headers('return=representation'),json={'area_id':area_id,'tipo_id':tipo_id,'nome':nome,'link':link},timeout=20)
    if not r.ok:return jsonify({'erro':r.text}),500
    return jsonify({'ok':True,'item':r.json()[0]})

@app.route('/api/portfolio/solucao/<int:item_id>',methods=['DELETE'])
def portfolio_solucao_apagar(item_id):
    r=requests.delete(url(SOLUCAO_TABLE),headers=sb_headers('return=minimal'),params={'id':f'eq.{item_id}'},timeout=20)
    if not r.ok:return jsonify({'erro':r.text}),500
    return jsonify({'ok':True})


@app.route('/api/diagnostico-grafico/<path:cnpj>',methods=['POST'])
def grafico_diagnostico(cnpj):
    f=request.files.get('grafico')
    if not f:return jsonify({'erro':'Gráfico não recebido.'}),400
    data=f.read()
    if not data or len(data)>3*1024*1024:return jsonify({'erro':'Gráfico inválido ou muito grande.'}),400
    import time
    clean=fmt(cnpj) or 'empresa'
    path=f'diagnosticos/{clean}_{int(time.time())}.png'
    storage=f'{SUPABASE_URL}/storage/v1/object/fachadas/{path}'
    h={'apikey':SUPABASE_KEY,'Authorization':f'Bearer {SUPABASE_KEY}','Content-Type':'image/png','x-upsert':'true'}
    r=requests.post(storage,headers=h,data=data,timeout=45)
    if not r.ok:return jsonify({'erro':f'Falha ao enviar gráfico: {r.text[:250]}'}),500
    public=f'{SUPABASE_URL}/storage/v1/object/public/fachadas/{path}'
    return jsonify({'ok':True,'url':public})

@app.route('/api/diagnostico/<path:cnpj>',methods=['GET'])
def obter_diagnostico(cnpj):
    c=fmt(cnpj)
    r=requests.get(url(DIAG_TABLE),headers=sb_headers(),params={'cnpj':f'eq.{c}','select':'*','order':'criado_em.desc','limit':'1'},timeout=20)
    if not r.ok:return jsonify({'erro':r.text}),500
    rows=r.json()
    return jsonify({'diagnostico':rows[0] if rows else None})

@app.route('/api/diagnostico',methods=['POST'])
def salvar_diagnostico():
    b=request.json or {}; c=fmt(b.get('cnpj'))
    if not get_one(c):return jsonify({'erro':'Empresa não encontrada.'}),404
    respostas=b.get('respostas',{}) or {}
    tributacao=b.get('tributacao',{}) or {}
    # Cópia de segurança dentro de respostas: mantém a tributação persistida mesmo antes do cache
    # do PostgREST reconhecer a coluna tributacao.
    respostas['_tributacao']=tributacao
    payload={'cnpj':c,'segmento':b.get('segmento',''),'respostas':respostas,'areas_criticas':b.get('areas_criticas',[]),'problemas':b.get('problemas',[]),'solucoes':b.get('solucoes',[]),'tributacao':tributacao,'mensagem':b.get('mensagem','')}
    r=requests.post(url(DIAG_TABLE),headers=sb_headers('return=representation'),json=payload,timeout=25)
    # Compatibilidade: se o PostgREST ainda não enxergar a coluna, salva sem ela em vez de bloquear o atendimento.
    if not r.ok and ('tributacao' in r.text and ('PGRST204' in r.text or 'schema cache' in r.text)):
        payload.pop('tributacao',None)
        r=requests.post(url(DIAG_TABLE),headers=sb_headers('return=representation'),json=payload,timeout=25)
    if not r.ok:return jsonify({'erro':r.text}),500
    return jsonify({'ok':True,'diagnostico':r.json()[0]})

@app.route('/exportar.csv')
def exportar():
    lista_id=request.args.get('lista_id');rows=[]
    if lista_id:
        lr=requests.get(url(LINK_TABLE),headers=sb_headers(),params={'lista_id':f'eq.{lista_id}','select':'cnpj'},timeout=25);lr.raise_for_status();cnpjs=[x['cnpj'] for x in lr.json()]
        if cnpjs:
            r=requests.get(url(TABLE),headers=sb_headers(),params={'cnpj':'in.('+','.join('"'+c+'"' for c in cnpjs)+')','select':'*','order':'nome_fantasia.asc.nullslast'},timeout=25);r.raise_for_status();rows=r.json()
    else:
        r=requests.get(url(TABLE),headers=sb_headers(),params={'select':'*','order':'nome_fantasia.asc.nullslast'},timeout=25);r.raise_for_status();rows=r.json()
    out=io.StringIO();w=csv.writer(out,delimiter=';')
    if rows:
        keys=list(rows[0].keys());w.writerow(keys)
        for row in rows:w.writerow([row.get(k,'') for k in keys])
    return Response('\ufeff'+out.getvalue(),mimetype='text/csv',headers={'Content-Disposition':'attachment; filename=prosense_atdm.csv'})

if __name__=='__main__':app.run(host='0.0.0.0',port=int(os.getenv('PORT','5000')),debug=os.getenv('FLASK_DEBUG')=='1')
