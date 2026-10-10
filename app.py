from flask import Flask, render_template, request, jsonify, Response, render_template_string
import os, requests, re, csv, io, time, unicodedata, uuid, mimetypes
from openpyxl import load_workbook
from datetime import datetime, timezone

app = Flask(__name__)
app.config['MAX_CONTENT_LENGTH'] = 20 * 1024 * 1024
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
ATEND_CONTABIL_TABLE = os.getenv('SUPABASE_ATEND_CONTABIL_TABLE', 'atendimento_contabil_fs')

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



def _fetch_paginated(table, params=None, page_size=1000, timeout=30):
    """Busca todas as linhas de uma tabela PostgREST sem depender de filtro IN gigante."""
    out=[]; start=0; params=dict(params or {})
    while True:
        h=sb_headers(); h['Range']=f'{start}-{start+page_size-1}'
        r=requests.get(url(table),headers=h,params=params,timeout=timeout)
        r.raise_for_status()
        part=r.json()
        out.extend(part)
        if len(part)<page_size: break
        start += page_size
        if start>100000: raise RuntimeError('Limite de segurança excedido ao ler registros.')
    return out

def _empresas_da_lista(lista_id):
    """Resolve uma lista no backend sem montar cnpj=in.(...) com CNPJs formatados."""
    links=_fetch_paginated(LINK_TABLE,{'lista_id':f'eq.{lista_id}','select':'cnpj'})
    wanted={fmt(x.get('cnpj','')) for x in links if x.get('cnpj')}
    if not wanted:return []
    todas=_fetch_paginated(TABLE,{'select':'*'})
    rows=[x for x in todas if fmt(x.get('cnpj','')) in wanted]
    rows.sort(key=lambda x:((x.get('nome_fantasia') or '').casefold() or '\uffff',(x.get('razao_social') or '').casefold()))
    return rows
def link_empresa(lista_id,cnpj):
    payload={'lista_id':lista_id,'cnpj':fmt(cnpj)}
    r=requests.post(url(LINK_TABLE),headers=sb_headers('resolution=ignore-duplicates,return=minimal'),params={'on_conflict':'lista_id,cnpj'},json=payload,timeout=25)
    if not r.ok: raise Exception(f'Erro ao vincular à lista: {r.text[:200]}')


def _norm_header(v):
    s=unicodedata.normalize('NFKD',str(v or '')).encode('ascii','ignore').decode('ascii').lower().strip()
    return re.sub(r'[^a-z0-9]+','_',s).strip('_')

def _cell_text(v):
    if v is None:return ''
    if isinstance(v,datetime):return v.strftime('%d/%m/%Y')
    if isinstance(v,float) and v.is_integer():return str(int(v))
    return str(v).strip()

def _parse_planilha(fs):
    nome=(fs.filename or '').strip(); ext=os.path.splitext(nome.lower())[1]
    raw=fs.read()
    if not raw:raise ValueError('A planilha está vazia.')
    linhas=[]
    if ext in ('.xlsx','.xlsm'):
        wb=load_workbook(io.BytesIO(raw),read_only=True,data_only=True)
        ws=wb.active
        it=ws.iter_rows(values_only=True)
        try:cab=[_cell_text(x) for x in next(it)]
        except StopIteration:raise ValueError('A planilha está vazia.')
        for vals in it:
            row=[_cell_text(x) for x in vals]
            if any(x for x in row):linhas.append(row)
    elif ext in ('.csv','.tsv','.txt'):
        try:texto=raw.decode('utf-8-sig')
        except UnicodeDecodeError:texto=raw.decode('latin-1')
        amostra=texto[:8192]
        if ext=='.tsv' or '\t' in amostra:delim='\t'
        elif ';' in amostra and amostra.count(';')>=amostra.count(','):delim=';'
        else:delim=','
        rr=csv.reader(io.StringIO(texto),delimiter=delim)
        try:cab=[_cell_text(x) for x in next(rr)]
        except StopIteration:raise ValueError('A planilha está vazia.')
        for vals in rr:
            row=[_cell_text(x) for x in vals]
            if any(x for x in row):linhas.append(row)
    else:
        raise ValueError('Formato não suportado. Use XLSX, XLSM, CSV ou TSV.')
    if len(linhas)>20000:raise ValueError('Limite de 20.000 linhas por importação.')
    if not cab:raise ValueError('Cabeçalho não encontrado.')
    norm=[_norm_header(x) for x in cab]
    if 'cnpj' not in norm:raise ValueError('A planilha precisa ter uma coluna CNPJ.')
    saida=[];erros=[];vistos=set()
    for idx,vals in enumerate(linhas,start=2):
        vals=(vals+['']*len(cab))[:len(cab)]
        original={cab[i] or f'coluna_{i+1}':vals[i] for i in range(len(cab))}
        d={norm[i]:vals[i] for i in range(len(norm))}
        c=digits(d.get('cnpj',''))
        if len(c)==13:c=c.zfill(14)
        if len(c)!=14:
            erros.append({'linha':idx,'erro':'CNPJ inválido ou ausente'})
            continue
        cf=fmt(c)
        if cf in vistos:
            erros.append({'linha':idx,'erro':'CNPJ repetido na planilha (ignorado)'})
            continue
        vistos.add(cf)
        tipo=(d.get('tipo_logradouro') or '').strip(); log=(d.get('logradouro') or '').strip()
        logfull=(' '.join(x for x in [tipo,log] if x)).strip()
        cnae=(d.get('cnae') or '').strip(); ramo=(d.get('ramo_de_atividade') or '').strip()
        cnaefull=(f'{cnae} - {ramo}' if cnae and ramo else cnae or ramo)
        tel=(d.get('telefone1_completo') or d.get('telefone_1_completo') or d.get('telefone1') or d.get('telefone') or d.get('telefone2_completo') or d.get('telefone2') or '').strip()
        ddd=(d.get('ddd1') or d.get('ddd2') or '').strip()
        if tel and ddd and not digits(tel).startswith(ddd): tel=f'({ddd}) {tel}'
        payload={
          'cnpj':cf,
          'razao_social':d.get('razao_social',''),
          'nome_fantasia':d.get('nome_fantasia') or d.get('nome_fantaria') or '',
          'situacao':d.get('situacao',''),
          'cnae':cnaefull,
          'logradouro':logfull,
          'numero':d.get('numero',''),
          'complemento':d.get('complemento',''),
          'bairro':d.get('bairro',''),
          'cep':digits(d.get('cep','')),
          'municipio':d.get('municipio',''),
          'uf':d.get('uf',''),
          'telefone':tel,
          'email':d.get('e_mail','') or d.get('email',''),
          'origem_importacao':nome,
          'dados_importados':original,
        }
        saida.append(payload)
    return saida,erros

def _merge_planilha_com_cadastro(rows, sobrescrever=False):
    # Uma leitura paginada evita milhares de consultas individuais à internet.
    atuais={fmt(e.get('cnpj','')):e for e in _fetch_paginated(TABLE, {'select':'*'})}
    novos=0; existentes=0; saida=[]
    campos={'razao_social','nome_fantasia','situacao','cnae','logradouro','numero',
            'complemento','bairro','cep','municipio','uf','telefone','email'}
    for d in rows:
        original=atuais.get(d['cnpj'])
        if not original:
            novos+=1; saida.append(d); continue
        existentes+=1
        patch={'cnpj':d['cnpj']}
        for k in campos:
            novo=d.get(k)
            if novo not in (None,'') and (sobrescrever or not original.get(k)):
                patch[k]=novo
        antigos=original.get('dados_importados') or {}
        if not isinstance(antigos,dict):antigos={}
        patch['dados_importados']={**antigos,**d['dados_importados']}
        if sobrescrever or not original.get('origem_importacao'):
            patch['origem_importacao']=d['origem_importacao']
        # Não modificar visitas, diagnósticos, fotos, observações, coordenadas ou status.
        saida.append(patch)
    return saida,novos,existentes

def _upsert_lote(rows, batch=200):
    """PostgREST exige exatamente as mesmas chaves em todos os objetos de um POST.

    Registros novos trazem todos os campos cadastrais; existentes trazem somente
    campos que devem ser atualizados. Agrupar por conjunto de chaves preserva
    os campos antigos sem enviar valores vazios ou nulos por engano.
    """
    grupos = {}
    for row in rows:
        chave = tuple(sorted(row.keys()))
        grupos.setdefault(chave, []).append(row)
    for chave, grupo in grupos.items():
        for i in range(0, len(grupo), batch):
            parte = grupo[i:i + batch]
            r = requests.post(
                url(TABLE),
                headers=sb_headers('resolution=merge-duplicates,return=minimal'),
                params={'on_conflict': 'cnpj'},
                json=parte,
                timeout=60,
            )
            if not r.ok:
                if 'dados_importados' in r.text or 'origem_importacao' in r.text:
                    raise Exception('Estrutura do banco ainda não atualizada. Verifique o SQL de migração de importação no Supabase.')
                raise Exception(f'Falha ao salvar lote no Supabase: {r.status_code} {r.text[:350]}')

def _link_lote(lista_id,cnpjs,batch=300):
    rows=[{'lista_id':lista_id,'cnpj':c} for c in cnpjs]
    for i in range(0,len(rows),batch):
        r=requests.post(url(LINK_TABLE),headers=sb_headers('resolution=ignore-duplicates,return=minimal'),params={'on_conflict':'lista_id,cnpj'},json=rows[i:i+batch],timeout=60)
        if not r.ok:raise Exception(f'Falha ao vincular empresas à lista: {r.status_code} {r.text[:300]}')

def _valid_coord(lat, lon):
    try:
        lat=float(lat); lon=float(lon)
        return (-34 <= lat <= 6 and -74 <= lon <= -28), lat, lon
    except Exception:
        return False, None, None

def _norm_lugar(v):
    v=unicodedata.normalize('NFKD',str(v or '')).encode('ascii','ignore').decode('ascii').lower()
    v=re.sub(r'[^a-z0-9]+',' ',v)
    return re.sub(r'\s+',' ',v).strip()

def _identificar_endereco(e):
    """Reconhece quadra/casa/setor mesmo quando misturados ao logradouro ou complemento."""
    partes=[str(e.get(c) or '') for c in ('logradouro','numero','complemento','bairro')]
    bruto=' | '.join(partes)
    rxq=r'\b(?:quadra|qd|qdr|qra|q)\.?\s*[-.:º°n]*\s*(\d+[a-z]?)\b'
    rxc=r'\b(?:casa|cs|lote|lt)\.?\s*[-.:º°n]*\s*(\d+[a-z]?)\b'
    rxs=r'\b(?:setor|set)\.?\s*[-.:º°n]*\s*([a-z0-9]+)\b'
    def ler(rx):
        m=re.search(rx,bruto,re.IGNORECASE)
        return m.group(1).upper() if m else ''
    quadra,casa,setor=ler(rxq),ler(rxc),ler(rxs)
    numero=str(e.get('numero') or '').strip()
    if not casa and numero.isdigit() and (quadra or 'casa' in _norm_lugar(bruto)):casa=numero
    return quadra,casa,setor

def _nominatim_candidatos(q):
    if not q:return []
    try:
        # Respeita o limite público aproximado de 1 consulta/s por aplicação.
        time.sleep(1.1)
        r=requests.get('https://nominatim.openstreetmap.org/search',params={
          'q':q,'format':'jsonv2','limit':5,'countrycodes':'br','addressdetails':1,
        },headers={'User-Agent':'ProSense-ATDM/2.0 (endereco estruturado; contato administrador ProSense)'},timeout=20)
        if r.ok:return r.json() or []
    except Exception:pass
    return []

def _nominatim(q):
    for item in _nominatim_candidatos(q):
        ok,lat,lon=_valid_coord(item.get('lat'),item.get('lon'))
        if ok:return lat,lon
    return None,None

def _evidencia_local(item, e, quadra, casa, setor):
    """Score + qualidade: ruas e números primeiro; setor/quadra como aproximação.
    Nunca confundir cidade ou bairro isolado com uma localização exata.
    """
    a=item.get('address') or {}
    raw=_norm_lugar(item.get('display_name'))
    contexto=' '.join([raw]+[_norm_lugar(v) for v in a.values() if isinstance(v,str)])
    mun=_norm_lugar(e.get('municipio') or 'Teresina')
    if mun and mun not in contexto:return None
    bairro=_norm_lugar(e.get('bairro'))
    bairro_ok=(not bairro or bairro in contexto)
    typ=_norm_lugar(item.get('type')); cl=_norm_lugar(item.get('class'))
    if typ in ('city','town','village','municipality','suburb','neighbourhood','postcode','administrative') or cl=='boundary':return None
    rua=_norm_lugar(e.get('logradouro'))
    rua=re.sub(r'\b(?:quadra|qd|qdr|qra|casa|cs|lote|lt|setor|set)\b.*','',rua).strip()
    rua=re.sub(r'^(rua|avenida|av|travessa|tv)\s+','',rua).strip()
    rua_ret=_norm_lugar(a.get('road') or a.get('pedestrian') or a.get('residential') or a.get('street') or '')
    rua_ok=bool(rua and rua_ret and len(rua)>3 and (rua in rua_ret or rua_ret in rua))
    numero=_norm_lugar(e.get('numero')); num_ret=_norm_lugar(a.get('house_number'))
    if rua_ok and numero and numero==num_ret and bairro_ok:return 100,'rua e número conferidos'
    if rua_ok and bairro_ok and not numero:
        return 75,'rua confirmada (sem número cadastrado)'
    # Rua sem o número pedido não equivale ao imóvel; tentar quadra ou setor.

    if not bairro_ok:return None
    def existe_rotulo(rotulos, valor):
        return bool(valor and re.search(r'\b(?:'+rotulos+r')\s*0*'+re.escape(str(valor).lower())+r'\b',contexto))
    q=existe_rotulo('quadra|qd|qdr|qra',quadra)
    st=existe_rotulo('setor|set',setor)
    ca=existe_rotulo('casa|cs|lote|lt',casa) or bool(casa and num_ret==str(casa).lower())
    if q and ca:return 92,'quadra e casa conferidas'
    if q and st:return 85,'quadra e setor conferidos; casa aproximada'
    if q:return 65,'quadra confirmada; casa aproximada'
    if st:return 55,'setor confirmado; quadra/casa aproximadas'
    return None

def _nominatim_reverse(lat, lon):
    ok,lat,lon=_valid_coord(lat,lon)
    if not ok:return None
    try:
        r=requests.get('https://nominatim.openstreetmap.org/reverse', params={
            'lat':lat,'lon':lon,'format':'jsonv2','addressdetails':1,'zoom':18
        }, headers={'User-Agent':'prosense_atdm/1.2 (field-service reverse-geocoder)'}, timeout=20)
        if not r.ok:return None
        j=r.json() or {}; a=j.get('address') or {}
        log=(a.get('road') or a.get('pedestrian') or a.get('residential') or a.get('footway') or a.get('path') or '')
        bairro=(a.get('suburb') or a.get('neighbourhood') or a.get('quarter') or a.get('city_district') or '')
        municipio=(a.get('city') or a.get('town') or a.get('municipality') or a.get('village') or '')
        uf=''
        iso=a.get('ISO3166-2-lvl4') or a.get('ISO3166-2-lvl6') or ''
        if isinstance(iso,str) and iso.startswith('BR-') and len(iso)>=5:uf=iso.split('-',1)[1][:2]
        return {'logradouro':log,'numero':a.get('house_number') or '','bairro':bairro,'cep':digits(a.get('postcode')),'municipio':municipio,'uf':uf,'display_name':j.get('display_name') or ''}
    except Exception:return None

# Geoapify operacional: sem limite diário artificial no ProSense.
# A franquia e a taxa de requisições continuam sujeitas às regras do Geoapify.

def _geo_operacional(e):
    key=os.getenv('GEOAPIFY_API_KEY','').strip()
    if not key:return None
    log=str(e.get('logradouro') or '').strip(); num=str(e.get('numero') or '').strip()
    bairro=str(e.get('bairro') or '').strip(); mun=str(e.get('municipio') or 'Teresina').strip()
    if not log:return None
    endereco=', '.join(x for x in [log,num,bairro,mun,str(e.get('uf') or 'PI'),'Brasil'] if x)
    try:
        r=requests.get('https://api.geoapify.com/v1/geocode/search',params={'text':endereco,'format':'geojson','filter':'countrycode:br','limit':5,'lang':'pt','apiKey':key},timeout=17)
        if not r.ok:return ('erro',f'Geoapify HTTP {r.status_code}')
        res=_geoapify_resultado(e,(r.json() or {}).get('features') or [])
        if res.get('status')=='numero_confirmado' and res.get('latitude') is not None:
            return ('exato',res)
        if res.get('status')=='rua_aproximada':
            return ('aproximado',res)
        return ('nao_localizado','Geoapify não confirmou rua e número')
    except Exception as exc:return ('erro',f'Geoapify: {str(exc)[:90]}')

def geocode_empresa(e):
    ok,lat,lon=_valid_coord(e.get('latitude'),e.get('longitude'))
    if ok:return lat,lon,'posição anteriormente salva'
    tentativa_geo=_geo_operacional(e)
    if tentativa_geo and tentativa_geo[0]=='exato':
        res=tentativa_geo[1]; lat=res['latitude'];lon=res['longitude']
        # Só persistimos quando há correspondência confirmada de rua e número.
        try:requests.patch(url(TABLE),headers=sb_headers('return=minimal'),params={'cnpj':f"eq.{e['cnpj']}"},json={'latitude':lat,'longitude':lon},timeout=15)
        except Exception:pass
        return lat,lon,'Geoapify: rua e número conferidos'
    # Buscas Nominatim ainda disponíveis para quadras/setores e instalações sem API.
    log=(e.get('logradouro') or '').strip();num=(e.get('numero') or '').strip()
    comp=(e.get('complemento') or '').strip();bairro=(e.get('bairro') or '').strip()
    mun=(e.get('municipio') or 'Teresina').strip();uf=(e.get('uf') or 'PI').strip()
    quadra,casa,setor=_identificar_endereco(e)
    base=[bairro,mun,uf,'Brasil']
    queries=[]
    # Priorizar endereço convencional antes de usar a quadra como indicação aproximada.
    if log and not re.match(r'^(?:quadra|qd|qdr|setor|set)\b',_norm_lugar(log)):
        queries.extend([', '.join(x for x in [log,num,bairro,mun,uf,'Brasil'] if x),
                        ', '.join(x for x in [log,bairro,mun,uf,'Brasil'] if x)])
    if quadra:
        queries.append(', '.join(x for x in [f'Quadra {quadra}',f'Setor {setor}' if setor else '',*base] if x))
        queries.append(', '.join(x for x in [f'Quadra {quadra}',*base] if x))
    if setor:queries.append(', '.join(x for x in [f'Setor {setor}',*base] if x))
    # Não pesquisar bairro sozinho: retornaria um ponto genérico enganoso.
    melhor=None;vistos=set()
    for q in queries:
        if q in vistos:continue
        vistos.add(q)
        for item in _nominatim_candidatos(q):
            ok,la,lo=_valid_coord(item.get('lat'),item.get('lon'))
            if not ok:continue
            evidencia=_evidencia_local(item,e,quadra,casa,setor)
            if evidencia and (melhor is None or evidencia[0]>melhor[0]):melhor=(evidencia[0],la,lo,evidencia[1])
        if melhor and melhor[0]>=100:break
    if melhor:
        _,lat,lon,descricao=melhor
        try:
            requests.patch(url(TABLE),headers=sb_headers('return=minimal'),params={'cnpj':f"eq.{e['cnpj']}"},json={'latitude':lat,'longitude':lon},timeout=15)
        except Exception:pass
        return lat,lon,descricao
    detalhe=(tentativa_geo[1] if tentativa_geo and isinstance(tentativa_geo[1],str) else ('Geoapify encontrou apenas a rua; número não confirmado' if tentativa_geo and tentativa_geo[0]=='aproximado' else ''))
    return None,None,(detalhe+' | ' if detalhe else '')+'Sem endereço exato verificável. Revisar manualmente: '+(' / '.join(queries[:2]) or ', '.join(base))

IMPORTAR_PLANILHA_HTML = '<!doctype html><html lang="pt-br"><head><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1"><title>ProSense | Importar planilha</title><style>\n:root{font:16px system-ui,Arial;color:#163047;background:#f2f6fa}*{box-sizing:border-box}body{margin:0}.wrap{max-width:1060px;margin:34px auto;padding:0 16px}.head{display:flex;justify-content:space-between;align-items:center;gap:12px}.logo{font-size:25px;font-weight:800;color:#155d9b}.panel{background:white;border:1px solid #dce6ee;border-radius:14px;padding:24px;margin-top:22px;box-shadow:0 8px 26px #093b5b0b}.grid{display:grid;grid-template-columns:1fr 1fr;gap:18px}label{display:block;font-weight:650;font-size:14px;margin-bottom:7px}input[type=text],input[type=file]{width:100%;padding:12px;border:1px solid #b9cbd9;border-radius:8px;font:inherit}button,.link{background:#1563a7;color:white;border:0;border-radius:8px;padding:12px 17px;font-size:15px;font-weight:700;cursor:pointer;text-decoration:none}button:disabled{opacity:.55;cursor:wait}.muted{color:#617b8f;font-size:13px}.actions{display:flex;align-items:center;gap:14px;flex-wrap:wrap;margin-top:16px}.status{padding:12px;border-radius:8px;margin-top:18px;background:#eaf2f8;white-space:pre-line}.error{background:#fff0ee;color:#a42e28}.success{background:#e7f5ec;color:#155e37}table{width:100%;border-collapse:collapse;font-size:13px}th,td{border-bottom:1px solid #e5edf2;padding:9px;text-align:left;vertical-align:top}th{background:#edf4fa} .tablewrap{overflow:auto}h1{font-size:22px;margin:0}h2{font-size:17px}input[type=checkbox]{transform:scale(1.2);margin-right:7px}@media(max-width:650px){.grid{grid-template-columns:1fr}.panel{padding:15px}}\n</style></head><body><div class="wrap"><div class="head"><div class="logo">ProSense · Importação</div><a href="/" class="link">Voltar ao painel</a></div><section class="panel"><h1>Importar empresas de uma planilha</h1><p class="muted">Importe dados prontos da planilha, sem consultar CNPJ na BrasilAPI. XLSX, XLSM, CSV ou TSV. Até 20 mil linhas e 20 MB.</p><div class="grid"><div><label for="lista">Nome da lista de destino</label><input id="lista" type="text" placeholder="Ex.: MEIs desenquadrados — Teresina" value="MEIs Desenquadrados"></div><div><label for="arquivo">Escolha a planilha</label><input id="arquivo" type="file" accept=".xlsx,.xlsm,.csv,.tsv,.txt"></div></div><p class="muted">Colunas como CNPJ, Razao Social, Nome Fantaria, Tipo Logradouro, Logradouro, Numero, Bairro, CEP, DDD1, Telefone1, E-mail e outros campos são reconhecidos. Os 33 campos originais ficam preservados no banco em <code>dados_importados</code>.</p><label style="font-weight:400"><input id="sobrescrever" type="checkbox">Substituir dados cadastrais já preenchidos nas empresas existentes (visitas, fotos, diagnósticos, observações e coordenadas nunca são substituídos).</label><div class="actions"><button id="validar">1. Conferir planilha</button><button id="salvar" disabled>2. Salvar empresas na lista</button><span id="passo" class="muted">Aguardando arquivo</span></div><div id="status" class="status" style="display:none"></div></section><section class="panel" id="previa" style="display:none"><h2>Prévia das primeiras empresas</h2><div class="tablewrap"><table><thead><tr><th>CNPJ</th><th>Razão social</th><th>Fantasia</th><th>Endereço</th><th>Bairro</th><th>Telefone</th></tr></thead><tbody id="rows"></tbody></table></div><h2>Problemas de importação (até 50)</h2><div id="erros" class="muted"></div></section></div><script>\nconst $=id=>document.getElementById(id);let validado=false;\nfunction mensagem(t,classe=\'\'){const e=$(\'status\');e.style.display=\'block\';e.className=\'status \'+classe;e.textContent=t}\nfunction escapeHtml(s){return String(s??\'\').replace(/[&<>"\']/g,c=>({\'&\':\'&amp;\',\'<\':\'&lt;\',\'>\':\'&gt;\',\'"\':\'&quot;\',"\'":\'&#39;\'}[c]))}\nfunction form(){const f=new FormData();f.append(\'arquivo\',$(\'arquivo\').files[0]);f.append(\'nome_lista\',$(\'lista\').value.trim());f.append(\'sobrescrever\',$(\'sobrescrever\').checked?\'true\':\'false\');return f}\nfunction ocupado(v){$(\'validar\').disabled=v;$(\'salvar\').disabled=v||!validado}\n$(\'arquivo\').addEventListener(\'change\',()=>{validado=false;$(\'salvar\').disabled=true;$(\'previa\').style.display=\'none\';$(\'passo\').textContent=\'Arquivo alterado: confira novamente\'});\n$(\'validar\').onclick=async()=>{if(!$(\'arquivo\').files.length)return mensagem(\'Selecione uma planilha.\',\'error\');validado=false;ocupado(true);mensagem(\'Verificando campos e CNPJs da planilha...\');try{const r=await fetch(\'/api/planilha/validar\',{method:\'POST\',body:form()});const j=await r.json();if(!r.ok)throw Error(j.erro||\'Falha na validação\');$(\'previa\').style.display=\'block\';$(\'rows\').innerHTML=j.amostra.map(x=>\'<tr>\'+[x.cnpj,x.razao_social,x.nome_fantasia,x.logradouro,x.bairro,x.telefone].map(v=>\'<td>\'+escapeHtml(v)+\'</td>\').join(\'\')+\'</tr>\').join(\'\');$(\'erros\').textContent=j.erros.length?j.erros.map(e=>\'Linha \'+e.linha+\': \'+e.erro).join(\' · \'):\'Nenhum erro detectado\';validado=j.validas>0;mensagem(j.validas+\' empresas válidas; \'+j.ignoradas+\' linhas ignoradas. Nenhum cadastro salvo ainda.\',\'success\');$(\'passo\').textContent=\'Conferência concluída\'}catch(e){mensagem(e.message,\'error\')}finally{ocupado(false)}};\n$(\'salvar\').onclick=async()=>{if(!validado)return;const destino=$(\'lista\').value.trim();if(!destino)return mensagem(\'Informe o nome da lista.\',\'error\');ocupado(true);mensagem(\'Gravando empresas no Supabase e vinculando à lista. Não feche esta guia durante o envio.\');try{const r=await fetch(\'/api/importar-planilha\',{method:\'POST\',body:form()});const j=await r.json();if(!r.ok)throw Error(j.erro||\'Falha ao importar\');mensagem(\'Importação concluída!\\nLista: \'+j.lista.nome+\'\\nEmpresas salvas/vinculadas: \'+j.salvas+\'\\nNovas empresas: \'+j.novas+\'\\nJá cadastradas: \'+j.ja_cadastradas+\'\\nLinhas ignoradas: \'+j.ignoradas,\'success\');$(\'passo\').textContent=\'Salvo com sucesso\';validado=false}catch(e){mensagem(e.message,\'error\')}finally{ocupado(false)}};\n</script></body></html>\n'

@app.route('/')
def home(): return render_template('index.html')


# Google Maps comparison mode: no automatic writes to Supabase and no mass calls.
# Only the 20 CNPJs explicitly chosen by the user are tested.
@app.route('/mapa-google')
def mapa_google():
    key=os.getenv('GOOGLE_MAPS_BROWSER_KEY','').strip()
    if not key:
        return render_template('mapa_google.html',google_key='',cnpjs='',configured=False)
    requested=[fmt(x) for x in (request.args.get('cnpjs') or '').split(',') if len(digits(x))==14]
    return render_template('mapa_google.html',google_key=key,cnpjs=','.join(dict.fromkeys(requested[:20])),configured=True)

@app.route('/health')
def health(): return jsonify({'ok':True,'database':'supabase','configured':bool(SUPABASE_URL and SUPABASE_KEY),'table':TABLE})

@app.route('/api/listas')
def listas():
    try:
        ls=_fetch_paginated(LIST_TABLE,{'select':'*','order':'criado_em.desc'})
        links=_fetch_paginated(LINK_TABLE,{'select':'lista_id'})
        counts={}
        for x in links: counts[x['lista_id']]=counts.get(x['lista_id'],0)+1
        for x in ls:x['total']=counts.get(x['id'],0)
        return jsonify(ls)
    except Exception as e:return jsonify({'erro':str(e)}),500

@app.route('/api/lista/<int:lista_id>',methods=['PUT','DELETE'])
def alterar_lista(lista_id):
    if request.method=='DELETE':
        # Exclui somente a lista e seus vínculos (cascade). As empresas permanecem no banco.
        r=requests.delete(url(LIST_TABLE),headers=sb_headers('return=representation'),params={'id':f'eq.{lista_id}'},timeout=25)
        if not r.ok:return jsonify({'erro':r.text}),500
        return jsonify({'ok':True,'lista_id':lista_id,'empresas_preservadas':True})
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
            rows=_empresas_da_lista(lista_id)
        else:
            rows=_fetch_paginated(TABLE,{'select':'*'})
            rows.sort(key=lambda x:((x.get('nome_fantasia') or '').casefold() or '\uffff',(x.get('razao_social') or '').casefold()))
        # Marca empresas que já possuem pelo menos um diagnóstico salvo.
        # Se a tabela de diagnósticos estiver temporariamente indisponível, a lista continua funcionando.
        try:
            diag_cnpjs={x.get('cnpj') for x in _fetch_paginated(DIAG_TABLE,{'select':'cnpj'},timeout=20)}
        except Exception:
            diag_cnpjs=set()
        for x in rows:x['diagnosticado']=x.get('cnpj') in diag_cnpjs
        # Situação comercial FS, independente do diagnóstico Sebrae.
        try:
            pm={x.get('cnpj'):x for x in _fetch_paginated(PROSP_TABLE,{'select':'cnpj,status,interesse,ultimo_contato,envios_count,ultimo_canal'},timeout=20)}
        except Exception: pm={}
        try:
            am={x.get('cnpj'):x for x in _fetch_paginated(ATEND_CONTABIL_TABLE,{'select':'cnpj,status,data,observacao'},timeout=20)}
        except Exception: am={}
        for x in rows:
            px=pm.get(x.get('cnpj'),{})
            x['prospeccao_status']=px.get('status') or 'nao_abordado'
            x['prospeccao_interesse']=px.get('interesse') or ''
            x['prospeccao_ultimo_contato']=px.get('ultimo_contato')
            x['prospeccao_envios_count']=int(px.get('envios_count') or 0)
            x['prospeccao_ultimo_canal']=px.get('ultimo_canal') or ''
            ax=am.get(x.get('cnpj'),{})
            x['atendimento_contabil_status']=ax.get('status') or 'nao_iniciado'
            x['atendimento_contabil_data']=ax.get('data')
            x['atendimento_contabil_observacao']=ax.get('observacao') or ''
        return jsonify(rows)
    except Exception as e:return jsonify({'erro':str(e)}),500

@app.route('/api/cnpj/consultar/<path:cnpj>')
def consultar_cnpj_unitario(cnpj):
    c=digits(cnpj)
    if len(c)!=14:return jsonify({'erro':'Informe um CNPJ válido com 14 dígitos.'}),400
    try:
        d=fetch_cnpj(c)
        existente=get_one(c)
        return jsonify({'ok':True,'empresa':d,'ja_cadastrada':bool(existente)})
    except Exception as e:
        return jsonify({'erro':str(e)}),502

@app.route('/api/importar-unitario',methods=['POST'])
def importar_unitario():
    b=request.json or {}; c=digits(b.get('cnpj')); lista_id=b.get('lista_id')
    if len(c)!=14:return jsonify({'erro':'Informe um CNPJ válido com 14 dígitos.'}),400
    try: lista_id=int(lista_id)
    except Exception:return jsonify({'erro':'Selecione a lista de destino.'}),400
    try:
        lr=requests.get(url(LIST_TABLE),headers=sb_headers(),params={'id':f'eq.{lista_id}','select':'id,nome','limit':'1'},timeout=20)
        lr.raise_for_status(); ld=lr.json()
        if not ld:return jsonify({'erro':'Lista de destino não encontrada.'}),404
        d=fetch_cnpj(c)
        origem='endereco_cadastral'
        if b.get('usar_localizacao_atual'):
            try: lat=float(b.get('latitude')); lon=float(b.get('longitude'))
            except Exception:return jsonify({'erro':'Não foi possível ler a localização atual.'}),400
            ok,lat,lon=_valid_coord(lat,lon)
            if not ok:return jsonify({'erro':'A localização capturada está fora da área válida.'}),400
            d['latitude']=lat; d['longitude']=lon; origem='localizacao_atual'
        upsert_empresa(d); link_empresa(lista_id,d['cnpj'])
        salvo=get_one(d['cnpj']) or d
        return jsonify({'ok':True,'empresa':salvo,'lista':ld[0],'origem_localizacao':origem})
    except Exception as e:
        return jsonify({'erro':str(e)}),500

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

@app.route('/api/planilha/validar',methods=['POST'])
def validar_planilha():
    fs=request.files.get('arquivo')
    if not fs or not fs.filename:return jsonify({'erro':'Selecione uma planilha.'}),400
    try:
        rows,erros=_parse_planilha(fs)
        return jsonify({'ok':True,'validas':len(rows),'ignoradas':len(erros),
                        'amostra':[{'cnpj':e['cnpj'],'razao_social':e['razao_social'],
                                   'nome_fantasia':e['nome_fantasia'],'logradouro':e['logradouro'],
                                   'bairro':e['bairro'],'telefone':e['telefone']}
                                  for e in rows[:8]],'erros':erros[:50],
                        'colunas_identificadas':list(rows[0]['dados_importados']) if rows else []})
    except ValueError as e:return jsonify({'erro':str(e)}),400
    except Exception as e:return jsonify({'erro':str(e)}),500

@app.route('/api/importar-planilha',methods=['POST'])
def importar_planilha():
    fs=request.files.get('arquivo'); nome=(request.form.get('nome_lista') or '').strip()
    if not nome:return jsonify({'erro':'Informe o nome da lista.'}),400
    if not fs or not fs.filename:return jsonify({'erro':'Selecione uma planilha.'}),400
    try:
        # Conferência apenas em /api/planilha/validar; esta rota salva de verdade.
        rows,erros=_parse_planilha(fs)
        if not rows:return jsonify({'erro':'Nenhuma linha válida com CNPJ foi encontrada.','erros':erros[:50]}),400
        sobrescrever=request.form.get('sobrescrever')=='true'
        merged,novos,existentes=_merge_planilha_com_cadastro(rows,sobrescrever)
        lista=get_or_create_list(nome)
        _upsert_lote(merged)
        _link_lote(lista['id'],[x['cnpj'] for x in rows])
        return jsonify({'ok':True,'lista':lista,'lidas':len(rows)+len(erros),
                        'salvas':len(rows),'novas':novos,'ja_cadastradas':existentes,
                        'ignoradas':len(erros),'erros':erros[:50]})
    except ValueError as e:return jsonify({'erro':str(e)}),400
    except Exception as e:return jsonify({'erro':str(e)}),500

@app.route('/importar-planilha')
def tela_importacao_planilha():
    return render_template_string(IMPORTAR_PLANILHA_HTML)

@app.route('/api/geolocalizacao/reverso',methods=['POST'])
def geolocalizacao_reverso():
    data=request.json or {}
    ok,lat,lon=_valid_coord(data.get('latitude'),data.get('longitude'))
    if not ok:return jsonify({'erro':'Coordenadas inválidas ou fora da área esperada.'}),400
    endereco=_nominatim_reverse(lat,lon)
    if endereco is None:return jsonify({'erro':'A localização foi obtida, mas não foi possível identificar o endereço agora.'}),502
    return jsonify({'ok':True,'latitude':lat,'longitude':lon,'endereco':endereco})

@app.route('/api/empresa/<path:cnpj>',methods=['PUT'])
def editar(cnpj):
    data=request.json or {};allowed=['razao_social','nome_fantasia','situacao','cnae','logradouro','numero','complemento','bairro','cep','municipio','uf','telefone','email','status_visita','dia_mes_sse','observacoes','latitude','longitude']
    patch={k:data[k] for k in allowed if k in data};status=patch.get('status_visita');now=datetime.now(timezone.utc).isoformat()
    old=get_one(cnpj)
    if not old:return jsonify({'erro':'CNPJ não encontrado'}),404
    if status is not None:
        try:status=int(status);patch['status_visita']=status
        except:return jsonify({'erro':'status_visita inválido'}),400
        if status==1 and not old.get('primeira_visita_em'):patch['primeira_visita_em']=now
        if status==2:
            if not old.get('primeira_visita_em'):patch['primeira_visita_em']=now
            if not old.get('segunda_visita_em'):patch['segunda_visita_em']=now
    tem_coords=('latitude' in patch and 'longitude' in patch)
    if tem_coords:
        ok,lat,lon=_valid_coord(patch.get('latitude'),patch.get('longitude'))
        if not ok:return jsonify({'erro':'Latitude/longitude inválidas'}),400
        patch['latitude']=lat;patch['longitude']=lon
    elif any(k in patch for k in ['logradouro','numero','bairro','cep','municipio','uf']):
        patch.update({'latitude':None,'longitude':None})
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

@app.route('/api/mapa/buscar-endereco',methods=['GET'])
def buscar_endereco_no_mapa():
    """Pesquisa pontual por ação do usuário; não escreve coordenadas automaticamente."""
    q=(request.args.get('q') or '').strip()
    if len(q)<3 or len(q)>180:return jsonify({'erro':'Informe entre 3 e 180 caracteres.'}),400
    query=q if 'teresina' in _norm_lugar(q) else q+', Teresina, Piauí, Brasil'
    resultados=[]
    try:
        # Pesquisa sob demanda: prioriza Geoapify e deixa Nominatim como alternativa.
        key=os.getenv('GEOAPIFY_API_KEY','').strip()
        if key:
            gr=requests.get('https://api.geoapify.com/v1/geocode/search',params={'text':query,'format':'geojson','filter':'countrycode:br','limit':5,'lang':'pt','apiKey':key},timeout=15)
            if gr.ok:
                for f in (gr.json() or {}).get('features',[])[:5]:
                    prop=f.get('properties') or {}; coords=(f.get('geometry') or {}).get('coordinates') or []
                    if len(coords)!=2:continue
                    valid,lat,lon=_valid_coord(coords[1],coords[0])
                    if valid:resultados.append({'nome':prop.get('formatted') or query,'latitude':lat,'longitude':lon})
                if resultados:return jsonify({'resultados':resultados,'fonte':'Geoapify'})
        # Endereço encontrado ainda depende de confirmação visual.
        for item in _nominatim_candidatos(query)[:5]:
            ok,lat,lon=_valid_coord(item.get('lat'),item.get('lon'))
            if not ok:continue
            addr=item.get('address') or {}
            cidade=_norm_lugar(addr.get('city') or addr.get('town') or addr.get('municipality') or '')
            if cidade and cidade!='teresina':continue
            resultados.append({'nome':item.get('display_name') or query,'latitude':lat,'longitude':lon})
        return jsonify({'resultados':resultados})
    except Exception as ex:return jsonify({'erro':str(ex)}),502

# Geoapify: comparação opt-in, nunca grava coordenadas, chave apenas no backend.
GEOAPIFY_KEY=os.getenv('GEOAPIFY_API_KEY','').strip()

@app.route('/mapa-geoapify')
def mapa_geoapify():
    return render_template('mapa_geoapify.html')

@app.route('/api/geoapify/status')
def geoapify_status():
    return jsonify({'configurado':bool(GEOAPIFY_KEY),'limite_teste':20})

def _geoapify_addr(row):
    return ', '.join(str(row.get(k) or '').strip() for k in ('logradouro','numero','complemento','bairro','municipio','uf') if str(row.get(k) or '').strip())+', Brasil'

def _geoapify_resultado(e,features):
    endereco=_geoapify_addr(e)
    expected=str(e.get('numero') or '').strip()
    requested_road=_norm_lugar(e.get('logradouro'))
    requested_road=re.sub(r'^(rua|av|avenida|travessa|tv)\s+','',requested_road).strip()
    requested_city=_norm_lugar(e.get('municipio') or 'Teresina')
    best=None
    for f in features[:5]:
        prop=f.get('properties') or {}; geom=f.get('geometry') or {}; coords=geom.get('coordinates') or []
        if len(coords)!=2:continue
        valid,la,lo=_valid_coord(coords[1],coords[0])
        if not valid:continue
        city=_norm_lugar(prop.get('city') or prop.get('municipality') or '')
        if city and requested_city and city!=requested_city:continue
        road=_norm_lugar(prop.get('street') or '')
        road_match=bool(requested_road and road and (requested_road==road or (len(requested_road)>5 and requested_road in road)))
        found_num=str(prop.get('housenumber') or '').strip()
        number_match=bool(expected and found_num and digits(found_num)==digits(expected))
        confidence=float((prop.get('rank') or {}).get('confidence') or 0)
        # Número exato depende também de correspondência de rua.
        status='numero_confirmado' if road_match and number_match else ('rua_aproximada' if road_match else 'nao_confirmado')
        score=(100 if status=='numero_confirmado' else 50 if status=='rua_aproximada' else 0)+min(confidence,1)*10
        result={'status':status,'endereco_consultado':endereco,'encontrado':prop.get('formatted') or '', 'latitude':la,'longitude':lo,'numero_informado':expected,'numero_encontrado':found_num,'rua_encontrada':prop.get('street') or '', 'confidence':round(confidence,2),'fonte':'Geoapify'}
        if best is None or score>best[0]:best=(score,result)
    return best[1] if best else {'status':'nao_localizado','endereco_consultado':endereco,'encontrado':'','latitude':None,'longitude':None,'fonte':'Geoapify'}

@app.route('/api/geoapify/comparar',methods=['POST'])
def comparar_geoapify():
    if not GEOAPIFY_KEY:return jsonify({'erro':'Configure GEOAPIFY_API_KEY no Railway para liberar o teste.'}),503
    body=request.get_json(silent=True) or {}; cnpjs=body.get('cnpjs')
    if not isinstance(cnpjs,list) or not (1<=len(cnpjs)<=20):return jsonify({'erro':'Escolha entre 1 e 20 empresas.'}),400
    ids=list(dict.fromkeys(fmt(str(c)) for c in cnpjs if len(digits(str(c)))==14))
    if not ids:return jsonify({'erro':'Nenhum CNPJ válido.'}),400
    out=[]
    for c in ids:
        try:
            e=get_one(c)
            if not e:
                out.append({'cnpj':c,'status':'erro','erro':'CNPJ não cadastrado'});continue
            params={'text':_geoapify_addr(e),'format':'geojson','filter':'countrycode:br','limit':5,'lang':'pt','apiKey':GEOAPIFY_KEY}
            r=requests.get('https://api.geoapify.com/v1/geocode/search',params=params,timeout=18)
            if not r.ok:
                out.append({'cnpj':c,'status':'erro','erro':f'Geoapify HTTP {r.status_code}'});continue
            res=_geoapify_resultado(e,(r.json() or {}).get('features') or [])
            res.update({'cnpj':c,'nome_fantasia':e.get('nome_fantasia') or e.get('razao_social') or c})
            out.append(res)
        except Exception as exc:out.append({'cnpj':c,'status':'erro','erro':str(exc)[:160]})
    return jsonify({'resultados':out,'consultadas':len(out),'gravadas':0})

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
              'complemento':e.get('complemento') or '',
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


@app.route('/api/mapa/relocalizar',methods=['POST'])
def relocalizar_empresas():
    """Ação explícita do usuário: limpar pontos antigos para nova busca. Não mexe nos outros dados."""
    cnpjs=(request.json or {}).get('cnpjs') or []
    if not isinstance(cnpjs,list) or not cnpjs or len(cnpjs)>150:
        return jsonify({'erro':'Selecione entre 1 e 150 empresas.'}),400
    atualizados=[];erros=[]
    for c in dict.fromkeys(fmt(v) for v in cnpjs if len(digits(v))==14):
        try:
            r=requests.patch(url(TABLE),headers=sb_headers('return=minimal'),params={'cnpj':f'eq.{c}'},json={'latitude':None,'longitude':None},timeout=20)
            if r.ok:atualizados.append(c)
            else:erros.append({'cnpj':c,'erro':r.text[:160]})
        except Exception as ex:erros.append({'cnpj':c,'erro':str(ex)})
    return jsonify({'ok':not erros,'atualizados':atualizados,'erros':erros})

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
    allowed=['status','interesse','ultimo_contato','proximo_contato','envios_count','ultimo_canal']
    payload={'cnpj':c,**{k:b.get(k) for k in allowed if k in b}}
    r=requests.post(url(PROSP_TABLE),headers=sb_headers('resolution=merge-duplicates,return=representation'),params={'on_conflict':'cnpj'},json=payload,timeout=20)
    if not r.ok:return jsonify({'erro':r.text}),500
    return jsonify(r.json()[0] if r.json() else payload)

@app.route('/api/prospeccao/enviar/<path:cnpj>',methods=['POST'])
def registrar_envio_prospeccao(cnpj):
    c=fmt(cnpj); b=request.json or {}; canal=(b.get('canal') or '').strip().lower()
    if canal not in ('whatsapp','sms'): return jsonify({'erro':'Canal inválido.'}),400
    try:
        r=requests.get(url(PROSP_TABLE),headers=sb_headers(),params={'cnpj':f'eq.{c}','select':'envios_count','limit':'1'},timeout=20)
        if not r.ok:
            if 'envios_count' in r.text and ('does not exist' in r.text or '42703' in r.text):
                return jsonify({'erro':'Banco ainda sem a coluna envios_count. Execute migracao_v44.sql no Supabase e tente novamente.'}),500
            return jsonify({'erro':r.text}),500
        d=r.json(); atual=int((d[0].get('envios_count') if d else 0) or 0)
        now=datetime.now(timezone.utc).isoformat(); payload={'cnpj':c,'status':'contato_realizado','interesse':b.get('interesse') or '','ultimo_contato':now,'envios_count':atual+1,'ultimo_canal':canal,'atualizado_em':now}
        r=requests.post(url(PROSP_TABLE),headers=sb_headers('resolution=merge-duplicates,return=representation'),params={'on_conflict':'cnpj'},json=payload,timeout=20)
        if not r.ok:return jsonify({'erro':r.text}),500
        return jsonify({'ok':True,'status':'contato_realizado','ultimo_contato':now,'envios_count':atual+1,'ultimo_canal':canal})
    except Exception as e:return jsonify({'erro':str(e)}),500

@app.route('/api/atendimento-contabil/<path:cnpj>',methods=['GET','PUT'])
def atendimento_contabil(cnpj):
    c=fmt(cnpj)
    if request.method=='GET':
        r=requests.get(url(ATEND_CONTABIL_TABLE),headers=sb_headers(),params={'cnpj':f'eq.{c}','select':'*','limit':'1'},timeout=20)
        if not r.ok:return jsonify({'erro':r.text}),500
        d=r.json(); return jsonify(d[0] if d else {'cnpj':c,'status':'nao_iniciado'})
    b=request.json or {}; allowed_status={'nao_iniciado','agendado','em_atendimento','proposta_enviada','fidelizado','retorno','nao_convertido'}
    status=(b.get('status') or 'nao_iniciado').strip()
    if status not in allowed_status:return jsonify({'erro':'Status inválido.'}),400
    payload={'cnpj':c,'status':status,'data':b.get('data') or None,'observacao':b.get('observacao') or '','atualizado_em':datetime.now(timezone.utc).isoformat()}
    r=requests.post(url(ATEND_CONTABIL_TABLE),headers=sb_headers('resolution=merge-duplicates,return=representation'),params={'on_conflict':'cnpj'},json=payload,timeout=20)
    if not r.ok:return jsonify({'erro':r.text}),500
    return jsonify(r.json()[0] if r.json() else payload)

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
        rows=_empresas_da_lista(lista_id)
    else:
        rows=_fetch_paginated(TABLE,{'select':'*'})
        rows.sort(key=lambda x:((x.get('nome_fantasia') or '').casefold() or '\uffff',(x.get('razao_social') or '').casefold()))
    out=io.StringIO();w=csv.writer(out,delimiter=';')
    if rows:
        keys=list(rows[0].keys());w.writerow(keys)
        for row in rows:w.writerow([row.get(k,'') for k in keys])
    return Response('\ufeff'+out.getvalue(),mimetype='text/csv',headers={'Content-Disposition':'attachment; filename=prosense_atdm.csv'})

if __name__=='__main__':app.run(host='0.0.0.0',port=int(os.getenv('PORT','5000')),debug=os.getenv('FLASK_DEBUG')=='1')