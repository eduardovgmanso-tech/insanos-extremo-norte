# -*- coding: utf-8 -*-
"""
Gera, a partir da planilha Insanos.xlsx:
  - "Extremo Norte - SP3 - Analise.html"        -> painel geral para compartilhar (sem telefones)
  - "Extremo Norte - SP3 - Analise (ADM).html"  -> seu painel, com botão de WhatsApp (NÃO compartilhar)
  - pasta "Individuais"                          -> um HTML por integrante, só com os dados dele
O painel geral e o ADM saem protegidos: entra quem digitar o E-MAIL de um integrante de AUTORIZADOS
(coluna de e-mail da aba Dados Integrantes) ou uma senha de SENHAS_GERAL.
O painel geral e o ADM também abrem com as senhas de SENHAS_GERAL.
Tudo na mesma pasta: Insanos.xlsx, maxresdefault.jpg, template_extremo_norte.html e este script.
"""
import base64, datetime, json, os, shutil, sys, tempfile
import re

PASTA     = os.path.dirname(os.path.abspath(__file__))
PLANILHA  = os.path.join(PASTA, "Insanos.xlsx")
IMAGEM    = os.path.join(PASTA, "maxresdefault.jpg")
TEMPLATE  = os.path.join(PASTA, "template_extremo_norte.html")
SAIDA     = os.path.join(PASTA, "Extremo Norte - SP3 - Analise.html")
SAIDA_ADM = os.path.join(PASTA, "Extremo Norte - SP3 - Analise (ADM).html")
PASTA_IND = os.path.join(PASTA, "Individuais")

# Quem pode abrir o painel geral e o ADM (apelido igual ao da aba Dados Integrantes).
# A senha de cada um é o CPF cadastrado na coluna "CPF" dessa aba (só os números).
AUTORIZADOS = ["Queijo", "Brawl", "Maycon", "Rick", "Edu", "Demonão"]

# Quem abre SÓ o painel geral com o próprio CPF (não abre o ADM, que tem telefones e pagamentos).
AUTORIZADOS_SO_GERAL = []

# Senha(s) extra(s) que abrem o painel geral E o ADM, além dos CPFs acima.
SENHAS_GERAL = [x for x in [os.environ.get("SENHA_COLETIVA", "")] if x]   # do cofre do GitHub
ITERACOES = 600_000
# HTMLs individuais (pasta Individuais): substituídos pela imagem gerada no ADM. True volta a gerar os arquivos.
GERAR_INDIVIDUAIS = False
INFO_LOGIN = ""   # linha da tela de CPF: última atualização e última verificação dos pagamentos
ULTIMA_ALTERACAO_TXT = ""   # na nuvem: "o que foi alterado · por quem · quando" (sem dados pessoais de outros irmãos)
# Aviso de manutenção: aparece no site enquanto existir o arquivo MANUTENCAO.txt nesta pasta
# (LigarManutencao.bat cria o arquivo; DesligarManutencao.bat apaga)
AVISO_MANUTENCAO = ('<div class="manut" role="status">🛠️ <b>Site em manutenção</b>: os dados estão sendo analisados '
                    'e podem conter erros durante a manutenção.</div>')
MANUT_HTML = ""
# Registro de acessos: o site manda cada entrada (ou tentativa sem acesso) para a planilha Google abaixo,
# e o .py copia os registros para a aba "Acessos" da Insanos.xlsx. Deixe ACESSOS_URL = "" para desligar.
ACESSOS_URL = os.environ.get("ACESSOS_URL", "")       # do cofre do GitHub
ACESSOS_CHAVE = os.environ.get("ACESSOS_CHAVE", "")   # do cofre do GitHub
# quem vê a tela "Acessos" no ADM (pelo e-mail da planilha). Os demais do comando não veem nem conseguem ler.
ACESSOS_SOMENTE = ["Edu"]
PLANILHA_URL = ""      # na nuvem: link da planilha do Google (atalho 📗 Planilha, trancado como a tela Acessos)
EMAIL_APELIDO = {}
NUVEM = None      # na nuvem: {"url": recebedor, "chave": ...} -> os lançamentos vão direto para a planilha do Google
EMAIL_APELIDO_INV = {}
# Fotos espalhadas no fundo da tela de CPF do ADM: arquivos da pasta "fotos_login" (webp, jpg ou png, até 6)
FOTOS_LOGIN = ""
POS_FOTOS = [("6%", "4%", -8), ("8%", "68%", 7), ("58%", "2%", 6), ("62%", "70%", -6), ("34%", "-4%", -3), ("36%", "80%", 4)]

def modelos_fly():
    """Fundos dos flys (sem as informações) e a fonte, da pasta "modelos_fly" (vão só no ADM)."""
    pasta = os.path.join(PASTA, "modelos_fly")
    out = {"fundos": {}, "fonte": ""}
    if not os.path.isdir(pasta):
        return out
    for a in sorted(os.listdir(pasta)):
        base_nome, ext = os.path.splitext(a.lower())
        with open(os.path.join(pasta, a), "rb") as f:
            b64 = base64.b64encode(f.read()).decode()
        if ext in (".jpg", ".jpeg", ".png", ".webp"):
            mime = {".jpg": "jpeg", ".jpeg": "jpeg"}.get(ext, ext[1:])
            out["fundos"][base_nome] = f"data:image/{mime};base64,{b64}"
        elif ext in (".ttf", ".otf", ".woff", ".woff2") and not out["fonte"]:
            out["fonte"] = f"data:font/{ext[1:]};base64,{b64}"
    return out

def fotos_galeria():
    """Fotos da pasta "fotos_galeria" para a tela Fotos do ADM (vão dentro do arquivo criptografado)."""
    pasta = os.path.join(PASTA, "fotos_galeria")
    if not os.path.isdir(pasta):
        return []
    fotos = []
    for a in sorted(x for x in os.listdir(pasta) if x.lower().endswith((".jpg", ".jpeg", ".png", ".webp"))):
        ext = a.rsplit(".", 1)[-1].lower()
        mime = {"jpg": "jpeg", "jpeg": "jpeg"}.get(ext, ext)
        with open(os.path.join(pasta, a), "rb") as f:
            fotos.append({"n": a, "d": f"data:image/{mime};base64," + base64.b64encode(f.read()).decode()})
    return fotos

def fotos_login_html():
    pasta = os.path.join(PASTA, "fotos_login")
    if not os.path.isdir(pasta):
        return ""
    arqs = sorted(a for a in os.listdir(pasta) if a.lower().endswith((".webp", ".jpg", ".jpeg", ".png")))[:len(POS_FOTOS)]
    partes = []
    for a, (top, left, rot) in zip(arqs, POS_FOTOS):
        ext = a.rsplit(".", 1)[-1].lower()
        mime = {"jpg": "jpeg", "jpeg": "jpeg"}.get(ext, ext)
        with open(os.path.join(pasta, a), "rb") as f:
            b64 = base64.b64encode(f.read()).decode()
        partes.append(f'<img class="fl" alt="" src="data:image/{mime};base64,{b64}" style="top:{top};left:{left};transform:rotate({rot}deg)">')
    return '<div class="fotos-login" aria-hidden="true">' + "".join(partes) + "</div>" if partes else ""

# Vencimento da mensalidade: todo dia 12. Se cair em sábado, domingo ou feriado, vai para o primeiro dia útil seguinte.
VENCIMENTO_DIA = 12
# Feriados locais extras (dia/mês), além dos nacionais. Ex.: ["25/01", "09/07"] (cidade e estado de São Paulo)
FERIADOS_EXTRAS = []

def _pascoa(ano):
    """Domingo de Páscoa (algoritmo de Meeus/Jones/Butcher)."""
    a, b, c = ano % 19, ano // 100, ano % 100
    d, e = b // 4, b % 4
    f = (b + 8) // 25
    g = (b - f + 1) // 3
    h = (19 * a + b - d - g + 15) % 30
    i, k = c // 4, c % 4
    l = (32 + 2 * e + 2 * i - h - k) % 7
    m = (a + 11 * h + 22 * l) // 451
    mes = (h + l - 7 * m + 114) // 31
    dia = ((h + l - 7 * m + 114) % 31) + 1
    return datetime.date(ano, mes, dia)

def feriados(ano):
    """Feriados nacionais (e bancários de Carnaval) do ano, mais os FERIADOS_EXTRAS."""
    fixos = ["01/01", "21/04", "01/05", "07/09", "12/10", "02/11", "15/11", "20/11", "25/12"] + list(FERIADOS_EXTRAS)
    dias = set()
    for dm in fixos:
        try:
            d, m = dm.split("/"); dias.add(datetime.date(ano, int(m), int(d)))
        except Exception:
            pass
    p = _pascoa(ano)
    for delta in (-48, -47, -2, 60):   # Carnaval (segunda e terça), Sexta-feira Santa, Corpus Christi
        dias.add(p + datetime.timedelta(days=delta))
    return dias

def vencimento(ano, mes):
    """Data de vencimento do mês: dia VENCIMENTO_DIA ou o primeiro dia útil depois dele."""
    d = datetime.date(ano, mes, VENCIMENTO_DIA)
    while d.weekday() >= 5 or d in feriados(d.year):   # 5 = sábado, 6 = domingo
        d += datetime.timedelta(days=1)
    return d

MESES = {"jan":1,"fev":2,"mar":3,"abr":4,"mai":5,"jun":6,"jul":7,"ago":8,"set":9,"out":10,"nov":11,"dez":12}

def falha(msg):
    print("\n[ERRO] " + msg)
    sys.exit(1)

def texto(v):
    return str(v).strip() if v is not None else ""

def data_iso(v, ano_padrao):
    if isinstance(v, (datetime.datetime, datetime.date)):
        return v.strftime("%Y-%m-%d")
    s = texto(v).lower()
    try:
        dia, mes = s.replace("/", "-").split("-")[:2]
        mes = MESES[mes[:3]] if not mes.isdigit() else int(mes)
        return f"{ano_padrao}-{mes:02d}-{int(dia):02d}"
    except Exception:
        return None

LOGIN = r"""<!DOCTYPE html>
<html lang="pt-BR"><head><meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1, viewport-fit=cover">
<title>__TITULO__</title>
<link rel="manifest" href="__MANIFEST__">
<meta name="theme-color" content="#0b0b0b">
<link rel="icon" href="icon-192.png"><link rel="apple-touch-icon" href="icon-192.png">
<meta name="apple-mobile-web-app-capable" content="yes"><meta name="mobile-web-app-capable" content="yes">
<meta name="apple-mobile-web-app-status-bar-style" content="black-translucent">
<script>if("serviceWorker" in navigator&&location.protocol==="https:")navigator.serviceWorker.register("sw.js").catch(()=>{});</script>
<style>
:root{--osso:#ECE7DC;--osso2:#BDB7AA;--sangue:#A3302F;box-sizing:border-box}
*,*::before,*::after{box-sizing:inherit}
html,body{height:100%;margin:0}
body{background:#000;color:var(--osso);font-family:"Barlow","Segoe UI",Roboto,Arial,sans-serif;display:grid;place-items:center;padding:20px}
body::before{content:"";position:fixed;inset:0;z-index:-1;background:url("data:image/jpeg;base64,__IMG__") right center/cover no-repeat;opacity:.25}
form{width:min(380px,100%);background:rgba(236,231,220,.10);backdrop-filter:blur(14px);-webkit-backdrop-filter:blur(14px);border:1px solid rgba(236,231,220,.18);border-radius:14px;padding:28px}
h1{font-family:"Big Shoulders Display","Arial Narrow",Impact,sans-serif;font-weight:900;font-size:40px;line-height:.95;margin:0 0 6px}
p{color:var(--osso2);font-size:14px;margin:0 0 20px}
label{display:block;font-size:13px;color:var(--osso2);margin-bottom:6px}
input{width:100%;background:rgba(0,0,0,.45);color:var(--osso);border:1px solid rgba(236,231,220,.3);border-radius:8px;padding:12px;font:inherit;font-size:18px;letter-spacing:.06em}
button{width:100%;min-height:48px;margin-top:14px;background:var(--osso);color:#111;border:0;border-radius:8px;padding:12px;font:inherit;font-weight:600;font-size:16px;cursor:pointer}
button:disabled{opacity:.6;cursor:wait}
input:focus-visible,button:focus-visible{outline:2px solid var(--osso);outline-offset:2px}
#erro{color:#E07472;font-size:14px;min-height:20px;margin:12px 0 0}
.solic{display:block;margin:10px 0 0;text-align:center;padding:12px;border-radius:10px;border:1px solid rgba(217,178,106,.7);color:#F0D9A8;text-decoration:none;font-weight:700;font-size:15px}
.solic[hidden]{display:none}
.manut{background:rgba(217,178,106,.16);border:1px solid rgba(217,178,106,.6);color:#F0D9A8;border-radius:10px;padding:10px 12px;font-size:13px;line-height:1.45;margin:0 0 14px}
.registro{font-size:12px;color:var(--osso2);text-align:center;margin:12px 0 0}
.registro:empty{display:none}
.atual{font-size:12px;color:var(--osso2);margin:10px 0 0;text-align:center;line-height:1.5}
.atual [hidden]{display:none!important}
.atual .novo{display:inline-block;width:7px;height:7px;border-radius:50%;background:#4CAF6A;margin-left:6px;vertical-align:middle;box-shadow:0 0 5px rgba(76,175,106,.8)}
.fotos-login{position:fixed;inset:0;overflow:hidden;z-index:0;pointer-events:none}
.fotos-login .fl{position:absolute;width:min(30vw,340px);height:auto;opacity:.17;filter:grayscale(45%) contrast(1.05);border-radius:8px}
#f,noscript form{position:relative;z-index:1}
@media (max-width:700px){.fotos-login .fl{width:44vw;opacity:.13}}
.fundo-video{position:fixed;inset:0;width:100%;height:100%;object-fit:contain;background:#000;z-index:-1;opacity:0;pointer-events:none;transition:opacity .8s}
html.com-video .fundo-video{opacity:.6}html.com-video body::before{opacity:0}
@media (prefers-reduced-motion:reduce){.fundo-video{display:none}}
</style></head><body><video class="fundo-video" src="fundo.mp4" autoplay muted loop playsinline preload="auto" aria-hidden="true" tabindex="-1"
  oncanplay="document.documentElement.classList.add('com-video')" onerror="this.remove()"></video>__FOTOS__
<noscript><style>#f{display:none}</style><form style="display:block"><h1>Extremo Norte<br>SP3</h1><p>Este visualizador não abre o painel. Toque em compartilhar e escolha abrir no Safari, Chrome ou Edge.</p></form></noscript>
<form id="f" autocomplete="off">
  __MANUT__
  <h1>Extremo Norte<br>SP3</h1>
  <p>__SUBTITULO__</p>
  <label for="cpf">__ROTULO__</label>
  <input id="cpf" type="text" inputmode="email" autocomplete="email" autocapitalize="none" spellcheck="false" placeholder="seu@email.com" maxlength="120" autofocus>
  <button id="b" type="submit">Entrar</button>
  <div id="erro" role="alert"></div>
  <a id="solic" class="solic" hidden target="_blank" rel="noopener">📩 Solicitar acesso</a>
  <p class="registro">__REGISTRO__</p>
  <p class="atual">__ATUAL__</p>
  <script>try{var np=document.getElementById("novoPag");if(np&&np.dataset.ts&&np.dataset.ts>(localStorage.getItem("pag_visto")||"")){np.hidden=false;document.getElementById("novoPagTxt").hidden=false}}catch(_){}</script>
</form>
<script>
const P=__PACOTE__;
const b64=s=>Uint8Array.from(atob(s),c=>c.charCodeAt(0));
const cpf=document.getElementById("cpf"),erro=document.getElementById("erro"),btn=document.getElementById("b");
const solic=document.getElementById("solic");
document.getElementById("f").addEventListener("submit",async ev=>{
  ev.preventDefault();erro.textContent="";
  solic.hidden=true;
  const d=cpf.value.trim().toLowerCase();
  const registrar=ok=>{try{const u="__ACESSOS_URL__";if(!u)return;
    const ua=navigator.userAgent,ap=/iPhone|iPad/.test(ua)?"iPhone":/Android/.test(ua)?"Celular Android":/Windows/.test(ua)?"Computador Windows":/Mac/.test(ua)?"Mac":"Outro";
    const corpo=JSON.stringify({email:d.includes("@")?d:"(senha coletiva)",ok,painel:/ADM/.test(document.title)?"ADM":"Geral",aparelho:ap});
    if(navigator.sendBeacon)navigator.sendBeacon(u,new Blob([corpo],{type:"text/plain"}));else fetch(u,{method:"POST",mode:"no-cors",body:corpo})}catch(_){}};
  if(!d){erro.textContent="Digite o seu e-mail.";return}
  if(!window.crypto||!crypto.subtle){erro.textContent="Este visualizador não permite abrir o painel. Abra o arquivo no Chrome, Edge ou Safari.";return}
  btn.disabled=true;btn.textContent="Verificando...";
  try{
    const base=await crypto.subtle.importKey("raw",new TextEncoder().encode(d),"PBKDF2",false,["deriveKey"]);
    const kek=await crypto.subtle.deriveKey({name:"PBKDF2",salt:b64(P.sal),iterations:P.it,hash:"SHA-256"},base,{name:"AES-GCM",length:256},false,["decrypt"]);
    let chave=null;
    for(const s of P.slots){try{chave=await crypto.subtle.decrypt({name:"AES-GCM",iv:b64(s.iv)},kek,b64(s.ct));break}catch(e){}}
    if(!chave)throw 0;
    const k=await crypto.subtle.importKey("raw",chave,"AES-GCM",false,["decrypt"]);
    const html=new TextDecoder().decode(await crypto.subtle.decrypt({name:"AES-GCM",iv:b64(P.iv)},k,b64(P.ct)));
    registrar(true);
    try{sessionStorage.setItem("painel_email",d)}catch(_){}
    document.open();document.write(html);document.close();
  }catch(e){if(d.includes("@"))registrar(false);erro.textContent="Este e-mail não tem acesso a este painel. Confira se digitou certo ou solicite o acesso.";btn.disabled=false;btn.textContent="Entrar";cpf.select();
    if(d.includes("@")){solic.href="https://wa.me/?text="+encodeURIComponent(`Olá! Quero solicitar acesso ao ${document.title.replace(/ - Analise/," - Painel Insanos")}. Meu e-mail: ${d}`);solic.hidden=false}}
});
</script></body></html>"""

def arred(x):
    """Arredonda .5 para cima, igual ao painel (Math.round), para os números baterem."""
    return int(x + 0.5)

def grupo_de(cargo):
    """Mesmo agrupamento do filtro do painel: Comando, Full, Meio, PP e Camiseta."""
    c = (cargo or "").strip().lower()
    for prefixo, nome in (("full", "Full"), ("meio", "Meio"), ("pp", "PP"), ("camiseta", "Camiseta")):
        if c.startswith(prefixo):
            return nome
    return "Comando"

def estatico(m, evs, rank):
    """Versão sem JavaScript do HTML individual (pré-visualização do iPhone/WhatsApp)."""
    from html import escape as e
    pr = sum(1 for x in evs if x["p"][0] == 1); au = sum(1 for x in evs if x["p"][0] == 0)
    pc = f"{arred(pr * 100 / (pr + au))}%" if pr + au else "–"
    r = rank.get("todos", {})
    pos = f'{r["pos"]}º entre {r["total"]} integrantes' if r.get("pos") else "sem posição"
    media = f' · média da divisão {r["media"]}%' if r.get("media") is not None else ""
    nasc_txt = ""
    if m.get("nasc"):
        n = m["nasc"]
        nasc_txt = f' · nascido em {n[8:10]}/{n[5:7]}/{n[:4]}' + (f' · {m["idade"]} anos' if m.get("idade") is not None else "")
    rot = {1: "Presente", 0: "Ausente", None: "Sem registro", "-": "Não conta"}
    cls = {1: "s1", 0: "s0", None: "sn", "-": "sn"}
    nomes_mes = ["Janeiro", "Fevereiro", "Março", "Abril", "Maio", "Junho", "Julho", "Agosto",
                 "Setembro", "Outubro", "Novembro", "Dezembro"]
    extra = ""
    if r.get("gpos"):
        extra += (f'<p class="nota">No grupo {e(r["grupo"])}: <b>{r["gpos"]}º de {r["gtotal"]}</b>'
                  f' · média do grupo {r["gmedia"]}%</p>')
    extra += '<h2 style="margin-top:20px">Evolução mês a mês</h2>'
    for mes in sorted(k for k in rank if k != "todos"):
        pm = sum(1 for x in evs if x["d"][:7] == mes and x["p"][0] == 1)
        am = sum(1 for x in evs if x["d"][:7] == mes and x["p"][0] == 0)
        med = rank[mes]["media"]
        extra += (f'<div class="ev"><div>{nomes_mes[int(mes[5:]) - 1]}<small>média da divisão '
                  f'{med if med is not None else "–"}%</small></div><b>{arred(pm * 100 / (pm + am)) if pm + am else "–"}%</b></div>'
                  if pm + am else
                  f'<div class="ev"><div>{nomes_mes[int(mes[5:]) - 1]}</div><b>–</b></div>')
    linhas = "".join(
        f'<div class="ev"><div>{e(x["e"])}<small>{x["d"][8:]}/{x["d"][5:7]}</small></div>'
        f'<span class="status {cls[x["p"][0]]}">{rot[x["p"][0]]}</span></div>'
        for x in reversed(evs))
    return (f'<noscript><style>.js{{display:none!important}}</style>'
            f'<section class="painel estatico"><div class="contato on" style="padding:0 0 12px;border:0;background:none">'
            f'<div class="quem">{e(m["n"])}<small>Nº {e(m["num"])} · {e(m["c"])}{nasc_txt}</small></div></div>'
            f'<div class="kpis"><div class="kpi"><div class="v">{pc}</div><div class="l">de adesão</div></div>'
            f'<div class="kpi"><div class="v">{pos.split(" ")[0]}</div><div class="l">{e(pos[len(pos.split(" ")[0])+1:])}{media}</div></div>'
            f'<div class="kpi ok"><div class="v">{pr}</div><div class="l">presenças</div></div>'
            f'<div class="kpi falta"><div class="v">{au}</div><div class="l">faltas</div></div></div>'
            f'{extra}<h2 style="margin-top:20px">Por evento</h2>{linhas}</section></noscript>')

def cifrar_para(obj, senhas):
    """Embaralha um conteúdo (JSON) com a mesma proteção do painel, que só abre com uma das senhas dadas."""
    from cryptography.hazmat.primitives.ciphers.aead import AESGCM
    import hashlib, secrets
    b64 = lambda x: base64.b64encode(x).decode()
    chave = AESGCM.generate_key(bit_length=256); iv = secrets.token_bytes(12); sal = secrets.token_bytes(16)
    ct = AESGCM(chave).encrypt(iv, json.dumps(obj, ensure_ascii=False).encode("utf-8"), None)
    slots = []
    for s_ in senhas:
        kek = hashlib.pbkdf2_hmac("sha256", s_.encode(), sal, ITERACOES, 32); iv_s = secrets.token_bytes(12)
        slots.append({"iv": b64(iv_s), "ct": b64(AESGCM(kek).encrypt(iv_s, chave, None))})
    return {"sal": b64(sal), "it": ITERACOES, "iv": b64(iv), "ct": b64(ct), "slots": slots}

def proteger(html, senhas, titulo, subtitulo, img, rotulo="E-mail"):
    """Criptografa o painel (AES-256-GCM). Cada senha libera a mesma chave do conteúdo."""
    from cryptography.hazmat.primitives.ciphers.aead import AESGCM
    import hashlib, secrets
    b64 = lambda x: base64.b64encode(x).decode()
    chave = AESGCM.generate_key(bit_length=256)
    iv = secrets.token_bytes(12)
    ct = AESGCM(chave).encrypt(iv, html.encode("utf-8"), None)
    sal = secrets.token_bytes(16)
    slots = []
    for s in senhas:
        kek = hashlib.pbkdf2_hmac("sha256", s.encode(), sal, ITERACOES, 32)
        iv_s = secrets.token_bytes(12)
        slots.append({"iv": b64(iv_s), "ct": b64(AESGCM(kek).encrypt(iv_s, chave, None))})
    secrets.SystemRandom().shuffle(slots)
    pacote = {"sal": b64(sal), "it": ITERACOES, "iv": b64(iv), "ct": b64(ct), "slots": slots}
    return (LOGIN.replace("__PACOTE__", json.dumps(pacote)).replace("__IMG__", img)
                 .replace("__TITULO__", titulo).replace("__SUBTITULO__", subtitulo).replace("__ROTULO__", rotulo)
                 .replace("__ATUAL__", INFO_LOGIN).replace("__MANUT__", MANUT_HTML).replace("__FOTOS__", FOTOS_LOGIN)
                 .replace("__MANIFEST__", "manifest-adm.json").replace("__REGISTRO__", "🔎 Os acessos a este painel são registrados pelo comando." if ACESSOS_URL else "")
                 .replace("__ACESSOS_URL__", ACESSOS_URL))

# ---------------------------------------------------------------------
#  Importação dos lançamentos feitos no painel ADM ("Lançar evento")
# ---------------------------------------------------------------------
PASTA_LANC = os.path.join(PASTA, "Lancamentos")

def _pasta_downloads():
    """Pasta Downloads real do Windows (funciona mesmo se ela foi movida para o OneDrive)."""
    pastas = []
    if os.name == "nt":
        try:
            import winreg
            chave = r"Software\Microsoft\Windows\CurrentVersion\Explorer\User Shell Folders"
            with winreg.OpenKey(winreg.HKEY_CURRENT_USER, chave) as k:
                v, _ = winreg.QueryValueEx(k, "{374DE290-123F-4565-9164-39C4925E467B}")
                pastas.append(os.path.expandvars(v))
        except Exception:
            pass
    casa = os.path.expanduser("~")
    pastas += [os.path.join(casa, "Downloads"), os.path.join(casa, "OneDrive", "Downloads")]
    try:
        pastas += [os.path.join(casa, d, "Downloads") for d in os.listdir(casa) if d.lower().startswith("onedrive")]
    except OSError:
        pass
    vistas, unicas = set(), []
    for x in pastas:
        k = os.path.normcase(os.path.abspath(x))
        if k not in vistas and os.path.isdir(x):
            vistas.add(k); unicas.append(x)
    return unicas

def _arquivos_lancamento():
    """Procura lancamento*.json na pasta do painel, em Lancamentos e na(s) pasta(s) Downloads."""
    import glob
    achados = []
    pastas = [PASTA, PASTA_LANC] + _pasta_downloads()
    for pasta in pastas:
        achados += glob.glob(os.path.join(pasta, "lancamento*.json"))
    if not achados:
        print("Nenhum lançamento novo encontrado. Procurei em:")
        for pasta in pastas:
            print("   " + pasta)
    return sorted(set(achados), key=lambda x: os.path.normcase(x))

def _letra(c):
    s = ""
    while c:
        c, r = divmod(c - 1, 26); s = chr(65 + r) + s
    return s

def _garantir_totais(ws, linha, ult_col, forcar=False):
    """Garante as fórmulas de Sim, Não e % da linha do evento (colunas logo depois do último integrante)."""
    ini, fim = "D", _letra(ult_col)
    c_sim, c_nao, c_pct = ult_col + 1, ult_col + 2, ult_col + 3
    ls, ln = _letra(c_sim), _letra(c_nao)
    rot_sim = f"${ls}$7" if str(ws.Cells(7, c_sim).Value or "").strip().lower() == "sim" else '"Sim"'
    rot_nao = f"${ln}$7" if str(ws.Cells(7, c_nao).Value or "").strip().lower() in ("não", "nao") else '"Não"'
    faixa = f"{ini}{linha}:{fim}{linha}"
    formulas = {c_sim: f'=IF(COUNTIF({faixa},{rot_sim})=0,"",COUNTIF({faixa},{rot_sim}))',
                c_nao: f'=IF(COUNTIF({faixa},{rot_nao})=0,"",COUNTIF({faixa},{rot_nao}))',
                c_pct: f'=IFERROR({ls}{linha}/({ls}{linha}+{ln}{linha}),"")'}
    for col, f in formulas.items():
        cel = ws.Cells(linha, col)
        try:
            if forcar or not cel.HasFormula:
                cel.Formula = f
        except Exception:
            cel.Formula = f
    try:
        ws.Cells(linha, c_pct).NumberFormat = "0.00%"
    except Exception:
        pass

def importar_lancamentos():
    arquivos = _arquivos_lancamento()
    if not arquivos:
        return
    lancs = []
    for arq in arquivos:
        try:
            with open(arq, encoding="utf-8") as f:
                d = json.load(f)
            if d.get("tipo") != "lancamento_presenca" or not d.get("data") or not d.get("evento"):
                continue
            datetime.date.fromisoformat(d["data"])
            lancs.append((arq, d))
        except Exception:
            print(f"Aviso: não consegui ler o lançamento {os.path.basename(arq)} (ignorado).")
    if not lancs:
        return
    print(f"Encontrei {len(lancs)} lançamento(s) de evento para importar na planilha...")
    if os.name != "nt":
        print("Aviso: a importação precisa do Windows com Excel. Lançamentos mantidos para depois.")
        return
    try:
        import win32com.client as w32
    except ImportError:
        print("Aviso: componente pywin32 ausente. Rode:  pip install pywin32   (lançamentos mantidos para depois)")
        return

    # cópia de segurança antes de gravar (mantém as 5 mais recentes)
    pasta_bkp = os.path.join(PASTA_LANC, "backup")
    os.makedirs(pasta_bkp, exist_ok=True)
    carimbo = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
    shutil.copy2(PLANILHA, os.path.join(pasta_bkp, f"backup_{carimbo}_{os.path.basename(PLANILHA)}"))
    bkps = sorted(x for x in os.listdir(pasta_bkp) if x.startswith("backup_"))
    for velho in bkps[:-5]:
        try: os.remove(os.path.join(pasta_bkp, velho))
        except OSError: pass

    xl, wb, fechar = None, None, False
    try:
        try:  # se a planilha já estiver aberta no Excel, grava nela mesma
            xl = w32.GetActiveObject("Excel.Application")
            for w in xl.Workbooks:
                if os.path.normcase(os.path.abspath(w.FullName)) == os.path.normcase(PLANILHA):
                    wb = w; break
        except Exception:
            xl = None
        if wb is None:
            xl = w32.DispatchEx("Excel.Application")
            xl.Visible = False
            xl.DisplayAlerts = False
            wb = xl.Workbooks.Open(PLANILHA)
            fechar = True
        if wb.ReadOnly:
            print("Aviso: a planilha está aberta só para leitura (talvez em outro computador). Lançamentos mantidos.")
            return
        # cada lançamento vai para a sua aba ("aba" no arquivo; sem ela, Presenca). Insights e Staff têm o mesmo formato.
        abas = {}
        def preparar(nome_aba):
            alvo = next((w for w in wb.Worksheets if str(w.Name).strip().lower() == nome_aba.strip().lower()), None)
            if alvo is None:
                return None
            if alvo.Name in abas:
                return abas[alvo.Name]
            nomes = {}
            c = 4
            while True:
                v = alvo.Cells(6, c).Value
                v = str(v).strip() if v is not None else ""
                if not v or v.lower().startswith("presen"):
                    break
                nomes[v.lower()] = c
                c += 1
            fim_b = None
            for r in range(8, 400):
                a = alvo.Cells(r, 1).Value
                if a is not None and str(a).strip().lower().startswith("ades"):
                    fim_b = r - 2          # a linha "Adesão" fica 2 linhas abaixo do último evento
                    break
            if fim_b is None:
                fim_b = 35 if alvo.Name.strip().lower() == "presenca" else 300
            abas[alvo.Name] = {"ws": alvo, "nomes": nomes, "ult": c - 1, "fim": fim_b}
            return abas[alvo.Name]

        def linha_existente(info, data_alvo, nome_ev):
            ws_ = info["ws"]
            for r in range(8, info["fim"] + 1):
                b, cc = ws_.Cells(r, 2).Value, ws_.Cells(r, 3).Value
                if cc is None or b is None:
                    continue
                try:
                    di = datetime.date(b.year, b.month, b.day).isoformat()
                except AttributeError:
                    di = globals()["data_iso"](b, int(data_alvo[:4])) if "data_iso" in globals() else None   # data escrita como texto ("11-jul")
                if di == data_alvo and str(cc).strip().lower() == nome_ev.strip().lower():
                    return r
            return None

        importados = []
        for arq, d in sorted(lancs, key=lambda x: x[1]["data"]):
            nome_aba = d.get("aba") or "Presenca"
            info = preparar(nome_aba)
            if info is None:
                print(f"  - {d['data']} {d['evento']}: a aba '{nome_aba}' não existe na planilha (arquivo mantido).")
                continue
            ws, nomes, ult_col = info["ws"], info["nomes"], info["ult"]
            linha = linha_existente(info, d["data"], d["evento"])
            if linha and not d.get("atualizar"):
                print(f"  - {d['data']} {d['evento']}: já está na planilha (ignorado).")
                importados.append(arq)
                continue
            novo_ev = linha is None
            if novo_ev:
                linha = next((r for r in range(8, info["fim"] + 1)
                              if ws.Cells(r, 2).Value is None and ws.Cells(r, 3).Value is None), None)
            if linha is None:
                # bloco cheio: insere uma linha dentro do bloco (Excel ajusta fórmulas e formatações)
                ws.Rows(info["fim"]).Insert()
                ws.Rows(info["fim"] + 1).Copy(ws.Rows(info["fim"]))      # a linha de cima vira cópia fiel da última
                info["fim"] += 1
                linha = info["fim"]
            dt = datetime.date.fromisoformat(d["data"])
            ws.Cells(linha, 2).Value = datetime.datetime(dt.year, dt.month, dt.day)
            ws.Cells(linha, 2).NumberFormat = "dd-mmm"
            ws.Cells(linha, 3).Value = d["evento"]
            for col in range(4, ult_col + 1):
                ws.Cells(linha, col).Value = None
            desconhecidos = []
            for ap, marca in d.get("marcacoes", {}).items():
                col = nomes.get(str(ap).strip().lower())
                if col is None:
                    desconhecidos.append(ap); continue
                if marca in ("Sim", "Não", "-", "Staff"):
                    ws.Cells(linha, col).Value = {"Sim": "SIM", "Não": "NÃO", "Staff": "STAFF"}.get(marca, marca)   # na planilha: SIM / NÃO / STAFF / -
            # totais da linha (Sim, Não e %): mantém/recoloca as fórmulas nas 3 colunas depois do último irmão
            _garantir_totais(ws, linha, ult_col)
            print(f"  + [{ws.Name}] {d['data']} {d['evento']}: {'gravado' if novo_ev else 'atualizado'} na linha {linha}.")
            if desconhecidos:
                print(f"    Atenção: não encontrei na linha 6 da aba {ws.Name}: " + ", ".join(desconhecidos))
            importados.append(arq)
        try:
            wb.Application.Calculate()
        except Exception:
            pass
        wb.Save()
    except Exception as e:
        print(f"Aviso: não consegui gravar na planilha pelo Excel ({e}). Lançamentos mantidos para depois.")
        return
    finally:
        try:
            if fechar and wb is not None:
                wb.Close(False)
            if fechar and xl is not None:
                xl.Quit()
        except Exception:
            pass

    # lançamento já gravado na planilha: o arquivo .json não é mais necessário e é apagado
    for arq in importados:
        try:
            os.remove(arq)
        except OSError:
            print(f"Aviso: não consegui apagar {os.path.basename(arq)}. Pode apagar à mão.")
    print(f"Planilha atualizada e arquivo(s) de lançamento apagado(s). Cópia de segurança em Lancamentos\\backup.\n")

# ---------------------------------------------------------------------
#  Verificação automática dos pagamentos (links da Iugu na aba PGTO)
# ---------------------------------------------------------------------
ARQ_VERIF = os.path.join(PASTA, "pagamentos_verificacao.json")
ARQ_LOG   = os.path.join(PASTA_LANC, "verificacao_pagamentos.log")

def _log(msg):
    """Mostra na janela e grava no log (a tarefa agendada roda sem janela)."""
    print(msg)
    try:
        os.makedirs(PASTA_LANC, exist_ok=True)
        with open(ARQ_LOG, "a", encoding="utf-8") as f:
            f.write(datetime.datetime.now().strftime("%d/%m/%Y %H:%M  ") + msg + "\n")
    except OSError:
        pass

def _ler_boleto(url):
    """Devolve (texto do selo de status ou None, data de vencimento do boleto ou None)."""
    import re, urllib.request, html as H
    req = urllib.request.Request(url, headers={"User-Agent": "Mozilla/5.0"})
    with urllib.request.urlopen(req, timeout=25) as r:
        bruto = r.read().decode("utf-8", "ignore")
    selo = re.search(r'class="status_warning"[^>]*>\s*([^<]+?)\s*<', bruto)
    texto = re.sub(r"\s+", " ", H.unescape(re.sub(r"<[^>]+>", " ", bruto)))
    v = re.search(r"Vencimento\D{0,40}(\d{2})/(\d{2})/(\d{4})", texto, re.I)
    venc = datetime.date(int(v.group(3)), int(v.group(2)), int(v.group(1))) if v else None
    return (selo.group(1).strip() if selo else None), venc

def verificar_pagamentos(forcar=False):
    hoje = datetime.date.today()
    ciclo = vencimento(hoje.year, hoje.month)
    if not forcar and hoje > ciclo:
        _log(f"Verificação de pagamentos: fora do período (vencimento foi {ciclo:%d/%m}). Nada a fazer.")
        return False
    if os.name != "nt":
        _log("Verificação de pagamentos: precisa do Windows com Excel.")
        return False
    try:
        import win32com.client as w32
    except ImportError:
        _log("Verificação de pagamentos: componente pywin32 ausente (pip install pywin32).")
        return False
    try:  # se a planilha estiver aberta no Excel, não mexe (pode ter edição em andamento)
        xl_ativo = w32.GetActiveObject("Excel.Application")
        for w in xl_ativo.Workbooks:
            if os.path.normcase(os.path.abspath(w.FullName)) == os.path.normcase(PLANILHA):
                _log("Verificação de pagamentos: a planilha está aberta no Excel. Rodada pulada; tento na próxima.")
                return False
    except Exception:
        pass

    _log(f"Verificando pagamentos (ciclo com vencimento em {ciclo:%d/%m/%Y})...")
    pasta_bkp = os.path.join(PASTA_LANC, "backup")
    os.makedirs(pasta_bkp, exist_ok=True)
    shutil.copy2(PLANILHA, os.path.join(pasta_bkp, f"backup_{datetime.datetime.now():%Y%m%d_%H%M%S}_{os.path.basename(PLANILHA)}"))
    for velho in sorted(x for x in os.listdir(pasta_bkp) if x.startswith("backup_"))[:-5]:
        try: os.remove(os.path.join(pasta_bkp, velho))
        except OSError: pass

    manual = _manual_carregar(ciclo)
    inativos = inativos_planilha()
    xl = wb = None
    itens, mudou, novos_pagos = [], 0, 0
    try:
        xl = w32.DispatchEx("Excel.Application")
        xl.Visible = False
        xl.DisplayAlerts = False
        wb = xl.Workbooks.Open(PLANILHA)
        if wb.ReadOnly:
            _log("Verificação de pagamentos: planilha somente leitura. Rodada pulada.")
            return False
        ws = wb.Worksheets("PGTO")
        for r in range(2, 300):
            link = ws.Cells(r, 10).Value
            if not link or not str(link).startswith("http"):
                continue
            rotulo = str(ws.Cells(r, 2).Value or "").strip()
            if any(rotulo.lower().endswith(ap) for ap in inativos):
                _log(f"  {rotulo}: afastado/excluído (não verificado)"); continue
            atual = str(ws.Cells(r, 3).Value or "").strip()
            item = {"rotulo": rotulo, "selo": None, "venc_boleto": None, "status": atual, "obs": ""}
            try:
                selo, vb = _ler_boleto(str(link))
            except Exception as e:
                item["obs"] = "não consegui abrir o boleto"
                itens.append(item); _log(f"  {rotulo}: não abriu ({e})"); continue
            item["selo"] = selo
            item["venc_boleto"] = vb.isoformat() if vb else None
            # regra: só é "Pago em Dia" o boleto DESTE mês com o selo "Pago" da Iugu; qualquer outra coisa é "Em aberto"
            boleto_do_mes = not vb or (vb.year, vb.month) == (ciclo.year, ciclo.month)
            if boleto_do_mes and selo and "pag" in selo.lower():
                novo = "Pago em Dia"
            else:
                novo = "Em aberto"
                if not boleto_do_mes:
                    item["obs"] = f"boleto com vencimento {vb:%d/%m/%Y}: o boleto deste mês ainda não foi cadastrado"
                elif selo:
                    item["obs"] = f"selo da Iugu: {selo}"
            ap_man = _manual_de(rotulo, manual)
            if not ap_man and novo != "Pago em Dia" and atual.lower().startswith("aguard"):
                # compensação colocada direto na planilha também vale (até o boleto ser pago)
                ap_man = rotulo.split()[-1] if rotulo else rotulo
                manual[ap_man] = "Aguardando Compensação"
            if novo == "Pago em Dia":
                if ap_man:
                    manual.pop(ap_man, None)          # boleto pago: vale o boleto, a marcação manual deixa de valer
                item["boleto_pago"] = True
            elif ap_man:
                novo = manual[ap_man]                 # boleto não pago: mantém o que foi marcado à mão no mês
                item["obs"] = (item.get("obs") + " · " if item.get("obs") else "") + "marcado à mão"
            item["status"] = novo
            if novo != atual:
                ws.Cells(r, 3).Value = novo
                mudou += 1
                _log(f"  {rotulo}: {atual or '(vazio)'} -> {novo}")
                if novo == "Pago em Dia":
                    novos_pagos += 1
                    item["novo_pago"] = True
            itens.append(item)
        if mudou:
            wb.Save()
        _manual_salvar(ciclo, manual)
    except Exception as e:
        _log(f"Verificação de pagamentos: erro ao usar o Excel ({e}). Nada foi salvo.")
        return False
    finally:
        try:
            if wb is not None: wb.Close(False)
            if xl is not None: xl.Quit()
        except Exception:
            pass
    quando = datetime.datetime.now().strftime("%d/%m/%Y às %H:%M")
    ultimo, pagos_em = None, {}
    try:
        with open(ARQ_VERIF, encoding="utf-8") as f:
            ant = json.load(f)
        ultimo = ant.get("ultimo_pagamento")
        if str(ant.get("ciclo") or "")[:7] == ciclo.isoformat()[:7]:
            pagos_em = dict(ant.get("pagos_em") or {})
    except Exception:
        pass
    agora_ts = datetime.datetime.now().strftime("%Y-%m-%dT%H:%M:%S")
    for it in itens:   # horário em que a verificação encontrou o pagamento (para ordenar e destacar no painel)
        rot = it.get("rotulo", "")
        if it.get("status") == "Pago em Dia":
            pagos_em.setdefault(rot, agora_ts if it.get("novo_pago") else "")
        else:
            pagos_em.pop(rot, None)
    if novos_pagos:
        ultimo = {"ts": datetime.datetime.now().strftime("%Y-%m-%dT%H:%M:%S"), "hora": datetime.datetime.now().strftime("%d/%m às %H:%M")}
    with open(ARQ_VERIF, "w", encoding="utf-8") as f:
        json.dump({"quando": quando, "ciclo": ciclo.isoformat(), "novos_pagos": novos_pagos, "ultimo_pagamento": ultimo,
                   "pagos_em": pagos_em, "itens": itens}, f, ensure_ascii=False, indent=1)
    _log(f"Verificação concluída: {len(itens)} boletos lidos, {mudou} status alterado(s).")
    novos = [str(it.get("rotulo", "")).split(" - ")[-1].strip() for it in itens if it.get("novo_pago")]
    if novos:
        hora = datetime.datetime.now().strftime("%d/%m às %H:%M")
        if len(novos) == 1:
            notificar("💰 Pagamento novo", f"{novos[0]} pagou a mensalidade (identificado {hora}).", "pagamento")
        else:
            notificar(f"💰 {len(novos)} pagamentos novos", ", ".join(novos) + f" (identificados {hora}).", "pagamento")
    return True

# ---------------------------------------------------------------------
#  Status de pagamento marcado à mão no painel: só "Em compensação" (irmão que mandou o comprovante)
#  Regra: o boleto vem primeiro. Boleto do mês pago na Iugu = "Pago em Dia" e não pode ser mudado.
# ---------------------------------------------------------------------
ARQ_MANUAL = os.path.join(PASTA, "pagamentos_manual.json")
STATUS_MANUAIS = ("Aguardando Compensação", "Em aberto")   # à mão só dá para marcar/desmarcar "Em compensação"

def _manual_carregar(ciclo):
    try:
        with open(ARQ_MANUAL, encoding="utf-8") as f:
            d = json.load(f)
        return dict(d.get("status", {})) if d.get("ciclo") == ciclo.isoformat() else {}
    except Exception:
        return {}

def _manual_salvar(ciclo, status):
    with open(ARQ_MANUAL, "w", encoding="utf-8") as f:
        json.dump({"ciclo": ciclo.isoformat(), "status": status}, f, ensure_ascii=False, indent=1)

def _manual_de(rotulo, manual):
    r = str(rotulo or "").strip().lower()
    return next((ap for ap in manual if r.endswith(ap.lower())), None)

# ---------------------------------------------------------------------
#  Integrantes: novo / afastar / reativar / excluir (arquivos "integrante_*.json" gerados no painel ADM)
# ---------------------------------------------------------------------
def _sem_acento(v):
    import unicodedata
    return "".join(ch for ch in unicodedata.normalize("NFD", str(v or "").lower()) if not unicodedata.combining(ch)).strip()

def importar_integrantes():
    import glob
    pastas = [PASTA, PASTA_LANC] + _pasta_downloads()
    arquivos = sorted({a for p in pastas for a in glob.glob(os.path.join(p, "integrante_*.json"))}, key=os.path.getmtime)
    pedidos = []
    for arq in arquivos:
        try:
            with open(arq, encoding="utf-8") as f:
                d = json.load(f)
            if d.get("tipo") == "integrante" and d.get("acao") in ("novo", "afastar", "reativar", "excluir") and d.get("apelido"):
                pedidos.append((arq, d))
        except Exception:
            print(f"Aviso: não consegui ler {os.path.basename(arq)} (ignorado).")
    if not pedidos:
        return
    print(f"Encontrei {len(pedidos)} alteração(ões) de integrante para aplicar na planilha...")
    if os.name != "nt":
        print("Aviso: precisa do Windows com Excel. Arquivos mantidos para depois."); return
    try:
        import win32com.client as w32
    except ImportError:
        print("Aviso: componente pywin32 ausente. Arquivos mantidos para depois."); return
    xl = wb = None; fechar = False; feitos = []
    try:
        try:
            xl = w32.GetActiveObject("Excel.Application")
            for w in xl.Workbooks:
                if os.path.normcase(os.path.abspath(w.FullName)) == os.path.normcase(PLANILHA):
                    wb = w; break
        except Exception:
            xl = None
        if wb is None:
            xl = w32.DispatchEx("Excel.Application"); xl.Visible = False; xl.DisplayAlerts = False
            wb = xl.Workbooks.Open(PLANILHA); fechar = True
        if wb.ReadOnly:
            print("Aviso: planilha somente leitura. Arquivos mantidos para depois."); return
        pasta_bkp = os.path.join(PASTA_LANC, "backup"); os.makedirs(pasta_bkp, exist_ok=True)
        shutil.copy2(PLANILHA, os.path.join(pasta_bkp, f"backup_{datetime.datetime.now():%Y%m%d_%H%M%S}_{os.path.basename(PLANILHA)}"))
        DI = wb.Worksheets("Dados Integrantes")
        tab = DI.ListObjects(1)
        cab = {}
        for c in range(1, 40):
            v = _sem_acento(DI.Cells(1, c).Value)
            if v: cab[v] = c
        def col(*inicios):
            for k, c in cab.items():
                if any(k.startswith(i) for i in inicios):
                    return c
            return None
        C = {"num": 1, "ap": 2, "cargo": 3, "lista": col("numero lista") or 4, "grau": col("grau") or 5,
             "tel": col("telefone"), "nome": col("nome completo"), "nasc": col("nascimento"), "cpf": col("cpf"),
             "colete": col("colete"), "tam": col("tamanho"), "email": col("email", "e-mail"),
             "sit": col("situa"), "desde": col("afastado")}
        if not C["sit"]:
            print("Aviso: a aba Dados Integrantes não tem a coluna 'Situação'. Arquivos mantidos para depois."); return
        ult_col_di = max(cab.values())
        def linha_de(ap):
            for r in range(2, 400):
                if _sem_acento(DI.Cells(r, 2).Value) == _sem_acento(ap):
                    return r
            return None
        def cabec_tab():
            return tab.HeaderRowRange.Row
        def mover_para_tabela(ap, apos):
            """Garante a linha de 'ap' dentro da tabela, logo depois de 'apos'. Devolve a linha final."""
            r_ap = linha_de(apos) if apos else None
            r_old = linha_de(ap)
            if r_ap is None:
                return r_old
            if r_old == r_ap + 1:
                return r_old
            pos = r_ap - cabec_tab() + 1          # posição dentro da tabela (1 = primeira linha de dados)
            nova = tab.ListRows.Add(pos).Range.Row
            r_old = linha_de(ap) if r_old else None    # a linha antiga pode ter descido uma posição
            if r_old:
                for c in range(1, ult_col_di + 1):
                    if not DI.Cells(nova, c).HasFormula:
                        DI.Cells(nova, c).Value = DI.Cells(r_old, c).Value
                t0, t1 = tab.Range.Row, tab.Range.Row + tab.Range.Rows.Count - 1
                if t0 < r_old <= t1:
                    tab.ListRows(r_old - cabec_tab()).Delete()
                else:
                    DI.Range(DI.Cells(r_old, 1), DI.Cells(r_old, ult_col_di)).ClearContents()
            return linha_de(ap)
        def colunas_nas_abas(ap, cargo, apos):
            """Cria a coluna do irmão (linha 6 = apelido, 7 = cargo) na Presenca, Insights e Staff, depois de 'apos'."""
            for nome_aba in ("Presenca", "Insights", "Staff"):
                ws = next((w for w in wb.Worksheets if _sem_acento(w.Name) == _sem_acento(nome_aba)), None)
                if ws is None:
                    continue
                cols = {}
                c = 4
                while True:
                    v = str(ws.Cells(6, c).Value or "").strip()
                    if not v or v.lower().startswith("presen"):
                        break
                    cols[_sem_acento(v)] = c; c += 1
                if _sem_acento(ap) in cols:
                    continue
                ult = c - 1
                pos = cols.get(_sem_acento(apos), ult) + 1
                ws.Columns(pos).Insert()
                # linha 7: cargo curto, sem a parte entre parênteses ("Camiseta (Grau X)" -> "Camiseta"), como nas outras colunas
                cargo_curto = re.sub(r"\s*\([^)]*\)\s*", " ", str(cargo or "")).strip()
                ws.Cells(6, pos).Value = ap; ws.Cells(7, pos).Value = cargo_curto
                ult += 1
                fim = None
                for r in range(8, 400):
                    a = ws.Cells(r, 1).Value
                    if a is not None and str(a).strip().lower().startswith("ades"):
                        fim = r; break
                ate = (fim - 2) if fim else 42
                for r in range(8, ate + 1):        # totais (Sim, Não e %) contando a coluna nova
                    _garantir_totais(ws, r, ult, forcar=True)
                if fim:                              # "Adesão Integrante" da coluna nova: mesma fórmula do vizinho
                    viz = pos - 1 if pos - 1 >= 4 else pos + 1
                    ws.Cells(fim, viz).Copy(ws.Cells(fim, pos))
                print(f"    + coluna de {ap} criada na aba {ws.Name}")
        def linha_pgto(ap, r_di):
            P = next((w for w in wb.Worksheets if _sem_acento(w.Name) == "pgto"), None)
            if P is None:
                return
            ult = 1
            for r in range(2, 400):
                v = str(P.Cells(r, 2).Value or "")
                if v.strip():
                    ult = r
                if v.lower().endswith(_sem_acento(ap)):
                    return       # já existe
            P.Rows(ult).Copy(P.Rows(ult + 1))
            P.Cells(ult + 1, 2).Formula = f"='Dados Integrantes'!F{r_di}&\" - \"&'Dados Integrantes'!B{r_di}"
            P.Cells(ult + 1, 3).Value = "Em aberto"
            P.Cells(ult + 1, 10).Value = None
            print(f"    + linha de {ap} criada na aba PGTO")
        for arq, d in pedidos:
            ap, acao = d["apelido"].strip(), d["acao"]
            r = linha_de(ap)
            if acao == "novo":
                if r:
                    print(f"  - {ap}: já está na Dados Integrantes (ignorado)."); feitos.append(arq); continue
                dd = d.get("dados", {})
                r_ap = linha_de(d.get("apos", "")) if d.get("apos") else None
                pos = (r_ap - cabec_tab() + 1) if r_ap else None
                r = (tab.ListRows.Add(pos) if pos else tab.ListRows.Add()).Range.Row
                vals = {"num": dd.get("num"), "ap": ap, "cargo": dd.get("cargo"), "grau": dd.get("grau"), "tel": dd.get("tel"),
                        "nome": dd.get("nome"), "cpf": dd.get("cpf"), "colete": dd.get("colete"), "tam": dd.get("tam"),
                        "email": dd.get("email"), "sit": "Ativo"}
                for k, v in vals.items():
                    if C.get(k) and v not in (None, ""):
                        DI.Cells(r, C[k]).Value = v
                if C.get("nasc") and dd.get("nasc"):
                    try:
                        DI.Cells(r, C["nasc"]).Value = datetime.datetime.strptime(dd["nasc"], "%Y-%m-%d")
                    except ValueError:
                        pass
                colunas_nas_abas(ap, dd.get("cargo", ""), d.get("apos", ""))
                linha_pgto(ap, r)
                print(f"  + {ap}: novo integrante incluído (linha {r} da Dados Integrantes).")
            elif r is None:
                print(f"  - {ap}: não encontrei na Dados Integrantes (ignorado)."); feitos.append(arq); continue
            elif acao == "afastar":
                DI.Cells(r, C["sit"]).Value = "Afastado"
                if C["desde"] and d.get("desde"):
                    DI.Cells(r, C["desde"]).Value = datetime.datetime.strptime(d["desde"], "%Y-%m-%d")
                print(f"  + {ap}: afastado desde {d.get('desde', '')}.")
            elif acao == "excluir":
                # apaga TUDO do integrante: linha da PGTO, colunas nas abas de eventos e a linha da Dados Integrantes
                P = next((w for w in wb.Worksheets if _sem_acento(w.Name) == "pgto"), None)
                if P is not None:
                    r_di = linha_de(ap)
                    for rp in range(2, 400):
                        txt = str(P.Cells(rp, 2).Value or "").strip()
                        frm = str(P.Cells(rp, 2).Formula or "").replace("$", "")
                        # pelo texto ("19. X - Pastor") ou pela fórmula que aponta para a linha dele na Dados Integrantes
                        if _sem_acento(txt).endswith(_sem_acento(ap)) or (r_di and "dados integrantes" in frm.lower() and re.search(rf"!B{r_di}(?!\d)", frm)):
                            P.Rows(rp).Delete(); print(f"    - linha de {ap} apagada da aba PGTO"); break
                for nome_aba in ("Presenca", "Insights", "Staff"):
                    ws = next((w for w in wb.Worksheets if _sem_acento(w.Name) == _sem_acento(nome_aba)), None)
                    if ws is None:
                        continue
                    alvo, c = None, 4
                    while True:
                        v = str(ws.Cells(6, c).Value or "").strip()
                        if not v or v.lower().startswith("presen"):
                            break
                        if _sem_acento(v) == _sem_acento(ap):
                            alvo = c
                        c += 1
                    if alvo is None:
                        continue
                    ws.Columns(alvo).Delete()
                    ult = c - 2                          # uma coluna a menos
                    fim = None
                    for rr in range(8, 400):
                        a_ = ws.Cells(rr, 1).Value
                        if a_ is not None and str(a_).strip().lower().startswith("ades"):
                            fim = rr; break
                    for rr in range(8, ((fim - 2) if fim else 42) + 1):
                        _garantir_totais(ws, rr, ult, forcar=True)
                    print(f"    - coluna de {ap} apagada da aba {ws.Name}")
                r = linha_de(ap)
                t0, t1 = tab.Range.Row, tab.Range.Row + tab.Range.Rows.Count - 1
                if t0 < r <= t1:
                    tab.ListRows(r - cabec_tab()).Delete()
                else:
                    DI.Rows(r).Delete()
                print(f"  + {ap}: excluído de toda a planilha (a numeração dos de baixo sobe).")
            elif acao == "reativar":
                r = mover_para_tabela(ap, d.get("apos", "")) or r
                DI.Cells(r, C["sit"]).Value = "Ativo"
                if C["desde"]:
                    DI.Cells(r, C["desde"]).Value = None
                colunas_nas_abas(ap, str(DI.Cells(r, 3).Value or ""), d.get("apos", ""))
                linha_pgto(ap, r)
                print(f"  + {ap}: reativado.")
            feitos.append(arq)
        try: wb.Application.Calculate()
        except Exception: pass
        wb.Save()
    except Exception as e:
        print(f"Aviso: não consegui aplicar as alterações de integrante ({e}). Arquivos mantidos; a planilha tem backup em Lancamentos\\backup.")
        return
    finally:
        try:
            if fechar and wb is not None: wb.Close(False)
            if fechar and xl is not None: xl.Quit()
        except Exception: pass
    for arq in feitos:
        try: os.remove(arq)
        except OSError: pass


ARQ_MIGR_STAFF = os.path.join(PASTA_LANC, "staff_migrado.json")

def migrar_staff():
    """Uma vez só: põe STAFF na lista suspensa da Presenca e passa o histórico da aba Staff para a Presenca (SIM -> STAFF)."""
    if os.path.exists(ARQ_MIGR_STAFF) or os.name != "nt":
        return
    try:
        import win32com.client as w32
    except ImportError:
        return
    xl = wb = None; fechar = False; ok = False
    try:
        try:
            xl = w32.GetActiveObject("Excel.Application")
            for w in xl.Workbooks:
                if os.path.normcase(os.path.abspath(w.FullName)) == os.path.normcase(PLANILHA):
                    wb = w; break
        except Exception:
            xl = None
        if wb is None:
            xl = w32.DispatchEx("Excel.Application"); xl.Visible = False; xl.DisplayAlerts = False
            wb = xl.Workbooks.Open(PLANILHA); fechar = True
        if wb.ReadOnly:
            return
        pasta_bkp = os.path.join(PASTA_LANC, "backup"); os.makedirs(pasta_bkp, exist_ok=True)
        shutil.copy2(PLANILHA, os.path.join(pasta_bkp, f"backup_{datetime.datetime.now():%Y%m%d_%H%M%S}_antes_staff_{os.path.basename(PLANILHA)}"))
        P = wb.Worksheets("Presenca")
        def colunas(ws):
            cols, c = {}, 4
            while True:
                v = str(ws.Cells(6, c).Value or "").strip()
                if not v or v.lower().startswith("presen"):
                    return cols, c - 1
                cols[_sem_acento(v)] = c; c += 1
        cp, ult = colunas(P)
        # 1) lista suspensa: SIM, NÃO, -, STAFF
        try:
            f1 = str(P.Range("D8").Validation.Formula1 or "")
        except Exception:
            f1 = ""
        m = re.match(r"=?\$?([A-Z]+)\$?(\d+):\$?([A-Z]+)\$?(\d+)", f1)
        valores = []
        if m:
            col_l, r1 = m.group(1), int(m.group(2))
            rr = r1
            while str(P.Range(f"{col_l}{rr}").Value or "").strip() and rr < r1 + 20:
                valores.append(str(P.Range(f"{col_l}{rr}").Value).strip().upper()); rr += 1
        if not m or "SIM" not in valores or not any(v in ("NÃO", "NAO") for v in valores):
            print("Aviso: não reconheci a lista suspensa da Presenca (SIM/NÃO). Coloque STAFF na lista à mão; o resto segue normal.")
            m = None
        if m:
            r = r1
            while str(P.Range(f"{col_l}{r}").Value or "").strip():
                if str(P.Range(f"{col_l}{r}").Value).strip().upper() == "STAFF":
                    break
                r += 1
            P.Range(f"{col_l}{r}").Value = "STAFF"
            fim = r
            while str(P.Range(f"{col_l}{fim + 1}").Value or "").strip():
                fim += 1
            alvo = P.Range(P.Cells(8, 4), P.Cells(42, ult))
            alvo.Validation.Delete()
            alvo.Validation.Add(3, 1, 1, f"=${col_l}${r1}:${col_l}${fim}")
            print(f"Lista suspensa da Presenca: agora com STAFF ({col_l}{r1}:{col_l}{fim}).")
        # 2) histórico: SIM na aba Staff vira STAFF no mesmo evento da Presenca
        S = next((w for w in wb.Worksheets if _sem_acento(w.Name) == "staff"), None)
        passados = 0
        if S is not None:
            cs, ult_s = colunas(S)
            ev_p = {}
            for r in range(8, 43):
                d, n = P.Cells(r, 2).Value, str(P.Cells(r, 3).Value or "").strip()
                if d:
                    ev_p.setdefault(str(d)[:10], []).append((r, _sem_acento(n)))
            for r in range(8, 43):
                d, n = S.Cells(r, 2).Value, str(S.Cells(r, 3).Value or "").strip()
                if not d:
                    continue
                cand = ev_p.get(str(d)[:10], [])
                lin = next((rp for rp, nn in cand if nn == _sem_acento(n)), None) or (cand[0][0] if len(cand) == 1 else None)
                if not lin:
                    continue
                for ap, c in cs.items():
                    if str(S.Cells(r, c).Value or "").strip().upper() == "SIM" and ap in cp:
                        P.Cells(lin, cp[ap]).Value = "STAFF"; passados += 1
        try: wb.Application.Calculate()
        except Exception: pass
        wb.Save(); ok = True
        print(f"Staff passado para a Presenca: {passados} marcação(ões) agora como STAFF.")
    except Exception as e:
        print(f"Aviso: não consegui passar o Staff para a Presenca ({e}). Vou tentar de novo na próxima vez.")
    finally:
        try:
            if fechar and wb is not None: wb.Close(False)
            if fechar and xl is not None: xl.Quit()
        except Exception: pass
    if ok:
        with open(ARQ_MIGR_STAFF, "w", encoding="utf-8") as f:
            json.dump({"feito": datetime.datetime.now().isoformat(timespec="minutes")}, f)

def importar_pagamentos_manuais():
    import glob
    pastas = [PASTA, PASTA_LANC] + _pasta_downloads()
    arquivos = sorted({a for p in pastas for a in glob.glob(os.path.join(p, "pagamento_manual*.json"))})
    pedidos = {}
    for arq in arquivos:
        try:
            with open(arq, encoding="utf-8") as f:
                d = json.load(f)
            if d.get("tipo") == "pagamento_manual":
                for ap, st in d.get("status", {}).items():
                    if st in STATUS_MANUAIS:
                        pedidos[ap] = st
        except Exception:
            print(f"Aviso: não consegui ler {os.path.basename(arq)} (ignorado).")
    if not pedidos:
        return
    print(f"Encontrei {len(pedidos)} alteração(ões) manual(is) de pagamento. Conferindo os boletos antes...")
    if os.name != "nt":
        print("Aviso: a gravação precisa do Windows com Excel. Alterações mantidas para depois."); return
    try:
        import win32com.client as w32
    except ImportError:
        print("Aviso: componente pywin32 ausente. Alterações mantidas para depois."); return
    hoje = datetime.date.today(); ciclo = vencimento(hoje.year, hoje.month)
    manual = _manual_carregar(ciclo)
    pendentes = {}
    xl = wb = None; fechar = False
    try:
        try:
            xl = w32.GetActiveObject("Excel.Application")
            for w in xl.Workbooks:
                if os.path.normcase(os.path.abspath(w.FullName)) == os.path.normcase(PLANILHA):
                    wb = w; break
        except Exception:
            xl = None
        if wb is None:
            xl = w32.DispatchEx("Excel.Application"); xl.Visible = False; xl.DisplayAlerts = False
            wb = xl.Workbooks.Open(PLANILHA); fechar = True
        if wb.ReadOnly:
            print("Aviso: planilha somente leitura. Alterações mantidas para depois."); return
        pasta_bkp = os.path.join(PASTA_LANC, "backup"); os.makedirs(pasta_bkp, exist_ok=True)
        shutil.copy2(PLANILHA, os.path.join(pasta_bkp, f"backup_{datetime.datetime.now():%Y%m%d_%H%M%S}_{os.path.basename(PLANILHA)}"))
        ws = wb.Worksheets("PGTO")
        for ap, st in pedidos.items():
            linha = None
            for r in range(2, 300):
                rot = str(ws.Cells(r, 2).Value or "").strip()
                if rot and rot.lower().endswith(ap.lower()):
                    linha = r; break
            if not linha:
                print(f"  - {ap}: não encontrado na aba PGTO (ignorado)."); continue
            link = str(ws.Cells(linha, 10).Value or "")
            if link.startswith("http"):
                try:
                    selo, vb = _ler_boleto(link)
                except Exception:
                    pendentes[ap] = st
                    print(f"  - {ap}: não consegui conferir o boleto agora. A alteração fica guardada para a próxima vez."); continue
                do_mes = not vb or (vb.year, vb.month) == (ciclo.year, ciclo.month)
                if do_mes and selo and "pag" in selo.lower():
                    ws.Cells(linha, 3).Value = "Pago em Dia"; manual.pop(ap, None)
                    print(f"  - {ap}: boleto já está PAGO na Iugu. Mantido 'Pago em Dia' (alteração manual ignorada)."); continue
            ws.Cells(linha, 3).Value = st
            if st == "Em aberto":
                manual.pop(ap, None)
            else:
                manual[ap] = st
            print(f"  + {ap}: {st}")
        try: wb.Application.Calculate()
        except Exception: pass
        wb.Save()
        _manual_salvar(ciclo, manual)
    except Exception as e:
        print(f"Aviso: não consegui gravar no Excel ({e}). Alterações mantidas para depois."); return
    finally:
        try:
            if fechar and wb is not None: wb.Close(False)
            if fechar and xl is not None: xl.Quit()
        except Exception: pass
    for arq in arquivos:
        try: os.remove(arq)
        except OSError: pass
    if pendentes:
        with open(os.path.join(PASTA_LANC, "pagamento_manual_pendente.json"), "w", encoding="utf-8") as f:
            json.dump({"tipo": "pagamento_manual", "status": pendentes}, f, ensure_ascii=False)

_ACESSOS_CACHE = {}
def buscar_acessos():
    """Lista de registros da planilha Google (busca uma vez por execução)."""
    if "linhas" in _ACESSOS_CACHE:
        return _ACESSOS_CACHE["linhas"]
    linhas = None
    if ACESSOS_URL:
        import urllib.request, urllib.parse
        try:
            with urllib.request.urlopen(ACESSOS_URL + "?chave=" + urllib.parse.quote(ACESSOS_CHAVE), timeout=25) as r:
                linhas = json.loads(r.read().decode("utf-8"))
        except Exception as e:
            print(f"Aviso: não consegui buscar os registros de acesso ({e}).")
    _ACESSOS_CACHE["linhas"] = linhas
    return linhas

def importar_acessos():
    """Busca os registros de acesso na planilha Google e grava na aba "Acessos" da Insanos.xlsx (mais recente primeiro)."""
    if not ACESSOS_URL or os.name != "nt":
        return
    linhas = buscar_acessos()
    if linhas is None:
        return
    if not isinstance(linhas, list) or len(linhas) < 2:
        return
    apel = {k.lower(): v for k, v in EMAIL_APELIDO.items()}
    nomes = {m.lower(): m for m in AUTORIZADOS + AUTORIZADOS_SO_GERAL}
    def hora_br(v):
        """'2026-10-06T19:29:08.000Z' (UTC) ou '2026-10-06 16:29:08' -> datetime no horário de Brasília."""
        t = str(v).strip()
        try:
            if t.endswith("Z") or "T" in t:
                d = datetime.datetime.strptime(t.replace("Z", "")[:19], "%Y-%m-%dT%H:%M:%S")
                return d - datetime.timedelta(hours=3)          # Brasília (sem horário de verão)
            return datetime.datetime.strptime(t[:19], "%Y-%m-%d %H:%M:%S")
        except ValueError:
            return None
    regs = []
    for l in linhas[1:]:
        quando, email, res, painel, aparelho = (list(l) + [""] * 5)[:5]
        ap = apel.get(str(email).lower(), "")
        dt = hora_br(quando)
        regs.append([dt or datetime.datetime.min, dt.strftime("%d/%m/%Y %H:%M:%S") if dt else str(quando),
                     nomes.get(ap.lower(), ap.title() if ap else "—"), str(email), str(res), str(painel), str(aparelho)])
    regs.sort(key=lambda x: x[0], reverse=True)
    regs = [r[1:] for r in regs]
    try:
        import win32com.client as w32
    except ImportError:
        return
    xl = wb = None; fechar = False
    try:
        try:
            xl = w32.GetActiveObject("Excel.Application")
            for w in xl.Workbooks:
                if os.path.normcase(os.path.abspath(w.FullName)) == os.path.normcase(PLANILHA):
                    wb = w; break
        except Exception:
            xl = None
        if wb is None:
            xl = w32.DispatchEx("Excel.Application"); xl.Visible = False; xl.DisplayAlerts = False
            wb = xl.Workbooks.Open(PLANILHA); fechar = True
        if wb.ReadOnly:
            return
        ws = next((w for w in wb.Worksheets if str(w.Name).strip().lower() == "acessos"), None)
        if ws is None:
            ws = wb.Worksheets.Add(After=wb.Worksheets(wb.Worksheets.Count)); ws.Name = "Acessos"
        ws.Cells.ClearContents()
        cab = [["Data/hora", "Irmão", "E-mail", "Resultado", "Painel", "Aparelho"]]
        dados = cab + regs
        ws.Range(ws.Cells(1, 1), ws.Cells(len(dados), 6)).Value = dados
        ws.Range("A1:F1").Font.Bold = True
        ws.Columns("A:F").AutoFit()
        wb.Save()
        print(f"Acessos:       {len(regs)} registro(s) na aba Acessos da planilha.")
    except Exception as e:
        print(f"Aviso: não consegui gravar a aba Acessos ({e}).")
    finally:
        try:
            if fechar and wb is not None: wb.Close(False)
            if fechar and xl is not None: xl.Quit()
        except Exception: pass

AFASTADOS = []
MESES_PT = ["janeiro", "fevereiro", "março", "abril", "maio", "junho", "julho", "agosto", "setembro", "outubro", "novembro", "dezembro"]

def info_afastamento(apelido, desde):
    """Afastado não paga as 3 faturas seguintes ao pedido; volta a ser cobrado na 4ª."""
    out = {"n": apelido, "desde": desde.isoformat() if desde else "", "volta": "", "fim": "", "terminou": False}
    if not desde:
        return out
    venc, ano, mes = [], desde.year, desde.month
    while len(venc) < 4:
        v = vencimento(ano, mes)
        if v > desde:
            venc.append(v)
        mes += 1
        if mes > 12:
            mes, ano = 1, ano + 1
    out["fim"] = venc[2].isoformat()
    out["volta"] = f"{MESES_PT[venc[3].month - 1]}/{venc[3].year}"
    out["terminou"] = datetime.date.today() > venc[2]
    return out

def renumerar_lista(mudancas):
    """Acerta o 'Numero Lista' (coluna D) da aba Dados Integrantes pelo Excel: ativos numerados, afastados/excluídos em branco."""
    if os.name != "nt":
        print("Aviso: a numeração da lista precisa ser acertada no Excel (só no Windows).")
        return
    try:
        import win32com.client as w32
    except ImportError:
        return
    xl = wb = None; fechar = False
    try:
        try:
            xl = w32.GetActiveObject("Excel.Application")
            for w in xl.Workbooks:
                if os.path.normcase(os.path.abspath(w.FullName)) == os.path.normcase(PLANILHA):
                    wb = w; break
        except Exception:
            xl = None
        if wb is None:
            xl = w32.DispatchEx("Excel.Application"); xl.Visible = False; xl.DisplayAlerts = False
            wb = xl.Workbooks.Open(PLANILHA); fechar = True
        if wb.ReadOnly:
            print("Aviso: planilha somente leitura; numeração da lista não foi acertada."); return
        pasta_bkp = os.path.join(PASTA_LANC, "backup"); os.makedirs(pasta_bkp, exist_ok=True)
        shutil.copy2(PLANILHA, os.path.join(pasta_bkp, f"backup_{datetime.datetime.now():%Y%m%d_%H%M%S}_{os.path.basename(PLANILHA)}"))
        ws = wb.Worksheets("Dados Integrantes")
        for linha, num, ap in mudancas:
            ws.Cells(linha, 4).Value = num if num else None
        try: wb.Application.Calculate()
        except Exception: pass
        wb.Save()
        print("Numero Lista acertado: " + ", ".join(f"{ap} = {num if num else '(afastado/excluído)'}" for _, num, ap in mudancas))
    except Exception as e:
        print(f"Aviso: não consegui acertar a numeração da lista ({e}).")
    finally:
        try:
            if fechar and wb is not None: wb.Close(False)
            if fechar and xl is not None: xl.Quit()
        except Exception: pass

def inativos_planilha():
    """Apelidos com Situação Afastado ou Excluído (para a verificação dos boletos pular)."""
    try:
        import openpyxl
        wb = openpyxl.load_workbook(PLANILHA, read_only=True, data_only=True)
        linhas = list(wb["Dados Integrantes"].iter_rows(min_row=1, max_col=60, values_only=True)); wb.close()
    except Exception:
        return set()
    cab = [str(x or "").lower() for x in linhas[0]]
    cs = next((k for k, v in enumerate(cab) if v.startswith("situa")), None)
    if cs is None:
        return set()
    return {str(r[1]).strip().lower() for r in linhas[1:] if r[1] and cs < len(r) and str(r[cs] or "").strip().lower()[:5] in ("afast", "exclu")}

# ---------------------------------------------------------------------
#  Notificações no celular (app "Insanos ADM"): pagamento novo, aniversário amanhã e fim de afastamento
#  Os celulares que ativaram ficam na planilha Google dos acessos (aba "Push"). Quem envia é este computador.
# ---------------------------------------------------------------------
ARQ_VAPID = os.path.join(PASTA_LANC, "push_vapid.json")          # chave secreta: fica só neste computador
ARQ_VAPID_PEM = os.path.join(PASTA_LANC, "push_vapid_privada.pem")
ARQ_AVISOS = os.path.join(PASTA_LANC, "notificacoes_enviadas.json")
HORA_AVISOS = "09:00"
HORAS_ANIVERSARIO = ["09:00", "13:00", "19:00"]   # aviso de aniversário amanhã: 3 vezes no dia

def chave_push_publica():
    """Cria (uma vez) o par de chaves das notificações e devolve a parte pública (vai no app)."""
    try:
        if os.path.exists(ARQ_VAPID) and os.path.exists(ARQ_VAPID_PEM):
            with open(ARQ_VAPID, encoding="utf-8") as f:
                return json.load(f)["pub"]
        from cryptography.hazmat.primitives.asymmetric import ec
        from cryptography.hazmat.primitives import serialization
        k = ec.generate_private_key(ec.SECP256R1())
        pub = k.public_key().public_bytes(serialization.Encoding.X962, serialization.PublicFormat.UncompressedPoint)
        os.makedirs(PASTA_LANC, exist_ok=True)
        with open(ARQ_VAPID_PEM, "wb") as f:
            f.write(k.private_bytes(serialization.Encoding.PEM, serialization.PrivateFormat.PKCS8, serialization.NoEncryption()))
        pub64 = base64.urlsafe_b64encode(pub).decode().rstrip("=")
        with open(ARQ_VAPID, "w", encoding="utf-8") as f:
            json.dump({"pub": pub64}, f)
        return pub64
    except Exception as e:
        print(f"Aviso: não consegui criar a chave das notificações ({e}).")
        return ""

def _pywebpush():
    try:
        import pywebpush
        return pywebpush
    except ImportError:
        import subprocess
        print("Instalando o componente das notificações (só na primeira vez)...")
        subprocess.run([sys.executable, "-m", "pip", "install", "--quiet", "pywebpush"], capture_output=True)
        try:
            import pywebpush
            return pywebpush
        except ImportError:
            print("Aviso: não consegui instalar o componente das notificações (pywebpush).")
            return None

def notificar(titulo, corpo, tag="aviso", zap=None):
    """Manda a notificação para todos os celulares que ativaram no app."""
    if not ACESSOS_URL or not os.path.exists(ARQ_VAPID_PEM):
        return
    import urllib.request, urllib.parse
    try:
        with urllib.request.urlopen(ACESSOS_URL + "?push=1&chave=" + urllib.parse.quote(ACESSOS_CHAVE), timeout=25) as r:
            inscritos = json.loads(r.read().decode("utf-8"))
    except Exception as e:
        _log(f"Notificação não enviada (não consegui buscar os celulares: {e})"); return
    if not isinstance(inscritos, list) or not inscritos:
        return
    wp = _pywebpush()
    if wp is None:
        return
    msg = {"title": titulo, "body": corpo, "tag": tag, "url": "adm.html"}
    if zap:      # mensagem pronta para o botão "📲 Enviar no WhatsApp" (a notificação tem limite de tamanho)
        msg["zap"] = zap
        if len(json.dumps(msg, ensure_ascii=False).encode("utf-8")) > 3400:
            msg.pop("zap"); _log("Aviso: a mensagem do WhatsApp ficou grande demais para ir na notificação (vai sem o botão).")
    dados = json.dumps(msg, ensure_ascii=False)
    enviados = 0
    for email, sub in inscritos:
        try:
            info = json.loads(sub) if isinstance(sub, str) else sub
            wp.webpush(subscription_info=info, data=dados, vapid_private_key=ARQ_VAPID_PEM,
                       vapid_claims={"sub": "mailto:painel@insanos.app"}, ttl=86400)
            enviados += 1
        except Exception as e:
            cod = getattr(getattr(e, "response", None), "status_code", None)
            if cod in (404, 410):          # celular desativou ou trocou: tira da lista
                try:
                    req = urllib.request.Request(ACESSOS_URL, data=json.dumps({"tipo": "push_remover", "chave": ACESSOS_CHAVE,
                                                 "endpoint": info.get("endpoint", "")}).encode("utf-8"), method="POST")
                    urllib.request.urlopen(req, timeout=20).read()
                except Exception:
                    pass
            else:
                _log(f"Notificação não entregue para {email}: {e}")
    _log(f"Notificação '{titulo}' enviada para {enviados} celular(es).")

def avisos_diarios(so_depois_das_9=False):
    """Aniversário amanhã e fim de afastamento: cada aviso sai uma vez só."""
    agora = datetime.datetime.now()
    if so_depois_das_9 and agora.strftime("%H:%M") < HORA_AVISOS:
        return
    try:
        with open(ARQ_AVISOS, encoding="utf-8") as f:
            feitos = json.load(f)
    except Exception:
        feitos = {}
    try:
        import openpyxl
        tmp = os.path.join(tempfile.gettempdir(), "insanos_avisos.xlsx")
        shutil.copy2(PLANILHA, tmp)
        wb = openpyxl.load_workbook(tmp, data_only=True, read_only=True)
        linhas = list(wb["Dados Integrantes"].iter_rows(min_row=1, max_col=60, values_only=True)); wb.close()
    except Exception as e:
        _log(f"Avisos diários: não consegui ler a planilha ({e})"); return
    cab = [_sem_acento(x) for x in linhas[0]]
    c_sit = next((k for k, v in enumerate(cab) if v.startswith("situa")), None)
    c_des = next((k for k, v in enumerate(cab) if "afastad" in v), None)
    hoje = agora.date(); amanha = hoje + datetime.timedelta(days=1)
    for r in linhas[1:]:
        ap = str(r[1] or "").strip()
        if not ap:
            continue
        sit = _sem_acento(r[c_sit]) if c_sit is not None and c_sit < len(r) else "ativo"
        nasc = r[9] if len(r) > 9 else None
        if isinstance(nasc, datetime.datetime):
            nasc = nasc.date()
        if (not sit or sit.startswith("ativo")) and isinstance(nasc, datetime.date) and (nasc.month, nasc.day) == (amanha.month, amanha.day):
            # aniversário amanhã: 3 avisos no dia (9h, 13h e 19h), cada um uma vez só
            hm = agora.strftime("%H:%M")
            horario = max((h for h in HORAS_ANIVERSARIO if h <= hm), default=None)
            chave = f"aniv:{amanha.isoformat()}:{ap}:{horario}"
            if horario and chave not in feitos:
                titulo = {"09:00": "🎂 Aniversário amanhã", "13:00": "🎂 Lembrete: aniversário amanhã",
                          "19:00": "🎂 Último lembrete: aniversário amanhã"}.get(horario, "🎂 Aniversário amanhã")
                notificar(titulo, f"Amanhã ({amanha:%d/%m}) é aniversário de {ap}: faz {amanha.year - nasc.year} anos. Hora de preparar o fly!", "aniversario")
                feitos[chave] = agora.isoformat(timespec="minutes")
        if sit.startswith("afast"):
            dd = r[c_des] if c_des is not None and c_des < len(r) else None
            if isinstance(dd, datetime.datetime):
                dd = dd.date()
            if isinstance(dd, datetime.date):
                inf = info_afastamento(ap, dd)
                chave = f"afast:{ap}:{inf['fim']}"
                if inf["terminou"] and chave not in feitos:
                    notificar("⚠️ Fim de afastamento", f"O afastamento de {ap} terminou em {datetime.date.fromisoformat(inf['fim']):%d/%m}. Reative ou exclua no Cadastro.", "afastamento")
                    feitos[chave] = agora.isoformat(timespec="minutes")
    try:
        with open(ARQ_AVISOS, "w", encoding="utf-8") as f:
            json.dump(feitos, f, ensure_ascii=False, indent=1)
    except Exception:
        pass

def garantir_tarefa_avisos():
    """Tarefa do Windows, todo dia às 9h, só para os avisos diários (criada sozinha uma vez)."""
    if os.name != "nt" or not ACESSOS_URL:
        return
    import subprocess
    nome = "Insanos - Avisos 9h"
    if subprocess.run(["schtasks", "/Query", "/TN", nome], capture_output=True, text=True).returncode == 0:
        return
    pyw = os.path.join(os.path.dirname(sys.executable), "pythonw.exe")
    if not os.path.exists(pyw):
        pyw = sys.executable
    r = subprocess.run(["schtasks", "/Create", "/F", "/SC", "DAILY", "/ST", HORA_AVISOS, "/TN", nome,
                        "/TR", f'"{pyw}" "{os.path.abspath(__file__)}" --avisos'], capture_output=True, text=True)
    print("Avisos diários (9h): tarefa criada." if r.returncode == 0 else f"Aviso: não consegui criar a tarefa das 9h ({(r.stderr or r.stdout).strip()}).")

def agenda_ativa():
    """Horários das tarefas de verificação que existem no Agendador do Windows."""
    if os.name != "nt":
        return []
    import subprocess
    ativos = []
    for h in HORARIOS_VERIFICACAO:
        r = subprocess.run(["schtasks", "/Query", "/TN", f"Insanos - Verificar pagamentos {h[:2]}h"], capture_output=True, text=True)
        if r.returncode == 0:
            ativos.append(h)
    return ativos

HORARIOS_VERIFICACAO = [f"{h:02d}:00" for h in range(8, 18)]          # de hora em hora, das 8h às 17h
HORARIOS_ANTIGOS = ["11:00", "13:00", "14:00", "17:00", "19:00", "20:00"]  # tarefas antigas, removidas ao ativar/desativar

def agendar(ativar=True):
    """Cria (ou remove) as tarefas do Agendador do Windows que verificam os pagamentos."""
    import subprocess
    pyw = os.path.join(os.path.dirname(sys.executable), "pythonw.exe")
    if not os.path.exists(pyw):
        pyw = sys.executable
    script = os.path.abspath(__file__)
    for h in HORARIOS_ANTIGOS:
        if h not in HORARIOS_VERIFICACAO:
            subprocess.run(["schtasks", "/Delete", "/F", "/TN", f"Insanos - Verificar pagamentos {h[:2]}h"], capture_output=True, text=True)
    for h in HORARIOS_VERIFICACAO:
        nome = f"Insanos - Verificar pagamentos {h[:2]}h"
        if ativar:
            cmd = ["schtasks", "/Create", "/F", "/SC", "DAILY", "/ST", h, "/TN", nome,
                   "/TR", f'"{pyw}" "{script}" --pagamentos']
        else:
            cmd = ["schtasks", "/Delete", "/F", "/TN", nome]
        r = subprocess.run(cmd, capture_output=True, text=True)
        ok = r.returncode == 0
        print(f"  {'Tarefa criada' if ativar else 'Tarefa removida'} ({h}): {'ok' if ok else 'falhou - ' + (r.stderr or r.stdout).strip()}")

# ---------------------------------------------------------------------
#  Publicação no link (GitHub Pages): https://eduardovgmanso-tech.github.io/insanos-extremo-norte/
# ---------------------------------------------------------------------
GITHUB_USUARIO = "eduardovgmanso-tech"
GITHUB_REPO = "insanos-extremo-norte"
SITE_URL = f"https://{GITHUB_USUARIO}.github.io/{GITHUB_REPO}/"
# No link, o ADM abre só com os CPFs do comando (a senha coletiva não vale lá). Troque para True se quiser liberar.
ADM_NO_LINK_ACEITA_SENHA_COLETIVA = False
# Nome completo e CPF dos integrantes no ADM do link (aba Kit/colete). True = vão para o link; False = só no computador.
LINK_COM_NOME_E_CPF = True
# A chave (token) fica só neste computador, fora da pasta do painel:
ARQ_TOKEN = os.path.join(os.environ.get("APPDATA") or os.path.expanduser("~"), "PainelInsanos", "github.json")
ARQ_PUBLICACAO = os.path.join(PASTA, "Lancamentos", "publicacao.json")

def _github(metodo, caminho, token, corpo=None):
    import urllib.request
    req = urllib.request.Request(f"https://api.github.com{caminho}", method=metodo,
        data=json.dumps(corpo).encode("utf-8") if corpo is not None else None,
        headers={"Authorization": f"Bearer {token}", "Accept": "application/vnd.github+json",
                 "X-GitHub-Api-Version": "2022-11-28", "User-Agent": "PainelInsanos",
                 "Content-Type": "application/json"})
    with urllib.request.urlopen(req, timeout=120) as r:
        return json.loads(r.read().decode("utf-8") or "{}")

def configurar_publicacao():
    import urllib.error
    print("Configuração da publicação no link " + SITE_URL)
    print("Cole abaixo a chave (token) que você gerou no GitHub e aperte Enter.")
    print("(no Prompt, para colar use o botão direito do mouse)\n")
    token = input("Chave: ").strip().strip('"').strip("'").replace(" ", "")
    if token.lower().startswith("bearer"):
        token = token[6:].strip()
    if len(token) < 20:
        print("Essa chave parece incompleta. Nada foi salvo."); return
    print(f"\nTestando a chave ({len(token)} caracteres, começa com {token[:11]}...)")
    try:
        repo = _github("GET", f"/repos/{GITHUB_USUARIO}/{GITHUB_REPO}", token)
        perm = repo.get("permissions") or {}
        if perm and not perm.get("push"):
            print("A chave abre o repositório, mas não tem permissão de escrita. Gere de novo com "
                  "Contents: Read and write. Nada foi salvo."); return
        print(f"Chave válida. Repositório encontrado: {repo.get('full_name')}")
    except urllib.error.HTTPError as e:
        if e.code == 401:
            print("O GitHub recusou essa chave (401). Gere outra e copie o código inteiro. Nada foi salvo.")
        elif e.code == 404:
            print("A chave foi aceita, mas não enxerga o repositório (404). Ao gerar a chave, escolha "
                  f"'Only select repositories' > {GITHUB_REPO}. Nada foi salvo.")
        else:
            print(f"O GitHub respondeu com erro {e.code}. Nada foi salvo.")
        return
    except Exception as e:
        print(f"Não consegui testar agora ({e}). Confira a internet e tente de novo. Nada foi salvo."); return
    os.makedirs(os.path.dirname(ARQ_TOKEN), exist_ok=True)
    with open(ARQ_TOKEN, "w", encoding="utf-8") as f:
        json.dump({"token": token, "usuario": GITHUB_USUARIO, "repo": GITHUB_REPO}, f)
    print(f"\nChave guardada em: {ARQ_TOKEN}")
    print("Ela fica só neste computador e não vai para a pasta do painel nem para os zips.")

def _assinatura_conteudo(dados_geral, dados_adm):
    """Resumo do conteúdo que importa (ignora data/hora de geração e da última verificação)."""
    import hashlib, copy
    adm = copy.deepcopy(dados_adm); adm.pop("molde", None)
    if isinstance(adm.get("pag"), dict):
        adm["pag"].pop("verif", None)
    with open(TEMPLATE, encoding="utf-8") as f:
        tpl = f.read()
    bruto = tpl + json.dumps(dados_geral, ensure_ascii=False, sort_keys=True) + json.dumps(adm, ensure_ascii=False, sort_keys=True)
    return hashlib.sha256(bruto.encode("utf-8")).hexdigest()

def publicar(arquivos):
    """Envia os arquivos {nome_no_site: conteúdo} para o repositório do GitHub (substitui a versão anterior)."""
    import urllib.error
    if not os.path.exists(ARQ_TOKEN):
        print("Publicação no link: ainda não configurada (rode o ConfigurarPublicacao.bat).")
        return False
    try:
        with open(ARQ_TOKEN, encoding="utf-8") as f:
            cfg = json.load(f)
        token = cfg["token"]
        base = f"/repos/{cfg.get('usuario', GITHUB_USUARIO)}/{cfg.get('repo', GITHUB_REPO)}/contents/"
        quando = datetime.datetime.now().strftime("%d/%m/%Y %H:%M")
        import hashlib
        for nome, conteudo in {**arquivos, ".nojekyll": ""}.items():
            dados = conteudo if isinstance(conteudo, bytes) else conteudo.encode("utf-8")   # texto ou imagem
            try:
                sha = _github("GET", base + nome, token).get("sha")
            except urllib.error.HTTPError as e:
                if e.code != 404:
                    raise
                sha = None
            if nome == ".nojekyll" and sha:
                continue
            if sha and sha == hashlib.sha1(b"blob %d\0" % len(dados) + dados).hexdigest():
                continue       # igual ao que já está no link: não precisa enviar de novo
            corpo = {"message": f"Atualiza {nome} ({quando})",
                     "content": base64.b64encode(dados).decode()}
            if sha:
                corpo["sha"] = sha
            _github("PUT", base + nome, token, corpo)
        print(f"Publicado no link: {SITE_URL}adm.html  (a versão nova aparece em 1 a 2 minutos)")
        return True
    except Exception as e:
        _log(f"Aviso: não consegui publicar no link ({e}). Os arquivos do computador foram gerados normalmente.")
        return False

# Publicar em toda geração (inclusive nas verificações automáticas), para o horário da tela de CPF ficar sempre atual.
# Uma vez por mês o histórico do GitHub é limpo (fica só a versão atual), para o repositório não crescer.
PUBLICAR_SEMPRE = True

def limpar_historico_github():
    """Deixa o repositório só com a versão atual (um único commit), mantendo os arquivos e o site."""
    try:
        with open(ARQ_TOKEN, encoding="utf-8") as f:
            cfg = json.load(f)
        token = cfg["token"]
        repo = f"/repos/{cfg.get('usuario') or GITHUB_USUARIO}/{cfg.get('repo') or GITHUB_REPO}"
        ref = _github("GET", f"{repo}/git/ref/heads/main", token)
        atual = _github("GET", f"{repo}/git/commits/{ref['object']['sha']}", token)
        novo = _github("POST", f"{repo}/git/commits", token,
                       {"message": f"Painel (histórico limpo em {datetime.date.today():%d/%m/%Y})",
                        "tree": atual["tree"]["sha"], "parents": []})
        _github("PATCH", f"{repo}/git/refs/heads/main", token, {"sha": novo["sha"], "force": True})
        print("Histórico do GitHub limpo (limpeza mensal): ficou só a versão atual.")
        return True
    except Exception as e:
        _log(f"Aviso: não consegui limpar o histórico do GitHub ({e}). Tento de novo na próxima publicação.")
        return False

APP_NOME = "Extremo Norte SP3"

# O link é só do comando (ADM). O endereço geral não tem mais o painel: só leva para a entrada do ADM, sem nenhum dado.
INDEX_SO_ADM = """<!DOCTYPE html>
<html lang="pt-BR"><head><meta charset="UTF-8"><meta name="viewport" content="width=device-width, initial-scale=1">
<meta name="robots" content="noindex"><title>Insanos ADM</title>
<link rel="manifest" href="manifest-adm.json"><meta name="theme-color" content="#0b0b0b"><link rel="icon" href="icon-192.png">
<meta http-equiv="refresh" content="0; url=adm.html"><script>location.replace("adm.html"+location.search)</script>
<style>body{background:#000;color:#ECE7DC;font-family:Segoe UI,Roboto,Arial,sans-serif;display:grid;place-items:center;height:100vh;margin:0}a{color:#8FD9A3}</style>
</head><body><p>Abrindo o painel do comando… <a href="adm.html">toque aqui</a> se não abrir.</p></body></html>
"""

def arquivos_app():
    """Arquivos que fazem o link virar "app" (instalável no celular): manifestos, ajudante (sw.js) e ícones."""
    def manifesto(nome, curto, inicio):
        return json.dumps({"id": inicio, "name": nome, "short_name": curto, "start_url": inicio, "scope": "./",
                           "display": "standalone", "orientation": "portrait", "background_color": "#000000", "theme_color": "#0b0b0b",
                           "lang": "pt-BR", "icons": [{"src": f"icon-{t}.png", "sizes": f"{t}x{t}", "type": "image/png", "purpose": "any maskable"} for t in (192, 512)]},
                          ensure_ascii=False, indent=1)
    sw = """// Painel Insanos: busca sempre a versão nova; sem internet, abre a última versão vista.
const CACHE = "insanos-v7";
self.addEventListener("install", e => self.skipWaiting());
self.addEventListener("activate", e => e.waitUntil(self.clients.claim()));
self.addEventListener("fetch", e => {
  const r = e.request;
  if (r.method !== "GET" || new URL(r.url).origin !== location.origin) return;
  if (/\\.mp4$/i.test(new URL(r.url).pathname)) return;      // vídeo de fundo: direto da rede (guardar vídeo no app trava a reprodução)
  e.respondWith(fetch(r, {cache: "no-store"}).then(resp => {
    if (resp.ok) { const c = resp.clone(); caches.open(CACHE).then(k => k.put(r, c)); }
    return resp;
  }).catch(() => caches.match(r, {ignoreSearch: true})));
});
self.addEventListener("push", e => {
  let d = {};
  try { d = e.data.json(); } catch (_) { d = {title: "Insanos ADM", body: e.data ? e.data.text() : ""}; }
  const op = {body: d.body || "", icon: "icon-192.png", badge: "icon-192.png", tag: d.tag || "aviso", renotify: true,
              data: {url: d.url || "adm.html", zap: d.zap || ""}};
  if (d.zap) op.actions = [{action: "zap", title: "📲 Enviar no WhatsApp"}];
  e.waitUntil(self.registration.showNotification(d.title || "Insanos ADM", op));
});
self.addEventListener("notificationclick", e => {
  e.notification.close();
  const dd = e.notification.data || {};
  // botão "Enviar no WhatsApp" (Android/computador) ou toque na notificação no iPhone (que não mostra botões):
  // abre o painel com a caixa da mensagem (o site do WhatsApp quebra os emojis; de lá o painel abre o app direto)
  if (dd.zap && (e.action === "zap" || /iPhone|iPad|iPod/.test(navigator.userAgent))) {
    e.waitUntil(caches.open("insanos-zap").then(c => c.put("zap", new Response(JSON.stringify({ts: Date.now(), txt: dd.zap})))))
      .then(() => clients.matchAll({type: "window", includeUncontrolled: true}))
      .then(ws => { for (const w of ws) if (w.url.includes("adm.html")) { w.postMessage({zap: 1}); return w.focus(); }
                    return clients.openWindow("adm.html#zap"); }));
    return;
  }
  e.waitUntil(clients.matchAll({type: "window", includeUncontrolled: true}).then(ws => {
    for (const w of ws) if (w.url.includes("adm.html")) return w.focus();
    return clients.openWindow((e.notification.data && e.notification.data.url) || "adm.html");
  }));
});
"""
    arq = {"manifest-adm.json": manifesto(f"Insanos ADM {APP_NOME}", "Insanos ADM", "./adm.html"),
           "sw.js": sw}
    for t in (192, 512):
        caminho = os.path.join(PASTA, "app", f"icon-{t}.png")
        if os.path.exists(caminho):
            with open(caminho, "rb") as f:
                arq[f"icon-{t}.png"] = f.read()
    return arq

def publicar_se_mudou(arquivos, assinatura, forcar=False):
    """Publica só quando o conteúdo mudou (ou quando forçado pelo PublicarAgora.bat)."""
    forcar = forcar or PUBLICAR_SEMPRE
    reg = {}
    if os.path.exists(ARQ_PUBLICACAO):
        try:
            with open(ARQ_PUBLICACAO, encoding="utf-8") as f:
                reg = json.load(f)
        except Exception:
            reg = {}
    if not forcar and reg.get("destino") == "github" and reg.get("assinatura") == assinatura:
        print("Link: nada mudou desde a última publicação. Publicação pulada.")
        return
    if publicar(arquivos):
        mes = datetime.date.today().strftime("%Y-%m")
        limpo = reg.get("limpeza")
        # 1ª publicação "só ADM": limpa o histórico na hora (tira do GitHub as versões antigas do painel geral)
        if (limpo != mes or not reg.get("so_adm")) and limpar_historico_github():
            limpo = mes
        os.makedirs(os.path.dirname(ARQ_PUBLICACAO), exist_ok=True)
        with open(ARQ_PUBLICACAO, "w", encoding="utf-8") as f:
            json.dump({"destino": "github", "assinatura": assinatura, "limpeza": limpo, "so_adm": True,
                       "quando": datetime.datetime.now().strftime("%d/%m/%Y %H:%M")}, f)

def main():
    try:
        import openpyxl
    except ImportError:
        falha("Biblioteca openpyxl não encontrada. Rode no prompt:  pip install openpyxl")
    try:
        import cryptography  # noqa: F401
    except ImportError:
        falha("Biblioteca cryptography não encontrada. Rode no prompt:  pip install cryptography")
    for arq in (PLANILHA, IMAGEM, TEMPLATE):
        if not os.path.exists(arq):
            falha(f"Arquivo não encontrado: {os.path.basename(arq)} (precisa estar na mesma pasta do .bat)")

    if "--configurar-publicacao" in sys.argv:
        configurar_publicacao()
        return
    if "--agendar" in sys.argv or "--desagendar" in sys.argv:
        agendar("--agendar" in sys.argv)
        return
    if "--testar-notificacao" in sys.argv:    # TestarNotificacao.bat
        if not os.path.exists(ARQ_VAPID_PEM):
            print("Rode o GerarPainel primeiro (ele cria a chave das notificações)."); return
        notificar("🔔 Teste do Insanos ADM", "As notificações estão funcionando neste aparelho. 💀", "teste")
        print("Teste enviado. Confira o celular (veja também Lancamentos\\verificacao_pagamentos.log).")
        return
    if "--avisos" in sys.argv:                # tarefa das 9h: só os avisos diários no celular
        avisos_diarios()
        return
    if "--pagamentos" in sys.argv:            # chamado pela tarefa agendada ou pelo VerificarPagamentos.bat
        forcar = "--forcar" in sys.argv
        importar_pagamentos_manuais()
        if not verificar_pagamentos(forcar=forcar) and not forcar:
            return                                # fora do período ou rodada pulada: não refaz o painel
    migrar_staff()
    importar_lancamentos()
    importar_integrantes()
    if "--pagamentos" not in sys.argv:
        importar_pagamentos_manuais()
    print("Lendo a planilha (pode levar alguns segundos)...")
    # copia para um temporário: funciona mesmo com a planilha aberta no Excel
    tmp = os.path.join(tempfile.gettempdir(), "insanos_leitura.xlsx")
    shutil.copy2(PLANILHA, tmp)
    wb = openpyxl.load_workbook(tmp, data_only=True, read_only=True)
    for aba in ("Presenca", "Dados Integrantes"):
        if aba not in wb.sheetnames:
            falha(f"Aba '{aba}' não encontrada na planilha.")

    # coluna de e-mail: procurada pelo título na linha 1 da aba Dados Integrantes ("E-mail", "Email"...)
    cab_di = next(wb["Dados Integrantes"].iter_rows(min_row=1, max_row=1, max_col=60, values_only=True))
    col_email = next((k for k, v in enumerate(cab_di) if v and "mail" in str(v).lower()), None)
    _semac = lambda v: "".join(ch for ch in __import__("unicodedata").normalize("NFD", str(v or "").lower()) if not __import__("unicodedata").combining(ch)).strip()
    col_sit = next((k for k, v in enumerate(cab_di) if _semac(v).startswith("situa")), None)
    col_desde = next((k for k, v in enumerate(cab_di) if "afastad" in _semac(v)), None)
    situ = {}          # apelido -> {"s": ativo|afastado|excluido, "desde": data, "linha": n, "d": Numero Lista, "grau": ...}
    ordem_di = []      # apelidos na ordem das linhas da aba
    emails = {}
    # Dados Integrantes: apelido -> (número, cargo, ordem da hierarquia)
    info = {}
    cpfs, extra_colete = {}, {}
    for n_linha, r_full in enumerate(wb["Dados Integrantes"].iter_rows(min_row=2, max_col=60, values_only=True), start=2):
        r = r_full[:14]
        apelido = texto(r[1])
        if apelido and apelido.lower() not in situ:
            sv = _semac(r_full[col_sit]) if col_sit is not None and col_sit < len(r_full) else ""
            dd = r_full[col_desde] if col_desde is not None and col_desde < len(r_full) else None
            if isinstance(dd, datetime.datetime):
                dd = dd.date()
            situ[apelido.lower()] = {"s": "afastado" if sv.startswith("afast") else "excluido" if sv.startswith("exclu") else "ativo",
                                     "desde": dd if isinstance(dd, datetime.date) else None, "linha": n_linha,
                                     "d": r[3], "grau": str(r[4]) if r[4] is not None else "", "apelido": apelido}
            ordem_di.append(apelido.lower())
        if apelido and apelido.lower() not in info:
            num = r[0]
            num = str(int(num)) if isinstance(num, (int, float)) else texto(num)
            tel = r[6]
            tel = str(int(tel)) if isinstance(tel, (int, float)) else "".join(ch for ch in texto(tel) if ch.isdigit())
            nasc = r[9]
            if isinstance(nasc, datetime.datetime):
                nasc = nasc.date()
            info[apelido.lower()] = (num, texto(r[2]), r[3] if isinstance(r[3], (int, float)) else 999, tel,
                                     nasc if isinstance(nasc, datetime.date) else None,
                                     str(r[5]) if r[5] is not None else "")   # "Lista/grau" (ex.: "01. VI "), sem cortar espaços
            cpfs[apelido.lower()] = "".join(ch for ch in texto(r[11]) if ch.isdigit())
            if col_email is not None:
                emails[apelido.lower()] = texto(r_full[col_email] if col_email < len(r_full) else "").lower()
            # dados para a Solicitação kit/colete (vão SÓ para o ADM do computador, nunca para o link)
            extra_colete[apelido.lower()] = {"nc": texto(r[7]), "cpf": cpfs[apelido.lower()],
                                             "col": texto(r[12] if len(r) > 12 else "").upper(),
                                             "tamc": texto(r[13] if len(r) > 13 else "").upper()}

    # acesso pelo E-MAIL (não mais pelo CPF)
    if col_email is None:
        falha("Não encontrei a coluna de e-mail na aba Dados Integrantes. Coloque o título 'E-mail' na linha 1 dessa coluna.")
    def email_de(nome):
        e = emails.get(nome.lower(), "")
        if "@" not in e:
            falha(f"E-mail de '{nome}' não encontrado ou inválido na aba Dados Integrantes (coluna '{cab_di[col_email]}').")
        return e
    EMAIL_APELIDO.clear()
    EMAIL_APELIDO.update({e: ap for ap, e in emails.items() if "@" in e})
    EMAIL_APELIDO_INV.clear()
    EMAIL_APELIDO_INV.update({ap: e for ap, e in emails.items() if "@" in e})
    senhas = [email_de(nome) for nome in AUTORIZADOS]
    senhas_so_geral = [e for e in (email_de(nome) for nome in AUTORIZADOS_SO_GERAL) if e not in senhas]

    linhas = list(wb["Presenca"].iter_rows(min_row=1, max_row=200, max_col=60, values_only=True))
    # abas com o mesmo formato da Presenca (Insights e Staff), se existirem na planilha
    abas_extras = {}
    for chave, nome_aba in (("ins", "insights"), ("staff", "staff")):
        aba = next((n for n in wb.sheetnames if n.strip().lower() == nome_aba), None)
        if aba:
            abas_extras[chave] = list(wb[aba].iter_rows(min_row=1, max_row=300, max_col=60, values_only=True))
    pgto = list(wb["PGTO"].iter_rows(min_row=2, max_row=250, max_col=10, values_only=True)) if "PGTO" in wb.sheetnames else []
    msg_pag = [r[0] for r in wb["Pagamentos"].iter_rows(min_row=1, max_row=80, min_col=3, max_col=3, values_only=True)] \
        if "Pagamentos" in wb.sheetnames else []
    wb.close()
    lin_nomes, lin_cargos = linhas[5], linhas[6]          # linhas 6 e 7 da aba

    # colunas de integrantes: da coluna D até antes de "Presença"
    cols = []
    for c in range(3, len(lin_nomes)):
        nome = texto(lin_nomes[c])
        if not nome or nome.lower().startswith("presen"):
            break
        cols.append(c)

    membros, avisos = [], []
    for c in cols:
        nome = texto(lin_nomes[c])
        num, cargo, ordem, tel, nasc, lg = info.get(nome.lower(), ("-", texto(lin_cargos[c]), 999, "", None, ""))
        if nome.lower() not in info:
            avisos.append(nome)
        m = {"n": nome, "c": cargo or texto(lin_cargos[c]), "num": num, "o": ordem, "tel": tel, "lg": lg,
             "x": extra_colete.get(nome.lower(), {})}
        if nasc:  # nascimento (coluna "Nascimento") e idade calculada na data de hoje
            hoje = datetime.date.today()
            m["nasc"] = nasc.isoformat()
            m["idade"] = hoje.year - nasc.year - ((hoje.month, hoje.day) < (nasc.month, nasc.day))
        membros.append(m)

    # ano padrão para datas digitadas como texto ("04-jul")
    anos = [r[1].year for r in linhas[7:] if isinstance(r[1], (datetime.datetime, datetime.date))]
    ano = max(anos) if anos else datetime.date.today().year

    eventos = []
    for r in linhas[7:]:
        if not r[1] or not r[2]:
            continue
        d = data_iso(r[1], ano)
        if not d:
            continue
        p, stf = [], []
        for c in cols:
            v = texto(r[c]).upper()
            # Sim = presença, STAFF = presença (fez staff), Não = falta, "-" = não conta, vazio = sem registro
            p.append(1 if v in ("SIM", "STAFF") else 0 if v in ("NÃO", "NAO") else "-" if v in ("-", "–", "—") else None)
            stf.append(1 if v == "STAFF" else 0)
        eventos.append({"d": d, "e": texto(r[2]), "p": p, "st": stf})

    # ordena integrantes pela hierarquia, levando junto as colunas de presença
    idx = sorted(range(len(membros)), key=lambda i: membros[i]["o"])
    # afastados e excluídos (coluna "Situação") ficam fora de tudo: análises, listas, pagamentos
    if col_sit is not None:
        idx = [i for i in idx if situ.get(membros[i]["n"].lower(), {}).get("s", "ativo") == "ativo"]
    membros = [membros[i] for i in idx]
    for e in eventos:
        e["p"] = [e["p"][i] for i in idx]
        e["st"] = [e.get("st", [0] * len(membros))[i] for i in idx]
    # numeração da lista: só os ativos, na ordem das linhas da Dados Integrantes (1, 2, 3...)
    renumerar = []
    if col_sit is not None:
        n_ativo = 0
        for ap in ordem_di:
            sx = situ[ap]
            certo = None
            if sx["s"] == "ativo":
                n_ativo += 1; certo = n_ativo
            atual = sx["d"] if isinstance(sx["d"], (int, float)) else None
            if (int(atual) if atual is not None else None) != certo:
                renumerar.append((sx["linha"], certo, sx["apelido"]))
            sx["certo"] = certo
        for m in membros:   # "Lista/grau" sempre com o número certo (mesmo antes de o Excel recalcular)
            sx = situ.get(m["n"].lower())
            if sx and sx.get("certo"):
                m["lg"] = f"{sx['certo']:02d}. {sx['grau']}"; m["o"] = sx["certo"]
    if renumerar:
        renumerar_lista(renumerar)
    AFASTADOS.clear()
    for ap in ordem_di:
        sx = situ[ap]
        if sx["s"] == "afastado":
            AFASTADOS.append(info_afastamento(sx["apelido"], sx["desde"]))

    def eventos_da_aba(lin):
        """Eventos de uma aba no formato da Presenca, alinhados aos integrantes pelo apelido."""
        if len(lin) < 8:
            return []
        pos = {m["n"].strip().lower(): k for k, m in enumerate(membros)}
        cols_aba = []
        for c in range(3, len(lin[5])):
            nm = texto(lin[5][c])
            if not nm or nm.lower().startswith("presen"):
                break
            if nm.lower() in pos:
                cols_aba.append((c, pos[nm.lower()]))
        anos_a = [r[1].year for r in lin[7:] if isinstance(r[1], (datetime.datetime, datetime.date))]
        ano_a = max(anos_a) if anos_a else ano
        evs = []
        for r in lin[7:]:
            if len(r) < 3 or not r[1] or not r[2]:
                continue
            d = data_iso(r[1], ano_a)
            if not d:
                continue
            pp = [None] * len(membros)
            for c, k in cols_aba:
                v = texto(r[c] if c < len(r) else "").upper()
                pp[k] = 1 if v in ("SIM", "STAFF") else 0 if v in ("NÃO", "NAO") else "-" if v in ("-", "–", "—") else None
            evs.append({"d": d, "e": texto(r[2]), "p": pp})
        return evs
    eventos_ins = eventos_da_aba(abas_extras.get("ins", []))
    eventos_staff = []        # staff agora é "STAFF" na aba Presenca (a aba Staff não é mais usada)

    with open(IMAGEM, "rb") as f:
        img = base64.b64encode(f.read()).decode()
    with open(TEMPLATE, encoding="utf-8") as f:
        html = f.read()
    gerado = datetime.datetime.now().strftime("%d/%m/%Y às %H:%M")
    global MANUT_HTML
    MANUT_HTML = AVISO_MANUTENCAO if os.path.exists(os.path.join(PASTA, "MANUTENCAO.txt")) else ""
    if MANUT_HTML:
        print("ATENÇÃO: aviso de manutenção LIGADO (para tirar, rode o DesligarManutencao.bat).")
    html = html.replace("__IMG__", img).replace("__GERADO__", gerado).replace("__MANUT__", MANUT_HTML)

    def gravar(caminho, dados, senha=None, sem_js="", extras=(), rotulo="E-mail", so_geral=False):
        pagina = html.replace("__NOSCRIPT__", sem_js, 1).replace("__DATA__", json.dumps(dados, ensure_ascii=False), 1)
        if senha:
            base = senhas + (senhas_so_geral if so_geral else [])
            pagina = proteger(pagina, base + [x for x in extras if x not in base],
                              os.path.splitext(os.path.basename(caminho or SAIDA_ADM))[0], senha, img, rotulo)
        if caminho:
            with open(caminho, "w", encoding="utf-8") as f:
                f.write(pagina)
        return pagina

    sem_tel = [{k: v for k, v in m.items() if k not in ("tel", "x")} for m in membros]
    # 1) geral (para compartilhar) e 2) ADM (com telefones)
    print("Protegendo o painel geral e o ADM com senha...")
    extras = ["".join(ch for ch in x if ch.isdigit()) for x in SENHAS_GERAL]
    global INFO_LOGIN
    try:
        with open(ARQ_VERIF, encoding="utf-8") as _f:
            _v = json.load(_f)
        _v["novos"] = _v.get("novos_pagos", 0)
    except Exception:
        _v = {}
    INFO_LOGIN = f"Última atualização: {gerado}"
    if ULTIMA_ALTERACAO_TXT:
        import html as _h
        INFO_LOGIN += f"<br>Última alteração: {_h.escape(ULTIMA_ALTERACAO_TXT)}"
    if _v.get("quando"):
        INFO_LOGIN += f"<br>Pagamentos verificados: {_v['quando']}"
    # tela de CPF do ADM: ponto verde quando há pagamento novo que você ainda não viu (some ao abrir a tela Pagamentos)
    _up = _v.get("ultimo_pagamento") or {}
    ULTIMO_PAG_TS = _up.get("ts", "")
    INFO_LOGIN_ADM = INFO_LOGIN + (
        f'<span id="novoPagTxt" hidden> · novo pagamento {_up.get("hora", "")}</span>'
        f'<span id="novoPag" class="novo" data-ts="{_up.get("ts", "")}" title="Pagamento novo: abra a tela Pagamentos" hidden></span>'
        if _up.get("ts") else "")
    pagina_geral = gravar(SAIDA, {"modo": "geral", "m": sem_tel, "ev": eventos, "ev_ins": eventos_ins, "ev_staff": eventos_staff}, "Painel de presença da divisão. Acesso restrito aos irmãos.",
           extras=[x for x in extras if x], rotulo="E-mail ou senha de acesso", so_geral=True)
    # o painel individual virou imagem (botão "Gerar imagem" no ADM): o molde do HTML individual não é mais necessário
    molde = base64.b64encode(html.encode("utf-8")).decode() if GERAR_INDIVIDUAIS else ""
    # pagamentos (só no ADM): status da aba PGTO, link do boleto e a mensagem pronta da aba Pagamentos
    pag = {"status": {}, "boleto": {}, "msg": "", "venc": None, "verif": None}
    for r in pgto:
        rotulo = texto(r[1])
        if not rotulo:
            continue
        for m in membros:
            ap = m["n"]
            if rotulo.lower().endswith(ap.lower()) and ap not in pag["status"]:
                pag["status"][ap] = texto(r[2]) or "Em aberto"
                link = texto(r[9])
                if link.startswith("http"):
                    pag["boleto"][ap] = link
                break
    while msg_pag and msg_pag[-1] in (None, ""):
        msg_pag.pop()
    pag["msg"] = "\n".join(texto(x) for x in msg_pag)
    hoje = datetime.date.today()
    pag["venc"] = vencimento(hoje.year, hoje.month).isoformat()
    import re as _re
    _v = vencimento(hoje.year, hoje.month)
    pag["msg"] = _re.sub(r"\(Vencimento\s*\d{1,2}/\d{1,2}/\d{2,4}\)", f"(Vencimento {_v:%d/%m/%Y})", pag["msg"])
    pag["agenda"] = agenda_ativa()
    pag["verif_antiga"] = False
    pag["horarios"] = HORARIOS_VERIFICACAO
    pag["travado"], pag["manual"], pag["boleto_mes"], pag["pago_em"] = {}, {}, {}, {}
    _man = _manual_carregar(vencimento(hoje.year, hoje.month))
    for m in membros:
        if m["n"] in _man:
            pag["manual"][m["n"]] = _man[m["n"]]
    pag["obs"] = {}
    if os.path.exists(ARQ_VERIF):                                   # criado pela verificação automática
        try:
            with open(ARQ_VERIF, encoding="utf-8") as f:
                v = json.load(f)
            pag["verif"] = {"quando": v.get("quando"), "ciclo": v.get("ciclo"), "novos": v.get("novos_pagos", 0)}
            mesmo_ciclo = str(v.get("ciclo") or "")[:7] == pag["venc"][:7]
            if not mesmo_ciclo:
                pag["verif_antiga"] = True     # a última verificação é de outro mês: nada fica travado como pago
            if mesmo_ciclo:
                for rot, ts in (v.get("pagos_em") or {}).items():
                    for m in membros:
                        if ts and str(rot).lower().endswith(m["n"].lower()):
                            pag["pago_em"][m["n"]] = ts; break
            for it in v.get("itens", []):
                selo_pago = bool(it.get("selo")) and "pag" in str(it.get("selo")).lower()
                if mesmo_ciclo and selo_pago and str(it.get("venc_boleto") or "")[:7] == pag["venc"][:7]:
                    for m in membros:
                        if str(it.get("rotulo", "")).lower().endswith(m["n"].lower()):
                            pag["travado"][m["n"]] = True; break
                vb = str(it.get("venc_boleto") or "")
                if vb:   # o link cadastrado é do mês atual? (para liberar o botão "Enviar boleto")
                    for m in membros:
                        if str(it.get("rotulo", "")).lower().endswith(m["n"].lower()):
                            pag["boleto_mes"][m["n"]] = vb[:7] == pag["venc"][:7]; break
                for m in membros:
                    if it.get("obs") and str(it.get("rotulo", "")).lower().endswith(m["n"].lower()):
                        pag["obs"][m["n"]] = it["obs"]; break
        except Exception:
            pass
    if pag.get("verif_antiga"):
        for ap, st in list(pag["status"].items()):
            if "pago" in str(st).lower():
                pag["status"][ap] = ""         # pagamento do mês anterior: aguarda a verificação deste mês
    INFO_LOGIN = INFO_LOGIN_ADM
    pag["ultimo_pag"] = ULTIMO_PAG_TS
    global FOTOS_LOGIN
    FOTOS_LOGIN = fotos_login_html()
    FOTOS_GAL = []      # galeria de fotos removida do projeto
    PUSH = {"pub": chave_push_publica(), "url": ACESSOS_URL} if ACESSOS_URL else None
    ACESSOS = None
    if ACESSOS_URL:
        _so = [EMAIL_APELIDO_INV.get(n.lower(), "") for n in ACESSOS_SOMENTE]
        _so = [e for e in _so if "@" in e]
        if _so:
            ACESSOS = {"cifrado": cifrar_para({"url": ACESSOS_URL, "chave": ACESSOS_CHAVE, "mapa": dict(EMAIL_APELIDO),
                                              "regs": (buscar_acessos() or [])[1:], "planilha": PLANILHA_URL}, _so)}
        else:
            print("Aviso: e-mail de " + ", ".join(ACESSOS_SOMENTE) + " não encontrado: a tela Acessos não vai aparecer.")
    FLY = {}      # "Gerar fly" removido: o fly oficial é feito no site do clube
    gravar(SAIDA_ADM, {"modo": "adm", "m": membros, "ev": eventos, "ev_ins": eventos_ins, "ev_staff": eventos_staff, "molde": molde, "pag": pag, "fotos": FOTOS_GAL, "fly": FLY, "acessos": ACESSOS, "afastados": AFASTADOS, "push": PUSH, "nuvem": NUVEM}, "Painel ADM - Acesso restrito ao Comando.",
           extras=[x for x in extras if x], rotulo="E-mail ou senha de acesso")
    dados_adm = {"modo": "adm", "m": membros, "ev": eventos, "ev_ins": eventos_ins, "ev_staff": eventos_staff, "molde": molde, "pag": pag, "fotos": FOTOS_GAL, "fly": FLY, "acessos": ACESSOS, "afastados": AFASTADOS, "push": PUSH, "nuvem": NUVEM}
    # no link: com nome completo e CPF dos integrantes (decisão do Eduardo, ciente do risco), para a Solicitação kit/colete.
    # Para voltar a tirar do link, troque LINK_COM_NOME_E_CPF para False no começo deste arquivo.
    membros_link = membros if LINK_COM_NOME_E_CPF else \
        [{**m, "x": {k: v for k, v in m.get("x", {}).items() if k in ("col", "tamc")}} for m in membros]
    pagina_adm_link = gravar(None, {**dados_adm, "m": membros_link}, "Painel ADM - Acesso restrito ao Comando.",
                             extras=[x for x in extras if x] if ADM_NO_LINK_ACEITA_SENHA_COLETIVA else [],
                             rotulo="E-mail ou senha de acesso" if ADM_NO_LINK_ACEITA_SENHA_COLETIVA else "E-mail")
    FOTOS_LOGIN = ""

    # 3) individuais: ranking e média da divisão por mês e no período todo
    periodos = {"todos": eventos}
    for e in eventos:
        periodos.setdefault(e["d"][:7], []).append(e)
    def ranking(evs, idxs):
        stats = []
        for i in idxs:
            pr = sum(1 for e in evs if e["p"][i] == 1)
            au = sum(1 for e in evs if e["p"][i] == 0)
            if pr + au:
                stats.append((i, pr, au, arred(pr * 100 / (pr + au))))
        stats.sort(key=lambda s: (-s[3], -s[1], membros[s[0]]["o"]))
        tot_p = sum(s[1] for s in stats); tot = sum(s[1] + s[2] for s in stats)
        return {"pos": {s[0]: k + 1 for k, s in enumerate(stats)}, "total": len(stats),
                "media": arred(tot_p * 100 / tot) if tot else None}

    grupos = {}
    for i, m in enumerate(membros):
        grupos.setdefault(grupo_de(m["c"]), []).append(i)
    todos_idx = list(range(len(membros)))
    rankings = {k: {"div": ranking(evs, todos_idx), "grp": {g: ranking(evs, ix) for g, ix in grupos.items()}}
                for k, evs in periodos.items()}

    if GERAR_INDIVIDUAIS:
        os.makedirs(PASTA_IND, exist_ok=True)
    proibidos = '<>:"/\\|?*'
    for i, m in (enumerate(sem_tel) if GERAR_INDIVIDUAIS else []):
        g = grupo_de(m["c"])
        rank = {k: {"pos": r["div"]["pos"].get(i), "total": r["div"]["total"], "media": r["div"]["media"],
                    "grupo": g, "gpos": r["grp"][g]["pos"].get(i), "gtotal": r["grp"][g]["total"],
                    "gmedia": r["grp"][g]["media"]} for k, r in rankings.items()}
        evs_i = [{"d": e["d"], "e": e["e"], "p": [e["p"][i]]} for e in eventos]
        nome_arq = "".join(ch for ch in m["n"] if ch not in proibidos)
        gravar(os.path.join(PASTA_IND, f"Extremo Norte - SP3 - {nome_arq}.html"),
               {"modo": "individual", "m": [m], "ev": evs_i, "rank": rank}, sem_js=estatico(m, evs_i, rank))

    print(f"OK: {len(membros)} integrantes e {len(eventos)} eventos.")
    sem = [m["n"] for m in membros if not m["tel"]]
    if sem:
        print("Sem telefone na aba Dados Integrantes: " + ", ".join(sem))
    if avisos:
        print("Atenção: não encontrei na aba Dados Integrantes (sem número/hierarquia): " + ", ".join(avisos))
    print(f"Painel geral:  {os.path.basename(SAIDA)}")
    print(f"Análises:      Presença {len(eventos)} evento(s) · Insights {len(eventos_ins)} · Staff {len(eventos_staff)}"
          + ("" if abas_extras else "  (abas Insights e Staff não encontradas na planilha)"))
    print(f"Painel ADM:    {os.path.basename(SAIDA_ADM)}")
    importar_acessos()
    garantir_tarefa_avisos()
    avisos_diarios(so_depois_das_9=True)     # se o computador estava desligado às 9h, o aviso do dia sai agora
    print(f"Individuais:   {len(membros)} arquivos na pasta Individuais (sem senha)" if GERAR_INDIVIDUAIS
          else "Individuais:   agora são imagens (no ADM, filtre o integrante e toque em 📸 Gerar imagem)")
    print("Acesso ao geral e ao ADM: e-mail de " + ", ".join(AUTORIZADOS) + " (ou a senha coletiva no geral)")
    if AUTORIZADOS_SO_GERAL:
        print("Só o painel geral, com o CPF: " + ", ".join(AUTORIZADOS_SO_GERAL))
    publicar_se_mudou({"index.html": INDEX_SO_ADM, "adm.html": pagina_adm_link, **arquivos_app()},
                      _assinatura_conteudo({"modo": "geral", "m": sem_tel, "ev": eventos}, dados_adm),
                      forcar="--publicar" in sys.argv)

if __name__ == "__main__":
    main()
