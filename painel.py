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
PASTA_IND = os.path.join(PASTA, "Individuais")

# ---------------------------------------------------------------------
# DIVISÃO: tudo o que é próprio de cada divisão fica no arquivo divisao.json (Kit Nova Divisão).
# Sem o arquivo, vale a Extremo Norte (os valores abaixo).
DIVISAO_PADRAO = {
    "nome": "Extremo Norte", "regional": "SP3", "estado": "SP", "comando_superior": "Comando VIII",
    "sede_local": "Sede Extremo Norte sp3 ( Quintos bar burg )",
    "sede_endereco": "Rua Silvio Brant correia, n° 313 - Vila Ema - SP.",
    "assinatura_local": "Extremo Norte SP3 - Vila Ema",
    "autorizados": ["Queijo", "Brawl", "Maycon", "Rick", "Edu", "Demonão"],
    "acessos_somente": ["Edu"],
}
def _carregar_divisao():
    d = dict(DIVISAO_PADRAO)
    arq = os.path.join(PASTA, "divisao.json")
    if os.path.exists(arq):
        try:
            with open(arq, encoding="utf-8") as f:
                d.update({k: v for k, v in json.load(f).items() if not str(k).startswith("_") and v not in (None, "")})
        except Exception as e:
            print(f"Aviso: não consegui ler o divisao.json ({e}). Usando os valores padrão.")
    return d
DIV = _carregar_divisao()
DIV_ROTULO = f"{DIV['nome']} - {DIV['regional']}"          # ex.: "Extremo Norte - SP3"
SAIDA     = os.path.join(PASTA, f"{DIV_ROTULO} - Analise.html")
SAIDA_ADM = os.path.join(PASTA, f"{DIV_ROTULO} - Analise (ADM).html")

def texto_divisao(html):
    """Troca os textos da Extremo Norte pelos da divisão do divisao.json (sem o arquivo, não muda nada)."""
    P, D = DIVISAO_PADRAO, DIV
    trocas = [(P["sede_local"], D["sede_local"]), (P["sede_endereco"], D["sede_endereco"]),
              (P["assinatura_local"], D["assinatura_local"]),
              ("${DIV_REG}-SP", "${DIV_REG}-" + D["estado"]),
              (P["comando_superior"], D["comando_superior"]), (P["comando_superior"].upper(), D["comando_superior"].upper()),
              (P["nome"], D["nome"]), (P["nome"].upper(), D["nome"].upper()), (P["nome"].lower(), D["nome"].lower()),
              (P["regional"], D["regional"])]
    for velho, novo in trocas:
        if velho != novo:
            html = html.replace(velho, novo)
    return html

# Quem pode abrir o painel geral e o ADM (apelido igual ao da aba Dados Integrantes).
# A senha de cada um é o CPF cadastrado na coluna "CPF" dessa aba (só os números).
AUTORIZADOS = list(DIV["autorizados"])     # do divisao.json

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
# Código por e-mail: com os dois segredos configurados, o painel passa a ser trancado com e-mail + chave secreta
# (só quem recebe o código no e-mail consegue a chave). Sem eles, funciona como antes (só o e-mail).
OTP_URL = os.environ.get("OTP_URL", "")
OTP_SEGREDO = os.environ.get("OTP_SEGREDO", "")

def otp_ativo():
    return bool(OTP_URL and OTP_SEGREDO)

def chave_de(senha):
    """O que tranca o painel de cada pessoa: com o código por e-mail ligado, e-mail + pedaço da chave secreta."""
    import hmac, hashlib
    s = str(senha).strip()
    if otp_ativo() and "@" in s:
        e = s.lower()
        return e + "|" + hmac.new(OTP_SEGREDO.encode(), e.encode(), hashlib.sha256).hexdigest()
    return s
# quem vê a tela "Acessos" no ADM (pelo e-mail da planilha). Os demais do comando não veem nem conseguem ler.
ACESSOS_SOMENTE = list(DIV["acessos_somente"])     # do divisao.json
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

MESES_NOME = ["JANEIRO", "FEVEREIRO", "MARÇO", "ABRIL", "MAIO", "JUNHO", "JULHO", "AGOSTO", "SETEMBRO", "OUTUBRO", "NOVEMBRO", "DEZEMBRO"]

def mensagem_pagamentos(venc, membros):
    """Molde da mensagem do status da mensalidade (WhatsApp). O painel troca a linha "01." pela lista dos irmãos
    e acerta o total e a %. A assinatura é de quem tem o cargo ADM (ou, se não houver, do Diretor)."""
    ordem = sorted(membros, key=lambda m: m.get("o", 999))
    assina = (next((m for m in ordem if re.search(r"\badm\b", str(m.get("c", "")), re.I)), None)
              or next((m for m in ordem if re.search(r"diretor", str(m.get("c", "")), re.I)), None))
    assinatura = f"*{assina['n']} - {assina['c']} Div {DIV['nome']}*" if assina else f"*ADM Div {DIV['nome']}*"
    return "\n".join([
        f"*🚨🚨🚨 DIVISÃO {DIV['nome'].upper()} - {DIV['regional'].upper()} 🚨🚨🚨*", "",
        f"*STATUS DA MENSALIDADE DE {MESES_NOME[venc.month - 1]} DE {venc.year} $$$ (Vencimento {venc:%d/%m/%Y})*", "",
        "*Srs. Bom Dia / Tarde / Noite*", "",
        "*Todos nós sabemos o quanto é importante cumprirmos com a nossa responsabilidade para com o clube. Portanto, vamos evitar deixar para o último dia.*", "",
        "*Legenda:*", "",
        "*❗- 1 MÊS DE ATRASO*",
        "*❗❗- 2 MESES DE ATRASO ( TEREMOS QUE RECOLHER CAMISETA OU COLETE E PRAZO DE 10 DIAS PARA QUITAR OU SERÁ DESLIGADO)*",
        "🏦 - PAGO",
        "⏰ - AGUARDANDO COMPENSAÇÃO", "",
        "01. (lista dos irmãos)", "",
        "*💀 Total de pagamentos efetuados: 0.💀 O que representa 0% dos boletos pagos.💀*", "",
        "*Conto com os senhores. Qualquer dúvida ou necessidade me chame no Privado!*", "",
        "*Somente ADM altera a listagem.*", "",
        assinatura])

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
<meta name="viewport" content="width=device-width, initial-scale=1, maximum-scale=1, user-scalable=no, viewport-fit=cover">
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
html{touch-action:manipulation}body{background:#000;color:var(--osso);font-family:"Barlow","Segoe UI",Roboto,Arial,sans-serif;display:grid;place-items:center;padding:20px}
/* a grade da entrada tem sempre a largura da tela: nenhum elemento mais largo consegue empurrar o quadro para o lado */
body{grid-template-columns:minmax(0,1fr);overflow-x:hidden}body>*{min-width:0;max-width:100%}html{overflow-x:hidden}
body::before{content:"";position:fixed;inset:0;z-index:-1;background:url("data:image/jpeg;base64,__IMG__") right center/cover no-repeat;opacity:.25}
form{width:min(290px,100%);background:rgba(236,231,220,.10);backdrop-filter:blur(14px);-webkit-backdrop-filter:blur(14px);border:1px solid rgba(236,231,220,.18);border-radius:12px;padding:20px}
h1{font-family:"Big Shoulders Display","Arial Narrow",Impact,sans-serif;font-weight:900;font-size:29px;line-height:.95;margin:0 0 5px}
p{color:var(--osso2);font-size:12px;margin:0 0 14px}
label{display:block;font-size:11.5px;color:var(--osso2);margin-bottom:5px}
input{width:100%;background:rgba(0,0,0,.45);color:var(--osso);border:1px solid rgba(236,231,220,.3);border-radius:8px;padding:9px 10px;font:inherit;font-size:16px;letter-spacing:.04em}
button{width:100%;min-height:40px;margin-top:11px;background:var(--osso);color:#111;border:0;border-radius:8px;padding:12px;font:inherit;font-weight:600;font-size:14px;cursor:pointer}
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
html.com-video .fundo-video{opacity:1}html.com-video body::before{opacity:0}
html.com-video body.entrou .fundo-video{opacity:.55}
body.abertura #f,body.abertura .manut{opacity:0;transform:translateY(8px);pointer-events:none}
#f,.manut{transition:opacity .9s ease,transform .9s ease}
/* fotos: durante a abertura vão surgindo uma a uma (sem mexer na inclinação de cada uma) */
.fotos-login .fl{transition:opacity .9s ease,scale .9s ease}
body.abertura:not(.fotos-on) .fotos-login .fl{opacity:0;scale:.85}
body.abertura.fotos-on .fotos-login .fl{opacity:.45}      /* durante o vídeo, bem visíveis; depois voltam a ficar clarinhas */
.fotos-login .fl:nth-child(2){transition-delay:.3s}.fotos-login .fl:nth-child(3){transition-delay:.6s}
.fotos-login .fl:nth-child(4){transition-delay:.9s}.fotos-login .fl:nth-child(5){transition-delay:1.2s}.fotos-login .fl:nth-child(6){transition-delay:1.5s}
body.entrou:not(.anim) .fotos-login .fl{transition:none}
body.entrou .fotos-login .fl{transition-delay:0s!important}     /* ao entrar, todas esmaecem juntas */
@media (prefers-reduced-motion:reduce){.fundo-video{display:none}}
</style></head><body class="abertura"><script>try{var _a=JSON.parse(localStorage.getItem("ajustes_app")||"{}");if(_a.eco)document.documentElement.classList.add("eco");if(_a.letra)document.documentElement.classList.add("letra-"+_a.letra)}catch(_){}</script><style>html.eco .fundo-video{display:none!important}html.letra-g body{zoom:1.12}html.letra-gg body{zoom:1.25}</style><video class="fundo-video" id="fv" src="fundo.mp4" autoplay muted loop playsinline preload="auto" aria-hidden="true" tabindex="-1"
  oncanplay="document.documentElement.classList.add('com-video')" onerror="this.remove();entrar()"></video>
<script>
// abertura do app: o vídeo aparece sozinho por ~2 s e depois a tela de entrada surge por cima
// (uma vez por visita; sem vídeo, internet fraca ou "menos movimento": aparece na hora)
function entrar(){document.body.classList.add("fotos-on");document.body.classList.remove("abertura");document.body.classList.add("entrou");try{sessionStorage.setItem("abertura_vista","1")}catch(_){}}
(function(){let vista=false;try{vista=sessionStorage.getItem("abertura_vista")==="1"}catch(_){}
  if(vista||matchMedia("(prefers-reduced-motion:reduce)").matches){entrar();return}
  const v=document.getElementById("fv");let ok=false;
  if(v)v.addEventListener("playing",()=>{if(ok)return;ok=true;document.body.classList.add("anim");
    setTimeout(()=>document.body.classList.add("fotos-on"),400);   // as fotos surgem uma a uma enquanto o logo se forma
    setTimeout(entrar,2600)},{once:true});
  setTimeout(()=>{if(!ok)entrar()},1500);       // o vídeo não começou a tempo: não deixa ninguém esperando
  setTimeout(entrar,6000)})();                   // garantia
</script>__FOTOS__
<noscript><style>#f{display:none}</style><form style="display:block"><h1>Extremo Norte<br>SP3</h1><p>Este visualizador não abre o painel. Toque em compartilhar e escolha abrir no Safari, Chrome ou Edge.</p></form></noscript>
<form id="f" autocomplete="off">
  __MANUT__
  <button id="sosAtivo" type="button" hidden style="width:100%;margin:0 0 12px;padding:12px;border-radius:14px;border:2px solid #ff5252;background:rgba(198,40,40,.85);color:#fff;font-weight:800;font-size:15px;text-align:left;line-height:1.35;animation:pisca 1.2s ease-in-out infinite"></button>
  <style>@keyframes pisca{50%{box-shadow:0 0 0 6px rgba(255,82,82,.35)}}</style>
  <button id="sosB" type="button" hidden style="width:100%;margin:0 0 14px;padding:14px;border-radius:14px;border:0;background:#c62828;color:#fff;font-weight:900;font-size:18px;letter-spacing:.5px;box-shadow:0 0 0 3px rgba(198,40,40,.25)">🆘 SOS · Preciso de ajuda</button>
  <div id="sosBox" hidden style="margin:0 0 14px;padding:14px;border-radius:14px;background:rgba(198,40,40,.12);border:1px solid rgba(198,40,40,.6);text-align:left">
    <p style="margin:0 0 10px;font-weight:800;font-size:16px">🆘 O que aconteceu?</p>
    <div style="display:grid;grid-template-columns:1fr 1fr;gap:8px">
      <button type="button" data-sos="pane" class="sosT">🛵 Pane/acidente</button><button type="button" data-sos="saude" class="sosT">🚑 Saúde</button>
      <button type="button" data-sos="seguranca" class="sosT">⚠️ Segurança</button><button type="button" data-sos="outro" class="sosT">🆘 Outro</button></div>
    <p style="margin:12px 0 6px;font-size:13px">⚠️ <b>Risco de vida?</b> Ligue primeiro:</p>
    <div style="display:flex;gap:6px"><a class="sosTel" href="tel:190">190 Polícia</a><a class="sosTel" href="tel:192">192 SAMU</a><a class="sosTel" href="tel:193">193 Bombeiros</a></div>
    <div id="sosConf" style="display:none;margin-top:12px;padding:12px;border-radius:12px;background:rgba(0,0,0,.35);border:1px solid rgba(198,40,40,.7)">
      <p id="sosConfTxt" style="margin:0 0 10px;font-size:15px;line-height:1.4"></p>
      <button type="button" id="sosConfSim" style="width:100%;padding:14px;border-radius:12px;border:0;background:#c62828;color:#fff;font-weight:900;font-size:16px">🆘 Sim, enviar o SOS</button>
      <button type="button" id="sosConfNao" class="sosS" style="width:100%;margin-top:8px">Cancelar</button></div>
    <p id="sosMsg" style="margin:12px 0 0;font-size:14px;line-height:1.4"></p>
    <div id="sosAcoes" style="display:none;flex-direction:column;gap:8px;margin-top:10px">
      <button type="button" id="sosZap" class="sosG">📤 Mandar nos grupos (WhatsApp)</button><button type="button" id="sosCopiar" class="sosS">📋 Copiar a mensagem</button></div>
    <a href="#" id="sosFechar" style="display:block;margin-top:12px;font-size:13px;color:#B8B2A6;text-align:center">Fechar</a></div>
  <div id="sosResp" hidden style="margin:0 0 14px;padding:14px;border-radius:14px;background:rgba(198,40,40,.12);border:1px solid rgba(198,40,40,.6);text-align:left">
    <p id="srTit" style="margin:0 0 4px;font-weight:900;font-size:17px"></p><p id="srSub" style="margin:0 0 10px;font-size:14px;color:#ECE7DC"></p>
    <p id="srDist" style="margin:0 0 12px;font-size:16px;font-weight:800;color:#E9B949">📍 Calculando a sua distância…</p>
    <div id="srMapaBox" style="display:none;margin:0 0 12px;border-radius:12px;overflow:hidden;border:1px solid rgba(236,231,220,.25)"><iframe id="srMapaF" title="Local do SOS" style="width:100%;height:280px;border:0;display:block" loading="lazy" referrerpolicy="no-referrer-when-downgrade"></iframe></div>
    <div style="display:flex;flex-direction:column;gap:8px">
      <button type="button" id="srIndo" class="sosG">🏍️ Estou indo (avisa e abre a rota)</button>
      <button type="button" id="srMapa" class="sosS">📍 Ver no mapa</button><button type="button" id="srNao" class="sosS">Não consigo ir</button></div>
    <p id="srMsg" style="margin:10px 0 0;font-size:13px;color:#B8B2A6"></p></div>
  <h1>Extremo Norte<br>SP3</h1>
  <p>__SUBTITULO__</p>
  <label for="cpf">__ROTULO__</label>
  <input id="cpf" type="text" inputmode="email" autocomplete="email" autocapitalize="none" spellcheck="false" placeholder="seu@email.com" maxlength="120" autofocus>
  <button id="b" type="submit">Entrar</button>
  <button id="bioB" type="button" hidden style="margin-top:10px;width:100%;padding:13px;border-radius:12px;border:1px solid #E9B949;background:rgba(233,185,73,.12);color:#ECE7DC;font-weight:800;font-size:15px">👆 Entrar com a digital</button>
  <style>.sosT{padding:12px 6px;border-radius:10px;border:1px solid rgba(236,231,220,.25);background:rgba(0,0,0,.35);color:#ECE7DC;font-weight:800;font-size:14px}
  .sosTel{flex:1;text-align:center;padding:9px 4px;border-radius:10px;background:#ECE7DC;color:#141210;font-weight:800;font-size:12.5px;text-decoration:none}
  .sosG{padding:13px;border-radius:12px;border:0;background:#1f8f4e;color:#fff;font-weight:800;font-size:15px}.sosS{padding:12px;border-radius:12px;border:1px solid rgba(236,231,220,.3);background:transparent;color:#ECE7DC;font-weight:700}</style>
  <div id="otpBox" hidden style="margin-top:12px;padding:12px;border-radius:12px;background:rgba(236,231,220,.06);border:1px solid rgba(236,231,220,.18);text-align:left">
    <p id="otpMsg" style="margin:0 0 10px;font-size:14px;line-height:1.4"></p>
    <input id="otpC" type="text" inputmode="numeric" autocomplete="one-time-code" maxlength="7" placeholder="000 000" style="width:100%;text-align:center;font-size:24px;letter-spacing:6px;padding:10px;border-radius:10px">
    <button id="otpOk" type="button" style="margin-top:10px;width:100%;padding:12px;border-radius:10px;border:0;background:#E9B949;color:#141210;font-weight:800">Confirmar código</button>
    <div style="display:flex;justify-content:space-between;margin-top:10px;font-size:13px"><a href="#" id="otpRe" style="color:#E9B949">Reenviar código</a><a href="#" id="otpVolta" style="color:#B8B2A6">Usar outro e-mail</a></div></div>
  <a href="#" id="emailLink" style="display:none;margin-top:10px;font-size:14px;color:#E9B949;text-align:center">Entrar com o e-mail</a>
  <div id="bioOf" hidden style="margin-top:12px;padding:12px;border-radius:12px;background:rgba(233,185,73,.10);border:1px solid rgba(233,185,73,.45);text-align:left">
    <p style="margin:0 0 10px;font-size:14px;line-height:1.4">👆 Quer entrar com a <b>digital</b> (ou <b>Face ID</b>) nas próximas vezes, neste celular?</p>
    <div style="display:flex;gap:8px"><button id="bioSim" type="button" style="flex:1;padding:11px;border-radius:10px;border:0;background:#E9B949;color:#141210;font-weight:800">Ativar</button>
    <button id="bioNao" type="button" style="flex:1;padding:11px;border-radius:10px;border:1px solid rgba(236,231,220,.3);background:transparent;color:#ECE7DC;font-weight:700">Agora não</button></div></div>
  <div id="erro" role="alert"></div>
  <style>#carrega{position:fixed;inset:0;z-index:99999;background:rgba(13,11,9,.94);display:flex;flex-direction:column;align-items:center;justify-content:center;gap:14px;color:#BDB7AA;font-size:15px}
  #carrega .cl-cav{position:relative;width:150px}#carrega .cl-cav img{display:block;width:100%;height:auto;filter:grayscale(1) brightness(.45)}
  #carrega .cl-cav img.ouro{position:absolute;inset:0;filter:sepia(1) saturate(3.2) hue-rotate(2deg) brightness(1.05);animation:clEnche 1.6s cubic-bezier(.45,.05,.3,1) infinite}
  @keyframes clEnche{0%{clip-path:inset(100% 0 0 0)}70%{clip-path:inset(0 0 0 0)}100%{clip-path:inset(0 0 0 0);opacity:.0}}
  @media (prefers-reduced-motion:reduce){#carrega .cl-cav img.ouro{animation:none;clip-path:inset(40% 0 0 0)}}</style>
  <a id="solic" class="solic" hidden target="_blank" rel="noopener">📩 Solicitar acesso</a>
  <button type="button" class="tut-abrir" id="tutLink" onclick="tutAbrir()">📲 Como instalar o app</button>
  <style>.avs{display:block;text-align:left;margin-top:6px}.av{border-top:1px solid rgba(236,231,220,.14);padding:8px 2px}.av summary{cursor:pointer;list-style:none;font-size:13px;color:#d9d2c4;white-space:nowrap;overflow:hidden;text-overflow:ellipsis}
  .av summary::-webkit-details-marker{display:none}.av summary::after{content:" ▾";color:#9c958a}.av[open] summary::after{content:" ▴"}.av summary b{color:#ECE7DC;font-weight:600}
  .av ul{margin:6px 0 2px;padding-left:18px;font-size:12.5px;color:#cfc8bb;line-height:1.5}.atual-dt{display:block;text-align:right;font-size:11px;color:#9c958a;margin-top:8px}
  .registro:empty{display:none}</style>
  <p class="registro">__REGISTRO__</p>
  <div class="atual">__ATUAL__</div>
  <script>try{var np=document.getElementById("novoPag");if(np&&np.dataset.ts&&np.dataset.ts>(localStorage.getItem("pag_visto")||"")){np.hidden=false;document.getElementById("novoPagTxt").hidden=false}}catch(_){}</script>
</form>
<script>
const P=__PACOTE__;
const b64=s=>Uint8Array.from(atob(s),c=>c.charCodeAt(0));
const cpf=document.getElementById("cpf"),erro=document.getElementById("erro"),btn=document.getElementById("b");
const solic=document.getElementById("solic");
// ---- entrar com a digital / Face ID: o e-mail fica guardado dentro da chave de acesso do celular (passkey),
//      e o celular só entrega depois que a digital confirma que é o dono. Nada sai do celular.
const bioB=document.getElementById("bioB"),bioOf=document.getElementById("bioOf");
const bioRnd=()=>crypto.getRandomValues(new Uint8Array(32));
let bioPode=false;
(async()=>{try{bioPode=!!(window.PublicKeyCredential&&await PublicKeyCredential.isUserVerifyingPlatformAuthenticatorAvailable());   // celular: digital/Face ID · computador: Windows Hello
  if(bioPode&&localStorage.getItem("bio_on")){bioB.hidden=false;if(/iPhone|iPad/.test(navigator.userAgent))bioB.textContent="🙂 Entrar com Face ID";
    if(!location.hash.startsWith("#sos="))soDigital()}}catch(_){}})();   // a digital abre só com o toque (o SOS fica sempre à mão)
// com a digital cadastrada neste celular: a entrada mostra só a digital; o e-mail fica a um toque (ou aparece se a digital falhar)
function soDigital(){["cpf","b"].forEach(id=>{const x=document.getElementById(id);if(x)x.hidden=true});
  document.querySelectorAll('label[for="cpf"]').forEach(x=>x.style.display="none");const l=document.getElementById("emailLink");if(l)l.style.display="block"}
function mostraEmail(){["cpf","b"].forEach(id=>{const x=document.getElementById(id);if(x)x.hidden=false});
  document.querySelectorAll('label[for="cpf"]').forEach(x=>x.style.display="");const l=document.getElementById("emailLink");if(l)l.style.display="none";
  const c=document.getElementById("cpf");if(c)setTimeout(()=>c.focus(),50)}
document.getElementById("emailLink").onclick=ev=>{ev.preventDefault();mostraEmail()};
(function(){let re="";try{re=sessionStorage.getItem("reentrar")||"";sessionStorage.removeItem("reentrar")}catch(_){}
  if(re&&re.includes("@")){cpf.value=re;setTimeout(()=>{try{document.getElementById("f").requestSubmit()}catch(_){}},250)}})();
const bioB64=x=>btoa(String.fromCharCode(...new Uint8Array(x))).replace(/\+/g,"-").replace(/\//g,"_").replace(/=+$/,"");
const bioDe64=x=>Uint8Array.from(atob(x.replace(/-/g,"+").replace(/_/g,"/")+"=".repeat((4-x.length%4)%4)),c=>c.charCodeAt(0));
async function bioEntrar(auto){erro.textContent="";try{
  let ids=[];try{ids=JSON.parse(localStorage.getItem("bio_ids")||"[]")}catch(_){}
  const pk={challenge:bioRnd(),userVerification:"required",timeout:60000,rpId:location.hostname};
  if(ids.length)pk.allowCredentials=ids.map(i=>({type:"public-key",id:bioDe64(i),transports:["internal"]}));   // só a digital deste celular (não procura outros aparelhos)   // a chave deste celular: vai direto para a digital
  let a=null;
  try{a=await navigator.credentials.get({publicKey:pk})}
  catch(e1){if(!ids.length||!/Android|iPhone|iPad|iPod/i.test(navigator.userAgent))throw e1;   // no computador não procura outros aparelhos (evita a janela do Chrome)
    delete pk.allowCredentials;pk.challenge=bioRnd();a=await navigator.credentials.get({publicKey:pk})}   // celular: a chave anotada sumiu, tenta qualquer chave deste site
  try{const id=bioB64(a.rawId);if(!ids.includes(id))localStorage.setItem("bio_ids",JSON.stringify([...ids,id].slice(-3)))}catch(_){}
  const em=new TextDecoder().decode(a.response.userHandle||new Uint8Array());
  if(!em.includes("@"))throw new Error("sem e-mail");
  cpf.value=em;window.__viaBio=true;document.getElementById("f").requestSubmit()}
  catch(e){if(auto)return;   // ao abrir o app (sem toque): se o celular não deixou ou a pessoa cancelou, o botão continua lá
    try{sessionStorage.setItem("bio_falhou","1")}catch(_){}   // se a chave sumiu do celular, depois do e-mail o app oferece cadastrar de novo
    mostraEmail();erro.textContent="Não deu certo com a digital. Tente de novo ou entre com o seu e-mail."}}
bioB.onclick=()=>bioEntrar(false);
// depois de entrar com o e-mail: oferece ativar a digital (uma vez por e-mail neste celular)
function bioPerguntar(d,seguir){
  let ja="";try{ja=(localStorage.getItem("bio_on")||"")+"|"+(localStorage.getItem("bio_nao")||"")}catch(_){}
  let falhou="";try{falhou=sessionStorage.getItem("bio_falhou")||""}catch(_){}
  if(!bioPode||!d.includes("@")||new TextEncoder().encode(d).length>64||(!falhou&&ja.split("|").includes(d))){seguir();return}
  ["cpf","b","bioB"].forEach(id=>{const x=document.getElementById(id);if(x)x.hidden=true});
  document.querySelectorAll('label[for="cpf"]').forEach(x=>x.style.display="none");bioOf.hidden=false;
  document.getElementById("bioNao").onclick=()=>{try{localStorage.setItem("bio_nao",d)}catch(_){}seguir()};
  document.getElementById("bioSim").onclick=async()=>{try{
    const nova=await navigator.credentials.create({publicKey:{challenge:bioRnd(),rp:{name:document.title.split(" - ")[0]||"Insanos",id:location.hostname},
      user:{id:new TextEncoder().encode(d),name:d,displayName:d},pubKeyCredParams:[{type:"public-key",alg:-7},{type:"public-key",alg:-257}],
      authenticatorSelection:{authenticatorAttachment:"platform",residentKey:"required",requireResidentKey:true,userVerification:"required"},timeout:60000}});
    try{localStorage.setItem("bio_on",d);localStorage.setItem("bio_ids",JSON.stringify([bioB64(nova.rawId)]));sessionStorage.removeItem("bio_falhou")}catch(_){}alert("👆 Pronto! Na próxima vez, é só tocar em \"Entrar com a digital\".")}
    catch(e){alert("Não consegui ativar a digital neste celular. Você continua entrando com o e-mail.")}seguir()}}
// ---- código por e-mail: a chave do painel = e-mail + um pedaço que só vem depois do código certo.
//      O celular lembra esse pedaço por 90 dias (com a digital, a pessoa só digita o código uma vez).
const otpCache=()=>{try{return JSON.parse(localStorage.getItem("otp_s")||"{}")}catch(_){return {}}};
const otpHoje=()=>{const x=new Date();return x.getFullYear()+"-"+(x.getMonth()+1)+"-"+x.getDate()};
function otpGuardar(d,s){const c=otpCache();c[d]={s,exp:Date.now()+90*864e5,dia:otpHoje()};try{localStorage.setItem("otp_s",JSON.stringify(c))}catch(_){}}
function otpEsquecer(d){const c=otpCache();delete c[d];try{localStorage.setItem("otp_s",JSON.stringify(c))}catch(_){}}
async function otpPost(o){const r=await fetch(P.otp,{method:"POST",headers:{"Content-Type":"text/plain;charset=utf-8"},body:JSON.stringify(o)});return r.json()}
function otpPedaco(d,viaDigital){const c=otpCache()[d];
  if(c&&c.exp>Date.now()&&(viaDigital||c.dia===otpHoje()))return Promise.resolve(c.s);   // e-mail: código 1x por dia · digital: 90 dias
  return new Promise(resolve=>{const box=document.getElementById("otpBox"),msg=document.getElementById("otpMsg"),inp=document.getElementById("otpC");
    const fim=s=>{box.hidden=true;["cpf","b"].forEach(id=>document.getElementById(id).hidden=false);document.querySelectorAll('label[for="cpf"]').forEach(x=>x.style.display="");resolve(s)};
    const pedir=async()=>{msg.textContent="Enviando o código…";try{const j=await otpPost({tipo:"otp_pedir",email:d});
        msg.innerHTML=j.ok?`📧 Enviamos um código para <b>${d.replace(/</g,"&lt;")}</b>. Digite abaixo. Não chegou? Confira o <b>spam</b>.`:"⚠️ "+(j.erro||"Não consegui enviar o código.");
        if(!j.ok&&/cadastrado/.test(j.erro||""))setTimeout(()=>fim(null),2500)}
      catch(e){msg.textContent="⚠️ Não consegui enviar o código agora. Confira a internet e toque em Reenviar."}};
    ["cpf","b","bioB"].forEach(id=>{const x=document.getElementById(id);if(x)x.hidden=true});document.querySelectorAll('label[for="cpf"]').forEach(x=>x.style.display="none");
    box.hidden=false;inp.value="";setTimeout(()=>inp.focus(),100);pedir();
    document.getElementById("otpRe").onclick=ev=>{ev.preventDefault();pedir()};
    document.getElementById("otpVolta").onclick=ev=>{ev.preventDefault();fim(null)};
    const conferir=async()=>{const cod=inp.value.replace(/\D/g,"");if(cod.length!==6){msg.textContent="Digite os 6 números do código.";return}
      msg.textContent="Conferindo…";try{const j=await otpPost({tipo:"otp_conferir",email:d,codigo:cod});
        if(j.ok&&j.s){otpGuardar(d,j.s);fim(j.s)}else msg.textContent="⚠️ "+(j.erro||"Código incorreto.")}catch(e){msg.textContent="⚠️ Não consegui conferir agora. Tente de novo."}};
    document.getElementById("otpOk").onclick=conferir;inp.oninput=()=>{if(inp.value.replace(/\D/g,"").length===6)conferir()}})}
// ---- SOS: só para quem já entrou neste celular (o celular guarda o pedaço da chave dele, que confirma quem é)
const sosIdbSet=(k,v)=>new Promise(r=>{try{const o=indexedDB.open("insanos",1);o.onupgradeneeded=()=>o.result.createObjectStore("kv");o.onsuccess=()=>{const t=o.result.transaction("kv","readwrite");t.objectStore("kv").put(v,k);t.oncomplete=()=>r(1);t.onerror=()=>r(0)};o.onerror=()=>r(0)}catch(e){r(0)}});
const sosKm=(a,b,c,d)=>{const R=6371,x=(c-a)*Math.PI/180,y=(d-b)*Math.PI/180,h=Math.sin(x/2)**2+Math.cos(a*Math.PI/180)*Math.cos(c*Math.PI/180)*Math.sin(y/2)**2;return 2*R*Math.asin(Math.sqrt(h))};
const sosGps=(alta)=>new Promise(r=>{if(!navigator.geolocation)return r(null);navigator.geolocation.getCurrentPosition(p=>{const l={lat:+p.coords.latitude.toFixed(5),lng:+p.coords.longitude.toFixed(5),ts:Date.now()};sosIdbSet("loc",l);r(l)},()=>r(null),{enableHighAccuracy:!!alta,timeout:alta?12000:8000,maximumAge:alta?0:600000})});
// guarda a última localização (só neste celular) para a notificação do SOS mostrar a distância
try{navigator.permissions&&navigator.permissions.query({name:"geolocation"}).then(x=>{if(x.state==="granted")sosGps(false)})}catch(_){}
const sosIdbGet=k=>new Promise(r=>{try{const o=indexedDB.open("insanos",1);o.onupgradeneeded=()=>o.result.createObjectStore("kv");o.onsuccess=()=>{try{const q=o.result.transaction("kv").objectStore("kv").get(k);q.onsuccess=()=>r(q.result||null);q.onerror=()=>r(null)}catch(e){r(null)}};o.onerror=()=>r(null)}catch(e){r(null)}});
// SOS em andamento (recebido nas últimas horas e ainda sem resposta deste celular): faixa no topo da entrada
(async()=>{if(location.hash.startsWith("#sos="))return;const a=await sosIdbGet("sos_ativo");if(!a||!a.d)return;
  const min=Math.round((Date.now()-a.ts)/60000);if(min>60||a.resp)return;
  const b=document.getElementById("sosAtivo");b.hidden=false;
  b.innerHTML=`🆘 <b>SOS em andamento</b> · ${String(a.d.n||"").replace(/</g,"&lt;")} precisa de ajuda (há ${min<1?"menos de 1":min} min)<br><span style="font-weight:600;font-size:13px">Toque para ver a distância, o mapa e responder</span>`;
  b.onclick=()=>{location.replace(a.url);location.reload()}})();
const sosQuem=()=>{const c=otpCache();let best=null;for(const e in c)if(c[e].exp>Date.now()&&(!best||c[e].exp>c[best].exp))best=e;return best?{email:best,s:c[best].s}:null};
const sosB=document.getElementById("sosB"),sosBox=document.getElementById("sosBox");
if(P.otp&&sosQuem())sosB.hidden=false;
if(location.hash==="#sos-abrir"){history.replaceState(null,"",location.pathname);   // veio pelo atalho do ícone do app
  if(P.otp&&sosQuem()){sosBox.hidden=false;sosB.hidden=true}else erro.textContent="Para usar o SOS, entre uma vez com o seu e-mail neste celular."}
let sosTexto="";
sosB.onclick=()=>{sosBox.hidden=false;sosB.hidden=true;document.getElementById("sosMsg").textContent="";document.getElementById("sosAcoes").style.display="none"};
document.getElementById("sosFechar").onclick=ev=>{ev.preventDefault();sosBox.hidden=true;sosB.hidden=!(P.otp&&sosQuem())};
let sosPend=null;
document.querySelectorAll(".sosT").forEach(b=>b.onclick=()=>{if(!sosQuem())return;
  sosPend={cat:b.dataset.sos,nome:b.textContent.trim(),gps:sosGps(true)};   // a localização já começa a ser pega enquanto confirma
  document.getElementById("sosConfTxt").innerHTML=`Confirmar o SOS de <b>${sosPend.nome}</b>?<br><span style="font-size:13px;color:#B8B2A6">Todos os irmãos com o app serão avisados, com a sua localização.</span>`;
  document.getElementById("sosConf").style.display="block";document.getElementById("sosMsg").textContent="";document.getElementById("sosAcoes").style.display="none"});
document.getElementById("sosConfNao").onclick=()=>{sosPend=null;document.getElementById("sosConf").style.display="none"};
document.getElementById("sosConfSim").onclick=async()=>{const q=sosQuem();if(!q||!sosPend)return;const {cat,nome}=sosPend;
  document.getElementById("sosConf").style.display="none";
  const m=document.getElementById("sosMsg");m.textContent="📍 Pegando a sua localização…";document.querySelectorAll(".sosT").forEach(x=>x.disabled=true);
  const l=await sosPend.gps;sosPend=null;const quando=new Date().toLocaleString("pt-BR",{day:"2-digit",month:"2-digit",hour:"2-digit",minute:"2-digit"});
  const monta=quem=>`🆘 *SOS · Insanos M.C.*\n${quem?`*${quem}*`:"Um irmão"} (Divisão ${document.title.split(" - ")[0]}) precisa de ajuda.\n${nome}\n`+(l?`📍 https://maps.google.com/?q=${l.lat},${l.lng}\n`:"📍 (sem localização)\n")+`🕒 ${quando}\nQuem estiver perto, por favor, dê um apoio! 🤝`+(window.fichaZap?fichaZap(q.email):"");
  sosTexto=monta(localStorage.getItem("sos_nome")||"");
  try{await navigator.clipboard.writeText(sosTexto)}catch(_){}
  m.textContent="Enviando o alerta para os irmãos…";
  try{const j=await otpPost({tipo:"sos",email:q.email,s:q.s,categoria:cat,lat:l?l.lat:null,lng:l?l.lng:null,ficha:(window.fichaSOS?fichaSOS(q.email):"")});
    if(j.ok&&j.nome){try{localStorage.setItem("sos_nome",j.nome)}catch(_){}sosTexto=monta(j.nome);try{await navigator.clipboard.writeText(sosTexto)}catch(_){}}
    m.innerHTML=j.ok?"✅ <b>Alerta enviado!</b> Os irmãos recebem em 1 a 3 minutos. Agora mande também nos grupos:":"⚠️ "+(j.erro||"Não consegui enviar o alerta.")+" Mande nos grupos:"}
  catch(e){m.innerHTML="⚠️ Não consegui enviar o alerta (sem internet?). <b>Mande nos grupos:</b>"}
  document.getElementById("sosAcoes").style.display="flex";document.querySelectorAll(".sosT").forEach(x=>x.disabled=false)};
document.getElementById("sosZap").onclick=async()=>{if(navigator.share){try{await navigator.share({text:sosTexto});return}catch(e){if(e&&e.name==="AbortError")return}}window.open("https://wa.me/?text="+encodeURIComponent(sosTexto),"_blank")};
document.getElementById("sosCopiar").onclick=async()=>{try{await navigator.clipboard.writeText(sosTexto);document.getElementById("sosCopiar").textContent="✅ Copiada: cole no WhatsApp"}catch(_){alert(sosTexto)}};
// ficha de emergência (guardada só neste celular): vai junto com o SOS, se o irmão deixou marcado
if(!window.fichaSOS)window.fichaSOS=function(em){try{const o=JSON.parse(localStorage.getItem("moto_"+em)||"{}").emg;if(!o||o.enviar===false)return "";
  const f={};["sangue","alerg","rem","obs","plano"].forEach(k=>{if(o[k])f[k]=String(o[k]).slice(0,120)});const c=(o.enviarCt===false?[]:(o.c||[])).filter(x=>x&&x.t).slice(0,2).map(x=>({n:String(x.n||"").slice(0,40),p:String(x.p||"").slice(0,20),t:String(x.t).slice(0,20)}));if(c.length)f.c=c;
  return Object.keys(f).length?JSON.stringify(f):""}catch(_){return ""}};
window.fichaZap=function(em){try{const j=(window.fichaSOS?fichaSOS(em):"");if(!j)return "";const f=JSON.parse(j),N=String.fromCharCode(10),L=[];
    if(f.sangue)L.push("🩸 Tipo sanguíneo: "+f.sangue);if(f.alerg)L.push("⚠️ Alergia: "+f.alerg);if(f.rem)L.push("💊 Remédios: "+f.rem);if(f.obs)L.push("ℹ️ "+f.obs);if(f.plano)L.push("🏥 Plano: "+f.plano);
    (f.c||[]).forEach(c=>L.push("📞 Contato: "+(c.n||"")+(c.p?" ("+c.p+")":"")+" · "+c.t));return L.length?N+N+"🩺 *Ficha de emergência*"+N+L.join(N):""}catch(_){return ""}};
// ---- resposta ao SOS (aberta pela notificação): distância exata, "Estou indo" com aviso, grupos e rota
(function(){const h=location.hash;if(!h.startsWith("#sos="))return;let S=null;
  try{const b=h.slice(5).split("&")[0].replace(/-/g,"+").replace(/_/g,"/");S=JSON.parse(new TextDecoder().decode(Uint8Array.from(atob(b+"=".repeat((4-b.length%4)%4)),c=>c.charCodeAt(0))))}catch(e){return}
  const box=document.getElementById("sosResp");box.hidden=false;sosB.hidden=true;
  ["cpf","b","bioB"].forEach(id=>{const x=document.getElementById(id);if(x)x.hidden=true});document.querySelectorAll('label[for="cpf"]').forEach(x=>x.style.display="none");
  document.getElementById("srTit").textContent=`🆘 SOS · ${S.n} precisa de ajuda`;document.getElementById("srSub").textContent=`${S.e||"🆘"} ${S.t} · ${S.q}`;
  if(S.f&&typeof S.f==="object"){const F=S.f,e2=t=>String(t||"").replace(/[&<>"]/g,c=>({"&":"&amp;","<":"&lt;",">":"&gt;",'"':"&quot;"}[c])),L=(r,t,al)=>t?`<div style="margin:6px 0;padding:8px 10px;border-radius:10px;background:${al?"rgba(179,38,30,.3);border:1px solid #E5484D":"rgba(255,255,255,.06)"}"><small style="display:block;opacity:.75;font-weight:700">${r}</small><b style="font-size:17px">${e2(t)}</b></div>`:"";
    document.getElementById("srSub").insertAdjacentHTML("afterend",`<div id="srFicha" style="margin:12px 0;text-align:left"><div style="font-weight:900;color:#ffb4ae;margin-bottom:4px">🩺 Ficha de emergência de ${e2(S.n)}</div>
      ${L("🩸 Tipo sanguíneo",F.sangue)}${L("⚠️ Alergias",F.alerg,1)}${L("💊 Remédios de uso contínuo",F.rem)}${L("ℹ️ Outras informações",F.obs)}${L("🏥 Plano de saúde",F.plano)}
      ${(F.c||[]).map(c=>`<a href="tel:${e2(String(c.t).replace(/[^0-9+]/g,""))}" style="display:block;margin:8px 0;padding:12px;border-radius:12px;background:#1E8E3E;color:#fff;text-align:center;font-weight:800;text-decoration:none">📞 Ligar para ${e2(c.n||"o contato")}${c.p?" ("+e2(c.p)+")":""} · ${e2(c.t)}</a>`).join("")}</div>`)}
  const tem=S.lat!==""&&S.lat!=null,rota=tem?`https://www.google.com/maps/dir/?api=1&destination=${S.lat}%2C${S.lng}&travelmode=driving`:"";
  // "Ver no mapa": o mapa aparece aqui mesmo (sem abrir o site ou o app do Google Maps, que às vezes se perde)
  const mostraMapa=()=>{if(!tem)return;const f=document.getElementById("srMapaF");if(!f.src)f.src=`https://maps.google.com/maps?q=${S.lat}%2C${S.lng}&z=16&hl=pt-BR&output=embed`;
    document.getElementById("srMapaBox").style.display="block";document.getElementById("srMapa").textContent="📍 Mapa aberto acima"};
  if(/&mapa=1/.test(h))mostraMapa();
  let km=null;const di=document.getElementById("srDist");
  if(!tem)di.textContent="📍 Ele não conseguiu mandar a localização. Combine pelo grupo.";
  else sosGps(true).then(l=>{if(!l){di.textContent="📍 Não consegui a sua localização. Toque em Ver no mapa.";return}
    km=sosKm(l.lat,l.lng,+S.lat,+S.lng);const min=Math.max(1,Math.round(km*2.2));
    di.textContent=`📍 Você está a ${km<1?Math.round(km*1000)+" m":km.toFixed(1).replace(".",",")+" km"} do ${S.n} · ~${min} min de moto`});
  document.getElementById("srMapa").onclick=mostraMapa;
  const respondeu=()=>sosIdbGet("sos_ativo").then(a=>{if(a&&a.d&&a.d.id===S.id){a.resp=true;sosIdbSet("sos_ativo",a)}});
  document.getElementById("srNao").onclick=()=>{respondeu();box.hidden=true;history.replaceState(null,"",location.pathname);["cpf","b"].forEach(id=>document.getElementById(id).hidden=false);document.querySelectorAll('label[for="cpf"]').forEach(x=>x.style.display="");if(P.otp&&sosQuem())sosB.hidden=false};
  document.getElementById("srIndo").onclick=async()=>{respondeu();const q=sosQuem(),dist=km==null?"":` · estou a ${km.toFixed(1).replace(".",",")} km (uns ${Math.max(1,Math.round(km*2.2))} min)`;
    const tx=`🏍️ Estou indo apoiar o *${S.n}* no SOS${dist}.`;
    try{await navigator.clipboard.writeText(tx)}catch(_){}
    if(q)otpPost({tipo:"sos_indo",email:q.email,s:q.s,id:S.id,dist:km==null?null:+km.toFixed(1)}).catch(()=>{});
    else document.getElementById("srMsg").textContent="Entre no app uma vez com o seu e-mail para o seu aviso chegar ao irmão.";
    if(navigator.share){try{await navigator.share({text:tx})}catch(_){}}
    if(rota)location.href=rota};
})();
// a caveira enchendo de dourado enquanto o painel abre (some se aparecer erro ou o pedido do código)
function carrega(mostra){let c=document.getElementById("carrega");
  if(!mostra){if(c)c.remove();return}
  if(c||!P.cav)return;c=document.createElement("div");c.id="carrega";
  c.innerHTML=`<div class="cl-cav"><img src="${P.cav}" alt=""><img class="ouro" src="${P.cav}" alt=""></div><p>Abrindo o painel…</p>`;document.body.appendChild(c)}
new MutationObserver(()=>{if(erro.textContent.trim())carrega(false)}).observe(erro,{childList:true,characterData:true,subtree:true});
(function(){const ob=document.getElementById("otpBox");if(ob)new MutationObserver(()=>{if(!ob.hidden)carrega(false)}).observe(ob,{attributes:true})})();
document.getElementById("f").addEventListener("submit",async ev=>{
  ev.preventDefault();erro.textContent="";carrega(true);
  solic.hidden=true;
  const d=cpf.value.trim().toLowerCase();
  const registrar=ok=>{try{const u="__ACESSOS_URL__";if(!u)return;
    const ua=navigator.userAgent,ap=(/iPhone|iPad/.test(ua)?"iPhone":/Android/.test(ua)?"Celular Android":/Windows/.test(ua)?"Computador Windows":/Mac/.test(ua)?"Mac":"Outro")
      +((matchMedia("(display-mode: standalone)").matches||navigator.standalone)?" · app":" · navegador");
    const corpo=JSON.stringify({email:d.includes("@")?d:"(senha coletiva)",ok,painel:/ADM/.test(document.title)?"ADM":"Geral",aparelho:ap});
    // fetch com keepalive: continua mesmo quando a página troca para o painel (o iPhone descarta o sendBeacon nessa hora)
    try{fetch(u,{method:"POST",mode:"no-cors",keepalive:true,headers:{"Content-Type":"text/plain;charset=utf-8"},body:corpo}).catch(()=>{try{navigator.sendBeacon&&navigator.sendBeacon(u,new Blob([corpo],{type:"text/plain"}))}catch(_){}})}
    catch(_){if(navigator.sendBeacon)navigator.sendBeacon(u,new Blob([corpo],{type:"text/plain"}))}}catch(_){}};
  if(!d){erro.textContent="Digite o seu e-mail.";return}
  if(!window.crypto||!crypto.subtle){erro.textContent="Este visualizador não permite abrir o painel. Abra o arquivo no Chrome, Edge ou Safari.";return}
  let entrada=d;
  const viaDigital=!!window.__viaBio;window.__viaBio=false;
  if(P.otp&&d.includes("@")){btn.disabled=true;btn.textContent="Aguardando o código…";const s=await otpPedaco(d,viaDigital);btn.disabled=false;btn.textContent="Entrar";
    if(!s)return;entrada=d+"|"+s}
  btn.disabled=true;btn.textContent="Verificando...";
  try{
    const base=await crypto.subtle.importKey("raw",new TextEncoder().encode(entrada),"PBKDF2",false,["deriveKey"]);
    const kek=await crypto.subtle.deriveKey({name:"PBKDF2",salt:b64(P.sal),iterations:P.it,hash:"SHA-256"},base,{name:"AES-GCM",length:256},false,["decrypt"]);
    let chave=null;
    const comoIrmao=(()=>{try{return sessionStorage.getItem("ver_irmao")==="1"}catch(_){return false}})();   // "👁️ Ver como irmão" (Gestão)
    const tentaSlots=async()=>{for(const s of P.slots){try{chave=await crypto.subtle.decrypt({name:"AES-GCM",iv:b64(s.iv)},kek,b64(s.ct));break}catch(e){}}};
    if(!comoIrmao)await tentaSlots();
    if(!chave){   // não é do comando: tenta o "Meu painel" do irmão
      for(const s of (P.pes||[])){try{const h=new TextDecoder().decode(await crypto.subtle.decrypt({name:"AES-GCM",iv:b64(s.iv)},kek,b64(s.ct)));
        try{const u="__ACESSOS_URL__";if(u){const c=JSON.stringify({email:d,ok:true,painel:s.t==="v"?"Visitante":"Meu painel",aparelho:(/Android/.test(navigator.userAgent)?"Celular Android":/iPhone|iPad/.test(navigator.userAgent)?"iPhone":"Outro")+((matchMedia("(display-mode: standalone)").matches||navigator.standalone)?" · app":" · navegador")});
          try{fetch(u,{method:"POST",mode:"no-cors",keepalive:true,headers:{"Content-Type":"text/plain;charset=utf-8"},body:c}).catch(()=>{try{navigator.sendBeacon&&navigator.sendBeacon(u,new Blob([c],{type:"text/plain"}))}catch(_){}})}
          catch(_){if(navigator.sendBeacon)navigator.sendBeacon(u,new Blob([c],{type:"text/plain"}))}}}catch(_){}
        try{sessionStorage.setItem("painel_email",d)}catch(_){}
        const abre=()=>{document.open();document.write(h.split("__CAV__").join(P.cav||""));document.close()};
        if(s.t==="v")abre();else bioPerguntar(d,abre);return}catch(e){}}
      if(comoIrmao){try{sessionStorage.removeItem("ver_irmao")}catch(_){}await tentaSlots()}   // sem Meu painel: volta ao ADM
      if(!chave)throw 0}
    const k=await crypto.subtle.importKey("raw",chave,"AES-GCM",false,["decrypt"]);
    const html=new TextDecoder().decode(await crypto.subtle.decrypt({name:"AES-GCM",iv:b64(P.iv)},k,b64(P.ct)));
    registrar(true);
    try{sessionStorage.setItem("painel_email",d)}catch(_){}
    bioPerguntar(d,()=>{document.open();document.write(html);document.close()});
  }catch(e){if(d.includes("@"))registrar(false);if(P.otp)otpEsquecer(d);erro.textContent="Este e-mail não tem acesso a este painel. Confira se digitou certo ou solicite o acesso.";btn.disabled=false;btn.textContent="Entrar";cpf.select();
    if(d.includes("@")){solic.href="https://wa.me/?text="+encodeURIComponent(`Olá! Quero solicitar acesso ao ${document.title.replace(/ - Analise/," - Painel Insanos")}. Meu e-mail: ${d}`);solic.hidden=false}}
});
</script>__TUTORIAL__
__VERSAOSNIP__</body></html>"""

# ---------------------------------------------------------------------
#  "Meu painel": cada irmão (fora do comando) com e-mail na planilha vê só os dados dele
# ---------------------------------------------------------------------
ADM_EXTRA = []         # irmãos liberados como ADM pela tela Acessos
GERAL_RAW = ""         # painel de presença sem dados pessoais (o que o visitante vê)
CAVEIRA_URL = ""       # imagem da caveira (vem do template): usada no "Meu painel"
MEU_PUSH = None        # {"pub", "url"}: para o irmão ativar os avisos no Meu painel
LEMBRETES = {}         # pendências de cada irmão (a nuvem manda os lembretes nos horários)
PAG_MEU = {}           # mensalidade de cada irmão para o Meu painel (situação, link do boleto, vencimento)
INFO_EV = {}           # informações dos eventos (aba Eventos): (data ISO, evento em minúsculas) -> tipo, horário, local, endereço...
ENQ = {}               # enquete do comando (abas Enquetes e Votos): a última; a contagem só vai para o painel do comando
JUST = {}              # motivos das faltas (aba Justificativas): (data ISO, evento em minúsculas) -> {irmão: motivo}
ZAP_CTX = {}           # o que a mensagem do evento para o WhatsApp precisa (integrantes, entrada de cada um na divisão)
HOJE_EVS = []          # eventos de hoje (nome, horário): o bom-dia das 9h cita
ZAP_PAG = ""           # mensagem pronta da situação dos pagamentos (vai dentro do aviso de pagamento novo: abre direto no WhatsApp)

CLIMA_GEO, CLIMA_PREV = {}, {}   # nesta rodada: endereço -> coordenadas, e (lugar, data) -> previsão
def lugar_evento(data_iso, nome):
    """O endereço do evento (aba Eventos: endereço ou local), ou o da sede; e o horário, se houver."""
    inf = INFO_EV.get((str(data_iso)[:10], str(nome).strip().lower()), {})
    end = (inf.get("endereco") or inf.get("local") or DIV.get("sede_endereco") or "").strip()
    return end, str(inf.get("hora") or "")

def clima_evento(data_iso, hora="", endereco=""):
    """Previsão para o dia e a hora do evento (Open-Meteo, gratuito; só para os próximos 15 dias): "☀️ 24°C, sem chuva"."""
    import urllib.request, urllib.parse, datetime as _dt, re as _r
    try:
        dia = _dt.date.fromisoformat(str(data_iso)[:10])
        if not endereco or not (0 <= (dia - _dt.date.today()).days <= 15):
            return ""
        if endereco not in CLIMA_GEO:
            u = "https://nominatim.openstreetmap.org/search?" + urllib.parse.urlencode({"format": "json", "limit": 1, "countrycodes": "br", "q": endereco})
            req = urllib.request.Request(u, headers={"User-Agent": "PainelInsanos/1.0 (painel da divisao)"})
            r = json.loads(urllib.request.urlopen(req, timeout=8).read().decode("utf-8"))
            if not r:   # endereço não achado: tenta só a cidade (depois do último hífen ou vírgula)
                cid = _r.split(r"[-,]", endereco)[-1].strip() or "São Paulo"
                req = urllib.request.Request("https://nominatim.openstreetmap.org/search?" + urllib.parse.urlencode({"format": "json", "limit": 1, "countrycodes": "br", "q": cid}),
                                             headers={"User-Agent": "PainelInsanos/1.0 (painel da divisao)"})
                r = json.loads(urllib.request.urlopen(req, timeout=8).read().decode("utf-8"))
            CLIMA_GEO[endereco] = (float(r[0]["lat"]), float(r[0]["lon"])) if r else None
        geo = CLIMA_GEO.get(endereco)
        if not geo:
            return ""
        k = (geo, dia.isoformat())
        if k not in CLIMA_PREV:
            u = "https://api.open-meteo.com/v1/forecast?" + urllib.parse.urlencode({"latitude": geo[0], "longitude": geo[1],
                "hourly": "temperature_2m,precipitation_probability,weathercode", "timezone": "America/Sao_Paulo",
                "start_date": dia.isoformat(), "end_date": dia.isoformat()})
            CLIMA_PREV[k] = json.loads(urllib.request.urlopen(u, timeout=8).read().decode("utf-8")).get("hourly") or {}
        hh = CLIMA_PREV[k]
        m = _r.match(r"\s*(\d{1,2})", str(hora or ""))
        h = min(23, int(m.group(1))) if m else 12
        t = hh.get("temperature_2m", [None] * 24)[h]; pr = (hh.get("precipitation_probability") or [0] * 24)[h] or 0
        c = (hh.get("weathercode") or [0] * 24)[h] or 0
        if t is None:
            return ""
        ico = "☀️" if c == 0 else "🌤️" if c in (1, 2) else "☁️" if c == 3 else "🌫️" if c in (45, 48) else "⛈️" if c >= 95 else "🌦️" if c in (80, 81, 82) else "🌧️" if c >= 51 else "☁️"
        chuva = f"chance alta de chuva ({pr}%)" if pr >= 60 else f"pode chover ({pr}%)" if pr >= 30 else "sem chuva"
        return f"{ico} {round(t)}°C, {chuva}"
    except Exception as e:
        print(f"Aviso: previsão do tempo indisponível ({type(e).__name__}).")
        return ""

FRASES_BOMDIA = [
    "A estrada não espera: acelera com fé e chega com os irmãos.",
    "Quem tem irmão do lado não anda sozinho em nenhuma estrada.",
    "Cada dia é uma nova estrada. Faz valer cada quilômetro.",
    "Respeito, lealdade e irmandade: o resto a gente conquista.",
    "Coragem não é não ter medo, é seguir em frente mesmo com ele.",
    "O vento no rosto é de graça. A irmandade, a gente constrói todo dia.",
    "Faz hoje o que o teu eu de amanhã vai agradecer.",
    "Grandes viagens começam com um simples ligar do motor.",
    "Gente de palavra vale mais que qualquer moto zero.",
    "A vida é curta demais pra rodar devagar nos sonhos.",
    "Família não é só sangue: é quem divide a estrada com você.",
    "Disciplina é fazer o que precisa ser feito, mesmo sem vontade.",
    "Não existe estrada ruim pra quem viaja com os irmãos certos.",
    "Seja a razão de alguém sorrir hoje.",
    "Força nos braços, fé no coração e respeito em tudo.",
    "Quem planta lealdade colhe irmandade.",
    "O caminho se faz rodando. Bora pra cima!",
    "Hoje é um ótimo dia pra ser uma versão melhor de você.",
    "Nenhuma curva dura pra sempre. Segura firme e segue.",
    "A humildade abre portas que a força nunca abriria.",
    "Trabalha em silêncio e deixa o resultado fazer o barulho.",
    "Um irmão de verdade aparece na hora que precisa.",
    "Sonho grande, pé no chão e a mão no acelerador.",
    "Quem tem propósito não se perde no caminho.",
    "Cuida da tua saúde: a estrada precisa de você inteiro.",
    "O respeito a gente não pede: a gente conquista.",
    "Gratidão por mais um dia de vida, de estrada e de irmandade.",
    "Dias difíceis fazem pilotos fortes.",
    "A união faz a força, e a irmandade faz a diferença.",
    "Seja firme nos princípios e leve no coração.",
    "Quem desiste não sente o vento da chegada.",
    "Mais um dia pra honrar o colete que você veste.",
    "Pequenos passos todo dia levam a grandes destinos.",
    "Atitude é o motor que move os sonhos.",
    "O melhor da viagem são as pessoas que vão com a gente.",
    "Hoje, faz alguém se sentir parte da família.",
    "Paciência na estrada, sabedoria na vida.",
    "Não espera o dia perfeito: faz o dia ser perfeito.",
    "Quem anda com os bons nunca anda sozinho.",
    "Bom dia, guerreiro. Hoje é dia de vencer.",
]

def bom_dia(eventos_hoje=None):
    """O bom-dia das 9h: a previsão do dia na cidade da sede (Open-Meteo), com a chuva por período, e uma frase do dia."""
    import urllib.request, urllib.parse, datetime as _dt
    hoje = _dt.date.today()
    frase = FRASES_BOMDIA[hoje.toordinal() % len(FRASES_BOMDIA)]   # muda todo dia, sem repetir em dias seguidos
    clima, ico, cidade = "", "☀️", ""
    try:
        end = (DIV.get("sede_endereco") or DIV.get("sede_local") or "São Paulo").strip()
        u = "https://nominatim.openstreetmap.org/search?" + urllib.parse.urlencode({"format": "json", "limit": 1, "countrycodes": "br", "addressdetails": 1, "q": end})
        r = json.loads(urllib.request.urlopen(urllib.request.Request(u, headers={"User-Agent": "PainelInsanos/1.0 (painel da divisao)"}), timeout=8).read().decode("utf-8"))
        if not r:
            r = json.loads(urllib.request.urlopen(urllib.request.Request("https://nominatim.openstreetmap.org/search?" + urllib.parse.urlencode(
                {"format": "json", "limit": 1, "countrycodes": "br", "addressdetails": 1, "q": "São Paulo"}), headers={"User-Agent": "PainelInsanos/1.0 (painel da divisao)"}), timeout=8).read().decode("utf-8"))
        lat, lon = float(r[0]["lat"]), float(r[0]["lon"])
        a = r[0].get("address", {})
        cidade = a.get("city") or a.get("town") or a.get("municipality") or "São Paulo"
        u = "https://api.open-meteo.com/v1/forecast?" + urllib.parse.urlencode({"latitude": lat, "longitude": lon, "timezone": "America/Sao_Paulo",
            "daily": "temperature_2m_max,temperature_2m_min,precipitation_probability_max,weathercode", "hourly": "precipitation_probability",
            "start_date": hoje.isoformat(), "end_date": hoje.isoformat()})
        w = json.loads(urllib.request.urlopen(u, timeout=8).read().decode("utf-8"))
        dd, hh = w.get("daily", {}), (w.get("hourly", {}).get("precipitation_probability") or [0] * 24)
        tmin, tmax = round(dd["temperature_2m_min"][0]), round(dd["temperature_2m_max"][0])
        pmax, cod = int(dd.get("precipitation_probability_max", [0])[0] or 0), int(dd.get("weathercode", [0])[0] or 0)
        ico = "☀️" if cod == 0 else "🌤️" if cod in (1, 2) else "☁️" if cod == 3 else "🌫️" if cod in (45, 48) else "⛈️" if cod >= 95 else "🌦️" if cod in (80, 81, 82) else "🌧️" if cod >= 51 else "☁️"
        per = [n for n, (a_, b_) in (("de manhã", (6, 12)), ("à tarde", (12, 18)), ("à noite", (18, 24))) if max((hh[a_:b_] or [0])) >= 50]
        if pmax >= 50:
            quando = (" e ".join(per) if len(per) <= 2 else "o dia todo") if per else "em algum momento do dia"
            clima = f"{tmin}° a {tmax}°C, com chance de chuva {quando} ({pmax}%). Leva a capa de chuva! 🧥"
            ico = "🌧️" if ico in ("☀️", "🌤️", "☁️") else ico
        elif pmax >= 30:
            clima = f"{tmin}° a {tmax}°C, pode chover ({pmax}%). Fica de olho no céu. 👀"
        else:
            clima = f"{tmin}° a {tmax}°C, sem chuva. " + ("Dia perfeito pra rodar! 🏍️ Hidrata, irmão." if tmax >= 28 else "Bom dia pra rodar! 🏍️")
    except Exception as e:
        print(f"Aviso: previsão do bom-dia indisponível ({type(e).__name__}).")
    corpo = (f"Hoje em {cidade}: {clima} " if clima else "")
    if eventos_hoje:
        corpo += "Hoje tem " + " e ".join(f"{n}{' às ' + h if h else ''}" for n, h in eventos_hoje) + ". "
    corpo += f"“{frase}” 💀"
    return f"{ico} Bom dia, irmão!", corpo

def _zap_formatar(txt):
    """Igual ao app: numeração "01 ▸" nas linhas da lista e as linhas de traços viram uma faixa ═ do tamanho do texto."""
    import re as _r
    out = [_r.sub(r"^(\s*)(\d{1,2})\.\s*", lambda m: f"{m.group(1)}{m.group(2).zfill(2)} ▸ ", l) if _r.match(r"^\s*\d{1,2}\.\s", l) else l
           for l in str(txt).split("\n")]
    sep = lambda l: bool(_r.match(r"^\s*[-–—_=═]{8,}\s*$", l))
    larg = max([len(l) for l in out if not sep(l) and "http" not in l] or [0])
    n = max(14, min(26, int(larg * 2 / 3 + 0.5)))
    return "\n".join("═" * n if sep(l) else l for l in out)

def zap_evento(d, quem_email=""):
    """A mesma mensagem da "Lista do evento para o WhatsApp" (Presença › Data específica), para um evento lançado:
    vai pronta dentro do aviso de "Painel atualizado" (abre direto no WhatsApp). Leva o motivo de cada falta."""
    import datetime as _dt
    membros = ZAP_CTX.get("membros") or []
    if not membros:
        return ""
    ins = str(d.get("aba") or "").strip().lower() == "insights"
    T = ({"leg1": "✅-Respondeu", "leg2": "❌-Não respondeu", "frase": "teve uma taxa de resposta de"} if ins else
         {"leg1": "✅-Presente", "leg2": "❌-Ausente", "leg3": "✅ (staff)-Presente no staff", "frase": "teve uma adesão de"})
    mv = {"Sim": 1, "Staff": "S", "Não": 0, "-": "-"}
    marca = {n: mv.get(str(v), None) for n, v in (d.get("marcacoes") or {}).items()}
    just = d.get("justificativas") or {}
    data = str(d.get("data") or "")[:10]
    def grupo(m):
        c = str(m.get("c") or "").strip().lower()
        return "Full" if c.startswith("full") else "Meio" if c.startswith("meio") else "PP" if c.startswith("pp") else "Camiseta" if c.startswith("camiseta") else "Comando"
    def nome(m):
        b = f"{m.get('lg') or ''}.{m['n']}"
        return b + f"-{m.get('c')}" if grupo(m) == "Comando" else b
    def mk(m):
        v = marca.get(m["n"])
        t = "✅" if v == 1 else "✅ (staff)" if v == "S" else "❌" if v == 0 else ""
        if v == 0 and str(just.get(m["n"]) or "").strip():
            t += f" ({str(just[m['n']]).strip()})"
        return t
    tit = str(d.get("evento") or "Evento").strip()
    reg = str(DIV.get("regional", ""))
    cab = [f"*💀 {tit} 💀*", f"*Divisão {DIV.get('nome', '')} - {reg[:1] + reg[1:].lower()}*", "-" * 35]
    try:
        cab.append(f"*🗓️ Data: {_dt.date.fromisoformat(data):%d/%m/%Y}*")
    except Exception:
        cab.append(f"*🗓️ Data: {data}*")
    cab += ["-" * 35, "*Legenda:*", T["leg1"], T["leg2"]]
    if not ins and any(v == "S" for v in marca.values()):
        cab.append(T["leg3"])
    ordem = sorted(membros, key=lambda m: m.get("o", 999))
    blocos = []
    for g_ in ("Comando", "Full", "Meio", "PP", "Camiseta"):
        b = [m for m in ordem if grupo(m) == g_]
        if b:
            blocos.append("\n".join(nome(m) + (" " + mk(m) if mk(m) else "") for m in b))
    ent = ZAP_CTX.get("entrada") or {}
    base = [m for m in membros if marca.get(m["n"]) is not None or (ent.get(m["n"]) and data >= ent[m["n"]])]
    pres = sum(1 for m in base if marca.get(m["n"]) in (1, "S"))
    pct = int(pres * 100 / len(base) + 0.5) if base else 0
    cmd = [m for m in ordem if grupo(m) == "Comando"]
    ap = (EMAIL_APELIDO.get(str(quem_email).strip().lower()) or "").lower()
    a = next((m for m in cmd if m["n"].lower() == ap), None) or next((m for m in cmd if str(m.get("c") or "").strip().lower().startswith("sub")), None)
    assin = ""
    if a:
        import re as _r
        cb = _r.sub(r"\s*\([^)]*\)\s*$", "", str(a.get("c") or "")).strip()
        local = f"{DIV.get('assinatura_local', '')} - {DIV.get('estado', '')}"
        assin = (f"\n\n*Somente {cb} Altera a Lista*\n*{a['n']} - {a.get('c')} - {local}*"
                 f"\n\n*A {DIV.get('nome', '')} {reg} {T['frase']} {pct}% nesse evento.*")
    return _zap_formatar("\n".join(cab) + "\n\n" + "\n\n".join(blocos) + assin)

def zap_pagamentos(pag, membros):
    """A mesma mensagem do botão "Copiar mensagem para o comando": a lista por grupo (Comando, Full, Meio, PP, Camiseta),
    🏦 quem pagou, ⏰ quem está em compensação e os totais atualizados."""
    import re as _r
    linhas = str(pag.get("msg") or "").split("\n")
    eh_irmao = lambda l: bool(_r.match(r"^\s*\d{1,2}\.\s", l))
    idx = [k for k, l in enumerate(linhas) if eh_irmao(l)]
    if not idx:
        return pag.get("msg") or ""
    i0, i1 = idx[0], idx[-1]
    st = pag.get("status", {}) or {}
    trav = pag.get("travado", {}) or {}
    def grupo(m):
        c = str(m.get("c") or "").strip().lower()
        return "Full" if c.startswith("full") else "Meio" if c.startswith("meio") else "PP" if c.startswith("pp") else "Camiseta" if c.startswith("camiseta") else "Comando"
    def pago(n):
        s_ = str(st.get(n, ""))
        return bool(_r.search("ades", s_, _r.I)) or bool(trav.get(n)) or bool(_r.search("pago", s_, _r.I))
    def rot(m):
        lg = str(m.get("lg") or "").rstrip() or str(m.get("num") or "")
        return f"{m.get('lg') or lg + ' '} - {m.get('c')} {m['n']}" if grupo(m) == "Comando" else f"{lg} - {m['n']}"
    def marca(m):
        return "🏦" if pago(m["n"]) else "⏰" if _r.search("aguard", str(st.get(m["n"], "")), _r.I) else ""
    ordem = sorted(membros, key=lambda m: m.get("o", 999))
    blocos = []
    for g_ in ("Comando", "Full", "Meio", "PP", "Camiseta"):
        b = [m for m in ordem if grupo(m) == g_]
        if b:
            blocos.append("\n".join(rot(m) + marca(m) for m in b))
    pagos = sum(1 for m in ordem if pago(m["n"]))
    pct = int(pagos * 100 / len(ordem) + 0.5) if ordem else 0   # arredonda como o app (62,5% → 63%)
    fim = [_r.sub(r"representa\s*\d+%", f"representa {pct}%",
                  _r.sub(r"(Total de pagamentos efetuados:\s*)\d+(\.?)", lambda m: f"{m.group(1)}{pagos}{m.group(2)}", l, flags=_r.I), flags=_r.I)
           for l in linhas[i1 + 1:]]
    out = "\n".join(linhas[:i0] + ["\n\n".join(blocos)] + fim).split("\n")
    out = [_r.sub(r"^(\s*)(\d{1,2})\.\s+", lambda m: f"{m.group(1)}{m.group(2).zfill(2)} ▸ ", l) for l in out]
    sep = lambda l: bool(_r.match(r"^\s*[-–—_=═]{8,}\s*$", l))
    larg = max([len(l) for l in out if not sep(l) and "http" not in l] or [0])
    n = max(14, min(26, int(larg * 2 / 3 + 0.5)))
    return "\n".join("═" * n if sep(l) else l for l in out)
# Quando uma atualização exige reinstalar o app (mudança no manifesto, como o Compartilhar), suba o "v" e escreva o motivo.
# "quem": "adm" (só o comando) ou "todos" (comando e irmãos). O aviso aparece só no Android, no app instalado, até a pessoa marcar "Já reinstalei".
AVISO_REINSTALAR = {"v": 1, "quem": "adm", "motivo": "receber os flys pelo Compartilhar do WhatsApp"}
PESSOAIS = []          # [(email, html)] trancados no mesmo cofre do painel, cada um com o e-mail do irmão
MEU = r"""<!DOCTYPE html><html lang="pt-BR"><head><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1,viewport-fit=cover">
<title>Meu painel · Insanos</title><meta name="theme-color" content="#0d0c0b">
<link rel="manifest" href="manifest-adm.json"><link rel="icon" href="icon-192.png"><link rel="apple-touch-icon" href="icon-192.png">
<meta name="apple-mobile-web-app-capable" content="yes"><meta name="mobile-web-app-capable" content="yes"><meta name="apple-mobile-web-app-status-bar-style" content="black-translucent"><style>
:root{--osso:#ECE7DC;--osso2:#B8B2A6;--linha:#2c2925;--ouro:#E9B949}
*{box-sizing:border-box}body{margin:0;background:#0d0c0b;color:var(--osso);font-family:Barlow,"Segoe UI",Roboto,Arial,sans-serif;padding:env(safe-area-inset-top,0) 0 env(safe-area-inset-bottom,0)}
.w{max-width:620px;margin:0 auto;padding:18px 14px 40px}.top{display:flex;align-items:center;gap:12px;border-bottom:1px solid var(--linha);padding-bottom:12px}
.top img{width:52px;height:52px}.top b{display:block;font-size:20px}.top span{color:var(--osso2);font-size:13px}
h1{font-size:30px;margin:18px 0 2px}.sub{color:var(--osso2);font-size:14px;margin:0 0 14px}
.card{border:1px solid var(--linha);border-radius:16px;background:#161412;padding:16px;margin:12px 0}.card h2{font-size:17px;margin:0 0 12px}
.cavs{display:flex;justify-content:space-around;gap:10px;text-align:center}
.ficha-tela{position:fixed;inset:0;z-index:10050;background:#0b0908;overflow:auto;color:#fff;font-family:inherit}
.ficha-cx{max-width:560px;margin:0 auto;padding:calc(18px + env(safe-area-inset-top,0px)) 18px calc(24px + env(safe-area-inset-bottom,0px))}
.ficha-top{background:#B3261E;color:#fff;font-weight:900;letter-spacing:.06em;text-align:center;padding:12px;border-radius:14px;font-size:18px}
.ficha-cx h2{font-size:30px;margin:16px 0 12px;text-align:center}
.ficha-sg,.ficha-al,.ficha-i{border-radius:14px;padding:12px 14px;margin-bottom:10px;background:rgba(255,255,255,.06)}
.ficha-sg span,.ficha-al span,.ficha-i span{display:block;font-size:14px;color:#BDB7AA;font-weight:700;margin-bottom:3px}
.ficha-sg b{font-size:44px;color:#fff}.ficha-al{background:rgba(179,38,30,.25);border:2px solid #E5484D}.ficha-al span{color:#ffb4ae}.ficha-al b,.ficha-i b{font-size:22px;line-height:1.3}
.ficha-lig{display:block;text-decoration:none;text-align:center;background:#1E8E3E;color:#fff;font-weight:800;font-size:20px;padding:14px;border-radius:14px;margin:10px 0}.ficha-lig small{display:block;font-size:15px;font-weight:600;opacity:.9}
.ficha-x{display:block;width:100%;margin-top:14px;padding:14px;border-radius:14px;border:1px solid rgba(255,255,255,.3);background:none;color:#fff;font:inherit;font-weight:700;font-size:17px}

.hist{position:relative;padding-left:6px}.hist::before{content:"";position:absolute;left:22px;top:8px;bottom:8px;width:2px;background:rgba(233,185,73,.35)}
.hi{position:relative;display:flex;gap:12px;align-items:flex-start;padding:7px 0}.hi i{font-style:normal;width:34px;height:34px;flex:none;border-radius:50%;background:#1b1714;border:1px solid rgba(233,185,73,.55);display:flex;align-items:center;justify-content:center;font-size:17px;z-index:1}
.hi b{display:block;font-size:15px}.hi small{display:block;color:#BDB7AA;font-size:13px;margin-top:2px}.hi.fut{opacity:.55}.hi.fut i{border-style:dashed}
.cart .compart{display:block;width:100%;margin-top:12px;padding:12px;border-radius:12px;border:0;background:#E9B949;color:#1a1206;font:inherit;font-weight:800;font-size:15px}
.foto-bt{width:54px;height:54px;border-radius:50%;border:2px dashed rgba(236,231,220,.4);background:rgba(236,231,220,.06);padding:0;flex:none;overflow:hidden;display:inline-flex;align-items:center;justify-content:center;color:#ECE7DC;font-size:20px;line-height:1;margin-left:8px;vertical-align:middle}
.foto-bt.tem{border:2px solid #E9B949}.foto-bt img{width:100%;height:100%;object-fit:cover;display:block}.foto-bt span{display:flex;flex-direction:column;align-items:center}.foto-bt small{font-size:12px;font-weight:800;margin-top:-2px}
.cart .moto-f{position:relative;margin:12px 0 0;border-radius:14px;overflow:hidden;height:120px;border:1px solid rgba(233,185,73,.45)}.cart .moto-f img{width:100%;height:100%;object-fit:cover;display:block}
.cart .moto-f span{position:absolute;left:0;right:0;bottom:0;padding:16px 10px 6px;background:linear-gradient(transparent,rgba(0,0,0,.75));font-size:13px;font-weight:700;color:#ECE7DC}
.cart .meio .foto{width:86px;height:86px;border-radius:50%;object-fit:cover;border:2px solid #E9B949;flex:none;box-shadow:0 4px 14px rgba(0,0,0,.5)}
.cart .fundos{display:flex;gap:8px;justify-content:center;margin:12px 0 4px;flex-wrap:wrap}.cart .fundos button{width:34px;height:34px;border-radius:50%;border:2px solid rgba(236,231,220,.3);padding:0}
.cart .fundos button[aria-pressed="true"]{border-color:#E9B949;box-shadow:0 0 0 2px rgba(233,185,73,.35)}.cart .fundos-t{text-align:center;font-size:12.5px;color:#BDB7AA;margin-top:10px}
.cart.f-couro,.fundos [data-fd="couro"]{background:radial-gradient(circle at 20% 15%,rgba(255,220,170,.10),transparent 40%),repeating-radial-gradient(circle at 30% 40%,rgba(0,0,0,.18) 0 2px,transparent 2px 5px),linear-gradient(160deg,#5a3520,#2b170c)!important}
.cart.f-metal,.fundos [data-fd="metal"]{background:repeating-linear-gradient(90deg,rgba(255,255,255,.04) 0 1px,transparent 1px 3px),linear-gradient(135deg,#2a2d31,#5d636b 45%,#24272b)!important}
.cart.f-carbono,.fundos [data-fd="carbono"]{background:repeating-linear-gradient(45deg,#151515 0 6px,#232323 6px 12px),#181818!important}
.cart.f-sangue,.fundos [data-fd="sangue"]{background:radial-gradient(circle at 70% 0%,rgba(255,80,80,.18),transparent 45%),linear-gradient(160deg,#5c0b0b,#1c0404)!important}
.fundos [data-fd="padrao"]{background:#1b1714!important}
.cart.f-imagem{background:linear-gradient(180deg,rgba(0,0,0,.55),rgba(0,0,0,.72)),var(--cimg) center/cover no-repeat,#1b1714!important}   /* a película escura deixa o texto legível */
.fundos [data-fd="imagem"]{background:rgba(236,231,220,.08);color:#ECE7DC;font-size:15px}.fundo-img-acoes[hidden]{display:none}

/* escala única de texto (o Tamanho da letra dos Ajustes aumenta tudo na mesma proporção) */
:root{--fs-campo:16px;--fs-rot:13.5px;--fs-nota:13px;--fs-btn:15px;--fs-h2:21px}
.painel input:not([type=checkbox]):not([type=radio]):not([type=range]),.painel select,.painel textarea,.lw-jtx,.lanc-just,.card input:not([type=checkbox]):not([type=radio]),.card select,.card textarea,.fa-cx input:not([type=checkbox]),.fa-cx select,.fa-cx textarea{font-size:var(--fs-campo)!important}
select{text-overflow:ellipsis;white-space:nowrap;overflow:hidden}
.painel label,.card label,.fa-cx label{font-size:var(--fs-rot)}
.painel .nota,.painel p.sub,.card .nota,.fa-cx .nota{font-size:var(--fs-nota)}
.painel h2,.card h2{font-size:var(--fs-h2)}
.btn,.qbts button{font-size:var(--fs-btn)}

html.eco .fundo-video{display:none!important}html.eco .hl,html.eco .hl-cobre{display:none!important}
html.eco .hl-env,html.eco .cav,html.eco .cav-dash,html.eco .pres .cv{filter:none!important}
html.eco *,html.eco *::before,html.eco *::after{animation-duration:.001s!important;animation-iteration-count:1!important;transition-duration:.001s!important}
html.letra-g body{zoom:1.12}html.letra-gg body{zoom:1.25}
.fa-veu{position:fixed;inset:0;z-index:9990;background:rgba(0,0,0,.7);display:flex;align-items:flex-end;justify-content:center}
.fa-cx{background:#1b1714;color:#ECE7DC;width:100%;max-width:520px;max-height:88vh;overflow:auto;box-sizing:border-box;border-radius:20px 20px 0 0;padding:18px 16px calc(18px + env(safe-area-inset-bottom,0px));border-top:3px solid #E9B949}
.fa-fechar{display:block;width:100%;margin-top:14px;padding:12px;border-radius:12px;border:1px solid rgba(236,231,220,.25);background:transparent;color:#ECE7DC;font:inherit;font-weight:700;font-size:15px}
.aj-op{display:flex;gap:12px;align-items:center;width:100%;text-align:left;padding:12px;border-radius:14px;border:1px solid rgba(236,231,220,.2);background:rgba(236,231,220,.04);color:inherit;font:inherit}
.aj-op span{flex:1}.aj-op small{display:block;color:#BDB7AA;font-size:12.5px;margin-top:3px;line-height:1.35}.aj-op i{font-style:normal;font-weight:800;font-size:13px;color:#8f887d;white-space:nowrap}
.aj-op[aria-pressed="true"]{border-color:#4CAF6A;background:rgba(76,175,106,.1)}.aj-op[aria-pressed="true"] i{color:#4CAF6A}
.aj-letras{display:grid;grid-template-columns:repeat(3,1fr);gap:8px}.aj-letras button{padding:12px 6px;border-radius:12px;border:1px solid rgba(236,231,220,.22);background:none;color:inherit;font:inherit;font-weight:700}
.aj-letras button[data-l="g"]{font-size:17px}.aj-letras button[data-l="gg"]{font-size:19px}.aj-letras button[aria-pressed="true"]{background:#ECE7DC;color:#111}

.hl,.hl-cobre{position:absolute;inset:0;pointer-events:none;z-index:2;-webkit-mask:var(--m,var(--cav)) center/contain no-repeat;mask:var(--m,var(--cav)) center/contain no-repeat;
  background:linear-gradient(115deg,transparent 20%,rgba(255,255,255,0) 30%,rgba(200,200,200,.55) 40%,rgba(255,255,255,.95) 48%,rgba(150,150,150,.6) 54%,rgba(235,235,235,.8) 60%,rgba(255,255,255,0) 70%,transparent 80%);
  background-size:300% 300%;background-position:var(--hx,50%) var(--hy,50%);mix-blend-mode:soft-light;opacity:.9}
.hl-cobre.fora{display:none}.hl-env,.cav,.cav-dash{filter:drop-shadow(var(--sx,0px) var(--sy,0px) 9px var(--g,transparent))}
.hl-env.trofeu,.cav.trofeu,.cav-dash.trofeu{filter:drop-shadow(var(--sx,0px) var(--sy,0px) 10px rgba(233,185,73,.85)) drop-shadow(0 0 24px rgba(233,185,73,.35))}
.trofeu .hl-cobre,.trofeu .hl{mix-blend-mode:overlay;opacity:1}.cav .hl,.cav-dash .hl{clip-path:inset(calc(100% - var(--p,100%)) 0 0 0)}.hl-env{position:relative;display:block;width:fit-content;margin:0 auto;line-height:0}.hl-cobre.cv{top:2.54%;bottom:2.54%}
.trofeu{filter:drop-shadow(0 0 10px rgba(233,185,73,.75)) drop-shadow(0 0 26px rgba(233,185,73,.35))}
.cv{display:flex;flex-direction:column;align-items:center;gap:2px}
.cv b{font-size:34px;font-weight:800;line-height:1}.cv span{font-size:14px}.cv small{font-size:12px;color:var(--osso2)}
.cav{position:relative;width:120px;height:146px;isolation:isolate}
.cav:before{content:"";position:absolute;inset:0;background:var(--cav) center/contain no-repeat;filter:grayscale(1) brightness(.75) contrast(1.1)}
.cav:after{content:"";position:absolute;inset:0;background:linear-gradient(to top,var(--c) var(--p),transparent var(--p));-webkit-mask:var(--cav) center/contain no-repeat;mask:var(--cav) center/contain no-repeat;mix-blend-mode:multiply}
.pos{display:grid;grid-template-columns:1fr 1fr;gap:10px}.pos div{background:rgba(236,231,220,.05);border-radius:12px;padding:12px;text-align:center}
.pos b{display:block;font-size:30px;color:var(--ouro)}.pos span{font-size:13px;color:var(--osso2)}
.cmp{margin:6px 0 14px}.cmp p{margin:0 0 6px;font-weight:700}.br{display:grid;grid-template-columns:110px 1fr 46px;gap:8px;align-items:center;font-size:13px;margin:5px 0}
.tr{height:10px;border-radius:6px;background:rgba(236,231,220,.08)}.tr i{display:block;height:10px;border-radius:6px}.br em{font-style:normal;font-weight:800;text-align:right}
.mes{display:grid;grid-template-columns:70px 1fr 1fr;gap:8px;align-items:center;font-size:13px;padding:7px 0;border-top:1px solid var(--linha)}
.mes:first-of-type{border-top:0}.ch{padding:3px 8px;border-radius:8px;font-weight:800;text-align:center}
details{border-top:1px solid var(--linha);padding-top:8px;margin-top:8px}summary{cursor:pointer;font-weight:700;padding:6px 0}
.ev{display:flex;gap:10px;font-size:13.5px;padding:6px 0;border-bottom:1px solid rgba(236,231,220,.06)}.ev i{font-style:normal;width:22px}.ev span{color:var(--osso2);width:44px;flex:0 0 44px}
.tabs{display:flex;gap:6px;margin:0 0 8px}.tabs button{flex:1;padding:9px 6px;border-radius:999px;border:1px solid var(--linha);background:transparent;color:var(--osso);font-weight:700;font-size:14px}
.tabs button.on{background:var(--ouro);color:#141210;border-color:var(--ouro)}.tabs.peq button{padding:6px;font-size:12.5px;font-weight:600}.tabs.peq button.on{background:rgba(233,185,73,.18);color:var(--ouro)}
.ev .q{border:1px solid rgba(240,182,180,.5);background:transparent;color:#F2A3A3;border-radius:999px;padding:4px 10px;font-size:12px;font-weight:700}
.qbox{background:rgba(236,231,220,.05);border-radius:12px;padding:10px;margin:4px 0 10px}.qbox pre{white-space:pre-wrap;font-family:inherit;font-size:13px;margin:0 0 8px}
.qbts{display:flex;gap:8px}.qbts button{flex:1;padding:9px;border-radius:999px;border:1px solid var(--linha);background:transparent;color:var(--osso);font-weight:700}.qbts .qz{background:#1f8f4e;border-color:#1f8f4e;color:#fff}
.mens-d summary{display:flex;align-items:center;gap:10px;cursor:pointer;list-style:none}.mens-d summary::-webkit-details-marker{display:none}
.mens-d summary b{display:block;font-size:17px}.mens-d summary small{display:block;font-size:13.5px;color:#E9B949;font-weight:700;margin-top:2px}
.mens-d summary i{margin-left:auto;font-style:normal;font-size:22px;color:#BDB7AA;transition:transform .2s}.mens-d[open] summary i{transform:rotate(90deg)}
.mens-d[open] summary{margin-bottom:12px}
.cart .pres{display:flex;align-items:center;gap:10px;margin-top:12px;padding:10px 12px;border-radius:14px;background:rgba(236,231,220,.05)}
.cart .pres .cv{position:relative;width:46px;flex:0 0 46px}.cart .pres .cv img{display:block;width:100%;height:auto;filter:grayscale(1) brightness(.45)}
.cart .pres .cv .o{position:absolute;inset:0;-webkit-mask:var(--cav) center/contain no-repeat;mask:var(--cav) center/contain no-repeat;
  background:var(--cav) center/contain no-repeat,linear-gradient(var(--c),var(--c));background-blend-mode:multiply}
.pres .cv{filter:drop-shadow(var(--sx,0px) var(--sy,0px) 6px var(--g,transparent))}.pres .cv.trofeu{filter:drop-shadow(var(--sx,0px) var(--sy,0px) 7px rgba(233,185,73,.85)) drop-shadow(0 0 16px rgba(233,185,73,.35))}
.cart .pres b{font-size:28px;line-height:1;color:#E9B949}.cart .pres span{display:block;font-size:12px;color:#BDB7AA}
.qbts a.qz{display:inline-flex;align-items:center;justify-content:center;padding:10px 14px;border-radius:12px;font-weight:700;font-size:14px;text-decoration:none;color:#fff}
.cart-bts{display:grid;grid-template-columns:1fr 1fr;gap:8px;margin:-2px 0 12px}
.cart-bts button{padding:11px 8px;border-radius:14px;border:1px solid rgba(233,185,73,.4);background:rgba(233,185,73,.07);color:#ECE7DC;font-weight:700;font-size:14px}
#conqN{display:inline-block;min-width:18px;padding:0 6px;border-radius:9px;background:#E9B949;color:#141210;font-size:11px;line-height:18px}
#conqBox{display:grid;grid-template-columns:1fr 1fr;gap:8px;margin:0 0 14px}#conqBox[hidden]{display:none}
.cq{padding:12px;border-radius:14px;background:rgba(236,231,220,.05);border:1px solid rgba(236,231,220,.1)}.cq b{display:block;font-size:22px;line-height:1.1}
.cq strong{display:block;font-size:13.5px;margin-top:4px}.cq span{display:block;font-size:12px;color:#BDB7AA;margin-top:2px}.cq.off{opacity:.45;filter:grayscale(1)}
.cart-veu{position:fixed;inset:0;z-index:9998;background:rgba(0,0,0,.78);display:flex;align-items:center;justify-content:center;padding:16px}
.cart{width:min(380px,100%);border-radius:22px;padding:18px;background:linear-gradient(150deg,#262019,#0f0d0b 70%);border:2px solid #E9B949;box-shadow:0 20px 60px rgba(0,0,0,.6);color:#ECE7DC;position:relative}
.cart .topo{display:flex;align-items:center;gap:10px;border-bottom:1px solid rgba(233,185,73,.35);padding-bottom:10px;margin-bottom:12px}
.cart .topo img{width:40px;height:auto}.cart .topo b{font-size:16px;letter-spacing:.06em}.cart .topo span{display:block;font-size:12px;color:#BDB7AA}
.cart .meio{display:flex;gap:14px;align-items:center}.cart .meio .pin{width:74px;height:auto;flex:0 0 74px}
.cart h3{margin:0;font-size:28px;line-height:1.05}.cart .c{font-size:14px;color:#E9B949;font-weight:700;margin-top:4px}.cart .d{font-size:12.5px;color:#BDB7AA;margin-top:4px}
.cart .qr{display:flex;align-items:center;gap:12px;margin-top:14px;padding-top:12px;border-top:1px solid rgba(233,185,73,.35)}
.cart .qr canvas{width:110px;height:110px;background:#fff;border-radius:10px;padding:6px;box-sizing:border-box}.cart .qr p{margin:0;font-size:12px;color:#BDB7AA;line-height:1.4}
.cart .fechar{display:block;width:100%;margin-top:14px;padding:11px;border-radius:12px;border:1px solid rgba(236,231,220,.3);background:transparent;color:#ECE7DC;font-weight:700;font-size:15px}
.msg-meu{display:block;width:100%;margin:-4px 0 14px;padding:13px;border-radius:16px;border:1px solid rgba(233,185,73,.45);background:rgba(233,185,73,.08);color:#ECE7DC;font-weight:800;font-size:16px}
.sos-meu{display:block;width:100%;margin:0 0 14px;padding:15px;border-radius:16px;border:0;background:#c62828;color:#fff;font-weight:900;font-size:17px;letter-spacing:.3px;box-shadow:0 0 0 3px rgba(198,40,40,.22)}
.nota{color:var(--osso2);font-size:12.5px}.sair{display:block;margin:18px auto 0;background:transparent;border:1px solid var(--linha);color:var(--osso);border-radius:999px;padding:10px 22px;font-weight:700}
</style></head><body><div class="w" id="app"></div><script>
const D=__DADOS__,CAV="__CAV__";
// honrarias: o pin pelo ano de entrada no MC (entregue no começo do ano em que completa 3, 6 ou 9 anos)
const PIN=(()=>{if(!D.ent)return null;const a=new Date().getFullYear()-(+String(D.ent).slice(0,4));
  return a>=9?["ouro","Pin Ouro"]:a>=6?["prata","Pin Prata"]:a>=3?["bronze","Pin Bronze"]:["caveira","Pin Caveira"]})();
const cor=v=>v==null?["transparent","#ECE7DC"]:v>=90?["#E9B949","#E9B949"]:v>75?["#9BE3AE","#8FD9A3"]:v>=50?["#F28B82","#F28B82"]:["#9B1111","#C0392B"];
const pc=a=>a&&a[1]?Math.round(a[0]*100/a[1]):null,es=s=>String(s).replace(/[&<>"]/g,c=>({"&":"&amp;","<":"&lt;",">":"&gt;",'"':"&quot;"}[c]));
const MS=["jan","fev","mar","abr","mai","jun","jul","ago","set","out","nov","dez"];
const MESES_N=["Janeiro","Fevereiro","Março","Abril","Maio","Junho","Julho","Agosto","Setembro","Outubro","Novembro","Dezembro"];
// Minha mensalidade: a situação do mês, o boleto a um toque e, se precisar, o pedido do boleto ao ADM
function mensQuadro(){const M=D.mens,st=String(M.st||""),mes=M.venc?MESES_N[+M.venc.slice(5,7)-1]:"",venc=M.venc?M.venc.slice(8,10)+"/"+M.venc.slice(5,7):"";
  const pago=/pag|ades/i.test(st)&&!/aberto/i.test(st),comp=/aguard|compens/i.test(st);
  if(pago)return `<div class="card" id="mensal"><h2 style="margin:0">💵 Minha mensalidade</h2><p style="margin:6px 0 0;font-weight:700;color:#4CAF6A">✅ ${mes} paga. Valeu, irmão!</p></div>`;
  return `<details class="card mens-d" id="mensal"><summary><span><b>💵 Minha mensalidade</b><small>${comp?`⏰ ${mes} em compensação`:`⏳ ${mes} em aberto${venc?" · vence "+venc:""}`}</small></span><i>›</i></summary>
    <div class="qbts">${M.lk?`<a class="qz" href="${es(M.lk)}" target="_blank" rel="noopener" style="text-decoration:none">📄 Abrir meu boleto</a>`:""}
    ${M.adm?`<button id="mensPedir">🙋 Pedir o boleto ao ${es(M.adm)}</button>`:""}</div>
    <p class="nota" id="mensTx" style="margin:8px 0 0">${M.lk?"O boleto também chega pelos avisos do app.":"O boleto deste mês ainda não está no painel. Se precisar, peça ao comando."}</p></details>`}
const cav=(v,r,sub)=>`<div class="cv"><span class="cav${v!=null&&v>=90?" trofeu":""}" style="--g:${corGlow(v)};--cav:url('${CAV}');--c:${cor(v)[0]};--p:${v||0}%"><i class="hl"></i></span><b style="color:${cor(v)[1]}">${v==null?"–":v+"%"}</b><span>${r}</span><small>${sub}</small></div>`;
const bar=(r,v)=>`<div class="br"><span>${r}</span><div class="tr"><i style="width:${v||0}%;background:${cor(v)[1]}"></i></div><em style="color:${cor(v)[1]}">${v==null?"–":v+"%"}</em></div>`;
const vp=pc(D.pres),vi=pc(D.ins);
const pos=(x,rot)=>x?`<div><b>${x[0]}º</b><span>de ${x[1]} ${rot}</span></div>`:`<div><b>🆕</b><span>${rot}: entra no ranking a partir de ${D.min} eventos</span></div>`;
document.getElementById("app").innerHTML=`<div class="top"><img src="${CAV}" alt=""><div><b>INSANOS M.C.</b><span>Divisão ${es(D.div)} · ${es(D.reg)}</span></div></div>
<h1 style="display:flex;align-items:center;gap:12px">${es(D.n)}${PIN?`<img src="pin-${PIN[0]}.png" alt="${PIN[1]}" title="${PIN[1]}" style="width:44px;height:auto">`:""}</h1><p class="sub">${es(D.c)} · Meu painel · ${D.ano}${PIN?" · "+PIN[1]:""}</p>
<div id="comoIrmao" hidden style="margin:0 0 12px;padding:10px 12px;border-radius:12px;background:rgba(233,185,73,.12);border:1px solid rgba(233,185,73,.45);font-size:14px">👁️ Você está vendo <b>como irmão</b> · <a href="#" id="voltarCmd" style="color:#E9B949;font-weight:800">Voltar ao painel do comando</a></div>
<div class="cart-bts"><button id="cartBt">🪪 Minha carteirinha</button><button id="conqBt" aria-expanded="false">🏆 Minhas conquistas <span id="conqN"></span></button></div>
<div id="conqBox" hidden></div>
<button class="sos-meu" onclick="abrirSOS()">🆘 SOS · Preciso de ajuda</button>
<button class="msg-meu" onclick="abrirMensagens({modo:'irmao',comando:D.cmd||[]})">✉️ Mensagens · falar com o comando</button>
${D.enq&&D.enq.ops?`<div class="card" id="enqC" style="border-color:rgba(233,185,73,.55)"><h2 style="margin:0 0 4px">📊 Enquete do comando</h2><p style="margin:0 0 10px;font-weight:600">${es(D.enq.p)}</p><div id="enqOps" class="qbts" style="flex-wrap:wrap"></div><p class="nota" id="enqNota" style="margin:8px 0 0;font-size:12.5px"></p></div>`:""}
<div class="tri" id="tri"></div>
${(D.prox||[]).length&&D.ru?`<details class="card mens-d" id="eventos"><summary><span><b>🗓️ Próximos eventos</b><small id="proxRes">${D.prox.length} evento${D.prox.length>1?"s":""}</small></span><i>›</i></summary><div id="proxL"></div>
  <p class="nota" style="margin-top:8px">A sua resposta vai para a lista da divisão (atualiza em até 30 minutos). Pode mudar de ideia até o dia do evento.</p></details>`:""}
${(D.anivs||[]).length?`<details class="card mens-d" id="anivC"><summary><span><b>🎂 Próximos aniversários</b><small id="anivRes"></small></span><i>›</i></summary><div id="anivL"></div></details>`:""}
${D.insp&&D.ru?`<div class="card" id="insightP"><h2>💬 ${es(D.insp.e)}</h2><p class="nota" id="insTx" style="margin:0 0 10px">Ainda não consta a sua resposta. Já respondeu no grupo? Toque abaixo para marcar.</p>
  <div class="qbts"><button class="qz" id="insBt">✅ Já respondi</button></div></div>`:""}
${D.mens&&D.mens.st?mensQuadro():""}
<div class="card" id="adesaoCard"><div style="display:flex;align-items:center;gap:8px;margin-bottom:6px"><h2 id="adTit" style="flex:1;margin:0">Minha adesão em ${D.ano}</h2><select id="perSel" aria-label="Período" style="max-width:46%;padding:7px 9px;border-radius:10px;border:1px solid rgba(236,231,220,.25);background:rgba(0,0,0,.35);color:inherit;font:inherit"><option value="">Ano todo (${D.ano})</option>${[...D.meses].reverse().map(m=>`<option value="${m.ym}">${MESES_N[+m.ym.slice(5)-1]}/${m.ym.slice(0,4)}</option>`).join("")}</select></div><div class="cavs" id="adCavs">${cav(vp,"Presença",D.pres[1]?`${D.pres[0]} de ${D.pres[1]} eventos`:"sem eventos")}${cav(vi,"Insights",D.ins[1]?`${D.ins[0]} de ${D.ins[1]} insights`:"sem insights")}</div></div>
<div class="card"><h2>Minha posição (presença) <span class="nota" style="font-weight:500">· no ano</span></h2><div class="pos">${pos(D.pf,"na função ("+es(D.g)+")")}${pos(D.pd,"na divisão")}</div></div>
<div class="card"><h2>Comparação</h2>
  <div class="cmp"><p>Presença</p>${bar("Eu",vp)}${bar("Média "+es(D.g),D.mf[0])}${bar("Média da divisão",D.md[0])}</div>
  <div class="cmp"><p>Insights</p>${bar("Eu",vi)}${bar("Média "+es(D.g),D.mf[1])}${bar("Média da divisão",D.md[1])}</div></div>
<div class="card"><h2>Mês a mês</h2><div class="mes nota"><span></span><span>Presença</span><span>Insights</span></div>${D.meses.map(m=>{const a=pc(m.p),b=pc(m.i);
  const ch=(v,x)=>x&&x[1]?`<span class="ch" style="color:${cor(v)[1]};background:rgba(236,231,220,.05)">${x[0]}/${x[1]} · ${v}%</span>`:'<span class="nota">–</span>';
  return `<div class="mes"><b>${MS[+m.ym.slice(5)-1]}/${m.ym.slice(2,4)}</b>${ch(a,m.p)}${ch(b,m.i)}</div>`}).join("")}
</div>
<div class="card"><h2 id="mpTit">📋 Minhas presenças em ${D.ano}</h2>
  <div class="tabs" id="tTipo"><button data-t="ev" class="on">🗓️ Eventos</button><button data-t="iv">💬 Insights</button></div>
  <div class="tabs peq" id="tFil"><button data-f="" class="on">Todos</button><button data-f="x">❌ Só onde faltei</button></div>
  <div id="lista"></div>
  <p class="nota" style="margin-top:10px">Acha que alguma falta está errada? Toque em <b>Questionar</b>: a mensagem sai pronta para você mandar no WhatsApp.</p></div>
<div class="card" id="instalar"><h2>📲 Instale o app</h2><p class="nota" style="margin:0 0 10px">Ícone na tela inicial, tela cheia e os avisos da divisão no seu celular.</p>
  <div class="qbts"><button class="qz" onclick="tutAbrir()">Ver como instalar</button></div></div>
<div class="card" id="avisos"><h2>🔔 Avisos no celular</h2><p class="nota" id="avSt" style="margin:0 0 10px">Receba lembretes das listas e da agenda da divisão, só para você.</p>
  <div class="qbts"><button id="avAt" class="qz">🔔 Ativar avisos</button><button id="avDes" hidden>Desativar</button></div>
  <div class="qbts" style="margin-top:8px"><button id="bdTg" style="font-size:13px">🔕 Não receber o bom-dia</button></div>
  <details style="margin-top:12px"><summary>🔊 Deixar o alerta do SOS bem alto (Android)</summary>
    <ol class="nota" style="margin:8px 0 0;padding-left:18px;line-height:1.5">
      <li>Abra as <b>Configurações</b> do celular › <b>Apps</b> (ou Aplicativos).</li>
      <li>Toque em <b>Insanos ADM</b>. Se não aparecer, toque em <b>Chrome</b>.</li>
      <li>Toque em <b>Notificações</b> e confira se estão <b>ativadas</b>.</li>
      <li>Toque na categoria das notificações do painel (no Chrome, em <b>Sites</b> › o endereço do painel, que termina em <b>github.io</b>).</li>
      <li>Em <b>Som</b>, escolha um toque <b>alto</b> (um de alarme é o ideal).</li>
      <li>Ative <b>Mostrar como pop-up</b> (ou <b>Prioridade alta / Alertar</b>) e <b>Vibração</b>.</li>
      <li>Ative <b>Ignorar Não perturbe</b> (ou <b>Exceções do Não perturbe</b>), para o SOS tocar mesmo no silencioso.</li>
      <li>Em <b>Tela de bloqueio</b>, escolha <b>Mostrar todo o conteúdo</b>.</li>
    </ol><p class="nota" style="margin:6px 0 0">Os nomes mudam um pouco conforme a marca (Samsung, Motorola, Xiaomi…). No iPhone, o som e a vibração são os padrões do aparelho.</p></details>
  <details style="margin-top:8px"><summary>💡 Fazer a tela piscar quando chegar um aviso</summary>
    <ol class="nota" style="margin:8px 0 0;padding-left:18px;line-height:1.5">
      <li><b>Android:</b> abra as <b>Configurações</b> › <b>Acessibilidade</b> › <b>Notificação com flash</b> (na Samsung: Acessibilidade › Configurações avançadas › Notificação com flash).</li>
      <li>Ligue <b>Flash da tela</b> (a tela pisca) e, se quiser, <b>Flash da câmera</b> (a luz da câmera pisca).</li>
      <li><b>iPhone:</b> Ajustes › Acessibilidade › Áudio e Visual › <b>Flash LED para Alertas</b>.</li>
    </ol><p class="nota" style="margin:6px 0 0">Vale para todos os avisos do celular, não só os do app.</p></details>
  <p class="nota" style="margin:12px 0 8px">🆘 Nos alertas de SOS, mostrar a que distância você está do irmão (a sua localização fica só neste celular):</p>
  <div class="qbts"><button id="locAt">📍 Usar minha localização no SOS</button></div></div>
<p class="nota" style="text-align:center">Atualizado em ${es(D.ger)} · Insanos Sempre, Sempre Insanos</p>
<button class="sair" onclick="try{sessionStorage.clear()}catch(e){}location.reload()">Sair</button>`;
// lista de presenças com filtro e "Questionar"
let TT="ev",FF="",ABERTO=-1,PER="";   // PER: mês escolhido no Período ("" = ano todo)
const br=d=>d.slice(8,10)+"/"+d.slice(5,7)+"/"+d.slice(0,4);
const texto=e=>TT==="ev"
  ?`Fala, comando! Sou o *${D.n}* (${D.c}).\n\nNo painel aparece que eu *não fui* ao evento *${e[1]}* de *${br(e[0])}*, mas eu estive presente.\n\nPodem conferir o lançamento? Obrigado! 🤝`
  :`Fala, comando! Sou o *${D.n}* (${D.c}).\n\nNo painel aparece que eu *não respondi* o insight *${e[1]}* de *${br(e[0])}*, mas eu respondi.\n\nPodem conferir o lançamento? Obrigado! 🤝`;
function desenha(){const L=(TT==="ev"?D.ev:D.iv).filter(e=>(!FF||e[2]!==1)&&(!PER||e[0].startsWith(PER)));
  document.getElementById("lista").innerHTML=L.length?L.map((e,k)=>`<div class="ev"><i>${e[2]===1?"✅":"❌"}</i><span>${e[0].slice(8,10)}/${e[0].slice(5,7)}</span><b style="flex:1;font-weight:500">${es(e[1])}${e[2]===0&&e[3]?`<small style="display:block;color:#E07472;font-weight:600;font-size:12.5px">Motivo: ${es(e[3])}</small>`:""}</b>
    ${e[2]===1?"":`<button class="q" data-k="${k}">Questionar</button>`}</div>${k===ABERTO?`<div class="qbox"><pre>${es(texto(e))}</pre>
      <div class="qbts"><button class="qc" data-k="${k}">📋 Copiar</button><button class="qz" data-k="${k}">Abrir no WhatsApp</button></div></div>`:""}`).join("")
    :`<p class="nota">${FF?"Nenhuma falta. 💪":"Nada lançado ainda."}</p>`;
  document.querySelectorAll("#lista [data-k]").forEach(b=>b.onclick=async()=>{const k=+b.dataset.k,e=L[k];
    if(b.classList.contains("q")){ABERTO=ABERTO===k?-1:k;desenha();return}
    if(b.classList.contains("qc")){try{await navigator.clipboard.writeText(texto(e));b.textContent="✅ Copiado"}catch(_){alert("Não consegui copiar: selecione o texto e copie.")}return}
    if(b.classList.contains("qz"))window.open("https://wa.me/?text="+encodeURIComponent(texto(e)),"_blank")})}
// próximos eventos: responder "Vou / Não vou" (vai para a aba Acessos com a assinatura; a nuvem marca na Presença)
// insight: "Já respondi" (vai para a aba Acessos com a assinatura; a nuvem marca SIM na aba Insights)
(function(){const b=document.getElementById("insBt");if(!b)return;let feito=false;try{feito=localStorage.getItem("insr_"+D.email)===D.insp.d+"|"+D.insp.e}catch(_){}
  const ok=()=>{b.textContent="✅ Marcado como respondido";b.disabled=true;document.getElementById("insTx").textContent="Valeu! A sua resposta entra na lista em até 30 minutos."};
  if(feito)ok();
  b.onclick=async()=>{try{await acordarDepois(fetch(D.ru,{method:"POST",mode:"no-cors",body:JSON.stringify({email:D.email,ok:true,painel:"✋|resp|"+D.insp.d+"|"+D.insp.tk+"|"+D.insp.e,aparelho:/iPhone|iPad/.test(navigator.userAgent)?"iPhone":/Android/.test(navigator.userAgent)?"Celular Android":"Outro"})}));
    try{localStorage.setItem("insr_"+D.email,D.insp.d+"|"+D.insp.e)}catch(_){}ok()}catch(_){alert("Não consegui enviar agora. Tente de novo em instantes.")}}})();
function agendaAdd(e){const hm=String(e.hora||"").match(/(\d{1,2})\s*[:hH]\s*(\d{2})?/);
  if(hm){   // com horário: evento de 3 horas, no horário de Brasília
    const ini=new Date(`${e.d}T${hm[1].padStart(2,"0")}:${(hm[2]||"00")}:00`),fim=new Date(ini.getTime()+3*36e5),f=x=>`${x.getFullYear()}${String(x.getMonth()+1).padStart(2,"0")}${String(x.getDate()).padStart(2,"0")}T${String(x.getHours()).padStart(2,"0")}${String(x.getMinutes()).padStart(2,"0")}00`;
    const tit=`Insanos · ${e.e}`,loc=[e.local||D.sede_n,e.endereco||D.sede].filter(Boolean).join(" · "),det=`Divisão ${D.div} ${D.reg}. Confirme sua presença no app.`;
    if(/iPhone|iPad|iPod/.test(navigator.userAgent)){const ics=["BEGIN:VCALENDAR","VERSION:2.0","PRODID:-//Insanos//Painel//PT","BEGIN:VEVENT",`UID:${f(ini)}-${encodeURIComponent(e.e)}@insanos`,
        `DTSTAMP:${new Date().toISOString().replace(/[-:]/g,"").slice(0,15)}Z`,`DTSTART;TZID=America/Sao_Paulo:${f(ini)}`,`DTEND;TZID=America/Sao_Paulo:${f(fim)}`,`SUMMARY:${tit}`,`LOCATION:${loc}`,`DESCRIPTION:${det}`,"END:VEVENT","END:VCALENDAR"].join("\r\n");
      const a=document.createElement("a");a.href="data:text/calendar;charset=utf-8,"+encodeURIComponent(ics);a.download="evento-insanos.ics";document.body.appendChild(a);a.click();a.remove();return}
    window.open(`https://calendar.google.com/calendar/render?action=TEMPLATE&text=${encodeURIComponent(tit)}&dates=${f(ini)}/${f(fim)}&ctz=America/Sao_Paulo&location=${encodeURIComponent(loc)}&details=${encodeURIComponent(det)}`,"_blank");return}
  const d1=e.d.replace(/-/g,""),dt=new Date(e.d+"T12:00:00");dt.setDate(dt.getDate()+1);
  const d2=`${dt.getFullYear()}${String(dt.getMonth()+1).padStart(2,"0")}${String(dt.getDate()).padStart(2,"0")}`;
  const tit=`Insanos · ${e.e}`,loc=[D.sede_n,D.sede].filter(Boolean).join(" · "),det=`Divisão ${D.div} ${D.reg}. Confirme sua presença no app.`;
  if(/iPhone|iPad|iPod/.test(navigator.userAgent)){   // iPhone: arquivo de agenda (abre direto no Calendário)
    const ics=["BEGIN:VCALENDAR","VERSION:2.0","PRODID:-//Insanos//Painel//PT","BEGIN:VEVENT",`UID:${d1}-${encodeURIComponent(e.e)}@insanos`,`DTSTAMP:${new Date().toISOString().replace(/[-:]/g,"").slice(0,15)}Z`,
      `DTSTART;VALUE=DATE:${d1}`,`DTEND;VALUE=DATE:${d2}`,`SUMMARY:${tit}`,`LOCATION:${loc}`,`DESCRIPTION:${det}`,"END:VEVENT","END:VCALENDAR"].join("\r\n");
    const a=document.createElement("a");a.href="data:text/calendar;charset=utf-8,"+encodeURIComponent(ics);a.download="evento-insanos.ics";document.body.appendChild(a);a.click();a.remove();return}
  window.open(`https://calendar.google.com/calendar/render?action=TEMPLATE&text=${encodeURIComponent(tit)}&dates=${d1}/${d2}&location=${encodeURIComponent(loc)}&details=${encodeURIComponent(det)}`,"_blank")}
function proxDesenha(){const el=document.getElementById("proxL");if(!el)return;let loc={};try{loc=JSON.parse(localStorage.getItem("rsvp_"+D.email)||"{}")}catch(_){}
  {const sem=D.prox.filter(e=>!(loc[e.d+"|"+e.e]||e.st)).length,r=document.getElementById("proxRes");
    if(r){r.textContent=`${D.prox.length} evento${D.prox.length>1?"s":""}`+(sem?` · ${sem} sem a sua resposta`:" · todos respondidos ✓");r.style.color=sem?"#E9B949":""}}
  el.innerHTML=D.prox.map((e,k)=>{const st=loc[e.d+"|"+e.e]||e.st,dd=e.d.slice(8,10)+"/"+e.d.slice(5,7);
    const txt=st==="sim"?"✅ Você vai":st==="nao"?"❌ Você não vai":"Ainda não respondeu";
    return `<div class="ev" style="flex-wrap:wrap"><span>${dd}</span><b style="flex:1;font-weight:600">${es(e.e)}</b>${e.clima?`<small style="width:100%;padding-left:54px;color:#BDB7AA">Previsão: ${es(e.clima)}</small>`:""}${e.hora||e.local?`<small style="width:100%;padding-left:54px;color:#E9B949;font-weight:600">${[e.hora&&"🕗 "+es(e.hora),e.local&&"📍 "+es(e.local)].filter(Boolean).join(" · ")}</small>`:""}<small class="nota" style="width:100%;padding-left:54px">${txt}</small>
      <div class="qbts" style="width:100%;padding-left:54px;margin-top:6px"><button data-r="vou" data-k="${k}" class="${st==="sim"?"qz":""}">✅ Vou</button><button data-r="nao" data-k="${k}"${st==="nao"?' style="background:#7a2e2c;border-color:#7a2e2c;color:#fff"':""}>❌ Não vou</button><button data-ag="${k}">📅 Agenda</button>${e.mapa?`<button data-mp="${k}">🗺️ Maps</button><button data-wz="${k}">🚙 Waze</button>`:""}</div></div>`}).join("");
  el.querySelectorAll("[data-ag]").forEach(b=>b.onclick=()=>agendaAdd(D.prox[+b.dataset.ag]));
el.querySelectorAll("[data-mp]").forEach(b=>b.onclick=()=>window.open("https://www.google.com/maps/search/?api=1&query="+encodeURIComponent(D.prox[+b.dataset.mp].mapa),"_blank"));
el.querySelectorAll("[data-wz]").forEach(b=>b.onclick=()=>window.open("https://waze.com/ul?navigate=yes&q="+encodeURIComponent(D.prox[+b.dataset.wz].mapa),"_blank"));
  el.querySelectorAll("[data-r]").forEach(b=>b.onclick=async()=>{const e=D.prox[+b.dataset.k],r=b.dataset.r;
    try{await acordarDepois(fetch(D.ru,{method:"POST",mode:"no-cors",body:JSON.stringify({email:D.email,ok:true,painel:"✋|"+r+"|"+e.d+"|"+e.tk+"|"+e.e,aparelho:/iPhone|iPad/.test(navigator.userAgent)?"iPhone":/Android/.test(navigator.userAgent)?"Celular Android":"Outro"})}));
      loc[e.d+"|"+e.e]=r==="vou"?"sim":"nao";try{localStorage.setItem("rsvp_"+D.email,JSON.stringify(loc))}catch(_){}proxDesenha();
      alert(r==="vou"?"✅ Presença confirmada! Valeu, irmão. 🤝":"👍 Resposta registrada. Valeu por avisar!")}catch(_){alert("Não consegui enviar agora. Tente de novo em instantes.")}})}
proxDesenha();if(location.hash==="#eventos"){const x=document.getElementById("eventos");if(x){x.open=true;x.scrollIntoView({behavior:"smooth"})}}
// ---- 🎂 próximos aniversários e os parabéns (WhatsApp + notificação para o aniversariante) ----
(function(){const el=document.getElementById("anivL");if(!el)return;
  const FR=["Parabéns, {n}! 🎂 Que Deus abençoe o seu novo ano com muita saúde, estrada livre e vento no rosto. Insanos sempre! 💀🏍️",
    "Feliz aniversário, irmão {n}! 🎉 Muitos anos de vida, de estrada e de irmandade. Conte sempre com a gente. 🤝",
    "{n}, parabéns pelo seu dia! 🎂 Que não falte gasolina no tanque, saúde no corpo e irmãos do seu lado. Insanos M.C.! 💀",
    "Salve, {n}! 🎉 Mais um ano rodando com a gente. Que venham muitas viagens, muitas histórias e muita paz. Parabéns, irmão! 🏍️",
    "Parabéns, {n}! 🎂 Hoje o asfalto é todo seu. Que o novo ano traga saúde, prosperidade e muitas estradas ao lado da irmandade. 🤝",
    "Feliz aniversário, {n}! 💀 Que a vida te dê sempre motivo pra acelerar e irmãos pra dividir o caminho. Tamo junto! 🏍️",
    "{n}, irmão, parabéns! 🎉 Que Deus te guarde em cada quilômetro e te dê muitos anos de vida e de Insanos. Forte abraço! 🤝",
    "Parabéns pelo seu dia, {n}! 🎂 Saúde, paz, família feliz e moto na estrada. O resto a gente resolve junto! 💀",
    "Hoje é dia de comemorar, {n}! 🎉 Muitas felicidades, muita estrada e muita irmandade. Insanos sempre! 🏍️",
    "Feliz aniversário, {n}! 🎂 Que o próximo ano seja de conquistas, saúde e muitos rolês com os irmãos. Parabéns! 🤝",
    "{n}, parabéns! 💀 Mais um ano de estrada e de lealdade. Que nunca falte vento no rosto e irmão do lado. 🏍️",
    "Parabéns, irmão {n}! 🎉 Que o seu novo ano seja leve como uma estrada vazia num domingo de sol. Felicidades! 🎂"];
  let ult=-1;const frase=n=>{let k;do{k=Math.floor(Math.random()*FR.length)}while(FR.length>1&&k===ult);ult=k;return FR[k].split("{n}").join(n)};
  const hoje=new Date();hoje.setHours(0,0,0,0);
  const dias=md=>{const [mm,dd]=md.split("-").map(Number);let d=new Date(hoje.getFullYear(),mm-1,dd);if(d<hoje)d=new Date(hoje.getFullYear()+1,mm-1,dd);return Math.round((d-hoje)/864e5)};
  const L=(D.anivs||[]).map(a=>({...a,d:dias(a.md)})).sort((a,b)=>a.d-b.d).slice(0,5);
  {const h=L.filter(a=>a.d===0);   // aniversário de irmão hoje: o app fica dourado, com a faixa de comemoração
    if(h.length){document.head.insertAdjacentHTML("beforeend",`<style>body.tema-aniv .fundo-video{filter:sepia(.85) saturate(2.3) hue-rotate(-8deg) brightness(.95)}
body.tema-aniv::after{content:"";position:fixed;inset:0;pointer-events:none;z-index:1;background:radial-gradient(120% 55% at 50% 0%,rgba(233,185,73,.16),transparent 70%)}
body.tema-aniv .painel,body.tema-aniv .card,body.tema-aniv .q3,body.tema-aniv .tri button{border-color:rgba(233,185,73,.55)!important}
body.tema-aniv h1,body.tema-aniv .painel h2,body.tema-aniv .card h2{color:#F1CF6E}
.aniv-faixa,body.tema-aniv .aniv-alerta{background:linear-gradient(90deg,#B8892A,#F1CF6E)!important;color:#1a1206!important;border:1px solid #F1CF6E!important;border-radius:14px;padding:12px 14px;margin:0 0 14px;font-size:15.5px;font-weight:600;box-shadow:0 6px 24px rgba(233,185,73,.35)}
.aniv-faixa b,body.tema-aniv .aniv-alerta b{color:#1a1206!important}
.aniv-faixa button{display:block;width:100%;margin-top:10px;padding:11px;border-radius:12px;border:0;background:#1a1206;color:#F1CF6E;font:inherit;font-weight:800}</style>`);document.body.classList.add("tema-aniv");
      const eu=h.find(a=>a.n===D.n),out=h.filter(a=>a.n!==D.n),fx=document.createElement("div");fx.className="aniv-faixa";
      fx.innerHTML=(eu?`🎂 <b>Feliz aniversário, ${es(D.n)}!</b> Todo o Insanos M.C. te deseja um ano de muita saúde, estrada e irmandade. 💀`:"")+
        (out.length?`${eu?"<br><br>":""}🎉 <b>Hoje é aniversário de ${out.map(a=>es(a.n)).join(" e ")}!</b>${out.some(a=>a.tel)?'<button type="button" id="anivFxBt">🎉 Mandar os parabéns</button>':""}`:"");
      const alvo=document.getElementById("enqC")||document.getElementById("tri")||el.closest(".card");if(alvo)alvo.before(fx);
      const fb=document.getElementById("anivFxBt");if(fb)fb.onclick=()=>{const tb=document.querySelector('#tri [data-tri="anivC"]'),d=document.getElementById("anivC");
        if(tb&&!d.open)tb.click();else if(d)d.open=true;setTimeout(()=>{const pb=document.querySelector("#anivL [data-pb]");if(pb)pb.click()},120)}}}
  el.innerHTML=L.map((a,k)=>{const eu=a.n===D.n,dm=a.md.slice(3,5)+"/"+a.md.slice(0,2);
    const qd=a.d===0?(eu?"🎉 Hoje é o seu dia!":"🎉 É hoje!"):a.d===1?"amanhã":`em ${a.d} dias`;
    return `<div class="ev" style="flex-wrap:wrap${a.d===0?";background:rgba(233,185,73,.10);border-radius:12px;padding:10px":""}"><span>${dm}</span><b style="flex:1;font-weight:600">${es(a.n)}${eu?" (você)":""}</b>
      <small class="nota" style="${a.d===0?"color:#E9B949;font-weight:700":""}">${qd}</small>
      ${a.d===0&&!eu&&a.tel?`<div class="qbts" style="width:100%;padding-left:54px;margin-top:6px"><button class="qz" data-pb="${k}">🎉 Mandar os parabéns</button></div>`:""}</div>`}).join("");
  {const r=document.getElementById("anivRes"),h=L.filter(a=>a.d===0),p=L.find(a=>a.d>0);
    if(r){if(h.length){r.textContent="🎉 Hoje: "+h.map(a=>a.n===D.n?"você!":a.n).join(", ");r.style.color="#E9B949";r.style.fontWeight="700"}
      else if(p)r.textContent=`Próximo: ${p.n} · ${p.d===1?"amanhã":"em "+p.d+" dias"}`}}
  el.onclick=e=>{const b=e.target.closest("[data-pb]");if(!b)return;const a=L[+b.dataset.pb];
    const v=document.createElement("div");v.className="sosm-veu";v.innerHTML=`<div class="sosm-cx" style="border-top-color:#E9B949" role="dialog" aria-modal="true">
      <h3 style="margin:0 0 6px">🎉 Parabéns para ${es(a.n)}</h3><p class="nota" style="margin:0 0 10px">Edite como quiser, ou toque em Outra frase.</p>
      <textarea id="pbTx" rows="6" maxlength="400" style="width:100%;box-sizing:border-box;border-radius:12px;padding:10px;background:rgba(0,0,0,.4);color:#ECE7DC;border:1px solid rgba(236,231,220,.25);font:inherit;font-size:15px"></textarea>
      <div class="qbts" style="margin-top:10px"><button id="pbOutra">🔄 Outra frase</button><button class="qz" id="pbZap">📲 Mandar pelo WhatsApp</button></div>
      <p class="nota" style="margin:10px 0 0;font-size:12.5px">${es(a.n)} também recebe a sua mensagem como notificação no app.</p>
      <div class="qbts" style="margin-top:8px"><button id="pbFechar">Fechar</button></div></div>`;
    document.body.appendChild(v);const tx=v.querySelector("#pbTx");tx.value=frase(a.n);
    const fecha=()=>v.remove();v.onclick=ev=>{if(ev.target===v||ev.target.id==="pbFechar")fecha()};
    v.querySelector("#pbOutra").onclick=()=>{tx.value=frase(a.n)};
    v.querySelector("#pbZap").onclick=async()=>{const t=tx.value.trim();if(!t){tx.focus();return}
      const q=typeof sosQuem==="function"?sosQuem():null;
      if(q&&window.SOS_URL){try{await Promise.race([fetch(window.SOS_URL,{method:"POST",headers:{"Content-Type":"text/plain;charset=utf-8"},
        body:JSON.stringify({tipo:"parabens",email:q.email,s:q.s,para:a.n,texto:t})}),new Promise(r=>setTimeout(r,3000))])}catch(_){}}
      fecha();location.href=`https://api.whatsapp.com/send/?phone=${a.tel}&text=${encodeURIComponent(t)}`}}})();
// ---- ✨ efeito holográfico das caveiras: um brilho prateado que corre conforme a inclinação do celular ----
(function(){if(window.__holo)return;window.__holo=1;const R=document.documentElement,quieto=window.matchMedia&&matchMedia("(prefers-reduced-motion: reduce)").matches;
  let ult=0,ax=0,ay=0;const poe=(x,y)=>{R.style.setProperty("--hx",(50+y*1.6).toFixed(1)+"%");R.style.setProperty("--hy",(50+x*1.6).toFixed(1)+"%");
    R.style.setProperty("--sx",(y*.32).toFixed(1)+"px");R.style.setProperty("--sy",(x*.32).toFixed(1)+"px")};   // a sombra colorida se desloca com a inclinaçãopoe(0,0);
  if(quieto||document.documentElement.classList.contains("eco"))return;
  const ori=e=>{if(e.beta==null)return;ult=performance.now();ax=Math.max(-30,Math.min(30,(e.beta||0)-45));ay=Math.max(-30,Math.min(30,e.gamma||0))};
  addEventListener("deviceorientation",ori);
  document.addEventListener("click",e=>{if(!e.target.closest(".cav,.cav-dash,.hl-cobre,#roda,#pagCav"))return;   // iPhone: pede a permissão ao tocar numa caveira
    if(typeof DeviceOrientationEvent!=="undefined"&&typeof DeviceOrientationEvent.requestPermission==="function")DeviceOrientationEvent.requestPermission().catch(()=>{})},true);
  let t0=performance.now(),q=0;const laco=t=>{if(!document.hidden&&(++q%2===0)){if(t-ult<2500)poe(ax,ay);else{const k=(t-t0)/1000;poe(Math.sin(k*.7)*16,Math.sin(k*.5)*24)}}requestAnimationFrame(laco)};
  requestAnimationFrame(laco)})();
function corGlow(v){return v==null?"transparent":v>=90?"rgba(233,185,73,.85)":v>75?"rgba(143,217,163,.65)":v>=50?"rgba(242,139,130,.65)":"rgba(185,20,20,.95)"}   // a sombra colorida da caveira (a cor da faixa)
// camada do brilho por cima de uma caveira desenhada (svg ou canvas), com o formato da caveira
function holoCobrir(el,mask){if(!el||el.__hl||!el.parentNode)return;   // envelope do mesmo tamanho da caveira: a camada do brilho nunca sai do lugar
  const w=document.createElement("span");w.className="hl-env";el.parentNode.insertBefore(w,el);w.appendChild(el);
  const s=document.createElement("span");s.className="hl-cobre"+(el.tagName==="CANVAS"?" cv":"");s.style.setProperty("--m",`url('${mask}')`);w.appendChild(s);el.__hl=s;
  // se a caveira já foi pintada antes do envelope existir: leva a cor da sombra, o dourado e o recorte do brilho
  const g=el.style.getPropertyValue("--g");if(g){w.style.setProperty("--g",g);el.style.removeProperty("--g")}
  if(el.classList.contains("trofeu")){w.classList.add("trofeu");el.classList.remove("trofeu")}
  if(el.dataset.clip)s.style.clipPath=el.dataset.clip}

// ---- 🏍️ Minha moto (só neste celular: documentos e abastecimentos; nada vai para a planilha, a nuvem ou o comando) ----
(function(){if(document.getElementById("motoV"))return;const K="moto_"+(D.email||"eu");
  const ler=()=>{try{return JSON.parse(localStorage.getItem(K)||"{}")}catch(_){return {}}},grava=o=>{try{localStorage.setItem(K,JSON.stringify(o))}catch(_){alert("Não consegui salvar neste celular.")}};
  let M=ler();M.docs=M.docs||{};M.abs=M.abs||[];
  const num=v=>{const n=parseFloat(String(v||"").trim().split(".").join("").split(",").join("."));return isFinite(n)?n:NaN};   // aceita 6,29 e 1.234,56
  const num2=v=>{const t=String(v||"").trim();return t.indexOf(",")>=0?num(t):parseFloat(t)};
  const R$=v=>"R$ "+Number(v).toFixed(2).split(".").join(","),dec=(v,c)=>Number(v).toFixed(c).split(".").join(",");
  const br=d=>d?d.slice(8,10)+"/"+d.slice(5,7)+"/"+d.slice(0,4):"";
  const hoje=new Date();hoje.setHours(0,0,0,0);const faltam=d=>Math.round((new Date(d+"T00:00:00")-hoje)/864e5);
  const DOCS=[["cnh","🪪 CNH"],["lic","📄 Licenciamento"]];
  document.head.insertAdjacentHTML("beforeend",`<style>
    body.aba-moto #app{display:none}body:not(.aba-moto) #motoV{display:none}
    #motoV .card h2{margin:0 0 6px}.priv{border-color:rgba(76,175,106,.55)!important;background:rgba(76,175,106,.08)!important}
    .mf{display:grid;grid-template-columns:1fr 1fr;gap:10px}.mf label{font-size:13px;color:#BDB7AA;display:block}
    .mf input,.mf select{display:block;width:100%;box-sizing:border-box;margin-top:4px;padding:10px;border-radius:10px;background:rgba(0,0,0,.4);color:#ECE7DC;border:1px solid rgba(236,231,220,.22);font:inherit;font-size:16px}
    .mf .inteiro{grid-column:1/-1}.litros{grid-column:1/-1;font-size:15px;color:#E9B949;font-weight:700}
    .doc{display:flex;align-items:center;gap:10px;padding:10px 0;border-bottom:1px solid rgba(236,231,220,.08)}.doc b{flex:1}.doc input{max-width:170px}
    .doc small{min-width:96px;text-align:right;font-weight:700}
    .tab-m{width:100%;border-collapse:collapse;font-size:14px}.tab-m th,.tab-m td{padding:8px 4px;border-bottom:1px solid rgba(236,231,220,.08);text-align:right}
    .tab-m th:first-child,.tab-m td:first-child{text-align:left}.tab-m th{color:#BDB7AA;font-weight:600;font-size:12.5px}
    .tab-m button{background:none;border:0;color:#8f887d;font-size:15px;padding:2px 4px}
    .res-m{display:grid;grid-template-columns:repeat(3,1fr);gap:8px;margin:4px 0 12px}.res-m div{background:rgba(236,231,220,.05);border-radius:12px;padding:10px;text-align:center}
    .res-m b{display:block;font-size:20px;color:#E9B949}.res-m span{font-size:12px;color:#BDB7AA}
    .nav-m{position:fixed;left:0;right:0;bottom:0;z-index:50;display:grid;grid-template-columns:repeat(5,1fr);align-items:end;gap:4px;padding:6px 8px calc(8px + env(safe-area-inset-bottom,0px));
      background:rgba(13,11,9,.94);backdrop-filter:blur(12px);border-top:1px solid rgba(236,231,220,.12)}
    .nav-m button{display:flex;flex-direction:column;align-items:center;gap:2px;padding:6px 2px;border-radius:14px;border:0;background:none;color:#BDB7AA;font:inherit;font-size:11.5px;font-weight:700}
    .nav-m button i{font-style:normal;font-size:22px;line-height:1}
    .nav-m button[aria-pressed="true"]{background:rgba(236,231,220,.12);color:#ECE7DC}
    .nav-m .nav-sos{transform:translateY(-14px);background:#c62828;color:#fff;border-radius:50%;width:66px;height:66px;margin:0 auto;justify-content:center;box-shadow:0 0 0 4px rgba(13,11,9,.95),0 6px 18px rgba(198,40,40,.55)}
    .nav-m .nav-sos i{font-size:24px}.nav-m .nav-sos span{font-size:11px;font-weight:900;letter-spacing:.5px}
    body{padding-bottom:96px}
    #app .cart-bts,#app .sos-meu,#app .msg-meu{display:none!important}
    #ajIco{width:44px;height:44px;border-radius:50%;border:1px solid rgba(236,231,220,.25);background:rgba(236,231,220,.06);font-size:19px;flex:none}
#msgIco{margin-left:auto;width:44px;height:44px;border-radius:50%;border:1px solid rgba(236,231,220,.25);background:rgba(236,231,220,.06);font-size:20px;flex:none}
    .tri button{aspect-ratio:auto!important;min-height:0;padding:8px 9px!important;gap:2px!important;border-radius:14px!important}
    .tri .ti{font-size:16px!important;display:inline}.tri .tt{font-size:10.5px!important}.tri .tp{font-size:12.5px!important;-webkit-line-clamp:1!important}.tri .ts{font-size:11px!important}
    .bols{display:grid;grid-template-columns:repeat(4,1fr);gap:8px;margin:0 0 14px}
    .priv-l{margin:0 0 12px;font-size:12.5px;color:#8f887d}.priv-l summary{cursor:pointer;list-style:none;text-align:center}.priv-l summary::-webkit-details-marker{display:none}.priv-l p{margin:6px 0 0;color:#BDB7AA;line-height:1.4}
    .emg-env{display:flex;gap:10px;align-items:flex-start;margin-top:12px;font-size:13.5px;color:#BDB7AA}.emg-env input{width:20px;height:20px;flex:none;margin-top:1px}
    .bol{position:relative;display:flex;flex-direction:column;align-items:center;gap:4px;padding:12px 4px 10px;border-radius:20px;border:1px solid rgba(236,231,220,.18);background:rgba(236,231,220,.05);color:inherit;font:inherit;text-align:center}
    .bol i{font-style:normal;font-size:24px;width:48px;height:48px;border-radius:50%;display:flex;align-items:center;justify-content:center;background:rgba(233,185,73,.12);border:1px solid rgba(233,185,73,.4)}
    .bol b{font-size:12.5px}.bol small{font-size:11px;color:#BDB7AA;line-height:1.25}.bol small.al{color:#E9B949;font-weight:700}
    .bol em{position:absolute;top:10px;right:14px;width:11px;height:11px;border-radius:50%;background:#E5484D;box-shadow:0 0 0 2px #1b1714}
    .bol.on{border-color:#E9B949;background:rgba(233,185,73,.1)}
    .moto-ft{border-radius:16px;overflow:hidden;aspect-ratio:16/9;background:rgba(236,231,220,.05);border:1px dashed rgba(236,231,220,.3);display:flex;align-items:center;justify-content:center;cursor:pointer}
    .moto-ft img{width:100%;height:100%;object-fit:cover;display:block}.moto-ft span{display:flex;flex-direction:column;align-items:center;font-size:30px;gap:6px}.moto-ft small{font-size:13.5px;color:#BDB7AA}
    .man{padding:12px 0;border-bottom:1px solid rgba(236,231,220,.08)}.man-t{display:flex;justify-content:space-between;gap:10px;margin-bottom:6px}.man-t small{font-weight:700;font-size:13px;text-align:right}.man-p{display:block;color:#8f887d;font-size:12.5px;margin-top:4px}
    .nav-m button{position:relative}.nav-bdg{position:absolute;top:2px;right:calc(50% - 26px);min-width:18px;height:18px;padding:0 5px;border-radius:9px;background:#E5484D;color:#fff;font-size:11px;font-weight:800;line-height:18px;text-align:center;box-shadow:0 0 0 2px rgba(13,11,9,.95)}.nav-bdg[hidden]{display:none}
    .aviso-doc{background:rgba(155,17,17,.18);border:1px solid rgba(220,80,80,.6);border-radius:14px;padding:12px 14px;margin:0 0 14px;font-weight:600}</style>`);
  const v=document.createElement("div");v.id="motoV";v.className="w";document.getElementById("app").after(v);
  // menu embaixo: Início · Carteirinha · SOS (no centro) · Conquistas · Minha moto (o topo fica enxuto e a caveira sobe)
  const nav=document.createElement("div");nav.className="nav-m";
  nav.innerHTML=`<button type="button" data-aba="ini" aria-pressed="true"><i>🏠</i><span>Início</span></button><button type="button" data-ac="cart"><i>🪪</i><span>Carteirinha</span></button>
    <button type="button" data-ac="sos" class="nav-sos" aria-label="SOS: preciso de ajuda"><i>🆘</i><span>SOS</span></button>
    <button type="button" data-ac="conq"><i>🏆</i><span>Conquistas</span></button><button type="button" data-aba="moto" aria-pressed="false"><i>🏍️</i><span>Minha moto</span></button>`;
  document.body.appendChild(nav);
  const vaiInicio=()=>{const b=nav.querySelector('[data-aba="ini"]');nav.querySelectorAll("[data-aba]").forEach(x=>x.setAttribute("aria-pressed",x===b));document.body.classList.remove("aba-moto")};
  nav.onclick=e=>{const b=e.target.closest("button");if(!b)return;
    if(b.dataset.ac==="sos"){if(typeof abrirSOS==="function")abrirSOS();return}
    if(b.dataset.ac==="cart"){const c=document.getElementById("cartBt");if(c)c.click();return}
    if(b.dataset.ac==="conq"){vaiInicio();const c=document.getElementById("conqBt"),bx=document.getElementById("conqBox");if(c&&(!bx||bx.hidden))c.click();setTimeout(()=>{if(bx)bx.scrollIntoView({behavior:"smooth",block:"start"})},80);return}
    nav.querySelectorAll("[data-aba]").forEach(x=>x.setAttribute("aria-pressed",x===b));
    document.body.classList.toggle("aba-moto",b.dataset.aba==="moto");scrollTo(0,0);if(b.dataset.aba==="moto")desenha()};
  // ✉️ Mensagens: um ícone ao lado do nome
  {const h=document.querySelector("#app h1");if(h&&!document.getElementById("msgIco")){const m=document.createElement("button");m.id="msgIco";m.type="button";m.setAttribute("aria-label","Mensagens: falar com o comando");m.textContent="✉️";
    m.onclick=()=>{const o=document.querySelector(".msg-meu");if(o)o.click()};h.appendChild(m);
    const aj=document.createElement("button");aj.id="ajIco";aj.type="button";aj.setAttribute("aria-label","Ajustes do app");aj.textContent="⚙️";aj.onclick=()=>ajustesAbrir();h.appendChild(aj)}}
  // atalhos do ícone do app: ao abrir e também com o app já aberto (o Android só muda o endereço)
  const atalho=()=>{const h=location.hash;if(!/^#(moto|eventos|sos-abrir)$/.test(h))return;try{history.replaceState(null,"",location.pathname)}catch(_){}
    if(h==="#moto"){nav.querySelector('[data-aba="moto"]').click();return}
    if(h==="#sos-abrir"){if(typeof abrirSOS==="function")abrirSOS();return}
    vaiInicio();const tb=document.querySelector('#tri [data-tri="eventos"]'),d=document.getElementById("eventos");
    if(tb&&d&&!d.open)tb.click();else if(d){d.open=true;d.scrollIntoView({behavior:"smooth"})}};
  setTimeout(atalho,250);addEventListener("hashchange",atalho);
  function consumo(){const L=[...M.abs].sort((a,b)=>a.km-b.km);L.forEach((a,i)=>{a.kml=i>0&&a.km>L[i-1].km&&a.l>0?(a.km-L[i-1].km)/a.l:null});return L}
  // ---- a aba em três bolinhas: Documentos, Abastecimento e Manutenção (por km) ----
  const MAN=[["rev","🔧 Revisão"],["oleo","🛢️ Troca de óleo"],["pneu","🛞 Pneus"],["rel","⛓️ Relação / corrente"]];
  let aberta=null;
  const kmAt=()=>M.abs.length?Math.max(...M.abs.map(a=>a.km)):null;
  const prox=k=>{const m=M.man||{};if(k==="rev")return +m.rev_km||null;const u=+m[k+"_ult"],n=+m[k+"_int"];return u>0&&n>0?u+n:null};
  const falta=k=>{const p=prox(k),a=kmAt();return p&&a!=null?p-a:null};
  const nKm=n=>Number(n).toLocaleString("pt-BR");
  function atualizaBadge(){let n=0;DOCS.forEach(([k])=>{const d=M.docs[k];if(d&&faltam(d)<=30)n++});MAN.forEach(([k])=>{const f=falta(k);if(f!=null&&f<=300)n++});
    const b=nav.querySelector('[data-aba="moto"]');if(!b)return;let s=b.querySelector(".nav-bdg");if(!s){s=document.createElement("b");s.className="nav-bdg";b.appendChild(s)}
    s.textContent=n;s.hidden=!n;b.setAttribute("aria-label","Minha moto"+(n?`: ${n} ${n===1?"item precisa":"itens precisam"} de atenção`:""));return n}
  function desenha(){const L=consumo(),comK=L.filter(a=>a.kml!=null),mes=new Date().toISOString().slice(0,7),doMes=L.filter(a=>a.d.slice(0,7)===mes);
    const media=comK.length?comK.reduce((s,a)=>s+a.kml*a.l,0)/comK.reduce((s,a)=>s+a.l,0):null;M.man=M.man||{};M.moto=M.moto||{};
    const docAl=DOCS.filter(([k])=>M.docs[k]&&faltam(M.docs[k])<=30).length,docOk=DOCS.filter(([k])=>M.docs[k]).length;
    M.emg=M.emg||{};const E=M.emg,emgOk=!!(E.sangue||E.alerg||(E.c||[]).some(c=>c&&c.t));
    const fs=MAN.map(([k])=>falta(k)).filter(f=>f!=null),manMin=fs.length?Math.min(...fs):null,manAl=fs.filter(f=>f<=300).length,ka=kmAt(),mf=motoFotoLer();
    const bol=(id,ic,r,st,al)=>`<button type="button" class="bol${aberta===id?" on":""}" data-bol="${id}" aria-expanded="${aberta===id}"><i>${ic}</i><b>${r}</b><small${al?' class="al"':""}>${st}</small>${al?'<em></em>':""}</button>`;
    v.innerHTML=`<details class="priv-l"><summary>🔒 Só neste celular · ninguém mais vê</summary><p>Tudo o que você coloca aqui fica <b>só neste celular</b>: não vai para a planilha, nem para a nuvem, nem para o comando. A única exceção é a <b>ficha de emergência</b>, que vai junto com o SOS se você deixar marcado. Se trocar de celular ou desinstalar o app, use o 💾 backup lá embaixo.</p></details>
    <div class="card moto-c"><div class="moto-ft" id="motoFt">${mf?`<img src="${mf}" alt="A minha moto">`:`<span>📸<small>Pôr a foto da moto</small></span>`}</div>
      <div class="mf" style="margin-top:12px"><label>Modelo<input type="text" id="mtMod" maxlength="40" placeholder="Ex.: Harley-Davidson Fat Boy" value="${es(M.moto.modelo||"")}"></label>
      <label>Ano<input type="text" inputmode="numeric" id="mtAno" maxlength="4" placeholder="Ex.: 2019" value="${es(M.moto.ano||"")}"></label></div>
      ${mf?'<div class="qbts" style="margin-top:8px"><button type="button" id="mtTroca">🖼️ Trocar a foto</button><button type="button" id="mtRem">🗑️ Remover</button></div>':""}
      <p class="nota" style="margin:8px 0 0">A foto da moto aparece também na sua carteirinha.</p></div>
    <div class="bols">${bol("doc","📄","Documentos",docAl?`⚠️ ${docAl} vence logo`:docOk?"Em dia":"Cadastrar",docAl)}${bol("abs","⛽","Abastecimento",media?dec(media,1)+" km/l":L.length?L.length+" registro":"Registrar",0)}${bol("man","🔧","Manutenção",manMin==null?(ka==null?"Abasteça 1ª vez":"Cadastrar"):manMin<0?`⚠️ passou ${nKm(-manMin)} km`:`faltam ${nKm(manMin)} km`,manAl)}${bol("emg","🆘","Emergência",emgOk?"Ficha pronta":"Preencher",0)}</div>
    <div id="bolC"></div>
    <div class="card"><h2>💾 Backup</h2><p class="nota" style="margin:0 0 8px">Guarde uma cópia (por exemplo, mandando para você mesmo no WhatsApp) para não perder os dados ao trocar de celular.</p>
      <div class="qbts"><button id="bkC">📋 Copiar backup</button><button id="bkR">📥 Restaurar backup</button></div></div>
    <input type="file" accept="image/*" hidden id="mtIn">`;
    const C=v.querySelector("#bolC");
    if(aberta==="doc")C.innerHTML=`<div class="card"><h2>📄 Documentos</h2><p class="nota" style="margin:0 0 6px">Coloque os vencimentos: o ícone da Minha moto avisa 30 dias antes.</p>
      ${DOCS.map(([k,r])=>{const d=M.docs[k]||"",f=d?faltam(d):null,cr=f==null?"#BDB7AA":f<0?"#E26D6B":f<=30?"#E9B949":"#8FD9A3";
        return `<div class="doc"><b>${r}</b><input type="date" data-doc="${k}" value="${d}"><small style="color:${cr}">${f==null?"–":f<0?"vencido":f===0?"vence hoje":"em "+f+" dias"}</small></div>`}).join("")}</div>`;
    if(aberta==="abs")C.innerHTML=`<div class="card"><h2>⛽ Novo abastecimento</h2><div class="mf">
      <label>Data<input type="date" id="abD" value="${new Date().toISOString().slice(0,10)}"></label>
      <label>Odômetro (km)<input type="text" inputmode="numeric" id="abK" placeholder="${ka!=null?"Último: "+nKm(ka):"Ex.: 23450"}"></label>
      <label>Valor total (R$)<input type="text" inputmode="decimal" id="abV" placeholder="Ex.: 85,00"></label>
      <label>Preço por litro (R$)<input type="text" inputmode="decimal" id="abP" placeholder="Ex.: 6,29"></label>
      <label class="inteiro">Combustível<select id="abT"><option>Gasolina comum</option><option>Gasolina aditivada</option><option>Gasolina premium</option><option>Etanol</option></select></label>
      <div class="litros" id="abL">Litros: –</div></div>
      <div class="qbts" style="margin-top:10px"><button class="qz" id="abS">💾 Registrar abastecimento</button></div></div>
      <div class="card"><h2>📊 Consumo</h2><div class="res-m"><div><b>${media?dec(media,1):"–"}</b><span>km/l (média)</span></div><div><b>${doMes.length?R$(doMes.reduce((s,a)=>s+a.v,0)):"–"}</b><span>gasto no mês</span></div><div><b>${L.length}</b><span>abastecimentos</span></div></div>
      ${L.length?`<table class="tab-m"><thead><tr><th>Data</th><th>Valor</th><th>Litros</th><th>km/l</th><th></th></tr></thead><tbody>${[...L].reverse().map(a=>`<tr><td>${br(a.d)}<br><small style="color:#8f887d">${nKm(a.km)} km · ${a.t||""}</small></td><td>${R$(a.v)}</td><td>${dec(a.l,2)}</td><td>${a.kml==null?'<small style="color:#8f887d">Primeiro registro</small>':"<b>"+dec(a.kml,1)+"</b>"}</td><td><button data-del="${a.id}" aria-label="Apagar">🗑️</button></td></tr>`).join("")}</tbody></table>`:'<p class="nota" style="margin:0">Nenhum abastecimento ainda. O primeiro registro serve de ponto de partida para o km/l.</p>'}</div>`;
    if(aberta==="emg"){const c=E.c||[],ct=k=>`<div class="mf"><label class="inteiro">Contato ${k+1}${k?" (opcional)":""}<input type="text" data-emc="${k}n" maxlength="40" placeholder="Nome" value="${es((c[k]||{}).n||"")}"></label>
        <label>Parentesco<input type="text" data-emc="${k}p" maxlength="20" placeholder="Ex.: esposa" value="${es((c[k]||{}).p||"")}"></label><label>Telefone<input type="tel" data-emc="${k}t" maxlength="20" placeholder="(11) 99999-0000" value="${es((c[k]||{}).t||"")}"></label></div>`;
      C.innerHTML=`<div class="card"><h2>🆘 Ficha de emergência</h2><p class="nota" style="margin:0 0 10px">Para quem for te socorrer. Mostre em tela cheia ou deixe ir junto com o SOS.</p>
      <div class="mf"><label>Tipo sanguíneo<select data-emg="sangue">${["","A+","A−","B+","B−","AB+","AB−","O+","O−","Não sei"].map(x=>`<option value="${x}"${(E.sangue||"")===x?" selected":""}>${x||"Escolha…"}</option>`).join("")}</select></label>
        <label>Plano de saúde<input type="text" data-emg="plano" maxlength="60" placeholder="Ex.: Unimed" value="${es(E.plano||"")}"></label>
        <label class="inteiro">⚠️ Alergias<input type="text" data-emg="alerg" maxlength="120" placeholder="Ex.: Dipirona" value="${es(E.alerg||"")}"></label>
        <label class="inteiro">💊 Remédios de uso contínuo<input type="text" data-emg="rem" maxlength="120" placeholder="Ex.: Losartana" value="${es(E.rem||"")}"></label>
        <label class="inteiro">ℹ️ Outras informações<input type="text" data-emg="obs" maxlength="120" placeholder="Ex.: diabético, usa marca-passo" value="${es(E.obs||"")}"></label></div>
      <h3 style="margin:14px 0 6px;font-size:16px">📞 Contatos de emergência</h3>${ct(0)}${ct(1)}
      <label class="emg-env"><input type="checkbox" id="emgEnv"${E.enviar===false?"":" checked"}> <span>Enviar a ficha junto com o <b>SOS</b>: vai na notificação dos irmãos e na <b>mensagem do WhatsApp</b> (tipo sanguíneo, alergias, remédios e plano)</span></label>
      <label class="emg-env"><input type="checkbox" id="emgCt"${E.enviarCt===false?"":" checked"}> <span>Incluir o <b>contato de emergência</b> (nome e telefone). Desmarcando, o contato não vai, e o resto da ficha continua indo</span></label>
      <div class="qbts" style="margin-top:12px"><button class="qz" id="emgVer">🩺 Ver a ficha em tela cheia</button></div></div>`}
    if(aberta==="man")C.innerHTML=`<div class="card"><h2>🔧 Manutenção</h2><p class="nota" style="margin:0 0 10px">Tudo por km. O app usa o odômetro do último abastecimento${ka!=null?` (<b>${nKm(ka)} km</b>)`:""} para mostrar quanto falta, e o ícone da Minha moto avisa quando faltar menos de 300 km.</p>
      ${MAN.map(([k,r])=>{const p=prox(k),f=falta(k),cr=f==null?"#BDB7AA":f<0?"#E26D6B":f<=300?"#E9B949":"#8FD9A3",m=M.man;
        return `<div class="man"><div class="man-t"><b>${r}</b><small style="color:${cr}">${p==null?"–":f==null?"próxima aos "+nKm(p)+" km":f<0?`passou ${nKm(-f)} km`:`faltam ${nKm(f)} km`}</small></div>
        ${k==="rev"?`<div class="mf"><label class="inteiro">Próxima revisão aos (km)<input type="text" inputmode="numeric" data-man="rev_km" value="${m.rev_km||""}" placeholder="Ex.: 25000"></label></div>`
        :`<div class="mf"><label>Última troca aos (km)<input type="text" inputmode="numeric" data-man="${k}_ult" value="${m[k+"_ult"]||""}" placeholder="Ex.: 21000"></label>
          <label>Trocar a cada (km)<input type="text" inputmode="numeric" data-man="${k}_int" value="${m[k+"_int"]||""}" placeholder="${k==="oleo"?"Ex.: 3000":k==="pneu"?"Ex.: 15000":"Ex.: 20000"}"></label></div>`}
        ${p!=null?`<small class="man-p">Próxima: aos ${nKm(p)} km</small>`:""}</div>`}).join("")}</div>`;
    // ações
    v.querySelectorAll("[data-bol]").forEach(b=>b.onclick=()=>{aberta=aberta===b.dataset.bol?null:b.dataset.bol;desenha();const c=document.getElementById("bolC");if(aberta&&c)c.scrollIntoView({behavior:"smooth",block:"start"})});
    const mtIn=v.querySelector("#mtIn"),pega=()=>mtIn.click();v.querySelector("#motoFt").onclick=pega;const tr=v.querySelector("#mtTroca");if(tr)tr.onclick=pega;
    mtIn.onchange=async()=>{const a=mtIn.files&&mtIn.files[0];if(!a)return;try{localStorage.setItem("motofoto_"+D.email,await cartImgProcessar(a));desenha()}catch(_){alert("Não consegui usar essa foto. Tente outra.")}};
    const rm=v.querySelector("#mtRem");if(rm)rm.onclick=()=>{if(!confirm("Remover a foto da moto?"))return;try{localStorage.removeItem("motofoto_"+D.email)}catch(_){}desenha()};
    v.querySelector("#mtMod").onchange=e=>{M.moto.modelo=e.target.value.trim();grava(M)};v.querySelector("#mtAno").onchange=e=>{M.moto.ano=e.target.value.trim();grava(M)};
    v.querySelectorAll("[data-man]").forEach(i=>i.onchange=()=>{const n=parseInt(String(i.value).split(".").join(""),10);if(n>0)M.man[i.dataset.man]=n;else delete M.man[i.dataset.man];grava(M);desenha();atualizaBadge()});
    v.querySelectorAll("[data-doc]").forEach(i=>i.onchange=()=>{M.docs[i.dataset.doc]=i.value;grava(M);desenha();atualizaBadge()});
    if(aberta==="emg"){M.emg.nome=D.n;
      v.querySelectorAll("[data-emg]").forEach(i=>i.onchange=()=>{const t=String(i.value||"").trim();if(t)M.emg[i.dataset.emg]=t;else delete M.emg[i.dataset.emg];grava(M)});
      v.querySelectorAll("[data-emc]").forEach(i=>i.onchange=()=>{const k=+i.dataset.emc[0],f=i.dataset.emc[1];M.emg.c=M.emg.c||[{},{}];M.emg.c[k]=M.emg.c[k]||{};M.emg.c[k][f]=String(i.value||"").trim();grava(M)});
      v.querySelector("#emgEnv").onchange=e=>{M.emg.enviar=e.target.checked;grava(M)};v.querySelector("#emgCt").onchange=e=>{M.emg.enviarCt=e.target.checked;grava(M)};
      v.querySelector("#emgVer").onclick=()=>{grava(M);if(window.fichaMostrar)fichaMostrar(D.email||"eu")}}
    if(aberta==="abs"){const calc=()=>{const vv=num2(v.querySelector("#abV").value),pp=num2(v.querySelector("#abP").value);v.querySelector("#abL").textContent=vv>0&&pp>0?"Litros: "+dec(vv/pp,2)+" L":"Litros: –"};
      ["#abV","#abP"].forEach(q=>v.querySelector(q).addEventListener("input",calc));
      v.querySelector("#abS").onclick=()=>{const d=v.querySelector("#abD").value,km=parseInt(String(v.querySelector("#abK").value).split(".").join(""),10),vv=num2(v.querySelector("#abV").value),pp=num2(v.querySelector("#abP").value);
        if(!d||!(km>0)||!(vv>0)||!(pp>0)){alert("Preencha a data, o odômetro, o valor total e o preço por litro.");return}
        const ult=consumo().slice(-1)[0];if(ult&&km<=ult.km&&!confirm(`O odômetro (${km} km) é menor ou igual ao último registro (${ult.km} km). Registrar mesmo assim?`))return;
        M.abs.push({id:Date.now(),d,km,v:vv,p:pp,l:vv/pp,t:v.querySelector("#abT").value});grava(M);desenha();atualizaBadge();try{navigator.vibrate&&navigator.vibrate(15)}catch(_){}};
      v.querySelectorAll("[data-del]").forEach(b=>b.onclick=()=>{if(!confirm("Apagar este abastecimento?"))return;M.abs=M.abs.filter(a=>String(a.id)!==b.dataset.del);grava(M);desenha();atualizaBadge()})}
    v.querySelector("#bkC").onclick=async()=>{const t="INSANOS-MOTO:"+btoa(unescape(encodeURIComponent(JSON.stringify({...M,foto:fotoLer(),fundo:fundoLer(),cartimg:cartImgLer(),motofoto:motoFotoLer()}))));
      try{await navigator.clipboard.writeText(t);alert("Backup copiado. Cole numa conversa sua no WhatsApp para guardar.")}catch(_){prompt("Copie o backup:",t)}};
    v.querySelector("#bkR").onclick=()=>{const t=prompt("Cole aqui o backup (começa com INSANOS-MOTO:)");if(!t)return;
      try{const o=JSON.parse(decodeURIComponent(escape(atob(t.trim().replace("INSANOS-MOTO:","")))));if(!o||typeof o!=="object")throw 0;M={docs:o.docs||{},abs:o.abs||[],man:o.man||{},moto:o.moto||{}};grava(M);
        try{if(o.foto)localStorage.setItem("foto_"+D.email,o.foto);if(o.fundo)localStorage.setItem("cartfundo_"+D.email,o.fundo);if(o.cartimg)localStorage.setItem("cartimg_"+D.email,o.cartimg);if(o.motofoto)localStorage.setItem("motofoto_"+D.email,o.motofoto)}catch(_){}
        fotoDesenha();desenha();atualizaBadge();alert("Backup restaurado.")}catch(_){alert("Esse texto não parece um backup válido.")}}}
  atualizaBadge()})();


// ---- ⚙️ ajustes do app (só neste celular): modo economia e tamanho da letra ----
function ajustesLer(){try{return JSON.parse(localStorage.getItem("ajustes_app")||"{}")}catch(_){return {}}}
function ajustesAplicar(){const a=ajustesLer(),R=document.documentElement;R.classList.toggle("eco",!!a.eco);R.classList.toggle("letra-g",a.letra==="g");R.classList.toggle("letra-gg",a.letra==="gg");
  if(a.eco)document.querySelectorAll("video.fundo-video").forEach(v=>{try{v.pause()}catch(_){}})}
ajustesAplicar();
function ajustesAbrir(){if(document.querySelector(".aj-cx"))return;const v=document.createElement("div");v.className="fa-veu";
  const tela=()=>{const a=ajustesLer();v.innerHTML=`<div class="fa-cx aj-cx" role="dialog" aria-label="Ajustes do app"><h3 style="margin:0 0 4px">⚙️ Ajustes do app</h3>
    <p class="nota" style="margin:0 0 14px">Valem só para este celular.</p>
    <button type="button" class="aj-op" id="ajEco" aria-pressed="${!!a.eco}"><span><b>🔋 Modo economia</b><small>Desliga o vídeo de fundo e os efeitos animados (brilho e sombras das caveiras). Bom para celular mais simples ou para poupar bateria e internet.</small></span><i>${a.eco?"Ligado":"Desligado"}</i></button>
    <p style="margin:16px 0 8px;font-weight:700">🔠 Tamanho da letra</p>
    <div class="aj-letras">${[["","Normal"],["g","Grande"],["gg","Muito grande"]].map(([k,r])=>`<button type="button" data-l="${k}" aria-pressed="${(a.letra||"")===k}">${r}</button>`).join("")}</div>
    <button type="button" class="fa-fechar">Fechar</button></div>`;
    v.querySelector("#ajEco").onclick=()=>{const b=ajustesLer();b.eco=!b.eco;try{localStorage.setItem("ajustes_app",JSON.stringify(b))}catch(_){}ajustesAplicar();if(!b.eco)document.querySelectorAll("video.fundo-video").forEach(x=>{try{x.play()}catch(_){}});tela()};
    v.querySelectorAll("[data-l]").forEach(b=>b.onclick=()=>{const c=ajustesLer();c.letra=b.dataset.l;try{localStorage.setItem("ajustes_app",JSON.stringify(c))}catch(_){}ajustesAplicar();tela()})};
  v.onclick=e=>{if(e.target===v||e.target.closest(".fa-fechar"))v.remove()};document.body.appendChild(v);tela()}

// ---- 📷 foto do irmão e 🎨 fundo da carteirinha (só neste celular: não vão para a planilha, a nuvem nem o comando) ----
function motoFotoLer(){try{return localStorage.getItem("motofoto_"+D.email)||""}catch(_){return ""}}
function motoNome(){try{const o=JSON.parse(localStorage.getItem("moto_"+(D.email||"eu"))||"{}").moto||{};return [o.modelo,o.ano].filter(Boolean).join(" · ")}catch(_){return ""}}
function fotoLer(){try{return localStorage.getItem("foto_"+D.email)||""}catch(_){return ""}}
function fundoLer(){try{return localStorage.getItem("cartfundo_"+D.email)||"padrao"}catch(_){return "padrao"}}
const CART_FUNDOS=[["padrao","Padrão"],["couro","Couro"],["metal","Metal"],["carbono","Fibra de carbono"],["sangue","Vermelho sangue"],["imagem","Minha imagem"]];
function cartImgLer(){try{return localStorage.getItem("cartimg_"+D.email)||""}catch(_){return ""}}
function cartImgProcessar(arq){return new Promise((ok,erro)=>{const r=new FileReader();r.onerror=erro;r.onload=()=>{const im=new Image();im.onerror=erro;im.onload=()=>{
  const k=Math.min(1,900/Math.max(im.width,im.height)),c=document.createElement("canvas");c.width=Math.round(im.width*k);c.height=Math.round(im.height*k);
  c.getContext("2d").drawImage(im,0,0,c.width,c.height);ok(c.toDataURL("image/jpeg",.75))};im.src=r.result};r.readAsDataURL(arq)})}
function fotoProcessar(arq){return new Promise((ok,erro)=>{const r=new FileReader();r.onerror=erro;r.onload=()=>{const im=new Image();im.onerror=erro;im.onload=()=>{
  const L=320,c=document.createElement("canvas");c.width=c.height=L;const g=c.getContext("2d"),s=Math.min(im.width,im.height);
  g.drawImage(im,(im.width-s)/2,(im.height-s)/2,s,s,0,0,L,L);ok(c.toDataURL("image/jpeg",.82))};im.src=r.result};r.readAsDataURL(arq)})}
function fotoDesenha(){const b=document.getElementById("fotoBt");if(!b)return;const f=fotoLer();
  b.innerHTML=f?`<img src="${f}" alt="Minha foto">`:`<span>📷<small>+</small></span>`;b.classList.toggle("tem",!!f)}
function fotoAbrir(){if(document.querySelector(".foto-cx"))return;const v=document.createElement("div");v.className="fa-veu";const f=fotoLer();
  v.innerHTML=`<div class="fa-cx foto-cx" role="dialog" aria-label="Minha foto"><h3 style="margin:0 0 6px">📷 Minha foto</h3>
    <p class="nota" style="margin:0 0 12px">🔒 A foto fica <b>só neste celular</b>. Ninguém mais vê, nem o comando. Ela aparece ao lado do Pin e na sua carteirinha.</p>
    ${f?`<img src="${f}" alt="" style="display:block;width:120px;height:120px;border-radius:50%;object-fit:cover;margin:0 auto 14px;border:2px solid #E9B949">`:""}
    <div class="qbts" style="flex-direction:column;gap:8px"><button type="button" id="ftCam">📸 Tirar foto</button><button type="button" id="ftGal">🖼️ Escolher da galeria</button>${f?'<button type="button" id="ftRem">🗑️ Remover a foto</button>':""}</div>
    <input type="file" accept="image/*" capture="user" id="ftInCam" hidden><input type="file" accept="image/*" id="ftInGal" hidden>
    <button type="button" class="fa-fechar">Fechar</button></div>`;
  document.body.appendChild(v);v.onclick=e=>{if(e.target===v||e.target.closest(".fa-fechar"))v.remove()};
  const usa=async inp=>{const a=inp.files&&inp.files[0];if(!a)return;try{const d=await fotoProcessar(a);localStorage.setItem("foto_"+D.email,d);fotoDesenha();v.remove()}catch(_){alert("Não consegui usar essa foto. Tente outra.")}};
  v.querySelector("#ftCam").onclick=()=>v.querySelector("#ftInCam").click();v.querySelector("#ftGal").onclick=()=>v.querySelector("#ftInGal").click();
  v.querySelector("#ftInCam").onchange=e=>usa(e.target);v.querySelector("#ftInGal").onchange=e=>usa(e.target);
  const rm=v.querySelector("#ftRem");if(rm)rm.onclick=()=>{if(!confirm("Remover a sua foto?"))return;try{localStorage.removeItem("foto_"+D.email)}catch(_){}fotoDesenha();v.remove()}}
// o círculo da foto, ao lado do Pin
(function(){const h=document.querySelector("#app h1");if(!h||document.getElementById("fotoBt"))return;const b=document.createElement("button");b.type="button";b.id="fotoBt";b.className="foto-bt";
  b.setAttribute("aria-label","Minha foto");b.onclick=fotoAbrir;const pin=h.querySelector("img");if(pin)pin.after(b);else h.appendChild(b);fotoDesenha()})();

// ---- ⏳ tempo de Insanos, 📤 compartilhar a carteirinha e 📜 minha história no clube ----
function tempoInsanos(){if(!D.ent)return "";const a=new Date(D.ent+"T00:00:00"),h=new Date();h.setHours(0,0,0,0);if(isNaN(a)||a>h)return "";
  let an=h.getFullYear()-a.getFullYear(),me=h.getMonth()-a.getMonth(),di=h.getDate()-a.getDate();
  if(di<0){me--;di+=new Date(h.getFullYear(),h.getMonth(),0).getDate()}if(me<0){an--;me+=12}
  const p=[];if(an)p.push(an+(an===1?" ano":" anos"));if(me)p.push(me+(me===1?" mês":" meses"));if(di||!p.length)p.push(di+(di===1?" dia":" dias"));
  return p.length>1?p.slice(0,-1).join(", ")+" e "+p[p.length-1]:p[0]}
async function cartCompartilhar(){
  const W=1080,H=1920,c=document.createElement("canvas");c.width=W;c.height=H;const g=c.getContext("2d");
  const ld=src=>new Promise(r=>{if(!src)return r(null);const im=new Image();im.onload=()=>r(im);im.onerror=()=>r(null);im.src=src});
  const [cav,pin,foto,bgi,mapa]=await Promise.all([ld(CAV),ld(PIN?`pin-${PIN[0]}.png`:""),ld(fotoLer()),ld(fundoLer()==="imagem"?cartImgLer():""),ld(D.mapa?"mapa-divisao.png?v="+encodeURIComponent(D.ger||""):"")]);const mfi=await ld(motoFotoLer());
  const rr=(x,y,w,h,r)=>{g.beginPath();g.moveTo(x+r,y);g.arcTo(x+w,y,x+w,y+h,r);g.arcTo(x+w,y+h,x,y+h,r);g.arcTo(x,y+h,x,y,r);g.arcTo(x,y,x+w,y,r);g.closePath()};
  const F="Inter, Roboto, Arial, sans-serif",fd=fundoLer();
  g.save();rr(0,0,W,H,56);g.clip();
  const lg=(a,b,cc)=>{const q=g.createLinearGradient(0,0,W,H);q.addColorStop(0,a);if(cc)q.addColorStop(.5,cc);q.addColorStop(1,b);g.fillStyle=q;g.fillRect(0,0,W,H)};
  if(fd==="imagem"&&bgi){const k=Math.max(W/bgi.width,H/bgi.height);g.drawImage(bgi,(W-bgi.width*k)/2,(H-bgi.height*k)/2,bgi.width*k,bgi.height*k);g.fillStyle="rgba(0,0,0,.62)";g.fillRect(0,0,W,H)}
  else if(fd==="couro")lg("#5a3520","#2b170c");else if(fd==="metal")lg("#2a2d31","#24272b","#5d636b");else if(fd==="sangue")lg("#5c0b0b","#1c0404");
  else if(fd==="carbono"){g.fillStyle="#181818";g.fillRect(0,0,W,H);g.strokeStyle="#232323";g.lineWidth=12;for(let x=-H;x<W;x+=24){g.beginPath();g.moveTo(x,0);g.lineTo(x+H,H);g.stroke()}}
  else{g.fillStyle="#1b1714";g.fillRect(0,0,W,H)}
  g.restore();g.lineWidth=8;g.strokeStyle="#E9B949";rr(4,4,W-8,H-8,52);g.stroke();
  if(cav){const h=170,w=h*cav.width/cav.height;g.drawImage(cav,80,60,w,h);
    g.fillStyle="#ECE7DC";g.font=`800 58px ${F}`;g.fillText("INSANOS M.C.",80+w+34,135);g.fillStyle="#BDB7AA";g.font=`500 36px ${F}`;g.fillText(`Divisão ${D.div} · ${D.reg}`,80+w+34,188)}
  g.fillStyle="rgba(233,185,73,.45)";g.fillRect(80,270,W-160,3);
  let x=80;
  if(foto){g.save();g.beginPath();g.arc(x+120,450,120,0,7);g.clip();g.drawImage(foto,x,330,240,240);g.restore();g.lineWidth=6;g.strokeStyle="#E9B949";g.beginPath();g.arc(x+120,450,120,0,7);g.stroke();x+=270}
  if(pin){const h=250,w=h*pin.width/pin.height;g.drawImage(pin,x,325,w,h);x+=w+30}
  const tw=W-80-x;let fs=74;g.font=`800 ${fs}px ${F}`;while(g.measureText(D.n).width>tw&&fs>40){fs-=4;g.font=`800 ${fs}px ${F}`}
  const cabe=(t,sz,pe)=>{let z=sz;g.font=`${pe} ${z}px ${F}`;while(g.measureText(t).width>tw&&z>20){z-=2;g.font=`${pe} ${z}px ${F}`}};   // a linha diminui para caber
  g.fillStyle="#ECE7DC";g.fillText(D.n,x,405);g.fillStyle="#E9B949";cabe(D.c||"",40,700);g.fillText(D.c||"",x,462);
  g.fillStyle="#BDB7AA";const ent=D.ent?D.ent.slice(8,10)+"/"+D.ent.slice(5,7)+"/"+D.ent.slice(0,4):"";
  if(ent){cabe("Insanos M.C. desde "+ent,32,500);g.fillText("Insanos M.C. desde "+ent,x,512)}const tp=tempoInsanos();if(tp){g.fillStyle="#ECE7DC";cabe("⏳ "+tp,32,500);g.fillText("⏳ "+tp,x,556)}
  const vp=pc(D.pres);
  if(vp!=null&&cav){g.fillStyle="rgba(255,255,255,.06)";rr(80,640,W-160,330,36);g.fill();
    const h=270,w=h*cav.width/cav.height,cx=130,cy=670,o=document.createElement("canvas");o.width=w;o.height=h;const og=o.getContext("2d");
    og.drawImage(cav,0,0,w,h);const hx=cor(vp)[0],cr=parseInt(hx.slice(1,3),16),cg=parseInt(hx.slice(3,5),16),cb=parseInt(hx.slice(5,7),16);
    const px=og.getImageData(0,0,w|0,h|0),dd=px.data,base=new ImageData(new Uint8ClampedArray(dd),px.width,px.height),bd=base.data;
    for(let k=0;k<dd.length;k+=4){const al=dd[k+3]<60?0:dd[k+3],lu=Math.min(1.15,(dd[k]+dd[k+1]+dd[k+2])/600+.25);   // ponto a ponto: o fundo quase transparente some
      dd[k]=Math.min(255,cr*lu);dd[k+1]=Math.min(255,cg*lu);dd[k+2]=Math.min(255,cb*lu);dd[k+3]=al;const cz=(bd[k]+bd[k+1]+bd[k+2])/3*.55;bd[k]=bd[k+1]=bd[k+2]=cz;bd[k+3]=al}
    og.putImageData(px,0,0);const ob=document.createElement("canvas");ob.width=w;ob.height=h;ob.getContext("2d").putImageData(base,0,0);
    g.globalAlpha=.45;g.drawImage(ob,cx,cy,w,h);g.globalAlpha=1;const t2=document.createElement("canvas");t2.width=w;t2.height=h;const tg=t2.getContext("2d");tg.beginPath();tg.rect(0,h*(1-vp/100),w,h*vp/100);tg.clip();tg.drawImage(o,0,0,w,h);
    g.save();if(vp>=90){g.shadowColor="rgba(233,185,73,.8)";g.shadowBlur=30}g.drawImage(t2,cx,cy,w,h);g.restore();   // o recorte primeiro, o brilho depois (sem cortar o brilho)
    if(mfi){const bx=W-80-30-300,by=670,bw=300,bh=270,k=Math.max(bw/mfi.width,bh/mfi.height);g.save();rr(bx,by,bw,bh,24);g.clip();
        g.drawImage(mfi,bx+(bw-mfi.width*k)/2,by+(bh-mfi.height*k)/2,mfi.width*k,mfi.height*k);g.restore();g.lineWidth=4;g.strokeStyle="rgba(233,185,73,.7)";rr(bx,by,bw,bh,24);g.stroke();
        g.fillStyle=cor(vp)[1];g.font=`900 120px ${F}`;g.fillText(vp+"%",cx+w+36,810);g.fillStyle="#BDB7AA";g.font=`500 30px ${F}`;g.fillText("de presença nos",cx+w+40,860);g.fillText(`eventos em ${D.ano}`,cx+w+40,898)}   // com a moto, o texto vai em duas linhas
    else{g.fillStyle=cor(vp)[1];g.font=`900 150px ${F}`;g.fillText(vp+"%",cx+w+50,830);g.fillStyle="#BDB7AA";g.font=`500 36px ${F}`;g.fillText(`de presença nos eventos em ${D.ano}`,cx+w+54,890)}}
  let yy=1010;let kmR=0;try{const ab=(JSON.parse(localStorage.getItem("moto_"+(D.email||"eu"))||"{}").abs||[]).map(a=>+a.km).filter(x=>x>0);if(ab.length>1)kmR=Math.max(...ab)-Math.min(...ab)}catch(_){}
  const kmTx=kmR>0?`🏍️ Eu já rodei ${kmR.toLocaleString("pt-BR")} km com a minha moto`:"";   // do primeiro ao último odômetro registrado na Minha moto
  if(mapa){g.fillStyle="rgba(255,255,255,.06)";rr(80,yy,W-160,530,36);g.fill();g.fillStyle="#ECE7DC";g.font=`800 40px ${F}`;g.fillText("📍 Por onde a Divisão já passou",110,yy+62);
    g.fillStyle="#BDB7AA";g.font=`500 30px ${F}`;g.fillText(`${D.mapa.lugares} ${D.mapa.lugares===1?"lugar":"lugares"} · ${D.mapa.eventos} eventos`,110,yy+104);
    if(kmTx){g.fillStyle="#E9B949";g.font=`700 28px ${F}`;g.textAlign="right";g.fillText(kmTx.replace("Eu já rodei ","").replace(" com a minha moto"," rodados"),W-110,yy+104);g.textAlign="left"}
    const mw=W-260,mh=mw*mapa.height/mapa.width;g.save();rr(130,yy+125,mw,mh,24);g.clip();g.drawImage(mapa,130,yy+125,mw,mh);g.restore();yy+=560}
  if(!mapa&&kmTx){g.fillStyle="#E9B949";g.font=`800 36px ${F}`;g.textAlign="center";g.fillText(kmTx,W/2,yy+40);g.textAlign="left";yy+=90}
  // 🤝 faça parte: o QR do cadastro de novos integrantes (o endereço sem o rastreador do Facebook)
  const URL_FP="https://marketing.insanosmc.com.br/faca-parte-do-maior-moto-clube-do-brasil-insanos-mc";
  g.fillStyle="rgba(233,185,73,.12)";rr(80,yy,W-160,250,36);g.fill();g.strokeStyle="rgba(233,185,73,.6)";g.lineWidth=3;rr(80,yy,W-160,250,36);g.stroke();
  try{await carregarQRLib();if(window.qrcode){const q=qrcode(0,"M");q.addData(URL_FP);q.make();const n=q.getModuleCount(),tam=200,m=tam/n,qx=115,qy=yy+25;
    g.fillStyle="#fff";rr(qx-10,qy-10,tam+20,tam+20,14);g.fill();g.fillStyle="#000";for(let r=0;r<n;r++)for(let k=0;k<n;k++)if(q.isDark(r,k))g.fillRect(qx+k*m,qy+r*m,Math.ceil(m),Math.ceil(m))}}catch(_){}
  g.fillStyle="#ECE7DC";g.font=`800 38px ${F}`;g.fillText("Quer fazer parte do maior",360,yy+72);g.fillText("moto clube do Brasil?",360,yy+118);
  g.fillStyle="#BDB7AA";g.font=`500 28px ${F}`;g.fillText("Aponte a câmera para o QR",360,yy+166);g.fillStyle="#E9B949";g.font=`800 32px ${F}`;g.fillText("insanosmc.com.br",360,yy+212);
  g.fillStyle="#8f887d";g.font=`600 28px ${F}`;g.textAlign="center";g.fillText(`Insanos M.C. · Divisão ${D.div} · ${D.reg}`,W/2,H-45);g.textAlign="left";
  c.toBlob(async b=>{if(!b){alert("Não consegui gerar a imagem.");return}const f=new File([b],"carteirinha-insanos.png",{type:"image/png"});
    try{if(navigator.canShare&&navigator.canShare({files:[f]})){await navigator.share({files:[f],title:"Minha carteirinha Insanos M.C."});return}}catch(_){return}
    const a=document.createElement("a");a.href=URL.createObjectURL(b);a.download="carteirinha-insanos.png";document.body.appendChild(a);a.click();a.remove()},"image/png")}
// 📜 minha história no clube (linha do tempo)
(function(){const app=document.getElementById("app");if(!app||document.getElementById("histC"))return;const H=D.hist||{},it=[],br=d=>d.slice(8,10)+"/"+d.slice(5,7)+"/"+d.slice(0,4);
  if(D.ent){it.push([D.ent,"🏍️","Entrada no Insanos M.C.","Recebeu o Pin Caveira"]);const a0=+D.ent.slice(0,4),hj=new Date().getFullYear();
    [[3,"🥉","Pin Bronze"],[6,"🥈","Pin Prata"],[9,"🥇","Pin Ouro"]].forEach(([n,ic,r])=>{const an=a0+n;it.push([an+"-01-01",ic,r,an<=hj?"Recebido em "+an:"Previsto para "+an,an>hj])})}
  if(H.pri)it.push([H.pri[0],"⭐","Primeira presença",H.pri[1]]);
  (H.marcos||[]).forEach(([n,d,e])=>it.push([d,"🎯",`${n}ª presença`,e]));
  Object.entries(H.anos||{}).sort().forEach(([a,[p,t]])=>{if(t)it.push([a+"-12-31","📅",`${a}: ${p} de ${t} eventos`,`${Math.round(p*100/t)}% de presença`,a==String(new Date().getFullYear())])});
  if(!it.length)return;it.sort((x,y)=>x[0]<y[0]?-1:x[0]>y[0]?1:0);
  const tp=tempoInsanos(),d=document.createElement("details");d.className="card";d.id="histC";
  d.innerHTML=`<summary style="cursor:pointer;font-weight:800;font-size:var(--fs-h2,21px)">📜 Minha história no clube</summary>
    <p class="nota" style="margin:8px 0 12px">${tp?`⏳ ${tp} de Insanos`:""}${H.tot?` · ${H.tot} presenças no total`:""}</p>
    <div class="hist">${it.map(([dt,ic,t,s2,fut])=>`<div class="hi${fut?" fut":""}"><i>${ic}</i><div><b>${es(t)}</b><small>${dt.slice(5)==="01-01"||dt.slice(5)==="12-31"?"":br(dt)+(s2?" · ":"")}${es(s2||"")}</small></div></div>`).join("")}</div>`;
  const ref=document.getElementById("avisos")||[...app.children].find(c=>/Atualizado em/.test(c.textContent||""));   // antes dos avisos e do rodapé
  if(ref&&ref.parentNode===app)app.insertBefore(d,ref);else app.appendChild(d)})();

// ---- ☀️ bom-dia das 9h: o irmão pode desligar só esse aviso (fica guardado na planilha, aba Preferências) ----
(function(){const b=document.getElementById("bdTg");if(!b)return;const k="bomdia_off_"+D.email;
  let off=false;try{off=localStorage.getItem(k)==="1"}catch(_){}
  const rot=()=>{b.textContent=off?"🔔 Voltar a receber o bom-dia":"🔕 Não receber o bom-dia"};rot();
  b.onclick=async()=>{const q=typeof sosQuem==="function"?sosQuem():null;if(!q||!window.SOS_URL){alert("Não consegui agora. Tente de novo em instantes.");return}
    b.disabled=true;try{await fetch(window.SOS_URL,{method:"POST",headers:{"Content-Type":"text/plain;charset=utf-8"},body:JSON.stringify({tipo:"bomdia",email:q.email,s:q.s,ativo:off})});
      off=!off;try{localStorage.setItem(k,off?"1":"0")}catch(_){}rot();alert(off?"🔕 Pronto: você não recebe mais o bom-dia (os outros avisos continuam).":"🔔 Pronto: o bom-dia volta amanhã, às 9h.")}
    catch(_){alert("Não consegui agora. Tente de novo em instantes.")}b.disabled=false}})();
// ---- 📊 enquete do comando: o irmão vota com um toque (pode trocar enquanto estiver aberta; o resultado só o comando vê) ----
(function(){const box=document.getElementById("enqOps");if(!box||!D.enq)return;const E=D.enq,k="enq_"+E.id+"_"+D.email;
  let meu=E.meu||"";try{meu=localStorage.getItem(k)||meu}catch(_){}
  const desenha=()=>{box.innerHTML=E.ops.map((o,i)=>`<button data-eo="${i}" class="${o===meu?"qz":""}" style="flex:1 1 40%">${o===meu?"✓ ":""}${es(o)}</button>`).join("");
    document.getElementById("enqNota").textContent=meu?"✓ Voto registrado. Dá para trocar enquanto a enquete estiver aberta.":"Toque na sua opção. O resultado só o comando vê."};
  desenha();
  box.onclick=async e=>{const b=e.target.closest("[data-eo]");if(!b)return;const o=E.ops[+b.dataset.eo];if(o===meu)return;
    const q=typeof sosQuem==="function"?sosQuem():null;if(!q||!window.SOS_URL){alert("Não consegui enviar agora. Tente de novo em instantes.");return}
    b.disabled=true;
    try{await fetch(window.SOS_URL,{method:"POST",headers:{"Content-Type":"text/plain;charset=utf-8"},body:JSON.stringify({tipo:"voto",email:q.email,s:q.s,id:E.id,opcao:o})});
      meu=o;try{localStorage.setItem(k,o)}catch(_){}desenha();try{navigator.vibrate&&navigator.vibrate(15)}catch(_){}}
    catch(_){b.disabled=false;alert("Não consegui enviar agora. Tente de novo em instantes.")}}})();
// ---- os 3 quadrados (eventos, aniversários, mensalidade): prévia; tocando, abre o bloco completo logo abaixo ----
(function(){const tri=document.getElementById("tri");if(!tri)return;
  document.head.insertAdjacentHTML("beforeend",`<style>
    .tri{display:grid;grid-template-columns:repeat(var(--n,3),minmax(0,1fr));gap:8px;margin:0 0 12px}
    .tri button{aspect-ratio:1/1;display:flex;flex-direction:column;align-items:flex-start;justify-content:space-between;gap:4px;padding:10px;border-radius:16px;
      border:1px solid rgba(236,231,220,.16);background:rgba(236,231,220,.05);color:#ECE7DC;font:inherit;text-align:left;cursor:pointer;min-width:0}
    .tri button[aria-pressed="true"]{border-color:#E9B949;background:rgba(233,185,73,.10)}
    .tri .ti{font-size:22px;line-height:1}.tri .tt{font-size:11.5px;font-weight:700;color:#BDB7AA;text-transform:uppercase;letter-spacing:.03em}
    .tri .tp{font-size:13px;font-weight:700;line-height:1.25;overflow:hidden;display:-webkit-box;-webkit-line-clamp:2;-webkit-box-orient:vertical;width:100%}
    .tri .ts{font-size:11.5px;color:#BDB7AA;width:100%;white-space:nowrap;overflow:hidden;text-overflow:ellipsis}
    details.em-tri:not([open]){display:none}details.em-tri>summary{display:none}</style>`);
  const hoje=new Date();hoje.setHours(0,0,0,0);const br=d=>d.slice(8,10)+"/"+d.slice(5,7);
  const Q=[];
  if(document.getElementById("eventos")){let loc={};try{loc=JSON.parse(localStorage.getItem("rsvp_"+D.email)||"{}")}catch(_){}
    const e=D.prox[0],st=e?(loc[e.d+"|"+e.e]||e.st):"";
    Q.push({id:"eventos",ico:"🗓️",t:"Eventos",p:e?`${br(e.d)} · ${e.e}`:"Nenhum",s:st==="sim"?"✅ Você vai":st==="nao"?"❌ Você não vai":"⏳ Responder",ouro:!st})}
  if(document.getElementById("anivC")){const dias=md=>{const [m,dd]=md.split("-").map(Number);let x=new Date(hoje.getFullYear(),m-1,dd);if(x<hoje)x=new Date(hoje.getFullYear()+1,m-1,dd);return Math.round((x-hoje)/864e5)};
    const L=(D.anivs||[]).map(a=>({...a,d:dias(a.md)})).sort((a,b)=>a.d-b.d),a=L[0];
    if(a)Q.push({id:"anivC",ico:"🎂",t:"Aniversários",p:a.n===D.n?"Você!":a.n,s:a.d===0?"🎉 Hoje!":`${a.md.slice(3,5)}/${a.md.slice(0,2)} · ${a.d===1?"amanhã":"em "+a.d+" dias"}`,ouro:a.d===0})}
  const ms=document.getElementById("mensal");
  if(ms){const sm=(ms.querySelector("summary small")||{}).textContent||"";Q.push({id:"mensal",ico:"💵",t:"Mensalidade",p:sm.replace(/^[^A-Za-zÀ-ÿ0-9]+/,"").split("·")[0].trim(),s:(sm.split("·")[1]||"").trim(),ouro:/aberto/i.test(sm)})}
  if(Q.length<2){tri.remove();return}   // com um bloco só, fica como está
  tri.style.setProperty("--n",Q.length);
  Q.forEach(q=>{const d=document.getElementById(q.id);if(d){d.classList.add("em-tri");d.open=false}});
  tri.innerHTML=Q.map(q=>`<button type="button" data-tri="${q.id}" aria-pressed="false"><span class="ti">${q.ico}</span><span class="tt">${q.t}</span>
    <span class="tp">${es(q.p)}</span><span class="ts" style="${q.ouro?"color:#E9B949;font-weight:700":""}">${es(q.s)}</span></button>`).join("");
  tri.onclick=ev=>{const b=ev.target.closest("[data-tri]");if(!b)return;const alvo=document.getElementById(b.dataset.tri),abrir=!alvo.open;
    Q.forEach(q=>{const d=document.getElementById(q.id);if(d)d.open=false});tri.querySelectorAll("button").forEach(x=>x.setAttribute("aria-pressed","false"));
    if(abrir){alvo.open=true;b.setAttribute("aria-pressed","true");setTimeout(()=>alvo.scrollIntoView({behavior:"smooth",block:"nearest"}),50)}};
  if(location.hash==="#eventos"){const b=tri.querySelector('[data-tri="eventos"]');if(b)b.click()}})();
// avisos no celular (lembretes individuais): mesmo cadastro do painel ADM, ligado ao e-mail do irmão
const b64u=s=>{const p="=".repeat((4-s.length%4)%4),b=atob((s+p).replace(/-/g,"+").replace(/_/g,"/"));return Uint8Array.from(b,c=>c.charCodeAt(0))};
const ehApp=()=>matchMedia("(display-mode: standalone)").matches||navigator.standalone,ehIph=()=>/iPhone|iPad|iPod/.test(navigator.userAgent);
async function avTela(){const st=document.getElementById("avSt"),at=document.getElementById("avAt"),de=document.getElementById("avDes");
  if(!D.push||!("serviceWorker" in navigator)||!("PushManager" in window)){if(ehIph()&&!ehApp()){st.textContent="No iPhone, os avisos só funcionam com o app instalado: Safari › Compartilhar › Adicionar à Tela de Início. Depois abra pelo ícone e entre de novo.";at.hidden=true;return}
    document.getElementById("avisos").hidden=true;return}
  let s=null;try{const r=await navigator.serviceWorker.ready;s=await r.pushManager.getSubscription()}catch(e){}
  const on=!!s&&Notification.permission==="granted";at.hidden=on;de.hidden=!on;
  st.textContent=on?"✅ Avisos ativados neste celular. Você recebe só os seus lembretes.":"Receba lembretes das listas e da agenda da divisão, só para você."}
document.getElementById("avAt").onclick=async()=>{try{localStorage.removeItem("push_off")}catch(_){}if(ehIph()&&!ehApp()){alert("No iPhone, instale o app primeiro: Safari › Compartilhar › Adicionar à Tela de Início.");return}
  const perm=await Notification.requestPermission();if(perm!=="granted"){alert("Permissão não concedida. Libere as notificações nas configurações do celular.");return}
  try{const r=await navigator.serviceWorker.ready;let s=await r.pushManager.getSubscription();
    if(!s)s=await r.pushManager.subscribe({userVisibleOnly:true,applicationServerKey:b64u(D.push.pub)});
    await fetch(D.push.url,{method:"POST",mode:"no-cors",body:JSON.stringify({tipo:"push",email:D.email,sub:JSON.stringify(s)})});
    alert("🔔 Avisos ativados neste celular!")}catch(e){alert("Não consegui ativar os avisos ("+(e.message||e)+").")}avTela()};
document.getElementById("avDes").onclick=async()=>{try{localStorage.setItem("push_off","1")}catch(_){}try{const r=await navigator.serviceWorker.ready,s=await r.pushManager.getSubscription();
  if(s){await fetch(D.push.url,{method:"POST",mode:"no-cors",body:JSON.stringify({tipo:"push_sair",endpoint:s.endpoint})});await s.unsubscribe()}}catch(e){}alert("Avisos desativados.");avTela()};
// SOS: guarda a localização só neste celular (a notificação calcula a distância sem mandar nada para ninguém)
const locIdb=(v)=>new Promise(r=>{try{const o=indexedDB.open("insanos",1);o.onupgradeneeded=()=>o.result.createObjectStore("kv");o.onsuccess=()=>{const t=o.result.transaction("kv","readwrite");t.objectStore("kv").put(v,"loc");t.oncomplete=()=>r(1)};o.onerror=()=>r(0)}catch(e){r(0)}});
const locPega=(avisa)=>navigator.geolocation&&navigator.geolocation.getCurrentPosition(p=>{locIdb({lat:+p.coords.latitude.toFixed(5),lng:+p.coords.longitude.toFixed(5),ts:Date.now()});
  const b=document.getElementById("locAt");if(b){b.textContent="✅ Localização guardada neste celular";b.classList.add("qz")}if(avisa)alert("📍 Pronto! Nos alertas de SOS, você vai ver a que distância está do irmão.")},
  ()=>{if(avisa)alert("Não consegui a localização. Libere a localização para este app nas configurações do celular.")},{enableHighAccuracy:false,timeout:10000,maximumAge:600000});
document.getElementById("locAt").onclick=()=>locPega(true);
try{navigator.permissions&&navigator.permissions.query({name:"geolocation"}).then(x=>{if(x.state==="granted")locPega(false)})}catch(_){}
avTela();
// avisos: com a permissão já dada, refaz o registro sozinho (app reinstalado, celular trocado, registro antigo)
async function avSinc(){try{if(!D.push||!("serviceWorker" in navigator)||!("PushManager" in window)||Notification.permission!=="granted")return;
  try{if(localStorage.getItem("push_off"))return}catch(_){}   // desligado de propósito
  const r=await navigator.serviceWorker.ready;let s=await r.pushManager.getSubscription();
  if(!s)s=await r.pushManager.subscribe({userVisibleOnly:true,applicationServerKey:b64u(D.push.pub)});
  let u={};try{u=JSON.parse(localStorage.getItem("push_sync")||"{}")}catch(_){}
  if(u.e===D.email&&u.ep===s.endpoint&&Date.now()-u.t<864e5)return;
  await fetch(D.push.url,{method:"POST",mode:"no-cors",body:JSON.stringify({tipo:"push",email:D.email,sub:JSON.stringify(s)})});
  try{localStorage.setItem("push_sync",JSON.stringify({e:D.email,ep:s.endpoint,t:Date.now()}))}catch(_){}avTela()}catch(e){}}
setTimeout(avSinc,2500);
// depois de um Vou / Não vou / Já respondi aqui no Meu painel: acorda o painel na hora (a lista atualiza em 1 a 3 minutos)
function acordarDepois(p){return Promise.resolve(p).then(r=>{try{if(window.SOS_URL)fetch(window.SOS_URL,{method:"POST",body:JSON.stringify({tipo:"rsvp_acordar",email:D.email})}).catch(()=>{})}catch(_){}return r})}
setTimeout(()=>{try{reinstAviso(D.reinst)}catch(_){}},1800);
// ---- carteirinha (abre com o QR de identificação) e conquistas (sequências e marcas, sem expor faltas)
function carregarQRLib(){return new Promise(ok=>{if(window.qrcode)return ok();const sc=document.createElement("script");sc.src="https://cdn.jsdelivr.net/npm/qrcode-generator@1.4.4/qrcode.js";sc.onload=()=>ok();sc.onerror=()=>ok();document.head.appendChild(sc)})}
function cartAbrir(){const ent=D.ent?D.ent.slice(8,10)+"/"+D.ent.slice(5,7)+"/"+D.ent.slice(0,4):"";
  const v=document.createElement("div");v.className="cart-veu";v.innerHTML=`<div class="cart f-${fundoLer()==="imagem"&&!cartImgLer()?"padrao":fundoLer()}" style="${cartImgLer()?`--cimg:url('${cartImgLer()}')`:""}"><div class="topo"><img src="${CAV}" alt=""><div><b>INSANOS M.C.</b><span>Divisão ${es(D.div)} · ${es(D.reg)}</span></div></div>
    <div class="meio">${fotoLer()?`<img class="foto" src="${fotoLer()}" alt="">`:""}${PIN?`<img class="pin" src="pin-${PIN[0]}.png" alt="${PIN[1]}">`:""}<div><h3>${es(D.n)}</h3><div class="c">${es(D.c||"")}</div>
      <div class="d">${ent?"Insanos M.C. desde "+ent:""}${PIN?" · "+PIN[1]:""}</div>${tempoInsanos()?`<div class="d" style="color:#ECE7DC">⏳ ${tempoInsanos()} de Insanos</div>`:""}</div></div>
    ${motoFotoLer()?`<div class="moto-f"><img src="${motoFotoLer()}" alt="A moto"><span>🏍️ ${es(motoNome()||"A minha moto")}</span></div>`:""}
    ${vp!=null?`<div class="pres"><div class="cv${vp>=90?" trofeu":""}" style="--g:${corGlow(vp)}"><img src="${CAV}" alt=""><span class="o" style="--cav:url('${CAV}');--c:${cor(vp)[0]};clip-path:inset(${100-vp}% 0 0 0)"></span></div><div><b style="color:${cor(vp)[1]}">${vp}%</b><span>de presença nos eventos em ${D.ano}${D.pres&&D.pres[1]?` · ${D.pres[0]} de ${D.pres[1]}`:""}</span></div></div>`:""}
    <div class="qr"><canvas width="220" height="220"></canvas><p>Código de identificação do irmão. Quem ler vê o nome, o cargo, a divisão e desde quando é Insanos.</p></div>
    <div class="fundos-t">🎨 Fundo da carteirinha${fotoLer()?"":` · <a href="#" class="cart-foto" style="color:#E9B949">📷 pôr minha foto</a>`}</div>
    <div class="fundos">${CART_FUNDOS.map(([k,r])=>`<button type="button" data-fd="${k}" aria-label="${r}" title="${r}" aria-pressed="${fundoLer()===k}"${k==="imagem"&&cartImgLer()?` style="background:url('${cartImgLer()}') center/cover!important"`:""}>${k==="imagem"&&!cartImgLer()?"📷":""}</button>`).join("")}</div>
    <div class="fundos-t fundo-img-acoes"${fundoLer()==="imagem"&&cartImgLer()?"":" hidden"}><a href="#" data-img-troca style="color:#E9B949">🖼️ Trocar imagem</a> · <a href="#" data-img-rem style="color:#BDB7AA">🗑️ Remover</a></div>
    <input type="file" accept="image/*" hidden class="cart-img-in">
    <button type="button" class="compart">📤 Compartilhar a carteirinha</button>
    <button class="fechar">Fechar</button></div>`;document.body.appendChild(v);
  const cx=v.querySelector(".cart"),inImg=v.querySelector(".cart-img-in"),acoes=v.querySelector(".fundo-img-acoes");
  v.querySelector(".compart").onclick=e=>{e.stopPropagation();cartCompartilhar()};
  const aplica=k=>{try{localStorage.setItem("cartfundo_"+D.email,k)}catch(_){}CART_FUNDOS.forEach(([f])=>cx.classList.remove("f-"+f));cx.classList.add("f-"+k);
    v.querySelectorAll("[data-fd]").forEach(x=>x.setAttribute("aria-pressed",x.dataset.fd===k));acoes.hidden=!(k==="imagem"&&cartImgLer())};
  v.querySelectorAll("[data-fd]").forEach(b=>b.onclick=e=>{e.stopPropagation();if(b.dataset.fd==="imagem"&&!cartImgLer()){inImg.click();return}aplica(b.dataset.fd)});
  inImg.onchange=async()=>{const a=inImg.files&&inImg.files[0];if(!a)return;try{const d=await cartImgProcessar(a);localStorage.setItem("cartimg_"+D.email,d);cx.style.setProperty("--cimg",`url('${d}')`);
      const sw=v.querySelector('[data-fd="imagem"]');sw.textContent="";sw.style.cssText=`background:url('${d}') center/cover!important`;aplica("imagem")}
    catch(_){alert("Não consegui usar essa imagem (pode ser grande demais). Tente outra.")}inImg.value=""};
  v.querySelector("[data-img-troca]").onclick=e=>{e.preventDefault();inImg.click()};
  v.querySelector("[data-img-rem]").onclick=e=>{e.preventDefault();if(!confirm("Remover a imagem de fundo?"))return;try{localStorage.removeItem("cartimg_"+D.email)}catch(_){}
    cx.style.removeProperty("--cimg");const sw=v.querySelector('[data-fd="imagem"]');sw.style.cssText="";sw.textContent="📷";aplica("padrao")};
  {const cf=v.querySelector(".cart-foto");if(cf)cf.onclick=e=>{e.preventDefault();v.remove();fotoAbrir()}}
  v.onclick=e=>{if(e.target===v||e.target.classList.contains("fechar"))v.remove()};
  carregarQRLib().then(()=>{if(!window.qrcode)return;const txt=`INSANOS M.C. · Divisão ${D.div} ${D.reg}\n${D.n} · ${D.c||""}${ent?"\nInsanos desde "+ent:""}${PIN?"\n"+PIN[1]:""}`;
    const q=qrcode(0,"M");q.addData(unescape(encodeURIComponent(txt)),"Byte");q.make();const cv=v.querySelector("canvas"),g=cv.getContext("2d"),n=q.getModuleCount(),m=cv.width/n;
    g.fillStyle="#fff";g.fillRect(0,0,cv.width,cv.height);g.fillStyle="#000";for(let r=0;r<n;r++)for(let c=0;c<n;c++)if(q.isDark(r,c))g.fillRect(Math.floor(c*m),Math.floor(r*m),Math.ceil(m),Math.ceil(m))})}
function conquistas(){const ev=[...(D.ev||[])].sort((a,b)=>a[0]<b[0]?1:-1),cron=[...ev].reverse();
  let atual=0;for(const e of ev){if(e[2]===1)atual++;else break}
  let maior=0,c=0;for(const e of cron){if(e[2]===1){c++;maior=Math.max(maior,c)}else c=0}
  const mesPerf=(D.meses||[]).filter(m=>m.p&&m.p[1]>=2&&m.p[0]===m.p[1]).length,insPerf=(D.meses||[]).filter(m=>m.i&&m.i[1]>=1&&m.i[0]===m.i[1]).length;
  const pa=D.pres&&D.pres[1]?Math.round(D.pres[0]*100/D.pres[1]):0;
  return [["🔥",atual>=3,`${atual} seguido${atual===1?"":"s"}`,"Sequência atual de presenças",atual>=3?"Continue assim!":"A partir de 3 seguidos"],
    ["🏅",maior>=5,`${maior} seguidos`,"Maior sequência no ano",maior>=5?"Conquistada":"Faça 5 seguidos"],
    ["📅",mesPerf>0,mesPerf?`${mesPerf} mês${mesPerf>1?"es":""}`:"0","Mês perfeito (100% de presença)","Todos os eventos de um mês"],
    ["💬",insPerf>0,insPerf?`${insPerf} mês${insPerf>1?"es":""}`:"0","Insights em dia (100% no mês)","Responder todos os insights de um mês"],
    ["⭐",pa>=75,`${pa}%`,"Frequência de ouro no ano","75% ou mais de presença"],
    ["💀",!!PIN,PIN?PIN[1]:"—","Pin Caveira",PIN&&PIN[0]!=="ouro"?"Próximo pin no começo do ano em que completar 3, 6 ou 9 anos":"Maior pin"]]}
(function(){const cb=document.getElementById("cartBt");if(cb)cb.onclick=cartAbrir;
  const L=conquistas(),box=document.getElementById("conqBox"),bt=document.getElementById("conqBt");if(!box||!bt)return;
  document.getElementById("conqN").textContent=L.filter(x=>x[1]).length;
  box.innerHTML=L.map(([i,ok,v,t,d])=>`<div class="cq${ok?"":" off"}"><b>${i} ${es(v)}</b><strong>${es(t)}</strong><span>${es(d)}</span></div>`).join("");
  bt.onclick=()=>{box.hidden=!box.hidden;bt.setAttribute("aria-expanded",!box.hidden)}})();
(function(){const b=document.getElementById("mensPedir");if(!b)return;const chave="bol_pedido_"+D.email,hoje=new Date().toISOString().slice(0,10);
  const feito=()=>{b.disabled=true;b.textContent="✅ Pedido enviado hoje";document.getElementById("mensTx").textContent=`O ${D.mens.adm} recebe o seu pedido em 1 a 3 minutos. A resposta aparece em ✉️ Mensagens.`};
  try{if(localStorage.getItem(chave)===hoje)feito()}catch(_){}
  b.onclick=async()=>{const mes=MESES_N[+String(D.mens.venc).slice(5,7)-1]||"";b.disabled=true;b.textContent="Enviando…";
    try{const r=await msgEnviar([D.mens.adm],"🧾 Pedido de boleto",`🧾 Fala, ${D.mens.adm}! Sou o ${D.n}. Pode me mandar o boleto da mensalidade de ${mes.toLowerCase()}? Valeu!`);
      if(!r||!r.ok)throw new Error((r&&r.erro)||"não enviado");try{localStorage.setItem(chave,hoje)}catch(_){}feito()}
    catch(e){b.disabled=false;b.textContent=`🙋 Pedir o boleto ao ${D.mens.adm}`;alert("Não consegui pedir agora ("+(e.message||e)+"). Entre uma vez no app com o seu e-mail e o código, se ainda não entrou.")}}})();
try{if(sessionStorage.getItem("ver_irmao")==="1"){const f=document.getElementById("comoIrmao");f.hidden=false;
  document.getElementById("voltarCmd").onclick=e=>{e.preventDefault();try{sessionStorage.removeItem("ver_irmao")}catch(_){}location.href=location.pathname}}}catch(_){}
if(location.hash==="#mensagens")setTimeout(()=>{try{abrirMensagens({modo:"irmao",comando:D.cmd||[]})}catch(_){}},900);
else setTimeout(()=>{try{msgNaoLidas({modo:"irmao",comando:D.cmd||[]})}catch(_){}},2500);document.addEventListener("visibilitychange",()=>{if(!document.hidden)avSinc()});
try{if(matchMedia("(display-mode: standalone)").matches||navigator.standalone===true)document.getElementById("instalar").hidden=true}catch(e){}
document.getElementById("tTipo").onclick=ev=>{const b=ev.target.closest("[data-t]");if(!b)return;TT=b.dataset.t;ABERTO=-1;document.querySelectorAll("#tTipo button").forEach(x=>x.classList.toggle("on",x===b));desenha()};
document.getElementById("tFil").onclick=ev=>{const b=ev.target.closest("[data-f]");if(!b)return;FF=b.dataset.f;ABERTO=-1;document.querySelectorAll("#tFil button").forEach(x=>x.classList.toggle("on",x===b));desenha()};
desenha();
document.getElementById("perSel").onchange=ev=>{PER=ev.target.value;const m=D.meses.find(x=>x.ym===PER);
  const nome=PER?MESES_N[+PER.slice(5)-1]+"/"+PER.slice(0,4):D.ano,P=PER?(m?m.p:[0,0]):D.pres,I=PER?(m?m.i:[0,0]):D.ins;
  document.getElementById("adTit").textContent="Minha adesão em "+nome;document.getElementById("mpTit").textContent="📋 Minhas presenças em "+nome;
  document.getElementById("adCavs").innerHTML=cav(pc(P),"Presença",P[1]?`${P[0]} de ${P[1]} eventos`:"sem eventos")+cav(pc(I),"Insights",I[1]?`${I[0]} de ${I[1]} insights`:"sem insights");
  ABERTO=-1;desenha()};
</script>__TUTORIAL____VERSAOSNIP____SOSMOD__</body></html>"""

RANK_MIN_MEU = 3       # mesmo mínimo do ranking do painel

def montar_pessoais(membros, eventos, eventos_ins, gerado):
    """Uma página por irmão (fora do comando) com e-mail: só os dados dele, médias e posições sem nomes."""
    PESSOAIS.clear()
    ano = str(datetime.date.today().year)
    ent = {}
    for evs in (eventos, eventos_ins):
        for e in evs:
            for i, v in enumerate(e["p"]):
                if v not in (None, "") and (i not in ent or e["d"] < ent[i]):
                    ent[i] = e["d"]
    def conta(v, i, d):
        if d > datetime.date.today().isoformat():
            return None          # evento que ainda não aconteceu: fora das estatísticas até o dia dele
        if v == "-":
            return 0
        if v in (None, ""):
            return 0 if (i in ent and d >= ent[i]) else None
        return 1 if v == 1 else 0
    def soma(evs, i, pref):
        p = t = 0
        for e in evs:
            if not e["d"].startswith(pref):
                continue
            w = conta(e["p"][i], i, e["d"])
            if w is not None:
                t += 1; p += (w == 1)
        return [p, t]
    grupos = [grupo_de(m.get("c")) for m in membros]
    est = [(soma(eventos, i, ano), soma(eventos_ins, i, ano)) for i in range(len(membros))]
    pct = lambda a: arred(a[0] * 100 / a[1]) if a[1] else None
    def media(idx, k):
        p = sum(est[i][k][0] for i in idx); t = sum(est[i][k][1] for i in idx)
        return arred(p * 100 / t) if t else None
    def posicao(i, idx):
        eleg = [j for j in idx if est[j][0][1] >= RANK_MIN_MEU]
        if i not in eleg:
            return None
        ordem = sorted(eleg, key=lambda j: (-(est[j][0][0] / est[j][0][1]), -est[j][0][0], membros[j].get("o", 999)))
        return [ordem.index(i) + 1, len(ordem)]
    todos = list(range(len(membros)))
    autoriz = {a.lower() for a in AUTORIZADOS + ADM_EXTRA}
    meses = sorted({e["d"][:7] for e in eventos + eventos_ins if e["d"].startswith(ano)})
    # próximos aniversários (Meu painel): dia/mês e apelido; o telefone só de quem faz aniversário hoje ou amanhã (para os parabéns)
    anivs_lista = []
    try:
        hj = datetime.date.today()
        for mm in membros:
            if not mm.get("nasc"):
                continue
            nd = datetime.date.fromisoformat(mm["nasc"])
            try:
                prox = nd.replace(year=hj.year)
            except ValueError:
                prox = datetime.date(hj.year, 3, 1)   # 29/02 em ano comum
            if prox < hj:
                try:
                    prox = nd.replace(year=hj.year + 1)
                except ValueError:
                    prox = datetime.date(hj.year + 1, 3, 1)
            x = {"n": mm["n"], "md": prox.strftime("%m-%d"), "dias": (prox - hj).days}
            tel = "".join(ch for ch in str(mm.get("tel") or "") if ch.isdigit())
            if x["dias"] <= 1 and tel:
                x["tel"] = tel if tel.startswith("55") else "55" + tel
            anivs_lista.append(x)
        anivs_lista = sorted(anivs_lista, key=lambda x: x["dias"])[:6]
    except Exception as e:
        print(f"Aviso: próximos aniversários não montados ({type(e).__name__}: {e}).")
    for i, m in enumerate(membros):
        email = EMAIL_APELIDO_INV.get(m["n"].lower(), "")
        if "@" not in email or (m["n"].lower() in autoriz and m["n"].lower() not in {a.lower() for a in ACESSOS_SOMENTE}):
            continue
        mesmo = [j for j in todos if grupos[j] == grupos[i]]
        def historia(k):
            """Linha do tempo do irmão (todos os anos): a primeira presença, os marcos de presença e o total por ano."""
            try:
                pres = [e for e in sorted(eventos, key=lambda x: x["d"]) if conta(e["p"][k], k, e["d"]) == 1]
                anos = {}
                for e in eventos:
                    c_ = conta(e["p"][k], k, e["d"])
                    if c_ is None:
                        continue
                    a_ = anos.setdefault(e["d"][:4], [0, 0]); a_[1] += 1; a_[0] += 1 if c_ == 1 else 0
                return {"pri": [pres[0]["d"], pres[0]["e"]] if pres else None,
                        "marcos": [[n_, pres[n_ - 1]["d"], pres[n_ - 1]["e"]] for n_ in (10, 25, 50, 100, 150, 200) if len(pres) >= n_],
                        "tot": len(pres), "anos": anos}
            except Exception:
                return None
        lista = lambda evs: [[e["d"], e["e"], conta(e["p"][i], i, e["d"])] + ([JUST[(e["d"], e["e"].strip().lower())][m["n"]]]
                              if m["n"] in JUST.get((e["d"], e["e"].strip().lower()), {}) else []) for e in sorted(evs, key=lambda x: x["d"], reverse=True)
                             if e["d"].startswith(ano) and conta(e["p"][i], i, e["d"]) is not None]
        dados = {"email": email, "push": MEU_PUSH, "n": m["n"], "c": m.get("c", ""), "ent": m.get("ent", ""), "reinst": AVISO_REINSTALAR if AVISO_REINSTALAR.get("quem") == "todos" else None,
                 "sede": str(DIV.get("sede_endereco", "")), "sede_n": str(DIV.get("sede_local", "")),
                 "mens": ({"st": PAG_MEU.get("st", {}).get(m["n"], ""), "lk": PAG_MEU.get("lk", {}).get(m["n"], ""), "venc": PAG_MEU.get("venc", ""),
                           "adm": next((x["n"] for x in membros if str(x.get("c", "")).strip().lower().startswith("adm")), "")} if PAG_MEU else None),
                 "cmd": [{"n": x["n"], "c": x.get("c", "")} for x in membros if x["n"].lower() in {a.lower() for a in AUTORIZADOS + ADM_EXTRA} and x["n"] != m["n"]], "g": grupos[i], "div": DIV["nome"], "reg": DIV["regional"], "ano": ano, "ger": gerado,
                 "min": RANK_MIN_MEU, "pres": est[i][0], "ins": est[i][1],
                 "mf": [media(mesmo, 0), media(mesmo, 1)], "md": [media(todos, 0), media(todos, 1)],
                 "pf": posicao(i, mesmo), "pd": posicao(i, todos),
                 "meses": [{"ym": ym, "p": soma(eventos, i, ym), "i": soma(eventos_ins, i, ym)} for ym in meses],
                 "ev": lista(eventos), "iv": lista(eventos_ins), "ru": ACESSOS_URL,
                 "anivs": anivs_lista,
                 "hist": historia(i), "mapa": MAPA_INFO,
                 "enq": ({"id": ENQ["id"], "p": ENQ["p"], "ops": ENQ["ops"], "meu": ENQ["votos"].get(email.strip().lower(), ("", ""))[1]}
                         if ENQ and ENQ.get("aberta") else None),
                 "prox": [{"d": e["d"], "e": e["e"], "tk": token_rsvp(email, e["d"], e["e"]),
                           **({k: INFO_EV[(e["d"], e["e"].strip().lower())][k] for k in ("hora", "local", "endereco")}
                              if (e["d"], e["e"].strip().lower()) in INFO_EV else {}),
                           "clima": clima_evento(e["d"], lugar_evento(e["d"], e["e"])[1], lugar_evento(e["d"], e["e"])[0]),
                           "mapa": lugar_evento(e["d"], e["e"])[0],
                           "st": "sim" if e["p"][i] == 1 else "nao" if e["p"][i] == 0 else ""}
                          for e in sorted(eventos, key=lambda x: x["d"]) if e["d"] >= datetime.date.today().isoformat()],
                 "insp": (lambda ui: {"d": ui["d"], "e": ui["e"], "tk": token_rsvp(email, ui["d"], ui["e"])}
                          if ui and ui["d"] >= (datetime.date.today() - datetime.timedelta(days=10)).isoformat() and ui["p"][i] != 1 else None)(
                          max([e for e in eventos_ins if e["d"] <= datetime.date.today().isoformat()], key=lambda x: x["d"], default=None))}
        PESSOAIS.append((email, MEU.replace("__DADOS__", json.dumps(dados, ensure_ascii=False)).replace("__TUTORIAL__", TUTORIAL_APP).replace("__VERSAOSNIP__", snip_versao(False))
                          .replace("__SOSMOD__", f'<script>window.SOS_URL={json.dumps(OTP_URL if otp_ativo() else "")};window.SOS_DIV={json.dumps(str(DIV.get("nome", "")) + " " + str(DIV.get("regional", "")))};</script>' + SOS_MODULO + MSG_MODULO + SELECT_MODULO + PTR_MODULO + REINST_MODULO)))
    print(f"Meu painel: {len(PESSOAIS)} irmão(s) com e-mail (fora do comando) podem ver os próprios dados.")

TUTORIAL_APP = r"""<style>
.tut-abrir{display:block;margin:12px auto 0;background:transparent;border:1px solid rgba(236,231,220,.25);color:#ECE7DC;border-radius:999px;padding:9px 18px;font-weight:700;font-size:14px;cursor:pointer}
#tut{position:fixed;inset:0;z-index:9999;background:rgba(0,0,0,.78);display:flex;align-items:flex-end;justify-content:center;padding:env(safe-area-inset-top,0) 0 0}
#tut[hidden]{display:none}#tut .cx{background:#161412;color:#ECE7DC;width:100%;max-width:560px;max-height:92vh;overflow:auto;border-radius:20px 20px 0 0;border:1px solid #2c2925;padding:16px 16px calc(18px + env(safe-area-inset-bottom,0));font-family:Barlow,"Segoe UI",Roboto,Arial,sans-serif}
#tut h2{margin:0 0 4px;font-size:20px}#tut .sb{color:#B8B2A6;font-size:13.5px;margin:0 0 12px}
#tut .tb{display:flex;gap:6px;margin:0 0 12px}#tut .tb button{flex:1;padding:10px;border-radius:999px;border:1px solid #3a3530;background:transparent;color:#ECE7DC;font-weight:700;font-size:14px}
#tut .tb button.on{background:#E9B949;color:#141210;border-color:#E9B949}
#tut .ps{display:flex;gap:12px;align-items:center;border-top:1px solid #2c2925;padding:12px 0}#tut .ps svg{flex:0 0 132px;width:132px;height:auto;border-radius:12px}
#tut .ps b.tt{display:block;font-size:15px;margin-bottom:3px}#tut .ps span b{color:#ECE7DC}#tut .ps span{font-size:13.5px;color:#B8B2A6;line-height:1.35}
#tut .n{display:inline-flex;width:24px;height:24px;border-radius:50%;background:#E9B949;color:#141210;font-weight:800;align-items:center;justify-content:center;font-size:13px;margin-right:6px}
#tut .fx{display:block;width:100%;margin-top:10px;padding:12px;border-radius:999px;border:1px solid #3a3530;background:transparent;color:#ECE7DC;font-weight:700}
#tut .ja{background:#1f8f4e;border:0;color:#fff;display:block;width:100%;padding:12px;border-radius:999px;font-weight:800;margin:0 0 10px}#tut .ja[hidden],#tut [hidden]{display:none!important}
</style>
<div id="tut" hidden onclick="if(event.target===this)tutFechar()"><div class="cx">
<h2>📲 Instale o app</h2><p class="sb">Fica com ícone na tela inicial, abre em tela cheia e recebe os avisos da divisão.</p>
<button class="ja" id="tutJa" hidden>⬇️ Instalar agora</button>
<div class="tb"><button data-so="and">🤖 Android</button><button data-so="ios">🍎 iPhone</button></div>
<div id="tutAnd">
 <div class="ps"><svg viewBox="0 0 132 92"><rect width="132" height="92" rx="12" fill="#0d0c0b"/><rect x="8" y="10" width="116" height="18" rx="9" fill="#2a2724"/><text x="18" y="23" font-size="9" fill="#9c958a">🔒 …/adm.html</text><circle cx="112" cy="19" r="9" fill="none" stroke="#E9B949" stroke-width="2.5"/><text x="109" y="23" font-size="12" font-weight="800" fill="#ECE7DC">⋮</text><rect x="8" y="36" width="116" height="48" rx="8" fill="#1d1b19"/></svg>
  <div><b class="tt"><span class="n">1</span>Abra o link no Chrome</b><span>Toque nos <b>três pontinhos ⋮</b>, no canto de cima à direita.</span></div></div>
 <div class="ps"><svg viewBox="0 0 132 92"><rect width="132" height="92" rx="12" fill="#0d0c0b"/><rect x="40" y="6" width="86" height="80" rx="8" fill="#2a2724"/><text x="48" y="22" font-size="8.5" fill="#9c958a">Nova guia</text><text x="48" y="36" font-size="8.5" fill="#9c958a">Favoritos</text><rect x="44" y="43" width="78" height="17" rx="6" fill="rgba(233,185,73,.22)" stroke="#E9B949" stroke-width="1.5"/><text x="48" y="55" font-size="8.5" font-weight="700" fill="#ECE7DC">⬇ Instalar app</text><text x="48" y="74" font-size="8.5" fill="#9c958a">Configurações</text></svg>
  <div><b class="tt"><span class="n">2</span>Toque em “Instalar app”</b><span>Em alguns celulares aparece <b>“Adicionar à tela inicial”</b>.</span></div></div>
 <div class="ps"><svg viewBox="0 0 132 92"><rect width="132" height="92" rx="12" fill="#0d0c0b"/><rect x="14" y="18" width="104" height="58" rx="10" fill="#2a2724"/><text x="24" y="36" font-size="9" font-weight="700" fill="#ECE7DC">Instalar app?</text><text x="24" y="50" font-size="8" fill="#9c958a">Insanos</text><rect x="72" y="56" width="38" height="14" rx="7" fill="#E9B949"/><text x="78" y="66" font-size="8" font-weight="800" fill="#141210">Instalar</text></svg>
  <div><b class="tt"><span class="n">3</span>Confirme em “Instalar”</b><span>O celular cria o ícone sozinho.</span></div></div>
 <div class="ps"><svg viewBox="0 0 132 92"><rect width="132" height="92" rx="12" fill="#0d0c0b"/><g fill="#2a2724"><rect x="14" y="14" width="22" height="22" rx="6"/><rect x="44" y="14" width="22" height="22" rx="6"/><rect x="74" y="14" width="22" height="22" rx="6"/><rect x="14" y="48" width="22" height="22" rx="6"/></g><rect x="42" y="46" width="26" height="26" rx="7" fill="#ECE7DC" stroke="#E9B949" stroke-width="2"/><text x="47" y="65" font-size="14">💀</text><text x="40" y="84" font-size="7" fill="#ECE7DC">Insanos</text></svg>
  <div><b class="tt"><span class="n">4</span>Pronto!</b><span>Abra pelo <b>ícone do Insanos</b> na tela inicial e entre com o seu e-mail.</span></div></div>
</div>
<div id="tutIos" hidden>
 <div class="ps"><svg viewBox="0 0 132 92"><rect width="132" height="92" rx="12" fill="#0d0c0b"/><rect x="8" y="8" width="116" height="50" rx="8" fill="#1d1b19"/><rect x="8" y="64" width="116" height="20" rx="8" fill="#2a2724"/><text x="18" y="78" font-size="10" fill="#9c958a">‹  ›</text><circle cx="66" cy="74" r="9" fill="none" stroke="#E9B949" stroke-width="2.5"/><path d="M66 69v9M62.5 72l3.5-3.5 3.5 3.5" stroke="#ECE7DC" stroke-width="1.6" fill="none"/><rect x="61" y="74" width="10" height="7" rx="1.5" fill="none" stroke="#ECE7DC" stroke-width="1.4"/></svg>
  <div><b class="tt"><span class="n">1</span>Abra o link no Safari</b><span>Toque no botão <b>Compartilhar</b> (o quadrado com a seta para cima), embaixo no meio.</span></div></div>
 <div class="ps"><svg viewBox="0 0 132 92"><rect width="132" height="92" rx="12" fill="#0d0c0b"/><rect x="10" y="8" width="112" height="78" rx="10" fill="#2a2724"/><text x="18" y="24" font-size="8.5" fill="#9c958a">Copiar</text><text x="18" y="38" font-size="8.5" fill="#9c958a">Adicionar aos Favoritos</text><rect x="13" y="45" width="106" height="17" rx="6" fill="rgba(233,185,73,.22)" stroke="#E9B949" stroke-width="1.5"/><text x="17" y="57" font-size="7.4" font-weight="700" fill="#ECE7DC">⊞ Adicionar à Tela de Início</text><text x="18" y="76" font-size="8.5" fill="#9c958a">Buscar na Página</text></svg>
  <div><b class="tt"><span class="n">2</span>“Adicionar à Tela de Início”</b><span>Role a lista para baixo até achar essa opção.</span></div></div>
 <div class="ps"><svg viewBox="0 0 132 92"><rect width="132" height="92" rx="12" fill="#0d0c0b"/><rect x="8" y="8" width="116" height="22" rx="8" fill="#2a2724"/><text x="16" y="23" font-size="8.5" fill="#9c958a">Cancelar</text><rect x="76" y="12" width="44" height="14" rx="7" fill="rgba(233,185,73,.22)" stroke="#E9B949" stroke-width="1.5"/><text x="81" y="22" font-size="8" font-weight="800" fill="#ECE7DC">Adicionar</text><rect x="16" y="40" width="22" height="22" rx="6" fill="#ECE7DC"/><text x="20" y="56" font-size="12">💀</text><text x="46" y="55" font-size="9" fill="#ECE7DC">Insanos</text></svg>
  <div><b class="tt"><span class="n">3</span>Toque em “Adicionar”</b><span>No canto de cima, à direita.</span></div></div>
 <div class="ps"><svg viewBox="0 0 132 92"><rect width="132" height="92" rx="12" fill="#0d0c0b"/><g fill="#2a2724"><rect x="14" y="14" width="22" height="22" rx="6"/><rect x="44" y="14" width="22" height="22" rx="6"/><rect x="74" y="14" width="22" height="22" rx="6"/><rect x="14" y="48" width="22" height="22" rx="6"/></g><rect x="42" y="46" width="26" height="26" rx="7" fill="#ECE7DC" stroke="#E9B949" stroke-width="2"/><text x="47" y="65" font-size="14">💀</text><text x="40" y="84" font-size="7" fill="#ECE7DC">Insanos</text></svg>
  <div><b class="tt"><span class="n">4</span>Abra pelo ícone</b><span>Entre de novo com o seu e-mail. No iPhone, <b>os avisos só funcionam pelo app instalado</b>.</span></div></div>
</div>
<button class="fx" onclick="tutFechar()">Fechar</button></div></div>
<script>
var tutEv=null;window.addEventListener("beforeinstallprompt",function(e){e.preventDefault();tutEv=e;var b=document.getElementById("tutJa"),a=document.getElementById("tutAnd");if(b)b.hidden=!(a&&!a.hidden)});
function tutSo(so){document.getElementById("tutAnd").hidden=so!=="and";var j=document.getElementById("tutJa");if(j)j.hidden=!(tutEv&&so==="and");document.getElementById("tutIos").hidden=so!=="ios";
  document.querySelectorAll("#tut .tb button").forEach(function(b){b.classList.toggle("on",b.dataset.so===so)})}
function tutAbrir(){tutSo(/iPhone|iPad|iPod/.test(navigator.userAgent)?"ios":"and");document.getElementById("tut").hidden=false}
function tutFechar(){document.getElementById("tut").hidden=true}
function tutInstalado(){return matchMedia("(display-mode: standalone)").matches||navigator.standalone===true}
document.querySelectorAll("#tut .tb button").forEach(function(b){b.onclick=function(){tutSo(b.dataset.so)}});
document.getElementById("tutJa").onclick=function(){if(tutEv){tutEv.prompt();tutEv.userChoice.then(function(){tutEv=null;tutFechar()})}};
</script>"""

SOS_MODULO = r"""<style>
.sosm-veu{position:fixed;inset:0;z-index:9998;background:rgba(0,0,0,.72);display:flex;align-items:flex-end;justify-content:center}
.sosm-cx{background:#1b1714;color:#ECE7DC;width:100%;max-width:520px;max-height:92vh;overflow:auto;border-radius:20px 20px 0 0;border-top:3px solid #c62828;padding:18px 16px calc(18px + env(safe-area-inset-bottom,0px));font-family:Barlow,"Segoe UI",Roboto,Arial,sans-serif}
.sosm-cx h3{margin:0 0 10px;font-size:19px}.sosm-cx p{margin:8px 0;font-size:14px;line-height:1.4;color:#d9d2c4}
.sosm-g{display:grid;grid-template-columns:1fr 1fr;gap:8px}
.sosm-t{padding:14px 6px;border-radius:12px;border:1px solid rgba(236,231,220,.25);background:rgba(0,0,0,.35);color:#ECE7DC;font-weight:800;font-size:15px}
.sosm-tel{display:flex;gap:6px}.sosm-tel a{flex:1;text-align:center;padding:9px 4px;border-radius:10px;background:#ECE7DC;color:#141210;font-weight:800;font-size:12.5px;text-decoration:none}
.sosm-b{display:block;width:100%;margin-top:10px;padding:13px;border-radius:12px;border:1px solid rgba(236,231,220,.3);background:transparent;color:#ECE7DC;font-weight:700;font-size:15px}
.sosm-b.verm{background:#c62828;border-color:#c62828;color:#fff;font-weight:900;font-size:16px}.sosm-b.verde{background:#1f8f4e;border-color:#1f8f4e;color:#fff;font-weight:800}
</style>
<script>
// SOS dentro do painel (ADM e Meu painel): o mesmo da tela de entrada. Quem pede é identificado pelo celular
// (o pedaço da chave guardado depois do código por e-mail), e o alerta vai para todos os irmãos.
(function(){
  const cache=()=>{try{return JSON.parse(localStorage.getItem("otp_s")||"{}")}catch(_){return {}}};
  const quem=()=>{const c=cache();let b=null;for(const e in c)if(c[e].exp>Date.now()&&(!b||c[e].exp>c[b].exp))b=e;return b?{email:b,s:c[b].s}:null};
  const gps=()=>new Promise(r=>{if(!navigator.geolocation)return r(null);navigator.geolocation.getCurrentPosition(p=>r({lat:+p.coords.latitude.toFixed(5),lng:+p.coords.longitude.toFixed(5)}),()=>r(null),{enableHighAccuracy:true,timeout:12000,maximumAge:0})});
  let veu=null;const fecha=()=>{if(veu){veu.remove();veu=null}};
  const abre=html=>{fecha();veu=document.createElement("div");veu.className="sosm-veu";veu.innerHTML=`<div class="sosm-cx" role="dialog" aria-modal="true">${html}</div>`;
    veu.onclick=e=>{if(e.target===veu)fecha()};document.body.appendChild(veu);return veu};
  const tel='<p style="margin-top:14px">⚠️ <b>Risco de vida?</b> Ligue primeiro:</p><div class="sosm-tel"><a href="tel:190">190 Polícia</a><a href="tel:192">192 SAMU</a><a href="tel:193">193 Bombeiros</a></div>';
  window.fichaMostrar=function(em){let o=null;try{o=JSON.parse(localStorage.getItem("moto_"+em)||"{}").emg}catch(_){}
    if(!o){alert("Você ainda não preencheu a ficha de emergência (Minha moto › Emergência).");return}
    const es2=t=>String(t||"").replace(/[&<>"]/g,c=>({"&":"&amp;","<":"&lt;",">":"&gt;",'"':"&quot;"}[c]));
    const v=document.createElement("div");v.className="ficha-tela";
    v.innerHTML=`<div class="ficha-cx"><div class="ficha-top">🩺 FICHA DE EMERGÊNCIA</div><h2>${es2(o.nome||"")}</h2>
      ${o.sangue?`<div class="ficha-sg"><span>Tipo sanguíneo</span><b>${es2(o.sangue)}</b></div>`:""}
      ${o.alerg?`<div class="ficha-al"><span>⚠️ ALERGIAS</span><b>${es2(o.alerg)}</b></div>`:""}
      ${o.rem?`<div class="ficha-i"><span>💊 Remédios de uso contínuo</span><b>${es2(o.rem)}</b></div>`:""}
      ${o.obs?`<div class="ficha-i"><span>ℹ️ Outras informações</span><b>${es2(o.obs)}</b></div>`:""}
      ${o.plano?`<div class="ficha-i"><span>🏥 Plano de saúde</span><b>${es2(o.plano)}</b></div>`:""}
      ${(o.c||[]).filter(c=>c&&c.t).map(c=>`<a class="ficha-lig" href="tel:${es2(String(c.t).replace(/[^0-9+]/g,""))}">📞 Ligar para ${es2(c.n||"o contato")}${c.p?" ("+es2(c.p)+")":""}<small>${es2(c.t)}</small></a>`).join("")}
      <button type="button" class="ficha-x">Fechar</button></div>`;
    v.querySelector(".ficha-x").onclick=()=>v.remove();document.body.appendChild(v)};
  window.fichaSOS=function(em){try{const o=JSON.parse(localStorage.getItem("moto_"+em)||"{}").emg;if(!o||o.enviar===false)return "";
    const f={};["sangue","alerg","rem","obs","plano"].forEach(k=>{if(o[k])f[k]=String(o[k]).slice(0,120)});const c=(o.enviarCt===false?[]:(o.c||[])).filter(x=>x&&x.t).slice(0,2).map(x=>({n:String(x.n||"").slice(0,40),p:String(x.p||"").slice(0,20),t:String(x.t).slice(0,20)}));if(c.length)f.c=c;
    return Object.keys(f).length?JSON.stringify(f):""}catch(_){return ""}};

  if(!window.fichaZap)window.fichaZap=function(em){try{const j=(window.fichaSOS?fichaSOS(em):"");if(!j)return "";const f=JSON.parse(j),N=String.fromCharCode(10),L=[];
    if(f.sangue)L.push("🩸 Tipo sanguíneo: "+f.sangue);if(f.alerg)L.push("⚠️ Alergia: "+f.alerg);if(f.rem)L.push("💊 Remédios: "+f.rem);if(f.obs)L.push("ℹ️ "+f.obs);if(f.plano)L.push("🏥 Plano: "+f.plano);
    (f.c||[]).forEach(c=>L.push("📞 Contato: "+(c.n||"")+(c.p?" ("+c.p+")":"")+" · "+c.t));return L.length?N+N+"🩺 *Ficha de emergência*"+N+L.join(N):""}catch(_){return ""}};
  window.abrirSOS=function(){
    const url=window.SOS_URL||"",q=quem();
    if(!url||!q){abre(`<h3>🆘 SOS</h3><p>Para usar o SOS por aqui, entre uma vez no app com o seu e-mail e o código que chega no e-mail, neste celular.</p>${tel}<button class="sosm-b" data-f>Fechar</button>`).querySelector("[data-f]").onclick=fecha;return}
    const v=abre(`<h3>🆘 O que aconteceu?</h3><div class="sosm-g">${[["pane","🛵 Pane/acidente"],["saude","🚑 Saúde"],["seguranca","⚠️ Segurança"],["outro","🆘 Outro"]].map(([k,t])=>`<button class="sosm-t" data-k="${k}">${t}</button>`).join("")}</div>${tel}<button class="sosm-b" data-f>Cancelar</button>`);
    v.querySelector("[data-f]").onclick=fecha;
    v.querySelectorAll("[data-k]").forEach(b=>b.onclick=()=>{const cat=b.dataset.k,nome=b.textContent.trim(),pg=gps();
      const c=abre(`<h3>Confirmar o SOS de ${nome}?</h3><p>Todos os irmãos com o app serão avisados, com a sua localização.</p><button class="sosm-b verm" data-s>🆘 Sim, enviar o SOS</button><button class="sosm-b" data-f>Cancelar</button>`);
      c.querySelector("[data-f]").onclick=fecha;
      c.querySelector("[data-s]").onclick=async()=>{const m=abre(`<h3>🆘 SOS</h3><p data-m>📍 Pegando a sua localização…</p>`).querySelector("[data-m]");
        const l=await pg,quando=new Date().toLocaleString("pt-BR",{day:"2-digit",month:"2-digit",hour:"2-digit",minute:"2-digit"});
        let nomeIrmao=localStorage.getItem("sos_nome")||"";
        const monta=n=>`🆘 *SOS · Insanos M.C.*\n${n?`*${n}*`:"Um irmão"} (Divisão ${window.SOS_DIV||""}) precisa de ajuda.\n${nome}\n`+(l?`📍 https://maps.google.com/?q=${l.lat},${l.lng}\n`:"📍 (sem localização)\n")+`🕒 ${quando}\nQuem estiver perto, por favor, dê um apoio! 🤝`+(window.fichaZap?fichaZap(q.email):"");
        let tx=monta(nomeIrmao);m.textContent="Enviando o alerta para os irmãos…";let ok=false;
        try{const r=await fetch(url,{method:"POST",headers:{"Content-Type":"text/plain;charset=utf-8"},body:JSON.stringify({tipo:"sos",email:q.email,s:q.s,categoria:cat,lat:l?l.lat:null,lng:l?l.lng:null,ficha:(window.fichaSOS?fichaSOS(q.email):"")})});
          const j=await r.json();ok=!!j.ok;if(j.nome){try{localStorage.setItem("sos_nome",j.nome)}catch(_){}tx=monta(j.nome)}if(!ok)m.textContent="⚠️ "+(j.erro||"Não consegui enviar o alerta.")}
        catch(e){m.textContent="⚠️ Não consegui enviar o alerta (sem internet?)."}
        try{await navigator.clipboard.writeText(tx)}catch(_){}
        const box=m.parentNode;if(ok)m.innerHTML="✅ <b>Alerta enviado!</b> Os irmãos recebem em 1 a 3 minutos. Agora mande também nos grupos:";else m.innerHTML+=" <b>Mande nos grupos:</b>";
        box.insertAdjacentHTML("beforeend",`<button class="sosm-b verde" data-z>📤 Mandar nos grupos (WhatsApp)</button><button class="sosm-b" data-c>📋 Copiar a mensagem</button>${tel}<button class="sosm-b" data-f>Fechar</button>`);
        box.querySelector("[data-z]").onclick=async()=>{if(navigator.share){try{await navigator.share({text:tx});return}catch(e){if(e&&e.name==="AbortError")return}}window.open("https://wa.me/?text="+encodeURIComponent(tx),"_blank")};
        box.querySelector("[data-c]").onclick=async e=>{try{await navigator.clipboard.writeText(tx);e.target.textContent="✅ Copiada: cole no WhatsApp"}catch(_){alert(tx)}};
        box.querySelector("[data-f]").onclick=fecha}})};
})();
</script>"""

MSG_MODULO = r"""<style>
.msgm-lbl{display:block;font-size:13px;color:#BDB7AA;margin:12px 0 6px}
.msgm-tabs{display:flex;gap:6px}.msgm-tabs button{flex:1;padding:10px 6px;border-radius:12px;border:1px solid rgba(236,231,220,.25);background:transparent;color:#ECE7DC;font:inherit;font-weight:700;font-size:14px}
.msgm-tabs button[aria-pressed="true"]{background:#ECE7DC;color:#141210}
.msgm-sel,.msgm-txt{width:100%;box-sizing:border-box;padding:11px;border-radius:12px;border:1px solid rgba(236,231,220,.25);background:rgba(0,0,0,.35);color:#ECE7DC;font:inherit;font-size:15px}
.msgm-txt{min-height:110px;resize:vertical}.msgm-cont{font-size:12px;color:#BDB7AA;text-align:right;margin-top:4px}
.msgm-it{padding:12px;border-radius:14px;background:rgba(236,231,220,.05);border:1px solid rgba(236,231,220,.1);margin-bottom:8px}
.msgm-it .q{display:flex;justify-content:space-between;gap:8px;font-size:12.5px;color:#BDB7AA}.msgm-it b{display:block;color:#ECE7DC;font-size:14.5px;margin-top:4px}
.msgm-it p{margin:6px 0 0;font-size:14.5px;line-height:1.4;white-space:pre-wrap}.msgm-it .rsp{margin-top:8px;font-size:13px;color:#E9B949}
.msgm-imp{display:flex;gap:10px;align-items:flex-start;margin-top:12px;padding:12px;border-radius:12px;border:1px solid rgba(233,185,73,.55);background:rgba(233,185,73,.1);font-size:14px;line-height:1.35;color:#ECE7DC}
.msgm-imp input{width:20px;height:20px;margin-top:2px;flex:0 0 auto}.msgm-imp b{color:#E9B949}
.msgm-it.imp{border-color:rgba(233,185,73,.55)}
.msgm-chips{display:flex;flex-wrap:wrap;gap:6px;margin-top:10px}.msgm-chips button{padding:7px 12px;border-radius:999px;border:1px solid rgba(236,231,220,.25);background:transparent;color:#ECE7DC;font:inherit;font-size:13px}
</style>
<script>
// Mensagens pelo app: o comando manda para todos, um grupo ou um irmão; o irmão manda para alguém do comando.
// Quem manda é identificado pelo celular (a mesma assinatura do SOS); as regras e o limite são conferidos na nuvem.
(function(){
  const cache=()=>{try{return JSON.parse(localStorage.getItem("otp_s")||"{}")}catch(_){return {}}};
  const quem=()=>{const c=cache();let b=null;for(const e in c)if(c[e].exp>Date.now()&&(!b||c[e].exp>c[b].exp))b=e;return b?{email:b,s:c[b].s}:null};
  const es=x=>String(x==null?"":x).replace(/[&<>"]/g,c=>({"&":"&amp;","<":"&lt;",">":"&gt;",'"':"&quot;"}[c]));
  const RESP=["✅ Recebido","📲 Te respondo no WhatsApp","📞 Me liga"];
  const post=async o=>{const r=await fetch(window.SOS_URL,{method:"POST",headers:{"Content-Type":"text/plain;charset=utf-8"},body:JSON.stringify(o)});return r.json()};
  let veu=null;const fecha=()=>{if(veu){veu.remove();veu=null}};
  const abre=html=>{fecha();veu=document.createElement("div");veu.className="sosm-veu";veu.innerHTML=`<div class="sosm-cx" style="border-top-color:#E9B949" role="dialog" aria-modal="true">${html}</div>`;
    veu.onclick=e=>{if(e.target===veu)fecha()};document.body.appendChild(veu);return veu.firstChild};
  const quando=iso=>{const d=new Date(iso),h=new Date(),ont=new Date(h-864e5),hm=d.toLocaleTimeString("pt-BR",{hour:"2-digit",minute:"2-digit"});
    return d.toDateString()===h.toDateString()?"hoje "+hm:d.toDateString()===ont.toDateString()?"ontem "+hm:d.toLocaleDateString("pt-BR",{day:"2-digit",month:"2-digit"})+" "+hm};
  window.abrirMensagens=async function(cfg){window.MSG_CFG=cfg=cfg||window.MSG_CFG||{};const q=quem();
    if(!window.SOS_URL||!q){abre(`<h3>📨 Mensagens</h3><p>Para usar as mensagens, entre uma vez no app com o seu e-mail e o código que chega no e-mail, neste celular.</p><button class="sosm-b" data-f>Fechar</button>`).querySelector("[data-f]").onclick=fecha;return}
    const cx=abre(`<h3>📨 Mensagens</h3><button class="sosm-b verde" data-n style="margin:0 0 12px">✉️ ${cfg.modo==="adm"?"Nova mensagem":"Mensagem para o comando"}</button><div data-h><p>Carregando…</p></div><button class="sosm-b" data-f>Fechar</button>`);
    cx.querySelector("[data-f]").onclick=fecha;cx.querySelector("[data-n]").onclick=()=>escrever(cfg);
    let j;try{j=await post({tipo:"msg_hist",email:q.email,s:q.s})}catch(_){j={ok:false,erro:"sem internet?"}}
    const h=cx.querySelector("[data-h]");if(!h)return;
    if(!j.ok){h.innerHTML=`<p>⚠️ Não consegui carregar as mensagens (${es(j.erro||"")}).</p>`;return}
    if(!j.itens.length){h.innerHTML='<p style="color:#BDB7AA">Nenhuma mensagem nos últimos 90 dias.</p>';return}
    h.innerHTML=j.itens.map(it=>{const jaResp=it.resp.some(r=>r.de==="Você"),imp=/^📌/.test(it.rot||"");
      return `<div class="msgm-it${imp?" imp":""}"><div class="q"><span>${imp?"📌 Importante · ":""}${it.minha?"Enviada":"Recebida"} · ${quando(it.q)}</span><span>${it.minha?(/^enviada/.test(it.sit)?"✅ "+es(it.sit)+(it.n?` · 👁️ leram ${it.nl||0} de ${it.n}`:""):/^recusada/.test(it.sit)?"⚠️ "+es(it.sit):"⏳ enviando…"):""}</span></div>
        <b>${it.minha?"Para: "+es(String(it.rot||it.dest).replace(/^📌\s*/,"")):"De "+es(it.de)}</b><p>${es(it.tx)}</p>
        ${it.resp.length?`<div class="rsp">↩️ ${it.resp.map(r=>es(r.de)+": "+es(r.tx)).join(" · ")}</div>`:""}
        ${!it.minha&&!jaResp?`<div class="msgm-chips">${RESP.map(r=>`<button data-r="${es(r)}" data-id="${es(it.id)}">${es(r)}</button>`).join("")}</div>`:""}</div>`}).join("");
    h.querySelectorAll("[data-r]").forEach(b=>b.onclick=async()=>{b.disabled=true;
      try{const r=await post({tipo:"msg_resp",email:q.email,s:q.s,id:b.dataset.id,resp:b.dataset.r});if(!r.ok)throw new Error(r.erro);
        b.parentNode.outerHTML=`<div class="rsp">↩️ Você: ${es(b.dataset.r)} (enviando…)</div>`}catch(e){b.disabled=false;alert("Não consegui responder agora ("+(e.message||e)+").")}})};
  window.msgEnviar=async function(para,rot,texto){const q=quem();if(!window.SOS_URL||!q)return {ok:false,erro:"sem identificação neste celular"};
    return post({tipo:"msg",email:q.email,s:q.s,para,rotulo:rot,texto})};
  window.msgNaoLidas=async function(cfg){try{if(navigator.clearAppBadge)navigator.clearAppBadge()}catch(_){}
    const q=quem();if(!window.SOS_URL||!q)return;let j;try{j=await post({tipo:"msg_hist",email:q.email,s:q.s})}catch(_){return}
    const novas=(j&&j.ok?j.itens:[]).filter(it=>!it.minha&&!it.lido).sort((a,b)=>(/^📌/.test(b.rot||"")?1:0)-(/^📌/.test(a.rot||"")?1:0));
    if(!novas.length||veu)return;
    const cx=abre(`<h3>📨 Você tem ${novas.length} mensage${novas.length>1?"ns novas":"m nova"}</h3>${novas.slice(0,3).map(it=>{const imp=/^📌/.test(it.rot||"");
      return `<div class="msgm-it${imp?" imp":""}"><div class="q"><span>${imp?"📌 Importante · ":""}${quando(it.q)}</span></div><b>De ${es(it.de)}</b><p>${es(it.tx)}</p></div>`}).join("")}
      <button class="sosm-b verde" data-a>Abrir as mensagens</button><button class="sosm-b" data-f>Depois</button>`);
    cx.parentNode.style.alignItems="center";cx.style.borderRadius="20px";
    cx.querySelector("[data-f]").onclick=fecha;cx.querySelector("[data-a]").onclick=()=>abrirMensagens(cfg)};
  function escrever(cfg){const q=quem();let modo=cfg.modo==="adm"?"todos":"um";
    const opts=cfg.modo==="adm"?null:(cfg.comando||[]);
    const cx=abre(`<h3>✉️ ${cfg.modo==="adm"?"Nova mensagem":"Mensagem para o comando"}</h3>
      ${cfg.modo==="adm"?`<label class="msgm-lbl">Para</label><div class="msgm-tabs"><button data-m="todos">📣 Todos</button><button data-m="grupo">👥 Grupo</button><button data-m="um">👤 Um irmão</button></div>`:""}
      <div data-alvo></div>
      <label class="msgm-lbl">Mensagem</label><textarea class="msgm-txt" maxlength="300" placeholder="Escreva a mensagem"></textarea><div class="msgm-cont">0 / 300</div>
      ${cfg.modo==="adm"?`<label class="msgm-imp"><input type="checkbox" data-imp><span><b>📌 Importante</b><br>Fica fixa na tela, vibra forte, com imagem, e avisa de novo quem não abrir em 30 min e em 2 h.</span></label>`:""}
      <button class="sosm-b verde" data-env>✉️ Enviar</button>
      <p style="font-size:12.5px;color:#BDB7AA;margin:8px 0 0">${cfg.modo==="adm"?"Chega em 1 a 3 minutos para quem tem o app com os avisos ativos.":"Você pode mandar até 5 mensagens por dia. As respostas aparecem em 📨 Mensagens."}</p>
      <button class="sosm-b" data-f>Cancelar</button>`);
    cx.querySelector("[data-f]").onclick=()=>abrirMensagens(cfg);
    const tx=cx.querySelector("textarea"),cont=cx.querySelector(".msgm-cont"),env=cx.querySelector("[data-env]"),alvo=cx.querySelector("[data-alvo]");
    tx.oninput=()=>{cont.textContent=tx.value.length+" / 300"};
    const destino=()=>{if(cfg.modo!=="adm"){const s=alvo.querySelector("select");return s&&s.value?{para:[s.value],rot:s.value}:null}
      if(modo==="todos")return {para:(cfg.todos||[]),rot:"📣 Todos"};
      const s=alvo.querySelector("select");if(!s||!s.value)return null;
      if(modo==="grupo"){const g=(cfg.grupos||{})[s.value]||[];return {para:g,rot:"👥 "+s.value}}
      return {para:[s.value],rot:s.value}};
    const impEl=cx.querySelector("[data-imp]"),ehImp=()=>!!(impEl&&impEl.checked);
    const conta=()=>{const d=destino(),pre=ehImp()?"📌 Enviar como importante":"✉️ Enviar";env.textContent=d&&d.para.length>1?`${pre} para ${d.para.length} irmãos`:pre};
    if(impEl)impEl.onchange=conta;
    const desenha=()=>{cx.querySelectorAll("[data-m]").forEach(b=>b.setAttribute("aria-pressed",b.dataset.m===modo));
      if(cfg.modo!=="adm")alvo.innerHTML=`<label class="msgm-lbl">Para quem</label><select class="msgm-sel"><option value="">Escolha…</option>${opts.map(o=>`<option value="${es(o.n)}">${es(o.n)}${o.c?" · "+es(o.c):""}</option>`).join("")}</select>`;
      else if(modo==="grupo")alvo.innerHTML=`<label class="msgm-lbl">Grupo</label><select class="msgm-sel"><option value="">Escolha…</option>${Object.entries(cfg.grupos||{}).filter(([,v])=>v.length).map(([k,v])=>`<option value="${es(k)}">${es(k)} (${v.length})</option>`).join("")}</select>`;
      else if(modo==="um")alvo.innerHTML=`<label class="msgm-lbl">Irmão</label><select class="msgm-sel"><option value="">Escolha…</option>${(cfg.nomes||[]).map(o=>`<option value="${es(o.n)}">${es(o.n)}${o.c?" · "+es(o.c):""}</option>`).join("")}</select>`;
      else alvo.innerHTML=`<p style="font-size:13px;color:#BDB7AA;margin:10px 0 0">Vai para os ${(cfg.todos||[]).length} irmãos ativos.</p>`;
      const s=alvo.querySelector("select");if(s)s.onchange=conta;conta()};
    cx.querySelectorAll("[data-m]").forEach(b=>b.onclick=()=>{modo=b.dataset.m;desenha()});desenha();
    env.onclick=async()=>{const d=destino(),t=tx.value.trim();if(!d||!d.para.length){alert("Escolha para quem enviar.");return}if(!t){alert("Escreva a mensagem.");return}
      if(d.para.length>1&&!confirm(`Enviar para ${d.para.length} irmãos?`))return;
      env.disabled=true;env.textContent="Enviando…";
      try{const r=await post({tipo:"msg",email:q.email,s:q.s,para:d.para,rotulo:(ehImp()?"📌 ":"")+d.rot,texto:t});if(!r.ok)throw new Error(r.erro);
        const c2=abre(`<h3>✅ Mensagem enviada</h3><p>Chega em 1 a 3 minutos para quem tem o app com os avisos ativos. As respostas aparecem em 📨 Mensagens.</p><button class="sosm-b" data-f>Ver mensagens</button>`);
        c2.querySelector("[data-f]").onclick=()=>abrirMensagens(cfg)}
      catch(e){env.disabled=false;conta();alert("Não consegui enviar agora ("+(e.message||e)+").")}}}
})();
</script>"""

PTR_MODULO = r"""<style>
#ptr{position:fixed;left:50%;top:0;z-index:9000;transform:translate(-50%,-60px);display:flex;align-items:center;gap:8px;padding:9px 16px;border-radius:999px;
  background:rgba(40,37,34,.85);-webkit-backdrop-filter:blur(14px);backdrop-filter:blur(14px);color:#ECE7DC;font-size:14px;font-weight:700;box-shadow:0 8px 24px rgba(0,0,0,.35);pointer-events:none}
#ptr i{font-style:normal;display:inline-block;transition:transform .2s}#ptr.pronto i{transform:rotate(180deg)}#ptr.vai i{animation:ptrGira .8s linear infinite}
@keyframes ptrGira{to{transform:rotate(360deg)}}
</style>
<script>
// puxar a tela para baixo (no topo) atualiza o painel; a entrada é feita de novo sozinha, nesta sessão
(function(){if(!("ontouchstart" in window))return;let g=null;const el=document.createElement("div");el.id="ptr";el.innerHTML="<i>↓</i><span>Puxe para atualizar</span>";
  document.addEventListener("DOMContentLoaded",()=>document.body.appendChild(el));if(document.body)document.body.appendChild(el);
  const aberto=()=>!!document.querySelector(".fa-veu,.sosm-veu,.selm,#folhaAt");
  document.addEventListener("touchstart",ev=>{g=null;if(ev.touches.length!==1||window.scrollY>2||aberto())return;g={y:ev.touches[0].clientY,x:ev.touches[0].clientX,dy:0}},{passive:true});
  document.addEventListener("touchmove",ev=>{if(!g)return;const t=ev.touches[0],dy=t.clientY-g.y,dx=t.clientX-g.x;
    if(dy<=0||Math.abs(dx)>dy||window.scrollY>2){if(dy<0||Math.abs(dx)>Math.abs(dy))g=null;el.style.transform="translate(-50%,-60px)";return}
    g.dy=dy;const p=Math.min(dy*0.45,90);el.style.transition="none";el.style.transform=`translate(-50%,${p-50}px)`;
    el.classList.toggle("pronto",dy>140);el.querySelector("span").textContent=dy>140?"Solte para atualizar":"Puxe para atualizar"},{passive:true});
  document.addEventListener("touchend",()=>{if(!g)return;const dy=g.dy;g=null;el.style.transition="transform .2s ease";
    if(dy>140){el.classList.remove("pronto");el.classList.add("vai");el.querySelector("i").textContent="↻";el.querySelector("span").textContent="Atualizando…";el.style.transform="translate(-50%,16px)";
      let em="";try{em=sessionStorage.getItem("painel_email")||""}catch(_){}if(!em&&typeof D!=="undefined"&&D&&D.email)em=D.email;
      try{if(em)sessionStorage.setItem("reentrar",em)}catch(_){}setTimeout(()=>location.reload(),350)}
    else el.style.transform="translate(-50%,-60px)"},{passive:true});
})();
</script>"""

REINST_MODULO = r"""<style>
#reinst{position:fixed;left:10px;right:10px;top:calc(10px + env(safe-area-inset-top,0px));z-index:9500;padding:12px 14px;border-radius:16px;background:rgba(40,34,24,.94);
  -webkit-backdrop-filter:blur(14px);backdrop-filter:blur(14px);border:1px solid rgba(233,185,73,.6);color:#ECE7DC;box-shadow:0 12px 30px rgba(0,0,0,.45);font-size:14px;line-height:1.4}
#reinst b{color:#E9B949}#reinst .bts{display:flex;flex-wrap:wrap;gap:6px;margin-top:10px}
#reinst button{padding:8px 12px;border-radius:999px;border:1px solid rgba(236,231,220,.3);background:rgba(236,231,220,.06);color:#ECE7DC;font:inherit;font-weight:700;font-size:13px}
#reinst button.ok{background:#E9B949;border-color:#E9B949;color:#141210}
#reinst ol{margin:8px 0 0;padding-left:20px}
</style>
<script>
// aviso de reinstalar o app (quando uma atualização muda o que o Android lê só na instalação)
window.reinstAviso=function(R){try{
  if(!R||!R.v||!/Android/.test(navigator.userAgent)||!(matchMedia("(display-mode: standalone)").matches))return;
  if(+(localStorage.getItem("reinst_v")||0)>=R.v)return;
  const ate=+(localStorage.getItem("reinst_depois")||0);if(ate>Date.now())return;
  if(document.getElementById("reinst"))return;
  const el=document.createElement("div");el.id="reinst";el.setAttribute("role","status");
  el.innerHTML=`<b style="display:block;font-size:15px;margin-bottom:4px">📲 Atualização importante do app</b>Houve uma atualização de <b>segurança e de funcionalidades</b>. Por favor, <b>desinstale e instale o app de novo</b> (leva 1 minuto).<div class="pass" hidden><ol>
    <li>Na tela inicial, <b>toque e segure</b> o ícone do app › <b>Desinstalar</b> (pode deixar "excluir os dados" desmarcado).</li>
    <li>Abra o <b>Chrome</b> no endereço do app, toque em <b>⋮</b> › <b>Instalar app</b>.</li><li>Abra o app pelo ícone novo e entre normalmente.</li></ol></div>
    <div class="bts"><button data-a="como">Como fazer</button><button data-a="ok" class="ok">Já reinstalei</button><button data-a="dep">Depois</button></div>`;
  document.body.appendChild(el);
  el.onclick=e=>{const b=e.target.closest("[data-a]");if(!b)return;const a=b.dataset.a;
    if(a==="como"){const p=el.querySelector(".pass");p.hidden=!p.hidden;b.textContent=p.hidden?"Como fazer":"Esconder"}
    else if(a==="ok"){try{localStorage.setItem("reinst_v",String(R.v))}catch(_){}el.remove()}
    else{try{localStorage.setItem("reinst_depois",String(Date.now()+864e5))}catch(_){}el.remove()}}}catch(_){}};
</script>"""

SELECT_MODULO = r"""<style>
.selm-veu{position:fixed;inset:0;z-index:10000;background:rgba(0,0,0,.62);display:flex;align-items:flex-end;justify-content:center;animation:selmVeu .18s ease both}
.selm{width:100%;max-width:560px;max-height:82vh;display:flex;flex-direction:column;background:#1b1714;color:#ECE7DC;border-radius:22px 22px 0 0;border-top:3px solid #E9B949;
  padding:10px 0 calc(14px + env(safe-area-inset-bottom,0px));box-shadow:0 -12px 40px rgba(0,0,0,.55);animation:selmSobe .24s cubic-bezier(.2,.9,.3,1) both;font-family:inherit}
.selm-veu.fora{animation:selmVeuFora .16s ease both}.selm-veu.fora .selm{animation:selmDesce .16s ease-in both}
@keyframes selmVeu{from{background:rgba(0,0,0,0)}}@keyframes selmVeuFora{to{background:rgba(0,0,0,0)}}
@keyframes selmSobe{from{transform:translateY(100%)}to{transform:none}}@keyframes selmDesce{to{transform:translateY(100%)}}
.selm .alca{width:44px;height:5px;border-radius:3px;background:rgba(236,231,220,.3);margin:0 auto 10px;flex:none}
.selm h3{margin:0 18px 10px;font-size:18px;flex:none}
.selm-busca{flex:none;margin:0 16px 8px;padding:11px 14px;border-radius:12px;background:rgba(0,0,0,.45);border:1px solid rgba(236,231,220,.2);color:#ECE7DC;font:inherit;font-size:16px}
.selm-l{overflow:auto;-webkit-overflow-scrolling:touch;flex:1 1 auto}
.selm-l button{display:flex;align-items:flex-start;gap:12px;width:100%;text-align:left;padding:13px 18px;border:0;border-bottom:1px solid rgba(236,231,220,.07);background:transparent;color:inherit;font:inherit;font-size:16px;line-height:1.35}
.selm-l button i{width:16px;flex:0 0 16px;font-style:normal;font-weight:800;color:#E9B949}
.selm-l button .dt{color:#E9B949;font-weight:700;min-width:46px;flex:none}
.selm-l button.on{background:rgba(233,185,73,.12);font-weight:700}.selm-l button:active{background:rgba(255,255,255,.08)}
.selm-l button[disabled]{opacity:.4}.selm-l button[hidden],.selm-l .grp[hidden]{display:none}
.selm-l .grp{padding:12px 18px 4px;font-size:12px;letter-spacing:.06em;color:#E9B949;font-weight:800}
.selm-fechar{flex:none;margin:10px 16px 0;padding:12px;border-radius:12px;border:1px solid rgba(236,231,220,.25);background:none;color:#ECE7DC;font:inherit;font-weight:600}
@media (prefers-reduced-motion:reduce){.selm,.selm-veu,.selm-veu.fora,.selm-veu.fora .selm{animation:none}}
</style>
<script>
// listas de escolha: em todo aparelho (Android, iPhone e computador), uma folha que sobe do pé da tela, com o título,
// o ✓ na opção atual e, nas listas longas, um campo de busca; eventos separados por mês.
(function(){
  let aberto=null;
  const fecha=()=>{if(!aberto)return;const v=aberto;aberto=null;v.classList.add("fora");setTimeout(()=>v.remove(),170)};
  const ACENTOS=new RegExp("["+String.fromCharCode(0x300)+"-"+String.fromCharCode(0x36f)+"]","g");
  const tira=t=>{try{return String(t).normalize("NFD").replace(ACENTOS,"").toLowerCase()}catch(_){return String(t).toLowerCase()}};
  const ESC=t=>String(t).replace(/[&<>"]/g,c=>({"&":"&amp;","<":"&lt;",">":"&gt;",'"':"&quot;"}[c]));
  const MES=["JANEIRO","FEVEREIRO","MARÇO","ABRIL","MAIO","JUNHO","JULHO","AGOSTO","SETEMBRO","OUTUBRO","NOVEMBRO","DEZEMBRO"];
  const titulo=sel=>{const l=sel.closest("label");if(l)for(const n of l.childNodes)if(n.nodeType===3&&n.textContent.trim())return n.textContent.trim();
    return sel.getAttribute("aria-label")||sel.title||"Escolha uma opção"};
  function abre(sel){
    if(aberto)fecha();
    const ops=[...sel.options];let html="",gAnt=null;
    ops.forEach((o,k)=>{const val=String(o.value),t=o.textContent;let g=null;
      if(o.parentElement&&o.parentElement.tagName==="OPTGROUP")g=o.parentElement.label;
      else if(/^[0-9]{4}-[0-9]{2}-[0-9]{2}$/.test(val))g=MES[+val.slice(5,7)-1]+" "+val.slice(0,4);
      if(g&&g!==gAnt){html+=`<div class="grp">${ESC(g)}</div>`;gAnt=g}
      let dt="",resto=t;if(t.length>8&&/^[0-9]{2}.[0-9]{2}$/.test(t.slice(0,5))&&t.slice(5,8)===" · "){dt=t.slice(0,5);resto=t.slice(8)}
      html+=`<button type="button" data-k="${k}" data-busca="${ESC(tira(t))}"${o.disabled?" disabled":""} class="${o.selected?"on":""}" role="option" aria-selected="${o.selected}"><i>${o.selected?"✓":""}</i>${dt?`<b class="dt">${ESC(dt)}</b>`:""}<span>${ESC(resto)}</span></button>`});
    const v=document.createElement("div");v.className="selm-veu";
    v.innerHTML=`<div class="selm" role="dialog" aria-label="${ESC(titulo(sel))}"><div class="alca"></div><h3>${ESC(titulo(sel))}</h3>${ops.length>8?`<input type="search" class="selm-busca" placeholder="🔎 Buscar…" aria-label="Buscar">`:""}<div class="selm-l" role="listbox">${html}</div><button type="button" class="selm-fechar">Fechar</button></div>`;
    document.body.appendChild(v);aberto=v;
    const L=v.querySelector(".selm-l"),on=L.querySelector(".on");if(on)on.scrollIntoView({block:"center"});
    const t0=Date.now();
    v.addEventListener("click",ev=>{if(ev.target===v){if(Date.now()-t0>350)fecha();return}
      if(ev.target.closest(".selm-fechar")){fecha();return}
      const b=ev.target.closest("button[data-k]");if(!b||b.disabled)return;const k=+b.dataset.k;
      if(sel.selectedIndex!==k){sel.selectedIndex=k;sel.dispatchEvent(new Event("input",{bubbles:true}));sel.dispatchEvent(new Event("change",{bubbles:true}))}
      try{navigator.vibrate&&navigator.vibrate(8)}catch(_){}fecha()});
    const bu=v.querySelector(".selm-busca");
    if(bu)bu.addEventListener("input",()=>{const q=tira(bu.value.trim());
      L.querySelectorAll("button[data-k]").forEach(b=>b.hidden=!!q&&!b.dataset.busca.includes(q));
      L.querySelectorAll(".grp").forEach(gr=>{let n=gr.nextElementSibling,vis=false;while(n&&!n.classList.contains("grp")){if(!n.hidden)vis=true;n=n.nextElementSibling}gr.hidden=!vis})})}
  const alvo=ev=>{const s=ev.target&&ev.target.closest&&ev.target.closest("select");return s&&!s.multiple&&!s.disabled&&s.size<=1&&!s.dataset.nativo?s:null};
  // o começo do toque só impede a janela padrão; a folha abre no fim do toque (clique), para o mesmo toque não fechá-la
  document.addEventListener("pointerdown",ev=>{if(alvo(ev))ev.preventDefault()},true);
  document.addEventListener("mousedown",ev=>{if(alvo(ev))ev.preventDefault()},true);
  document.addEventListener("click",ev=>{const s=alvo(ev);if(!s)return;ev.preventDefault();s.blur();abre(s)},true);
  document.addEventListener("keydown",ev=>{if(ev.key==="Escape"){fecha();return}
    const s=ev.target&&ev.target.tagName==="SELECT"?alvo(ev):null;if(s&&(ev.key==="Enter"||ev.key===" "||ev.key==="ArrowDown")){ev.preventDefault();abre(s)}});
  addEventListener("popstate",fecha);
})();
</script>"""

FILA_RECENTE = []   # resumo das últimas linhas da aba Fila (a nuvem preenche): "Meus envios" da página nova
def info_login(gerado, verif):
    """Embaixo da entrada: duas linhas que abrem ao toque (eventos e insights; pagamentos) e a data pequena no canto."""
    import html as _h
    def _pag(t):
        return bool(re.search(r"pagament|boleto|mensalid", t, re.I))
    itens = [x for x in FILA_RECENTE if not re.search(r"erro", x.get("sit", ""), re.I)]
    ger = [x for x in itens if not _pag(x.get("txt", ""))][:5]
    pagl = [x for x in itens if _pag(x.get("txt", ""))][:4]
    def curto(t):     # "Insight · Insight 018 (02/11)" -> "💬 Insight 018 (02/11)"; "Lista · Bonde (15/11)" -> "🗓️ Bonde (15/11)"
        t = str(t)
        if t.startswith("Insight · "):
            return "💬 " + t[10:]
        if t.startswith("Lista · "):
            return "🗓️ " + t[8:]
        return t
    li = lambda x: f'<li><b>{_h.escape(x.get("q", ""))}</b> · {_h.escape(curto(x.get("txt", "")))}</li>'
    if ger:
        corpo1, resumo1 = "".join(li(x) for x in ger), curto(ger[0].get("txt", ""))
    elif ULTIMA_ALTERACAO_TXT:
        corpo1, resumo1 = f"<li>{_h.escape(ULTIMA_ALTERACAO_TXT)}</li>", ULTIMA_ALTERACAO_TXT
    else:
        corpo1, resumo1 = "<li>Nenhuma novidade recente</li>", "sem novidades"
    corpo2 = (f"<li>Verificados: <b>{_h.escape(str(verif.get('quando')))}</b></li>" if verif.get("quando") else "") + "".join(li(x) for x in pagl)
    resumo2 = f"verificados {verif.get('quando')}" if verif.get("quando") else (pagl[0].get("txt", "") if pagl else "sem novidades")
    m = re.match(r"(\d{2})/(\d{2})/\d{4}(.*)", str(gerado))
    abrev = f"{m.group(1)}/{['jan','fev','mar','abr','mai','jun','jul','ago','set','out','nov','dez'][int(m.group(2)) - 1]}{m.group(3)}" if m else str(gerado)
    return ('<span class="avs">'
            f'<details class="av"><summary>📋 Eventos e insights · <b>{_h.escape(resumo1[:48])}</b></summary><ul>{corpo1}</ul></details>'
            f'<details class="av"><summary>💵 Pagamentos · <b>{_h.escape(str(resumo2)[:48])}</b></summary><ul>{corpo2 or "<li>Nenhuma novidade recente</li>"}</ul></details>'
            f'</span><span class="atual-dt">Atualizado {_h.escape(abrev)}</span>')

VERSAO_ATUAL = ""   # assinatura do conteúdo publicado: muda só quando o painel muda (versao.json)

def snip_versao(auto):
    """Confere se há versão nova quando o app volta para a tela (app que ficou "dormindo" mostra a tela antiga)."""
    return ('<script>(function(){var V="__VERSAO__",AUTO=' + ("true" if auto else "false") + ';if(!V||V.indexOf("__")===0)return;var b=null,ult=0,iv=0;'
            'function faixa(t,f){if(!b){b=document.createElement("div");b.id="verFaixa";b.style.cssText="position:fixed;left:10px;right:10px;bottom:calc(84px + env(safe-area-inset-bottom,0px));z-index:99999;background:#E9B949;color:#141210;border-radius:14px;padding:12px 14px;font:700 14px Barlow,Arial,sans-serif;box-shadow:0 6px 24px rgba(0,0,0,.5);text-align:center;cursor:pointer";document.body.appendChild(b)}b.textContent=t;b.onclick=f;b.style.display="block"}'
            'function some(){if(b)b.style.display="none"}'
            'function ver(){if(AUTO&&!document.getElementById("cpf")){clearInterval(iv);return}'   # recarregar sozinho só na tela de entrada
            'var ag=Date.now();if(ag-ult<60000)return;ult=ag;fetch("versao.json?t="+ag,{cache:"no-store"}).then(function(r){return r.ok?r.json():null}).then(function(j){if(!j||!j.v)return;'
            'if(j.v!==V){if(AUTO){location.reload();return}faixa("🔄 Tem uma versão nova do painel · toque para atualizar",function(){location.reload()})}else some()})'
            '.catch(function(){faixa("📴 Sem internet · você está vendo a última versão guardada",some)})}'
            'document.addEventListener("visibilitychange",function(){if(!document.hidden)ver()});'
            'addEventListener("focus",ver);addEventListener("online",function(){ult=0;ver()});'
            'try{navigator.serviceWorker&&navigator.serviceWorker.addEventListener("message",function(e){if(e.data&&e.data.tipo==="verificar_versao"){ult=0;setTimeout(ver,1500)}})}catch(_){}'
            'iv=setInterval(function(){if(!document.hidden)ver()},120000);setTimeout(ver,3000)})();</script>')

def montar_lembretes(membros, eventos, eventos_ins, pag):
    """Pendências de cada irmão com e-mail: insights sem resposta (últimos 7 dias), eventos sem confirmação
    (de hoje em diante) e a mensalidade (se ainda não aparece como paga). A nuvem manda os lembretes nos horários."""
    hoje = datetime.date.today().isoformat()
    sete = (datetime.date.today() - datetime.timedelta(days=7)).isoformat()
    ent = {}
    for evs in (eventos, eventos_ins):
        for e in evs:
            for i, v in enumerate(e["p"]):
                if v not in (None, "") and (i not in ent or e["d"] < ent[i]):
                    ent[i] = e["d"]
    def conta_irmao(i, d):
        return i in ent and d >= ent[i]
    passados = [e for e in eventos_ins if e["d"] <= hoje]
    ult_ins = max(passados, key=lambda e: e["d"]) if passados else None   # quando sai um insight novo, os anteriores param de avisar
    pessoas = []
    for i, m in enumerate(membros):
        email = EMAIL_APELIDO_INV.get(m["n"].lower(), "")
        if "@" not in email:
            continue
        ins = []
        if ult_ins and sete <= ult_ins["d"] and conta_irmao(i, ult_ins["d"]) and ult_ins["p"][i] != 1:   # só o insight mais recente
            ins.append({"d": ult_ins["d"], "nome": ult_ins["e"], "resp": sum(1 for v in ult_ins["p"] if v == 1)})
        evs = []
        for e in eventos:
            if e["d"] >= hoje and e["p"][i] in (None, "", "-") and (conta_irmao(i, e["d"]) or i in ent):
                evs.append({"d": e["d"], "nome": e["e"]})
        st = str(pag.get("status", {}).get(m["n"], "") or "")
        pago = bool(re.search(r"pago|aguard|ades", st, re.I))
        pessoas.append({"n": m["n"], "email": email.strip().lower(), "ins": ins, "ev": evs,
                        "pago": pago, "boleto": pag.get("boleto", {}).get(m["n"], "")})
    LEMBRETES.clear()
    LEMBRETES.update({"hoje": hoje, "venc": pag.get("venc"), "pessoas": pessoas})

def _json_estado(nome, padrao):
    try:
        with open(os.path.join(PASTA, nome), encoding="utf-8") as f:
            return json.load(f)
    except Exception:
        return padrao

def _quando(txt):
    """Data/hora de um registro de acesso (vários formatos possíveis)."""
    t = str(txt or "").strip()
    for f in ("%Y-%m-%dT%H:%M:%S", "%Y-%m-%d %H:%M:%S", "%Y-%m-%d %H:%M", "%d/%m/%Y %H:%M:%S", "%d/%m/%Y %H:%M"):
        try:
            return datetime.datetime.strptime(t[:19].replace("Z", ""), f)
        except ValueError:
            pass
    return None

def atualizar_visitantes():
    """Visitantes (uma entrada, com validade): marca quem já entrou, tira os vencidos e monta a gaveta dos que ainda podem entrar."""
    lista = [v for v in _json_estado("visitantes.json", []) if isinstance(v, dict) and "@" in str(v.get("email", ""))]
    if not lista:
        return []
    regs = (buscar_acessos() or [])[1:]
    agora = datetime.datetime.now()
    for v in lista:
        if v.get("usado"):
            continue
        criado = _quando(v.get("criado")) or agora
        for r in regs:
            r = list(r) + [""] * 5
            q = _quando(r[0])
            if str(r[1]).strip().lower() == v["email"] and str(r[2]).strip() == "Entrou" and q and q >= criado - datetime.timedelta(minutes=5):
                v["usado"] = q.strftime("%Y-%m-%d %H:%M"); break
    ativos = 0
    for v in lista:
        vence = _quando(v.get("expira"))
        v["status"] = "usado" if v.get("usado") else ("vencido" if vence and vence < agora else "aguardando")
        if v["status"] == "aguardando" and GERAL_RAW:
            PESSOAIS.append((v["email"], GERAL_RAW, "v")); ativos += 1
    # guarda só os últimos 30 dias de histórico
    corte = (agora - datetime.timedelta(days=30)).strftime("%Y-%m-%d")
    lista = [v for v in lista if str(v.get("criado", ""))[:10] >= corte]
    try:
        with open(os.path.join(PASTA, "visitantes.json"), "w", encoding="utf-8") as f:
            json.dump([{k: x for k, x in v.items() if k != "status"} for v in lista], f, ensure_ascii=False, indent=1)
    except Exception as e:
        print(f"Aviso: não consegui guardar os visitantes ({e}).")
    if ativos:
        print(f"Visitantes: {ativos} acesso(s) de uma entrada ativo(s).")
    return lista

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
        kek = hashlib.pbkdf2_hmac("sha256", chave_de(s).encode(), sal, ITERACOES, 32)
        iv_s = secrets.token_bytes(12)
        slots.append({"iv": b64(iv_s), "ct": b64(AESGCM(kek).encrypt(iv_s, chave, None))})
    secrets.SystemRandom().shuffle(slots)
    pes = []
    for item in PESSOAIS:              # "Meu painel" (irmão) ou visitante: a página trancada com o e-mail dele
        email, pag = item[0], item[1].replace("__VERSAO__", VERSAO_ATUAL)
        kek = hashlib.pbkdf2_hmac("sha256", chave_de(email).encode(), sal, ITERACOES, 32)
        iv_p = secrets.token_bytes(12)
        g_ = {"iv": b64(iv_p), "ct": b64(AESGCM(kek).encrypt(iv_p, pag.encode("utf-8"), None))}
        if len(item) > 2:
            g_["t"] = item[2]
        pes.append(g_)
    secrets.SystemRandom().shuffle(pes)
    pacote = {"sal": b64(sal), "it": ITERACOES, "iv": b64(iv), "ct": b64(ct), "slots": slots, "pes": pes, "cav": CAVEIRA_URL, "otp": OTP_URL if otp_ativo() else ""}
    return (texto_divisao(LOGIN).replace("__PACOTE__", json.dumps(pacote)).replace("__IMG__", img)
                 .replace("__TITULO__", titulo).replace("__SUBTITULO__", subtitulo).replace("__ROTULO__", rotulo)
                 .replace("__ATUAL__", INFO_LOGIN).replace("__MANUT__", MANUT_HTML).replace("__FOTOS__", FOTOS_LOGIN)
                 .replace("__MANIFEST__", "manifest-adm.json").replace("__TUTORIAL__", TUTORIAL_APP).replace("__VERSAOSNIP__", snip_versao(True)).replace("__VERSAO__", VERSAO_ATUAL).replace("__REGISTRO__", "")
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

ULTIMA_PAGINA_BOLETO = ""

def _ler_boleto(url):
    """Devolve (texto do selo de status ou None, data de vencimento do boleto ou None)."""
    import re, urllib.request, html as H
    req = urllib.request.Request(url, headers={"User-Agent": "Mozilla/5.0"})
    with urllib.request.urlopen(req, timeout=25) as r:
        bruto = r.read().decode("utf-8", "ignore")
        global ULTIMA_PAGINA_BOLETO
        ULTIMA_PAGINA_BOLETO = bruto   # diagnóstico: a verificação mostra o trecho quando não acha a situação
    selo = re.search(r'class="status_warning"[^>]*>\s*([^<]+?)\s*<', bruto)          # página da Iugu
    texto = re.sub(r"\s+", " ", H.unescape(re.sub(r"<[^>]+>", " ", bruto)))
    v = re.search(r"Vencimento\D{0,40}(\d{2})/(\d{2})/(\d{4})", texto, re.I)
    venc = datetime.date(int(v.group(3)), int(v.group(2)), int(v.group(1))) if v else None
    selo = selo.group(1).strip() if selo else None
    if not selo:   # página do Sistema Insanos (FaturaCopiaCola): a situação é um botão colorido ao lado do vencimento
        SIT = r"(Pago em Dia|Pago Atrasado|Pago|A Vencer|Vencid[ao]|Em Aberto|Atrasad[ao]|Cancelad[ao]|Isent[ao])"
        # 1) dentro do botão/campo (o texto fica no atributo value, title ou alt, e não aparece como texto da página)
        m = re.search(r'(?:value|title|alt)\s*=\s*["\']\s*' + SIT + r'\s*["\']', bruto, re.I)
        # 2) como texto, logo depois do vencimento
        if not m:
            m = re.search(r"Vencimento:?\s*\d{2}/\d{2}/\d{4}\s*" + SIT, texto, re.I)
        # 3) no código da página, logo depois do vencimento (ex.: escrito por um script ou em outro tipo de elemento)
        if not m:
            m = re.search(r"Vencimento.{0,1500}?[>\"'\s]" + SIT + r"[<\"'\s]", bruto, re.I | re.S)
        selo = m.group(1).strip() if m else None
    return selo, venc

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
            notificar("💵 Pagamento novo", f"{novos[0]} pagou a mensalidade (identificado {hora}).", "pagamento", icone="notif-pag.png")
        else:
            notificar(f"💵 {len(novos)} pagamentos novos", ", ".join(novos) + f" (identificados {hora}).", "pagamento", icone="notif-pag.png")
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
             "tel": col("telefone"), "nome": col("nome completo"), "nasc": col("nascimento"),
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
                        "nome": dd.get("nome"), "colete": dd.get("colete"), "tam": dd.get("tam"),
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

def token_rsvp(email, d, nome):
    """Assinatura da resposta "Vou / Não vou": só quem tem a chave da planilha consegue gerar (ninguém responde por outro)."""
    import hmac, hashlib
    k = ((ACESSOS_CHAVE or "insanos") + "|rsvp").encode()
    return hmac.new(k, f"{str(email).strip().lower()}|{d}|{str(nome).strip()}".encode(), hashlib.sha256).hexdigest()[:16]

ADM_EMAILS = set()     # e-mails do comando (fixos + liberados): os avisos gerais vão só para eles

def aparelhos_com_avisos():
    """Quantos aparelhos cada e-mail tem com os avisos ativos (para a tela Acessos)."""
    if not ACESSOS_URL:
        return {}
    import urllib.request, urllib.parse
    try:
        with urllib.request.urlopen(ACESSOS_URL + "?push=1&chave=" + urllib.parse.quote(ACESSOS_CHAVE), timeout=25) as r:
            inscritos = json.loads(r.read().decode("utf-8"))
    except Exception:
        return {}
    cont = {}
    for item in inscritos if isinstance(inscritos, list) else []:
        em = str(item[0] if isinstance(item, (list, tuple)) and item else "").strip().lower()
        if "@" in em:
            cont[em] = cont.get(em, 0) + 1
    return cont

def notificar(titulo, corpo, tag="aviso", zap=None, icone=None, so_para=None, link=None, link_rot=None, rsvp=None,
              todos=False, excluir=None, extra=None):
    """Manda a notificação. Sem so_para: só para os celulares do comando (ADM).
    Com so_para: só para os celulares daquele e-mail (lembretes individuais dos irmãos)."""
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
    if extra:    # dados especiais da notificação (ex.: o SOS)
        msg.update(extra)
    if link:     # botão da notificação que abre um endereço (ex.: "Abrir meu boleto")
        msg["link"] = link; msg["lrot"] = link_rot or "Abrir"
    if icone:    # imagem do aviso (pasta app: notif-pag.png, notif-evento.png, ...)
        msg["icon"] = icone
    if zap:      # mensagem pronta para o botão "📲 Enviar no WhatsApp" (a notificação tem limite de tamanho)
        msg["zap"] = zap
        if len(json.dumps(msg, ensure_ascii=False).encode("utf-8")) > 3400:
            msg.pop("zap"); _log("Aviso: a mensagem do WhatsApp ficou grande demais para ir na notificação (vai sem o botão).")
    dados = json.dumps(msg, ensure_ascii=False)
    enviados = 0
    for email, sub in inscritos:
        em = str(email).strip().lower()
        if so_para and em != str(so_para).strip().lower():
            continue            # aviso individual (ou de teste): só para aquele e-mail
        if excluir and em in ({str(x).strip().lower() for x in excluir} if isinstance(excluir, (list, set, tuple)) else {str(excluir).strip().lower()}):
            continue            # ex.: quem pediu o SOS não recebe o próprio alerta
        if not so_para and not todos and ADM_EMAILS and em not in ADM_EMAILS:
            continue            # aviso geral: só para o comando (os irmãos recebem só os lembretes deles); o SOS vai para todos
        try:
            info = json.loads(sub) if isinstance(sub, str) else sub
            extra = {"para": em, "ru": ACESSOS_URL} if ACESSOS_URL and "@" in em else {}
            if rsvp and extra:   # botões ✅ Vou / ❌ Não vou: a resposta leva a assinatura daquele irmão
                extra["rsvp"] = {"d": rsvp["d"], "e": rsvp["e"], "tk": token_rsvp(em, rsvp["d"], rsvp["e"]), "m": rsvp.get("m", "ev"),
                                 "ou": OTP_URL if otp_ativo() else ""}   # para acordar o painel na hora (a lista atualiza em 1 a 3 minutos)
            dados_i = json.dumps(dict(json.loads(dados), **extra), ensure_ascii=False) if extra else dados
            wp.webpush(subscription_info=info, data=dados_i, vapid_private_key=ARQ_VAPID_PEM,
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
                notificar(titulo, f"Amanhã ({amanha:%d/%m}) é aniversário de {ap}: faz {amanha.year - nasc.year} anos. Hora de preparar o fly!", "aniversario", icone="notif-niver.png")
                feitos[chave] = agora.isoformat(timespec="minutes")
        if sit.startswith("afast"):
            dd = r[c_des] if c_des is not None and c_des < len(r) else None
            if isinstance(dd, datetime.datetime):
                dd = dd.date()
            if isinstance(dd, datetime.date):
                inf = info_afastamento(ap, dd)
                chave = f"afast:{ap}:{inf['fim']}"
                if inf["terminou"] and chave not in feitos:
                    notificar("⚠️ Fim de afastamento", f"O afastamento de {ap} terminou em {datetime.date.fromisoformat(inf['fim']):%d/%m}. Reative ou exclua no Cadastro.", "afastamento", icone="notif-aviso.png")
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

APP_NOME = f"{DIV['nome']} {DIV['regional']}"

# O link é só do comando (ADM). O endereço geral não tem mais o painel: só leva para a entrada do ADM, sem nenhum dado.
INDEX_SO_ADM = """<!DOCTYPE html>
<html lang="pt-BR"><head><meta charset="UTF-8"><meta name="viewport" content="width=device-width, initial-scale=1">
<meta name="robots" content="noindex"><title>Insanos ADM</title>
<link rel="manifest" href="manifest-adm.json"><meta name="theme-color" content="#0b0b0b"><link rel="icon" href="icon-192.png">
<meta http-equiv="refresh" content="0; url=adm.html"><script>location.replace("adm.html"+location.search)</script>
<style>body{background:#000;color:#ECE7DC;font-family:Segoe UI,Roboto,Arial,sans-serif;display:grid;place-items:center;height:100vh;margin:0}a{color:#8FD9A3}</style>
</head><body><p>Abrindo o painel do comando… <a href="adm.html">toque aqui</a> se não abrir.</p></body></html>
"""

MAPA_PNG, MAPA_INFO = None, None   # o mapa da divisão (lugares dos eventos), para a imagem compartilhada da carteirinha
def montar_mapa(eventos):
    """Mapa com os lugares onde a divisão já esteve: o endereço do evento (aba Eventos) ou, sem ele, a sede.
    Coordenadas pelo OpenStreetMap (Nominatim, 1 por segundo) e o mapa com os blocos do OpenStreetMap, no tom escuro do app."""
    global MAPA_PNG, MAPA_INFO
    import urllib.request, urllib.parse, math, io, time
    from PIL import Image, ImageDraw, ImageOps, ImageEnhance
    UA = {"User-Agent": "PainelInsanos/1.0 (painel da divisao)"}
    sede = (DIV.get("sede_endereco") or DIV.get("sede_local") or "").strip()
    cont = {}
    for e in eventos:
        inf = INFO_EV.get((e["d"], e["e"].strip().lower()), {})
        end = (inf.get("endereco") or inf.get("local") or sede or "").strip()
        if end:
            cont[end] = cont.get(end, 0) + 1
    if not cont:
        return
    pts = []
    for end, n in sorted(cont.items(), key=lambda x: -x[1])[:25]:
        geo = CLIMA_GEO.get(end)
        if end not in CLIMA_GEO:
            try:
                for q in (end, end.split(" - ")[-1] if " - " in end else ""):
                    if not q:
                        continue
                    r = json.loads(urllib.request.urlopen(urllib.request.Request("https://nominatim.openstreetmap.org/search?" + urllib.parse.urlencode(
                        {"format": "json", "limit": 1, "countrycodes": "br", "q": q}), headers=UA), timeout=8).read().decode("utf-8"))
                    time.sleep(1.1)
                    if r:
                        geo = (float(r[0]["lat"]), float(r[0]["lon"])); break
            except Exception:
                geo = None
            CLIMA_GEO[end] = geo
        if geo:
            pts.append((geo[0], geo[1], n))
    if not pts:
        return
    junta = {}   # o mesmo ponto (endereços diferentes que caem no mesmo lugar) soma
    for la, lo, n in pts:
        k = (round(la, 4), round(lo, 4)); junta[k] = junta.get(k, 0) + n
    pts = [(la, lo, n) for (la, lo), n in junta.items()]
    W, H, PAD = 920, 440, 60
    px = lambda la, lo, z: ((lo + 180) / 360 * 256 * 2 ** z, (1 - math.log(math.tan(math.radians(la)) + 1 / math.cos(math.radians(la))) / math.pi) / 2 * 256 * 2 ** z)
    z = 14
    while z > 3:
        xs = [px(la, lo, z)[0] for la, lo, _ in pts]; ys = [px(la, lo, z)[1] for la, lo, _ in pts]
        if max(xs) - min(xs) <= W - 2 * PAD and max(ys) - min(ys) <= H - 2 * PAD:
            break
        z -= 1
    xs = [px(la, lo, z)[0] for la, lo, _ in pts]; ys = [px(la, lo, z)[1] for la, lo, _ in pts]
    cx, cy = (max(xs) + min(xs)) / 2, (max(ys) + min(ys)) / 2
    x0, y0 = cx - W / 2, cy - H / 2
    tela = Image.new("RGB", (W, H), (20, 18, 16))
    try:
        for tx in range(int(x0 // 256), int((x0 + W) // 256) + 1):
            for ty in range(int(y0 // 256), int((y0 + H) // 256) + 1):
                u = f"https://tile.openstreetmap.org/{z}/{tx % (2 ** z)}/{ty}.png"
                t = Image.open(io.BytesIO(urllib.request.urlopen(urllib.request.Request(u, headers=UA), timeout=10).read())).convert("RGB")
                tela.paste(t, (int(tx * 256 - x0), int(ty * 256 - y0)))
    except Exception as e:
        print(f"Aviso: mapa sem os blocos do OpenStreetMap ({type(e).__name__}); só os pontos.")
    tela = ImageOps.colorize(ImageEnhance.Contrast(ImageOps.grayscale(tela)).enhance(1.1), black=(10, 9, 8), white=(120, 104, 80))   # tom escuro do app
    d = ImageDraw.Draw(tela)
    for la, lo, n in sorted(pts, key=lambda p: p[2]):
        x, y = px(la, lo, z); x -= x0; y -= y0; r = 9 + min(14, n * 2)
        d.ellipse((x - r - 4, y - r - 4, x + r + 4, y + r + 4), fill=(233, 185, 73, 90))
        d.ellipse((x - r, y - r, x + r, y + r), fill=(233, 185, 73), outline=(20, 18, 16), width=3)
    d.text((W - 8, H - 6), "© OpenStreetMap", fill=(200, 190, 170), anchor="rd")
    b = io.BytesIO(); tela.save(b, "PNG", optimize=True); MAPA_PNG = b.getvalue()
    MAPA_INFO = {"lugares": len(pts), "eventos": sum(n for _, _, n in pts)}

def arquivos_app():
    """Arquivos que fazem o link virar "app" (instalável no celular): manifestos, ajudante (sw.js) e ícones."""
    def manifesto(nome, curto, inicio):
        return json.dumps({"id": inicio, "name": nome, "short_name": curto, "start_url": inicio, "scope": "./",
                           "display": "standalone", "orientation": "portrait", "background_color": "#000000", "theme_color": "#0b0b0b",
                           "lang": "pt-BR", "icons": [{"src": f"icon-{t}.png", "sizes": f"{t}x{t}", "type": "image/png", "purpose": "any maskable"} for t in (192, 512)],
                           # compartilhar: no WhatsApp, Compartilhar › Insanos ADM manda a imagem do fly direto para a tela Lista (Android)
                           "share_target": {"action": "compartilhar-fly", "method": "POST", "enctype": "multipart/form-data",
                                            "params": {"title": "title", "text": "text", "files": [{"name": "fly", "accept": ["image/*"]}]}},
                           # atalhos: segurando o ícone do app (Android) aparecem estas opções
                           "shortcuts": [
                               {"name": "🆘 SOS · Preciso de ajuda", "short_name": "SOS", "description": "Pedir ajuda aos irmãos",
                                "url": inicio + "#sos-abrir", "icons": [{"src": "atalho-sos.png", "sizes": "192x192", "type": "image/png"}]},
                               {"name": "Próximos eventos", "short_name": "Eventos", "description": "Responder Vou / Não vou",
                                "url": inicio + "#eventos", "icons": [{"src": "atalho-eventos.png", "sizes": "192x192", "type": "image/png"}]},
                               {"name": "Minha moto", "short_name": "Minha moto", "description": "Documentos e abastecimentos",
                                "url": inicio + "#moto", "icons": [{"src": "atalho-moto.png", "sizes": "192x192", "type": "image/png"}]},
                               {"name": "Lançar/Alterar (comando)", "short_name": "Lançar/Alterar", "description": "Lançar ou alterar um evento ou insight",
                                "url": inicio + "#lancar", "icons": [{"src": "atalho-lancar.png", "sizes": "192x192", "type": "image/png"}]},
                               {"name": "Lista (comando)", "short_name": "Lista", "description": "Montar a lista do WhatsApp",
                                "url": inicio + "#lista", "icons": [{"src": "atalho-lista.png", "sizes": "192x192", "type": "image/png"}]}]},
                          ensure_ascii=False, indent=1)
    sw = """// Painel Insanos: busca sempre a versão nova; sem internet, abre a última versão vista.
const CACHE = "insanos-v27";
self.addEventListener("install", e => self.skipWaiting());
self.addEventListener("activate", e => e.waitUntil(self.clients.claim()));
self.addEventListener("fetch", e => {
  if (e.request.method === "POST" && new URL(e.request.url).pathname.endsWith("/compartilhar-fly")) {   // fly compartilhado pelo WhatsApp
    e.respondWith((async () => {
      try { const fd = await e.request.formData(), f = fd.get("fly");
        if (f) { const c = await caches.open("fly-compartilhado"); await c.put("fly", new Response(f, {headers: {"Content-Type": f.type || "image/jpeg"}})); } } catch (_) {}
      return Response.redirect(new URL("adm.html#fly-compartilhado", self.location.href).href, 303);   // endereço completo (alguns celulares exigem)
    })());
    return;
  }
  const r = e.request;
  if (r.method !== "GET" || new URL(r.url).origin !== location.origin) return;
  if (/\\.mp4$/i.test(new URL(r.url).pathname)) return;      // vídeo de fundo: direto da rede (guardar vídeo no app trava a reprodução)
  e.respondWith(fetch(r, {cache: "no-store"}).then(resp => {
    if (resp.ok) { const c = resp.clone(); caches.open(CACHE).then(k => k.put(r, c)); }
    return resp;
  }).catch(() => caches.match(r, {ignoreSearch: true})));
});
// número de mensagens não lidas no ícone do app (as notificações de mensagem que ainda estão na tela)
function marcarIcone() {
  return self.registration.getNotifications().then(ns => { const n = ns.filter(x => String(x.tag || "").startsWith("msg-") && !String(x.tag).startsWith("msg-r-")).length;
    if (self.navigator.setAppBadge) return n ? self.navigator.setAppBadge(n) : self.navigator.clearAppBadge(); }).catch(() => {});
}
self.addEventListener("notificationclose", e => { e.waitUntil(marcarIcone()); });
const aoPush = e => {
  e.waitUntil(clients.matchAll({type: "window", includeUncontrolled: true}).then(cs => cs.forEach(c => c.postMessage({tipo: "verificar_versao"}))).catch(() => {}));
  let d = {};
  try { d = e.data.json(); } catch (_) { d = {title: "Insanos ADM", body: e.data ? e.data.text() : ""}; }
  const teste = !!d.teste;   // aviso de teste (Gestão › Notificações › Testar): aparece igual, mas os botões não enviam nada
  const op = {body: d.body || "", icon: d.icon || "icon-192.png", badge: "notif-badge.png", tag: d.tag || "aviso", renotify: true,
              data: {url: d.url || "adm.html", zap: d.zap || ""}};
  if (d.zap) op.actions = [{action: "zap", title: "📲 Enviar no WhatsApp"}];
  if (d.link) { op.data.link = d.link; op.actions = [{action: "link", title: d.lrot || "Abrir"}]; }
  if (d.msg) { op.data.msg = d.msg; op.actions = (d.msg.rs || []).map(([a, t]) => ({action: a, title: t}));
    if (d.msg.imp) { op.requireInteraction = true; op.renotify = true; op.vibrate = [700, 250, 700, 250, 1200]; op.image = "notif-importante.png"; } }
  if (d.rsvp) { op.data.rsvp = d.rsvp; op.requireInteraction = true;
    op.actions = d.rsvp.m === "ins" ? [{action: "resp", title: "✅ Já respondi"}] : [{action: "vou", title: "✅ Vou"}, {action: "nao", title: "❌ Não vou"}]; }
  op.data.para = d.para || ""; op.data.ru = d.ru || ""; op.data.tt = d.title || ""; op.data.teste = teste ? 1 : 0;
  if (d.sos) {   // SOS: distância aproximada pela última localização guardada neste celular, e botões
    op.data.sosUrl = d.url || "adm.html"; op.requireInteraction = true; op.renotify = true;
    op.vibrate = [1000, 250, 1000, 250, 1000, 250, 1000, 250, 1000];   // alarme (onde o celular respeitar o padrão)
    const S = d.sos, temLoc = S.lat !== "" && S.lat != null;
    if (temLoc) op.data.mapa = (d.url || "adm.html") + "&mapa=1";   // o mapa abre dentro do painel (o app do Maps às vezes se perde)
    op.actions = [{action: "indo", title: "🏍️ Estou indo"}].concat(temLoc ? [{action: "mapa", title: "📍 Ver no mapa"}] : []);
    const comDist = temLoc ? idbGet("loc").then(l => {
      if (!l) return;
      const R = 6371, x = (S.lat - l.lat) * Math.PI / 180, y = (S.lng - l.lng) * Math.PI / 180;
      const h = Math.sin(x / 2) ** 2 + Math.cos(l.lat * Math.PI / 180) * Math.cos(S.lat * Math.PI / 180) * Math.sin(y / 2) ** 2, km = 2 * R * Math.asin(Math.sqrt(h));
      const velha = Date.now() - l.ts > 12 * 3600e3;
      op.body += " · a ~" + (km < 1 ? Math.round(km * 1000) + " m" : km.toFixed(1).replace(".", ",") + " km") + " de você" + (velha ? " (localização antiga)" : "");
    }).catch(() => {}) : Promise.resolve();
    const guarda = teste ? Promise.resolve() : idbSet("sos_ativo", {d: S, url: op.data.sosUrl, ts: Date.now(), titulo: d.title || ""});   // para a faixa "SOS em andamento"
    e.waitUntil(Promise.all([comDist.then(() => self.registration.showNotification(d.title || "🆘 SOS", op)), registrarRecebeu(d), guarda]));
    return;
  }
  // vibração própria de cada tipo de aviso (Android; o iPhone ignora): dá para reconhecer sem olhar o celular
  const VIB = {pag: [120, 90, 120], evento: [450], social: [450], pub: [450], insight: [150],
               cadastro: [200, 100, 200, 100, 200], niver: [100, 70, 100, 70, 100], aviso: [450, 120, 120]};
  const tipo = ((d.icon || "").match(/notif-([a-z]+)\\.png/) || [])[1] || (d.tag === "pagamento" ? "pag" : d.tag === "aniversario" ? "niver" : d.tag === "afastamento" ? "aviso" : "");
  if (!(d.msg && d.msg.imp)) op.vibrate = VIB[tipo] || [200];
  // controle de entrega: registra na aba Acessos da planilha que o aviso chegou a este celular
  e.waitUntil(Promise.all([self.registration.showNotification(d.title || "Insanos ADM", op).then(() => d.msg ? marcarIcone() : null), registrarRecebeu(d)]));
};
self.addEventListener("push", aoPush);
// teste das notificações: o painel manda um aviso de exemplo e ele aparece exatamente como os de verdade
self.addEventListener("message", e => {
  if (!e.data || e.data.tipo !== "teste_aviso") return;
  const d = Object.assign({}, e.data.d, {teste: 1, ru: "", para: ""});
  aoPush({data: {json: () => d, text: () => ""}, waitUntil: p => e.waitUntil(p)});
});
function registrarRecebeu(d) {
  return (d.ru && d.para) ? fetch(d.ru, {method: "POST", mode: "no-cors", body: JSON.stringify({email: d.para, ok: true,
      painel: "🔔 Recebeu: " + (d.title || ""), aparelho: aparelhoSW()})}).catch(() => {}) : Promise.resolve();
}
function idbSet(k, v) {
  return new Promise(r => { try { const o = indexedDB.open("insanos", 1); o.onupgradeneeded = () => o.result.createObjectStore("kv");
    o.onsuccess = () => { try { const t = o.result.transaction("kv", "readwrite"); t.objectStore("kv").put(v, k); t.oncomplete = () => r(1); t.onerror = () => r(0); } catch (e) { r(0); } };
    o.onerror = () => r(0); } catch (e) { r(0); } });
}
// SOS dispensado (jogado para o lado): registra na aba Acessos, para o comando saber quem viu e não respondeu
self.addEventListener("notificationclose", e => {
  const dd = e.notification.data || {};
  if (dd.sosUrl && dd.ru && dd.para)
    e.waitUntil(fetch(dd.ru, {method: "POST", mode: "no-cors", body: JSON.stringify({email: dd.para, ok: true,
      painel: "👋 Dispensou: " + (e.notification.title || ""), aparelho: aparelhoSW()})}).catch(() => {}));
});
// última localização guardada neste celular (pela tela de entrada ou pelo Meu painel), para a distância do SOS
function idbGet(k) {
  return new Promise(r => { try { const o = indexedDB.open("insanos", 1); o.onupgradeneeded = () => o.result.createObjectStore("kv");
    o.onsuccess = () => { try { const q = o.result.transaction("kv").objectStore("kv").get(k); q.onsuccess = () => r(q.result || null); q.onerror = () => r(null); } catch (e) { r(null); } };
    o.onerror = () => r(null); } catch (e) { r(null); } });
}
function aparelhoSW() { const u = navigator.userAgent || ""; return /iPhone|iPad|iPod/.test(u) ? "iPhone" : /Android/.test(u) ? "Celular Android" : "Computador"; }
// depois de um Vou / Não vou / Já respondi: acorda o painel na hora (a lista atualiza em 1 a 3 minutos)
function acordarPainel(r, email) {
  return (r && r.ou) ? fetch(r.ou, {method: "POST", body: JSON.stringify({tipo: "rsvp_acordar", email: email})}).catch(() => {}) : Promise.resolve();
}
self.addEventListener("notificationclick", e => {
  const dd = e.notification.data || {};
  if (dd.teste) {   // aviso de teste: nenhum botão envia nada
    e.notification.close();
    if (e.action) e.waitUntil(self.registration.showNotification("🧪 Era só um teste", {body: "Nada foi enviado. Nos avisos de verdade, esse botão funciona.", icon: "icon-192.png", badge: "notif-badge.png", tag: "teste-ok"}));
    else e.waitUntil(clients.openWindow("adm.html"));
    return;
  }
  const soMapa = !!(dd.sosUrl && e.action === "mapa");
  if (!soMapa) e.notification.close();   // SOS · "Ver no mapa": a notificação fica, para depois tocar em "Estou indo"
  if (dd.ru && dd.para)    // controle: o irmão tocou no aviso
    e.waitUntil(fetch(dd.ru, {method: "POST", mode: "no-cors", body: JSON.stringify({email: dd.para, ok: true,
      painel: "👆 Abriu: " + (dd.tt || "") + (e.action === "link" ? " (" + (e.notification.actions && e.notification.actions[0] ? e.notification.actions[0].title : "botão") + ")" : ""),
      aparelho: aparelhoSW()})}).catch(() => {}));
  // botão "Enviar no WhatsApp" (Android/computador) ou toque na notificação no iPhone (que não mostra botões):
  // abre o painel com a caixa da mensagem (o site do WhatsApp quebra os emojis; de lá o painel abre o app direto)
  if (dd.sosUrl) {   // SOS: "Ver no mapa" abre o mapa; "Estou indo" ou o toque abrem a tela com a distância exata
    if (e.action === "mapa" && dd.mapa) {
      const n = e.notification;   // alguns celulares fecham a notificação ao tocar no botão: ela volta, em silêncio
      e.waitUntil(Promise.all([clients.openWindow(dd.mapa),
        self.registration.showNotification(n.title, {body: n.body, icon: n.icon, badge: n.badge, tag: n.tag, data: dd,
          actions: n.actions, requireInteraction: true, renotify: false, silent: true})]));
      return;
    }
    e.waitUntil(clients.openWindow(dd.sosUrl + (e.action === "indo" ? "&indo=1" : ""))); return;
  }
  if (dd.msg) e.waitUntil(new Promise(r => setTimeout(r, 300)).then(marcarIcone));
  if (dd.msg && e.action && dd.para) {   // mensagem: resposta rápida direto na notificação
    const m = dd.msg, txt = ((m.rs || []).find(x => x[0] === e.action) || [])[1] || "";
    e.waitUntil(fetch(m.ou, {method: "POST", mode: "no-cors", headers: {"Content-Type": "text/plain;charset=utf-8"},
        body: JSON.stringify({tipo: "msg_resp", email: dd.para, id: m.id, tk: m.tk, resp: txt})}).catch(() => {})
      .then(() => self.registration.showNotification("↩️ Resposta enviada", {body: txt, icon: "icon-192.png", badge: "notif-badge.png", tag: "msg-ok"})));
    return;
  }
  if (dd.msg && !e.action) { e.waitUntil(clients.openWindow("adm.html#mensagens")); return; }
  if (dd.rsvp && e.action === "resp" && dd.ru && dd.para) {   // insight: "Já respondi" direto na notificação
    const r = dd.rsvp;
    e.waitUntil(fetch(dd.ru, {method: "POST", mode: "no-cors", body: JSON.stringify({email: dd.para, ok: true,
        painel: "✋|resp|" + r.d + "|" + r.tk + "|" + r.e, aparelho: aparelhoSW()})}).catch(() => {}).then(() => acordarPainel(r, dd.para))
      .then(() => self.registration.showNotification("✅ Marcado como respondido", {body: "Valeu, irmão! A sua resposta no " + r.e + " entra na lista.", icon: "icon-192.png", badge: "notif-badge.png", tag: "rsvp-ok"})));
    return;
  }
  if (dd.rsvp && dd.rsvp.m === "ins" && !e.action) { e.waitUntil(clients.openWindow("adm.html#insight")); return; }
  if (dd.rsvp && (e.action === "vou" || e.action === "nao") && dd.ru && dd.para) {   // respondeu direto na notificação
    const r = dd.rsvp, vou = e.action === "vou";
    e.waitUntil(fetch(dd.ru, {method: "POST", mode: "no-cors", body: JSON.stringify({email: dd.para, ok: true,
        painel: "✋|" + e.action + "|" + r.d + "|" + r.tk + "|" + r.e, aparelho: aparelhoSW()})}).catch(() => {}).then(() => acordarPainel(r, dd.para))
      .then(() => self.registration.showNotification(vou ? "✅ Presença confirmada" : "👍 Resposta registrada",
        {body: vou ? "Valeu, irmão! Você vai na " + r.e + ". 🤝" : "Tranquilo, irmão. Registramos que você não vai na " + r.e + ".",
         icon: "icon-192.png", badge: "notif-badge.png", tag: "rsvp-ok"})));
    return;
  }
  if (dd.rsvp && !e.action) {   // tocou na notificação (no iPhone não há botões): abre o painel para responder
    e.waitUntil(clients.openWindow("adm.html#eventos")); return;
  }
  if (dd.link && (e.action === "link" || /iPhone|iPad|iPod/.test(navigator.userAgent))) {   // ex.: "Abrir meu boleto"
    e.waitUntil(clients.openWindow(dd.link)); return;
  }
  if (dd.zap && dd.zap !== "__STATUS_PAGAMENTOS__" && (e.action === "zap" || /iPhone|iPad|iPod/.test(navigator.userAgent))) {
    e.notification.close();   // a mensagem já veio pronta: abre direto o WhatsApp, sem precisar entrar no app
    // api.whatsapp.com direto: o link curto wa.me passa por um redirecionamento que troca os emojis por "�"
    e.waitUntil(clients.openWindow("https://api.whatsapp.com/send/?text=" + encodeURIComponent(dd.zap)));
    return;
  }
  if (dd.zap && (e.action === "zap" || /iPhone|iPad|iPod/.test(navigator.userAgent))) {
    e.waitUntil(caches.open("insanos-zap").then(c => c.put("zap", new Response(JSON.stringify({ts: Date.now(), txt: dd.zap}))))
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
           "sw.js": sw,
           "versao.json": json.dumps({"v": VERSAO_ATUAL})}
    for t in (192, 512):
        caminho = os.path.join(PASTA, "app", f"icon-{t}.png")
        if os.path.exists(caminho):
            with open(caminho, "rb") as f:
                arq[f"icon-{t}.png"] = f.read()
    pasta_app = os.path.join(PASTA, "app")
    for nome in (sorted(os.listdir(pasta_app)) if os.path.isdir(pasta_app) else []):
        if (nome.startswith("notif-") or nome.startswith("atalho-") or nome.startswith("pin-")) and nome.endswith(".png"):     # notificações, atalhos e pins (honrarias)
            with open(os.path.join(pasta_app, nome), "rb") as f:
                arq[nome] = f.read()
    if MAPA_PNG:
        arq["mapa-divisao.png"] = MAPA_PNG
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
    col_ent = next((k for k, v in enumerate(cab_di) if _semac(v).startswith("entrada")), None)   # "Entrada no MC" (honrarias)
    entradas = {}
    situ = {}          # apelido -> {"s": ativo|afastado|excluido, "desde": data, "linha": n, "d": Numero Lista, "grau": ...}
    ordem_di = []      # apelidos na ordem das linhas da aba
    emails = {}
    # Dados Integrantes: apelido -> (número, cargo, ordem da hierarquia)
    info = {}
    extra_colete = {}
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
            if col_ent is not None and col_ent < len(r_full):
                ve = r_full[col_ent]
                if isinstance(ve, datetime.datetime):
                    ve = ve.date()
                if isinstance(ve, (int, float)) and not isinstance(ve, bool) and 20000 < ve < 80000:   # data vinda do Google como número
                    ve = datetime.date(1899, 12, 30) + datetime.timedelta(days=int(ve))
                if isinstance(ve, datetime.date):
                    entradas[apelido.lower()] = ve.isoformat()
                else:
                    mm = re.match(r"\s*(\d{1,2})/(\d{1,2})/(\d{2,4})", texto(ve))
                    if mm:
                        a = int(mm.group(3)); a = a + 2000 if a < 100 else a
                        try:
                            entradas[apelido.lower()] = datetime.date(a, int(mm.group(2)), int(mm.group(1))).isoformat()
                        except ValueError:
                            pass
            if col_email is not None:
                emails[apelido.lower()] = texto(r_full[col_email] if col_email < len(r_full) else "").lower()
            # dados para a Solicitação kit/colete (vão SÓ para o ADM do computador, nunca para o link)
            extra_colete[apelido.lower()] = {"nc": texto(r[7]), 
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
    # ADM liberados pela tela Acessos (além dos fixos do divisao.json): só quem tem e-mail na planilha
    ADM_EXTRA[:] = [n for n in _json_estado("adm_extra.json", []) if isinstance(n, str) and n.lower() not in [a.lower() for a in AUTORIZADOS]]
    for n in ADM_EXTRA:
        e = emails.get(n.lower(), "")
        if "@" in e and e not in senhas:
            senhas.append(e)
        elif "@" not in e:
            print(f"Aviso: {n} foi liberado como ADM, mas não tem e-mail na aba Dados Integrantes.")
    ADM_EMAILS.clear(); ADM_EMAILS.update(e.strip().lower() for e in senhas if "@" in e)
    senhas_so_geral = [e for e in (email_de(nome) for nome in AUTORIZADOS_SO_GERAL) if e not in senhas]

    linhas = list(wb["Presenca"].iter_rows(min_row=1, max_row=200, max_col=60, values_only=True))
    # abas com o mesmo formato da Presenca (Insights e Staff), se existirem na planilha
    abas_extras = {}
    for chave, nome_aba in (("ins", "insights"), ("staff", "staff")):
        aba = next((n for n in wb.sheetnames if n.strip().lower() == nome_aba), None)
        if aba:
            abas_extras[chave] = list(wb[aba].iter_rows(min_row=1, max_row=300, max_col=60, values_only=True))
    # motivos das faltas (aba "Justificativas", gravada pelo Lançar): só para o comando e para o próprio irmão
    global JUST
    JUST = {}
    try:
        if "Justificativas" in wb.sheetnames:
            for r in wb["Justificativas"].iter_rows(min_row=2, max_col=5, values_only=True):
                dj, ej, aj, nj, mj = (list(r) + [None] * 5)[:5]
                if not (dj and ej and nj and mj) or str(aj or "Presenca").strip().lower() != "presenca":
                    continue
                dj = dj.date().isoformat() if hasattr(dj, "date") else str(dj).strip()[:10]
                JUST.setdefault((dj, str(ej).strip().lower()), {})[str(nj).strip()] = str(mj).strip()
    except Exception as e:
        print(f"Aviso: justificativas não lidas ({type(e).__name__}: {e}).")
    # informações dos eventos (aba "Eventos", gravada pelo Lançar a partir da Lista): tipo, horário, local, endereço, campos da lista
    global INFO_EV
    INFO_EV = {}
    try:
        if "Eventos" in wb.sheetnames:
            for r in wb["Eventos"].iter_rows(min_row=2, max_col=8, values_only=True):
                dj, ej, tj, hj, lj, enj, oj, fj = (list(r) + [None] * 8)[:8]
                if not (dj and ej):
                    continue
                dj = dj.date().isoformat() if hasattr(dj, "date") else str(dj).strip()[:10]
                try:
                    det = json.loads(fj) if fj else {}
                except Exception:
                    det = {}
                inf_ev = {"tnome": str(tj or ""), "hora": str(hj or ""), "local": str(lj or ""), "endereco": str(enj or ""), "obs": str(oj or "")}
                inf_ev.update({k: det[k] for k in ("tipo", "f", "mapas") if k in det})
                INFO_EV[(dj, str(ej).strip().lower())] = inf_ev   # (nome próprio: "info" é o dos integrantes)
    except Exception as e:
        print(f"Aviso: informações dos eventos não lidas ({type(e).__name__}: {e}).")
    # enquete do comando: a última da aba Enquetes, com os votos da aba Votos (um por irmão; vale o último)
    global ENQ
    ENQ = {}
    try:
        if "Enquetes" in wb.sheetnames:
            lin_e = [r for r in wb["Enquetes"].iter_rows(min_row=2, max_col=6, values_only=True) if r and r[0]]
            if lin_e:
                r = lin_e[-1]
                try:
                    ops_e = [str(x) for x in json.loads(r[4] or "[]")]
                except Exception:
                    ops_e = []
                ENQ = {"id": str(r[0]), "criada": str(r[1] or ""), "por": str(r[2] or ""), "p": str(r[3] or ""), "ops": ops_e,
                       "aberta": str(r[5] or "").strip().lower() == "aberta", "votos": {}}
                if "Votos" in wb.sheetnames:
                    for vv in wb["Votos"].iter_rows(min_row=2, max_col=5, values_only=True):
                        if vv and str(vv[1] or "") == ENQ["id"] and vv[2]:
                            ENQ["votos"][str(vv[2]).strip().lower()] = (str(vv[3] or "").strip(), str(vv[4] or "").strip())
    except Exception as e:
        print(f"Aviso: enquete não lida ({type(e).__name__}: {e}).")
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
             "x": extra_colete.get(nome.lower(), {}), "ent": entradas.get(nome.lower(), "")}
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
        html = texto_divisao(f.read())
    gerado = datetime.datetime.now().strftime("%d/%m/%Y às %H:%M")
    global MANUT_HTML
    MANUT_HTML = AVISO_MANUTENCAO if os.path.exists(os.path.join(PASTA, "MANUTENCAO.txt")) else ""
    if MANUT_HTML:
        print("ATENÇÃO: aviso de manutenção LIGADO (para tirar, rode o DesligarManutencao.bat).")
    html = html.replace("__IMG__", img).replace("__GERADO__", gerado).replace("__MANUT__", MANUT_HTML)

    # versão do vídeo de fundo (pelo conteúdo do arquivo): vídeo novo = endereço novo, e o celular baixa sozinho
    try:
        import hashlib
        with open(os.path.join(PASTA, "fundo.mp4"), "rb") as fv:
            ver_video = hashlib.md5(fv.read()).hexdigest()[:10]
    except Exception:
        ver_video = ""
    com_versao = lambda t: t.replace('src="fundo.mp4"', f'src="fundo.mp4?v={ver_video}"') if ver_video else t

    def gravar(caminho, dados, senha=None, sem_js="", extras=(), rotulo="E-mail", so_geral=False):
        pagina = com_versao(html.replace("__NOSCRIPT__", sem_js, 1).replace("__DATA__", json.dumps(dados, ensure_ascii=False), 1))
        pagina = pagina.replace("</body>", snip_versao(False) + "</body>", 1).replace("__VERSAO__", VERSAO_ATUAL)
        if dados.get("modo") == "adm":       # SOS dentro do painel ADM (quadrado do topo)
            pagina = pagina.replace("</body>", f'<script>window.SOS_URL={json.dumps(OTP_URL if otp_ativo() else "")};window.SOS_DIV={json.dumps(str(DIV.get("nome", "")) + " " + str(DIV.get("regional", "")))};</script>' + SOS_MODULO + MSG_MODULO + SELECT_MODULO + PTR_MODULO + REINST_MODULO + "</body>", 1)
        if dados.get("modo") == "geral" and caminho == SAIDA:
            global GERAL_RAW
            GERAL_RAW = pagina
        if senha:
            base = senhas + (senhas_so_geral if so_geral else [])
            pagina = proteger(pagina, base + [x for x in extras if x not in base],
                              os.path.splitext(os.path.basename(caminho or SAIDA_ADM))[0], senha, img, rotulo)
            pagina = com_versao(pagina)
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
    INFO_LOGIN = info_login(gerado, _v)
    # tela de CPF do ADM: ponto verde quando há pagamento novo que você ainda não viu (some ao abrir a tela Pagamentos)
    _up = _v.get("ultimo_pagamento") or {}
    ULTIMO_PAG_TS = _up.get("ts", "")
    INFO_LOGIN_ADM = INFO_LOGIN + (
        f'<span id="novoPagTxt" hidden> · novo pagamento {_up.get("hora", "")}</span>'
        f'<span id="novoPag" class="novo" data-ts="{_up.get("ts", "")}" title="Pagamento novo: abra a tela Pagamentos" hidden></span>'
        if _up.get("ts") else "")
    global CAVEIRA_URL
    _mc = re.search(r'const CAVEIRA="(data:image/[^"]+)"', html)
    CAVEIRA_URL = _mc.group(1) if _mc else ""
    global MEU_PUSH
    try:
        MEU_PUSH = {"pub": chave_push_publica(), "url": ACESSOS_URL} if ACESSOS_URL else None
    except Exception:
        MEU_PUSH = None
    global PAG_MEU
    try:   # mensalidade de cada irmão (Meu painel › Minha mensalidade): situação, link do boleto e vencimento
        PAG_MEU = {"venc": vencimento(datetime.date.today().year, datetime.date.today().month).isoformat(), "st": {}, "lk": {}}
        for r in pgto:
            rot = texto(r[1])
            m_ = next((m for m in membros if rot and rot.lower().endswith(m["n"].lower())), None)
            if m_ and m_["n"] not in PAG_MEU["st"]:
                PAG_MEU["st"][m_["n"]] = texto(r[2]) or "Em aberto"
                if texto(r[9]).startswith("http"):
                    PAG_MEU["lk"][m_["n"]] = texto(r[9])
    except Exception as e:
        PAG_MEU = {}; print(f"Aviso: mensalidade do Meu painel não montada ({type(e).__name__}: {e}).")
    try:
        try:
            montar_mapa(eventos)
        except Exception as e:
            print(f"Aviso: mapa da divisão não montado ({type(e).__name__}: {e}).")
        montar_pessoais(membros, eventos, eventos_ins, gerado)
    except Exception as e:
        PESSOAIS.clear(); print(f"Aviso: o 'Meu painel' dos irmãos não foi montado ({type(e).__name__}: {e}).")
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
    hoje = datetime.date.today()
    venc = vencimento(hoje.year, hoje.month)
    pag["venc"] = venc.isoformat()
    pag["msg"] = mensagem_pagamentos(venc, membros)   # o texto é fixo: só mudam o mês/vencimento, a lista e a assinatura
    import re as _re
    _v = vencimento(hoje.year, hoje.month)
    pag["msg"] = _re.sub(r"\(Vencimento\s*\d{1,2}/\d{1,2}/\d{2,4}\)", f"(Vencimento {_v:%d/%m/%Y})", pag["msg"])
    global ZAP_PAG
    pag["agenda"] = agenda_ativa()
    pag["verif_antiga"] = False
    pag["horarios"] = HORARIOS_VERIFICACAO
    pag["travado"], pag["manual"], pag["boleto_mes"], pag["pago_em"] = {}, {}, {}, {}
    _man = _manual_carregar(vencimento(hoje.year, hoje.month))
    for m in membros:
        if m["n"] in _man:
            pag["manual"][m["n"]] = _man[m["n"]]
    pag["obs"] = {}
    arq_ant = os.path.join(PASTA, "pagamentos_antecipados.json")     # pagos antes do dia 1º (boletos do mês seguinte)
    if os.path.exists(arq_ant):
        try:
            with open(arq_ant, encoding="utf-8") as f:
                ant = json.load(f)
            hoje_m = datetime.date.today().strftime("%Y-%m")
            if str(ant.get("mes", "")) > hoje_m and ant.get("pagos"):
                lista = []
                for rot, quando in ant["pagos"].items():
                    nome = next((m["n"] for m in membros if str(rot).lower().endswith(m["n"].lower())), str(rot).split(" - ")[-1].strip())
                    lista.append({"n": nome, "quando": quando})
                pag["antecipados"] = {"mes": ant["mes"], "lista": lista}
        except Exception as e:
            print(f"Aviso: não consegui ler os pagamentos antecipados ({e}).")
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
    try:
        montar_lembretes(membros, eventos, eventos_ins, pag)
    except Exception as e:
        print(f"Aviso: lembretes não montados ({type(e).__name__}: {e}).")
    ACESSOS = None
    if ACESSOS_URL:
        _so = [EMAIL_APELIDO_INV.get(n.lower(), "") for n in ACESSOS_SOMENTE]   # a tela Acessos é só de quem está em acessos_somente
        _so = [e for e in _so if "@" in e]
        visitantes = atualizar_visitantes()
        if _so:
            ACESSOS = {"cifrado": cifrar_para({"url": ACESSOS_URL, "chave": ACESSOS_CHAVE, "mapa": dict(EMAIL_APELIDO),
                                              "regs": (buscar_acessos() or [])[1:], "planilha": PLANILHA_URL,
                                              "fixos": AUTORIZADOS, "extra": ADM_EXTRA, "vis": visitantes,
                                              "comEmail": sorted(ap for ap, e in EMAIL_APELIDO_INV.items() if "@" in e),
                                              "push": aparelhos_com_avisos()}, _so)}
        else:
            print("Aviso: e-mail de " + ", ".join(ACESSOS_SOMENTE) + " não encontrado: a tela Acessos não vai aparecer.")
    FLY = {}      # "Gerar fly" removido: o fly oficial é feito no site do clube
    try:   # a mensagem pronta para o aviso de pagamento novo (abre direto no WhatsApp, sem entrar no app)
        ZAP_PAG = zap_pagamentos(pag, membros)
    except Exception as e:
        ZAP_PAG = ""; print(f"Aviso: mensagem dos pagamentos para o WhatsApp não montada ({type(e).__name__}: {e}).")
    global ZAP_CTX, HOJE_EVS
    try:
        _hj = datetime.date.today().isoformat()
        HOJE_EVS = [(e["e"], str(INFO_EV.get((e["d"], e["e"].strip().lower()), {}).get("hora") or "")) for e in eventos if e["d"] == _hj]
    except Exception:
        HOJE_EVS = []
    try:
        _ent = {}
        for e in eventos + eventos_ins:
            for k, v in enumerate(e["p"]):
                if v not in (None, "") and k < len(membros):
                    n_ = membros[k]["n"]
                    if n_ not in _ent or e["d"] < _ent[n_]:
                        _ent[n_] = e["d"]
        ZAP_CTX = {"membros": membros, "entrada": _ent}
    except Exception as e:
        print(f"Aviso: contexto da mensagem do evento não montado ({type(e).__name__}: {e}).")
    def _ev_adm(e):
        k_ = (e["d"], e["e"].strip().lower()); x = dict(e)
        if k_ in JUST: x["j"] = JUST[k_]
        if k_ in INFO_EV: x["info"] = INFO_EV[k_]
        return x
    eventos_adm = [_ev_adm(e) for e in eventos]
    enq_adm = None
    if ENQ:
        nomes_ = {m["n"].lower(): m["n"] for m in membros}
        res_ = {o: [] for o in ENQ["ops"]}
        for em_, (ap_, op_) in ENQ["votos"].items():
            nm_ = nomes_.get(ap_.lower()) or nomes_.get(str(EMAIL_APELIDO.get(em_, "")).lower()) or ap_ or em_.split("@")[0]
            res_.setdefault(op_, []).append(nm_)
        enq_adm = {k: ENQ[k] for k in ("id", "criada", "por", "p", "ops", "aberta")}
        _por = str(ENQ.get("por") or "").strip().lower()
        _ap = str(EMAIL_APELIDO.get(_por, "") or (_por if "@" not in _por else _por.split("@")[0]))
        enq_adm["por"] = nomes_.get(_ap.lower(), _ap)   # o apelido de quem criou
        enq_adm.update({"res": res_, "total": len(membros)})
    gravar(SAIDA_ADM, {"modo": "adm", "resp": (ACESSOS_SOMENTE or [""])[0], "enq": enq_adm, "m": membros, "ev": eventos_adm, "ev_ins": eventos_ins, "ev_staff": eventos_staff, "molde": molde, "pag": pag, "fotos": FOTOS_GAL, "fly": FLY, "acessos": ACESSOS, "afastados": AFASTADOS, "push": PUSH, "nuvem": NUVEM}, "Painel ADM - Acesso restrito ao Comando.",
           extras=[x for x in extras if x], rotulo="E-mail ou senha de acesso")
    dados_adm = {"modo": "adm", "reg_adm": str(DIV.get("regional_adm", "")), "reinst": AVISO_REINSTALAR,
                 "enq": enq_adm, "resp": (ACESSOS_SOMENTE or [""])[0], "quem": {e: ap for e, ap in EMAIL_APELIDO.items() if e in ADM_EMAILS}, "m": membros, "ev": eventos_adm, "ev_ins": eventos_ins, "ev_staff": eventos_staff, "molde": molde, "pag": pag, "fotos": FOTOS_GAL, "fly": FLY, "acessos": ACESSOS, "afastados": AFASTADOS, "push": PUSH, "nuvem": NUVEM}
    global VERSAO_ATUAL
    import hashlib as _h
    VERSAO_ATUAL = _h.md5(json.dumps(_assinatura_conteudo({"modo": "geral", "m": sem_tel, "ev": eventos}, dados_adm), sort_keys=True, default=str).encode()).hexdigest()[:12]
    # no link: com nome completo e CPF dos integrantes (decisão do Eduardo, ciente do risco), para a Solicitação kit/colete.
    # Para voltar a tirar do link, troque LINK_COM_NOME_E_CPF para False no começo deste arquivo.
    membros_link = membros if LINK_COM_NOME_E_CPF else \
        [{**m, "x": {k: v for k, v in m.get("x", {}).items() if k in ("col", "tamc")}} for m in membros]
    pagina_adm_link = gravar(None, {**dados_adm, "m": membros_link}, "Presença, insights e a caminhada de cada irmão. 💀",
                             extras=[x for x in extras if x] if ADM_NO_LINK_ACEITA_SENHA_COLETIVA else [],
                             rotulo="Seu e-mail ou a senha de acesso" if ADM_NO_LINK_ACEITA_SENHA_COLETIVA else "Seu e-mail cadastrado na divisão")
    # PROTÓTIPO DO NOVO LAYOUT (novo.html): mesmos dados e mesma entrada, só para olhar. O painel atual não muda.
    # Para tirar: apague o arquivo template_novo.html do repositório.
    pagina_novo = None
    tpl_novo = os.path.join(PASTA, "template_novo.html")
    if os.path.exists(tpl_novo):
        _html_atual = html
        try:
            with open(tpl_novo, encoding="utf-8") as f:
                html = f.read().replace("__DIV_NOME__", str(DIV.get("nome", ""))).replace("__DIV_REG__", str(DIV.get("regional", "")))
            pagina_novo = gravar(None, {**dados_adm, "m": membros_link, "fila": FILA_RECENTE}, "Novo layout (protótipo) · só para olhar",
                                 extras=[], rotulo="Seu e-mail cadastrado na divisão")
        except Exception as e:
            print(f"Aviso: protótipo do novo layout não gerado ({type(e).__name__}: {e}).")
            pagina_novo = None
        finally:
            html = _html_atual
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
        gravar(os.path.join(PASTA_IND, f"{DIV_ROTULO} - {nome_arq}.html"),
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
    publicar_se_mudou({"index.html": INDEX_SO_ADM, "adm.html": pagina_adm_link, **arquivos_app(), **({"novo.html": pagina_novo} if pagina_novo else {})},
                      _assinatura_conteudo({"modo": "geral", "m": sem_tel, "ev": eventos}, dados_adm),
                      forcar="--publicar" in sys.argv)

if __name__ == "__main__":
    main()
