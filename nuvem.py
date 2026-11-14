"""
Painel Insanos na nuvem (GitHub Actions).
1) Lê a planilha do Google (conta de serviço) e monta uma planilha temporária só com as abas que o painel usa.
2) Roda o mesmo programa do painel (painel.py) e publica neste repositório.
Nada de dados pessoais fica no repositório: a planilha temporária existe só durante a execução.
"""
import os, sys, json, datetime, unicodedata, re, base64
import urllib.parse

PASTA = os.path.dirname(os.path.abspath(__file__))
ABAS = ["Presenca", "Insights", "Dados Integrantes", "PGTO"]   # a mensagem dos pagamentos é gerada pelo programa
ABAS_OPC = ["Eventos", "Justificativas", "Enquetes", "Votos", "APH"]   # gravadas pelo Lançar: lidas quando existem (informações dos eventos e motivos das faltas)
COLS_DATA = {"Presenca": [1], "Insights": [1], "Dados Integrantes": [9, 17]}   # colunas de data (0 = A): B, J e R

def _serial_para_data(v):
    try:
        return datetime.datetime(1899, 12, 30) + datetime.timedelta(days=float(v))
    except Exception:
        return v

def montar_planilha(valores_por_aba, destino):
    """valores_por_aba: {aba: [[células...]]} (valores já calculados) -> .xlsx temporário para o painel ler."""
    import openpyxl
    wb = openpyxl.Workbook(); wb.remove(wb.active)
    for aba, linhas in valores_por_aba.items():
        ws = wb.create_sheet(aba)
        cd = COLS_DATA.get(aba, [])
        for r, linha in enumerate(linhas, start=1):
            for c, v in enumerate(linha, start=1):
                if v == "" or v is None:
                    continue
                if (c - 1) in cd and isinstance(v, (int, float)) and r >= (2 if aba == "Dados Integrantes" else 8) and 20000 < v < 80000:
                    v = _serial_para_data(v)
                ws.cell(r, c).value = v
    wb.save(destino)

def api_google(cred_json):
    from google.oauth2 import service_account
    from googleapiclient.discovery import build
    cred = service_account.Credentials.from_service_account_info(json.loads(cred_json),
           scopes=["https://www.googleapis.com/auth/spreadsheets"])
    return build("sheets", "v4", credentials=cred, cache_discovery=False)

def ler_google(planilha_id, cred_json, api=None):
    api = api or api_google(cred_json)
    meta = api.spreadsheets().get(spreadsheetId=planilha_id, fields="properties.title,sheets.properties.title").execute()
    existentes = [s["properties"]["title"] for s in meta["sheets"]]
    print(f"Planilha do Google: \"{meta.get('properties', {}).get('title', '?')}\" com {len(existentes)} abas: " + ", ".join(f'"{x}"' for x in existentes))
    # acha cada aba sem ligar para espaços, acentos ou maiúsculas ("Presença " = "Presenca")
    norm = lambda x: "".join(ch for ch in unicodedata.normalize("NFD", str(x).strip().lower()) if not unicodedata.combining(ch))
    real = {}
    for a in ABAS + ABAS_OPC:
        achou = next((e for e in existentes if norm(e) == norm(a)), None)
        if achou:
            real[a] = achou
    faltam = [a for a in ABAS if a not in real]
    if faltam:
        print("Aviso: abas não encontradas na planilha do Google:", ", ".join(faltam))
    if "Presenca" not in real or "Dados Integrantes" not in real:
        print("[ERRO] A planilha precisa ter as abas Presenca e Dados Integrantes. Confira se o PLANILHA_ID é o da planilha certa.")
        sys.exit(1)
    nomes = list(real)
    res = api.spreadsheets().values().batchGet(spreadsheetId=planilha_id, ranges=[f"'{real[a]}'!A1:J5000" if a in ABAS_OPC else f"'{real[a]}'!A1:BZ500" for a in nomes],
          valueRenderOption="UNFORMATTED_VALUE", dateTimeRenderOption="SERIAL_NUMBER").execute()
    return {a: vr.get("values", []) for a, vr in zip(nomes, res.get("valueRanges", []))}

# ---------------------------------------------------------------------
#  Fila: lançamentos que o painel mandou direto para a planilha (aba "Fila")
#  Colunas: A data/hora · B quem · C tipo · D dados (JSON) · E situação (pendente/feito/erro) · F resultado
# ---------------------------------------------------------------------
def _sa(x):
    return "".join(ch for ch in unicodedata.normalize("NFD", str(x or "").strip().lower()) if not unicodedata.combining(ch))

def _col(n):   # 1 -> A
    s = ""
    while n:
        n, r = divmod(n - 1, 26); s = chr(65 + r) + s
    return s

def _aba_real(api, pid, nome):
    meta = api.spreadsheets().get(spreadsheetId=pid, fields="sheets.properties.title").execute()
    return next((s["properties"]["title"] for s in meta["sheets"] if _sa(s["properties"]["title"]) == _sa(nome)), None)

def _ler(api, pid, rng):
    return api.spreadsheets().values().get(spreadsheetId=pid, range=rng, valueRenderOption="UNFORMATTED_VALUE",
                                           dateTimeRenderOption="SERIAL_NUMBER").execute().get("values", [])

def _gravar(api, pid, dados):
    if dados:
        api.spreadsheets().values().batchUpdate(spreadsheetId=pid, body={"valueInputOption": "USER_ENTERED",
            "data": [{"range": r, "values": [[v]]} for r, v in dados]}).execute()

def migrar_entradas(api, pid):
    """Honrarias: cria a coluna "Entrada no MC" na aba Dados Integrantes (no fim, sem mexer nas outras) e preenche
    as datas do arquivo entradas_iniciais.json, só nas células vazias. Roda em toda rodada, mas só escreve uma vez."""
    if not os.path.exists("entradas_iniciais.json"):
        return
    try:
        ini = {_sa(k): v for k, v in json.load(open("entradas_iniciais.json", encoding="utf-8")).items() if not str(k).startswith("_")}
        aba = _aba_real(api, pid, "Dados Integrantes")
        if not aba or not ini:
            return
        v = _ler(api, pid, f"'{aba}'!A1:BZ300")
        cab = [str(x).strip() for x in (v[0] if v else [])]
        col = next((i + 1 for i, x in enumerate(cab) if _sa(x).startswith("entrada")), None)
        gr = []
        if col is None:
            col = max([i + 1 for i, x in enumerate(cab) if x] or [0]) + 1
            gr.append((f"'{aba}'!{_col(col)}1", "Entrada no MC"))
        for r in range(2, len(v) + 1):
            lin = list(v[r - 1]) + [""] * 80
            d = ini.get(_sa(lin[1]))
            if d and not str(lin[col - 1]).strip():
                gr.append((f"'{aba}'!{_col(col)}{r}", d))
        if gr:
            _gravar(api, pid, gr)
            print(f"Honrarias: coluna 'Entrada no MC' ({_col(col)}) com {len(gr) - (1 if gr and gr[0][1] == 'Entrada no MC' else 0)} data(s) preenchida(s).")
    except Exception as e:
        print(f"Aviso: datas de entrada não preenchidas ({type(e).__name__}: {e}).")

def gravar_justificativas(api, pid, d, quem=""):
    """Motivos das faltas do evento (vindos do Lançar): aba "Justificativas" (criada sozinha). Ao salvar o mesmo evento
    de novo, os motivos dele são trocados pelos novos (os dos outros eventos ficam como estão)."""
    if d.get("justificativas") is None or str(d.get("aba") or "Presenca").strip().lower() != "presenca":
        return
    aba = _aba_real(api, pid, "Justificativas")
    if not aba:
        _estrutura(api, pid, [{"addSheet": {"properties": {"title": "Justificativas"}}}]); aba = "Justificativas"
    cab = ["Data", "Evento", "Aba", "Irmão", "Motivo", "Lançado por", "Quando"]
    v = api.spreadsheets().values().get(spreadsheetId=pid, range=f"'{aba}'!A1:G5000").execute().get("values", [])
    data, ev = str(d.get("data") or "")[:10], str(d.get("evento") or "").strip()
    outros = [list(r) + [""] * (7 - len(r)) for r in v[1:] if r and not (str(r[0])[:10] == data and str((r + ["", ""])[1]).strip().lower() == ev.lower())]
    agora = datetime.datetime.now().strftime("%d/%m/%Y %H:%M")
    novos = [[data, ev, "Presenca", n, str(m).strip()[:80], str(quem or "").split("@")[0], agora]
             for n, m in (d.get("justificativas") or {}).items() if str(m or "").strip()]
    api.spreadsheets().values().clear(spreadsheetId=pid, range=f"'{aba}'!A1:G5000").execute()
    api.spreadsheets().values().update(spreadsheetId=pid, range=f"'{aba}'!A1", valueInputOption="RAW",
                                       body={"values": [cab] + outros + novos}).execute()

def gravar_info_evento(api, pid, d, quem=""):
    """Informações do evento vindas da Lista (tipo, horário, local, endereço, observação e os campos da lista):
    aba "Eventos" (criada sozinha), uma linha por evento; salvar o mesmo evento de novo atualiza a linha dele."""
    info = d.get("info")
    if not info or str(d.get("aba") or "Presenca").strip().lower() != "presenca":
        return
    aba = _aba_real(api, pid, "Eventos")
    if not aba:
        _estrutura(api, pid, [{"addSheet": {"properties": {"title": "Eventos"}}}]); aba = "Eventos"
    cab = ["Data", "Evento", "Tipo", "Horário", "Local", "Endereço", "Observação", "Detalhes da lista", "Lançado por", "Quando"]
    v = api.spreadsheets().values().get(spreadsheetId=pid, range=f"'{aba}'!A1:J3000").execute().get("values", [])
    data, ev = str(d.get("data") or "")[:10], str(d.get("evento") or "").strip()
    outros = [list(r) + [""] * (10 - len(r)) for r in v[1:] if r and not (str(r[0])[:10] == data and str((r + ["", ""])[1]).strip().lower() == ev.lower())]
    det = json.dumps({"tipo": info.get("tipo", ""), "f": info.get("f", {}), "mapas": bool(info.get("mapas"))}, ensure_ascii=False)[:45000]
    linha = [data, ev, str(info.get("tnome") or info.get("tipo") or "")[:40], str(info.get("hora") or "")[:20], str(info.get("local") or "")[:200],
             str(info.get("endereco") or "")[:300], str(info.get("obs") or "")[:500], det, str(quem or "").split("@")[0],
             datetime.datetime.now().strftime("%d/%m/%Y %H:%M")]
    api.spreadsheets().values().clear(spreadsheetId=pid, range=f"'{aba}'!A1:J3000").execute()
    api.spreadsheets().values().update(spreadsheetId=pid, range=f"'{aba}'!A1", valueInputOption="RAW",
                                       body={"values": [cab] + outros + [linha]}).execute()

def aplicar_enquete(api, pid, d, quem=""):
    """Enquete do comando (aba Enquetes, criada sozinha): uma linha por enquete; a nova encerra a que estava aberta.
    enquete_fim: encerra a enquete pelo id."""
    aba = _aba_real(api, pid, "Enquetes")
    if not aba:
        _estrutura(api, pid, [{"addSheet": {"properties": {"title": "Enquetes"}}}]); aba = "Enquetes"
    v = api.spreadsheets().values().get(spreadsheetId=pid, range=f"'{aba}'!A1:F500").execute().get("values", [])
    cab = ["ID", "Criada", "Por", "Pergunta", "Opções", "Situação"]
    linhas = [list(r) + [""] * (6 - len(r)) for r in v[1:] if r]
    if d.get("tipo") == "enquete_fim":
        for r in linhas:
            if str(r[0]) == str(d.get("id")):
                r[5] = "encerrada"
        res = "enquete encerrada"
    else:
        ops = [str(x).strip()[:60] for x in (d.get("opcoes") or []) if str(x).strip()][:4]
        perg = str(d.get("pergunta") or "").strip()[:200]
        if not perg or len(ops) < 2:
            return "enquete sem pergunta ou com menos de 2 opções"
        for r in linhas:
            if r[5] == "aberta":
                r[5] = "encerrada"
        idn = str(d.get("id") or "E" + datetime.datetime.now().strftime("%Y%m%d%H%M%S"))[:40]
        linhas.append([idn, datetime.datetime.now().strftime("%d/%m/%Y %H:%M"), str(quem or "").strip().lower(), perg, json.dumps(ops, ensure_ascii=False), "aberta"])
        ENQ_NOVA.append(perg)
        res = "enquete criada"
    api.spreadsheets().values().clear(spreadsheetId=pid, range=f"'{aba}'!A1:F500").execute()
    api.spreadsheets().values().update(spreadsheetId=pid, range=f"'{aba}'!A1", valueInputOption="RAW", body={"values": [cab] + linhas[-200:]}).execute()
    return res

def aplicar_evento(api, pid, d):
    """Lançar evento / Insights / Salvar no Excel: data + nome + marcações na aba (Presenca ou Insights)."""
    aba = _aba_real(api, pid, d.get("aba") or "Presenca")
    if not aba:
        return f"aba {d.get('aba')} não encontrada"
    v = _ler(api, pid, f"'{aba}'!A1:BZ60")
    lin = lambda r: v[r - 1] if len(v) >= r else []
    cols, c = {}, 4
    while True:
        nome = str((lin(6) + [""] * 80)[c - 1]).strip()
        if not nome or nome.lower().startswith("presen"):
            break
        cols[_sa(nome)] = c; c += 1
    alvo = datetime.datetime.strptime(d["data"], "%Y-%m-%d").date()
    def data_da_linha(r):
        x = (lin(r) + ["", "", ""])[1]
        if isinstance(x, (int, float)) and x > 20000:
            return (datetime.datetime(1899, 12, 30) + datetime.timedelta(days=float(x))).date()
        return None
    linha = None
    if d.get("atualizar", True):
        linha = next((r for r in range(8, 43) if data_da_linha(r) == alvo and _sa((lin(r) + ["", "", ""])[2]) == _sa(d.get("evento", ""))), None)
    if linha is None:
        linha = next((r for r in range(8, 43) if not str((lin(r) + ["", ""])[1]).strip()), None)
    if linha is None:
        return "aba cheia (linhas 8 a 42)"
    gr = [(f"'{aba}'!B{linha}", alvo.strftime("%d/%m/%Y")), (f"'{aba}'!C{linha}", d.get("evento", ""))]
    conv = {"sim": "SIM", "nao": "NÃO", "não": "NÃO", "staff": "STAFF", "-": "-", "": ""}
    ignorados = []
    for nome, marca in (d.get("marcacoes") or {}).items():
        cc = cols.get(_sa(nome))
        if not cc:
            ignorados.append(nome); continue
        gr.append((f"'{aba}'!{_col(cc)}{linha}", conv.get(_sa(marca), str(marca))))
    _gravar(api, pid, gr)
    return f"gravado na aba {aba}, linha {linha}" + (f" (sem coluna: {', '.join(ignorados)})" if ignorados else "")

def aplicar_pagamento(api, pid, d):
    """Compensação marcada no painel: PGTO coluna C (nunca mexe em quem já está Pago em Dia)."""
    aba = _aba_real(api, pid, "PGTO")
    if not aba:
        return "aba PGTO não encontrada"
    v = _ler(api, pid, f"'{aba}'!A1:C400")
    gr, feitos = [], []
    for nome, st in (d.get("status") or {}).items():
        for i, row in enumerate(v[1:], start=2):
            rot = str((row + ["", ""])[1]); atual = str((row + ["", "", ""])[2])
            if _sa(rot).endswith(_sa(nome)):
                if "pago" in _sa(atual) and "aguard" not in _sa(atual) and "ades" not in _sa(atual) and "ades" not in _sa(st or ""):
                    break
                gr.append((f"'{aba}'!C{i}", st or "")); feitos.append(nome); break
    _gravar(api, pid, gr)
    return "PGTO: " + (", ".join(feitos) if feitos else "nada alterado")

# ---------------------------------------------------------------------
#  Cadastro (novo, afastar, reativar, excluir) na planilha do Google — mesma lógica do PC.
#  Cópia de segurança: o Histórico de versões do Google Planilhas (Arquivo > Histórico de versões).
# ---------------------------------------------------------------------
def _ids(api, pid):
    meta = api.spreadsheets().get(spreadsheetId=pid, fields="sheets.properties(sheetId,title)").execute()
    return {s["properties"]["title"]: s["properties"]["sheetId"] for s in meta["sheets"]}

def _ler_f(api, pid, rng):
    """Valores com as FÓRMULAS (para saber quais células são fórmula)."""
    return api.spreadsheets().values().get(spreadsheetId=pid, range=rng, valueRenderOption="FORMULA").execute().get("values", [])

def _estrutura(api, pid, reqs):
    if reqs:
        api.spreadsheets().batchUpdate(spreadsheetId=pid, body={"requests": reqs}).execute()

def _faixa(sid, r0, r1, c0, c1):   # índices começando em 1, fim incluído
    return {"sheetId": sid, "startRowIndex": r0 - 1, "endRowIndex": r1, "startColumnIndex": c0 - 1, "endColumnIndex": c1}

def _di_mapa(api, pid):
    aba = _aba_real(api, pid, "Dados Integrantes")
    v = _ler(api, pid, f"'{aba}'!A1:AN400")
    cab = {}
    for c, x in enumerate((v[0] if v else []), start=1):
        if _sa(x):
            cab[_sa(x)] = c
    def col(*inicios):
        return next((c for k, c in cab.items() if any(k.startswith(i) for i in inicios)), None)
    C = {"num": 1, "ap": 2, "cargo": 3, "lista": col("numero lista") or 4, "grau": col("grau") or 5,
         "tel": col("telefone"), "nome": col("nome completo"), "nasc": col("nascimento"),
         "colete": col("colete"), "tam": col("tamanho"), "email": col("email", "e-mail"),
         "sit": col("situa"), "desde": col("afastado")}
    linhas = {}
    fim_tab = 1
    for r, row in enumerate(v[1:], start=2):
        ap = str((list(row) + ["", ""])[1]).strip()
        if ap:
            linhas[_sa(ap)] = r
    for r in range(2, len(v) + 1):          # bloco contínuo de integrantes (a "tabela"): até a primeira linha vazia
        if str((list(v[r - 1]) + ["", ""])[1]).strip():
            fim_tab = r
        else:
            break
    return aba, C, linhas, fim_tab, max(cab.values() or [1])

def _inserir_linha_di(api, pid, aba, sid, depois_de, ult_col):
    """Insere uma linha logo depois de 'depois_de' copiando as FÓRMULAS da linha de cima; devolve a linha nova."""
    nova = depois_de + 1
    _estrutura(api, pid, [
        {"insertDimension": {"range": {"sheetId": sid, "dimension": "ROWS", "startIndex": nova - 1, "endIndex": nova}, "inheritFromBefore": True}},
        {"copyPaste": {"source": _faixa(sid, depois_de, depois_de, 1, ult_col), "destination": _faixa(sid, nova, nova, 1, ult_col),
                       "pasteType": "PASTE_NORMAL"}}])
    # limpa o que não é fórmula (os dados copiados do vizinho)
    f = (_ler_f(api, pid, f"'{aba}'!A{nova}:{_col(ult_col)}{nova}") or [[]])[0]
    limpar = [(f"'{aba}'!{_col(c)}{nova}", "") for c in range(1, ult_col + 1) if not str((list(f) + [""] * ult_col)[c - 1]).startswith("=")]
    _gravar(api, pid, limpar)
    return nova

def _totais_linha(aba, v, linha, ult):
    """Fórmulas de Sim, Não e % da linha do evento (logo depois do último integrante), como no PC."""
    c_sim, c_nao, c_pct = ult + 1, ult + 2, ult + 3
    ls, ln = _col(c_sim), _col(c_nao)
    l7 = list(v[6]) if len(v) > 6 else []
    rot_sim = f"${ls}$7" if _sa((l7 + [""] * 90)[c_sim - 1]) == "sim" else '"Sim"'
    rot_nao = f"${ln}$7" if _sa((l7 + [""] * 90)[c_nao - 1]) in ("nao",) else '"Não"'
    fx = f"D{linha}:{_col(ult)}{linha}"
    return [(f"'{aba}'!{ls}{linha}", f'=IF(COUNTIF({fx},{rot_sim})=0,"",COUNTIF({fx},{rot_sim}))'),
            (f"'{aba}'!{ln}{linha}", f'=IF(COUNTIF({fx},{rot_nao})=0,"",COUNTIF({fx},{rot_nao}))'),
            (f"'{aba}'!{_col(c_pct)}{linha}", f'=IFERROR({ls}{linha}/({ls}{linha}+{ln}{linha}),"")')]

def _abas_eventos(api, pid):
    return [a for a in (_aba_real(api, pid, n) for n in ("Presenca", "Insights", "Staff")) if a]

def _cols_integrantes(v):
    cols, c = {}, 4
    l6 = list(v[5]) if len(v) > 5 else []
    while True:
        nome = str((l6 + [""] * 90)[c - 1]).strip()
        if not nome or nome.lower().startswith("presen"):
            return cols, c - 1
        cols[_sa(nome)] = c; c += 1

def _linha_adesao(v):
    return next((r for r in range(8, len(v) + 1) if _sa((list(v[r - 1]) + [""])[0]).startswith("ades")), None)

ABAS_SEM_USO = ["Staff", "Capa", "Sobrecapa", "Colete"]   # o painel não lê nenhuma delas (pedido do comando: remover)

def remover_abas_sem_uso(api, pid):
    """Apaga da planilha as abas que o painel não usa (só se ainda existirem). Dá para recuperar pelo Histórico de versões."""
    ids, pedidos, nomes = _ids(api, pid), [], []
    for n in ABAS_SEM_USO:
        real = _aba_real(api, pid, n)
        if real and real in ids:
            pedidos.append({"deleteSheet": {"sheetId": ids[real]}}); nomes.append(real)
    if pedidos and len(ids) - len(pedidos) >= 1:
        _estrutura(api, pid, pedidos)
        print("Abas sem uso apagadas da planilha: " + ", ".join(nomes))
    return nomes

def limpar_totais(api, pid):
    """Apaga as fórmulas de Sim, Não e % ao lado dos integrantes nas abas de eventos.
    O painel não usa essas contas (calcula sozinho) e, na planilha em português, elas davam #ERROR!."""
    feito = 0
    for aba in _abas_eventos(api, pid):
        v = _ler(api, pid, f"'{aba}'!A1:CZ60")
        cols, ult = _cols_integrantes(v)
        if ult < 4:
            continue
        fim = _linha_adesao(v)
        ate = (fim - 2) if fim else 42
        f = _ler_f(api, pid, f"'{aba}'!{_col(ult + 1)}8:{_col(ult + 3)}{ate}") or []
        if any(str(c).strip() for linha in f for c in linha):
            _gravar(api, pid, [(f"'{aba}'!{_col(c)}{r}", "") for r in range(8, ate + 1) for c in (ult + 1, ult + 2, ult + 3)])
            print(f"Aba {aba}: contas de Sim, Não e % apagadas (o painel calcula sozinho).")
            feito += 1
    return feito

def _colunas_nas_abas(api, pid, ap, cargo, apos, log):
    ids = _ids(api, pid)
    for aba in _abas_eventos(api, pid):
        v = _ler(api, pid, f"'{aba}'!A1:CZ60")
        cols, ult = _cols_integrantes(v)
        if _sa(ap) in cols:
            continue
        pos = cols.get(_sa(apos), ult) + 1
        sid = ids[aba]
        viz = pos - 1 if pos - 1 >= 4 else pos + 1
        _estrutura(api, pid, [
            {"insertDimension": {"range": {"sheetId": sid, "dimension": "COLUMNS", "startIndex": pos - 1, "endIndex": pos}, "inheritFromBefore": pos > 4}},
            {"copyPaste": {"source": _faixa(sid, 8, 42, viz if viz < pos else pos + 1, viz if viz < pos else pos + 1),
                           "destination": _faixa(sid, 8, 42, pos, pos), "pasteType": "PASTE_DATA_VALIDATION"}}])
        cargo_curto = re.sub(r"\s*\([^)]*\)\s*", " ", str(cargo or "")).strip()
        _gravar(api, pid, [(f"'{aba}'!{_col(pos)}6", ap), (f"'{aba}'!{_col(pos)}7", cargo_curto)])
        v = _ler(api, pid, f"'{aba}'!A1:CZ60"); ult += 1
        fim = _linha_adesao(v)
        if fim:   # "Adesão Integrante" da coluna nova: mesma fórmula do vizinho
            vz = pos - 1 if pos - 1 >= 4 else pos + 1
            _estrutura(api, pid, [{"copyPaste": {"source": _faixa(sid, fim, fim, vz, vz), "destination": _faixa(sid, fim, fim, pos, pos), "pasteType": "PASTE_FORMULA"}}])
        log.append(f"coluna criada na aba {aba}")

def _linha_pgto(api, pid, ap, r_di, log):
    aba = _aba_real(api, pid, "PGTO")
    if not aba:
        return
    v = _ler(api, pid, f"'{aba}'!A1:J400")
    ult = 1
    for r, row in enumerate(v[1:], start=2):
        rot = str((list(row) + ["", ""])[1])
        if rot.strip():
            ult = r
        if _sa(rot).endswith(_sa(ap)):
            return
    sid = _ids(api, pid)[aba]
    _estrutura(api, pid, [{"copyPaste": {"source": _faixa(sid, ult, ult, 1, 26), "destination": _faixa(sid, ult + 1, ult + 1, 1, 26), "pasteType": "PASTE_NORMAL"}}])
    di = _aba_real(api, pid, "Dados Integrantes")
    _gravar(api, pid, [(f"'{aba}'!B{ult + 1}", f"='{di}'!F{r_di}&\" - \"&'{di}'!B{r_di}"),
                       (f"'{aba}'!C{ult + 1}", "Em aberto"), (f"'{aba}'!J{ult + 1}", "")])
    log.append("linha criada na aba PGTO")

def _mudar_cargo(api, pid, d, aba, C, linhas, fim_tab, ult_col):
    """Mudança de cargo: novo cargo e grau e, se pedido, nova posição na lista ('apos').
    Dados Integrantes: os dados trocam de linha (as fórmulas ficam); PGTO: situação e links acompanham cada irmão;
    Presença e Insights: a coluna do irmão muda de lugar com todas as marcações."""
    ap = str(d.get("apelido", "")).strip()
    cargo, grau, apos = str(d.get("cargo") or "").strip(), str(d.get("grau") or "").strip(), str(d.get("apos") or "").strip()
    r = linhas.get(_sa(ap))
    if not r or r > fim_tab:
        return f"{ap} não está entre os integrantes (nada feito)"
    log = []
    # 1) Dados Integrantes: nova ordem das linhas
    rows = list(range(2, fim_tab + 1))
    lst = rows[:]
    r_apos = linhas.get(_sa(apos)) if apos else None
    if r_apos and r_apos != r and r_apos <= fim_tab:
        lst = [x for x in rows if x != r]
        lst.insert(lst.index(r_apos) + 1, r)
    ult = _col(ult_col)
    v = _ler(api, pid, f"'{aba}'!A1:{ult}{fim_tab}")
    f = _ler_f(api, pid, f"'{aba}'!A1:{ult}{fim_tab}") or []
    cel = lambda grade, rr, c: str((list(grade[rr - 1]) + [""] * ult_col)[c - 1]) if rr - 1 < len(grade) else ""
    membro_antes = {rr: _sa(cel(v, rr, 2)) for rr in rows}
    gr, movidos = [], 0
    for alvo, origem in zip(rows, lst):
        if alvo == origem:
            continue
        movidos += 1
        for c in range(1, ult_col + 1):
            if cel(f, alvo, c).startswith("="):
                continue                                  # fórmula da linha fica onde está
            gr.append((f"'{aba}'!{_col(c)}{alvo}", (list(f[origem - 1]) + [""] * ult_col)[c - 1] if origem - 1 < len(f) else ""))
    r_novo = rows[lst.index(r)]
    if cargo:
        gr.append((f"'{aba}'!{_col(C['cargo'])}{r_novo}", cargo))
    if grau and C.get("grau"):
        gr.append((f"'{aba}'!{_col(C['grau'])}{r_novo}", grau))
    _gravar(api, pid, gr)
    membro_depois = {alvo: membro_antes[origem] for alvo, origem in zip(rows, lst)}
    log.append(f"cargo: {cargo or 'mantido'}" + (f", grau {grau}" if grau else "") + (f"; agora fica depois de {apos}" if movidos else ""))
    # 2) PGTO: situação (C) e links (J, K) acompanham cada irmão
    pg = _aba_real(api, pid, "PGTO")
    if pg and movidos:
        vv, ff = _ler(api, pid, f"'{pg}'!A1:K400"), _ler_f(api, pid, f"'{pg}'!A1:B400") or []
        ref = {}
        for p_ in range(2, len(vv) + 1):
            m = re.search(r"Dados Integrantes'?!\$?[A-Z]+\$?(\d+)", str((list(ff[p_ - 1]) + ["", ""])[1]) if p_ - 1 < len(ff) else "")
            if m:
                ref[p_] = int(m.group(1))
        guardado = {}
        for p_, k in ref.items():
            linha = list(vv[p_ - 1]) + [""] * 11
            if k in membro_antes:
                guardado[membro_antes[k]] = (linha[2], linha[9], linha[10])
        gp = []
        for p_, k in ref.items():
            novo_m = membro_depois.get(k)
            if novo_m is not None and novo_m != membro_antes.get(k) and novo_m in guardado:
                c_, j_, k_ = guardado[novo_m]
                gp += [(f"'{pg}'!C{p_}", c_), (f"'{pg}'!J{p_}", j_), (f"'{pg}'!K{p_}", k_)]
        _gravar(api, pid, gp)
        if gp:
            log.append("PGTO acompanhou a nova ordem")
    # 3) Presença e Insights: o conteúdo das colunas é reescrito na ordem nova (nome, cargo e marcações).
    #    Não "arrasta" a coluna: o Google não deixa mover coluna que atravessa células mescladas do cabeçalho.
    curto = re.sub(r"\s*\([^)]*\)\s*", " ", cargo).strip()
    for ae in _abas_eventos(api, pid):
        vv = _ler(api, pid, f"'{ae}'!A1:CZ60")
        cols, ult = _cols_integrantes(vv)
        c_ap, c_apos = cols.get(_sa(ap)), cols.get(_sa(apos)) if apos else None
        if not c_ap:
            continue
        fim = _linha_adesao(vv)
        ate = (fim - 1) if fim else 43
        if c_apos and c_apos not in (c_ap, c_ap - 1):
            ordem = [c for c in range(4, ult + 1) if c != c_ap]
            ordem.insert(ordem.index(c_apos) + 1, c_ap)
            alvo_ini, alvo_fim = min(c_ap, c_apos + 1), max(c_ap, c_apos)
            ff = _ler_f(api, pid, f"'{ae}'!A1:{_col(ult)}{ate}") or []
            cel = lambda r, c: (list(ff[r - 1]) + [""] * (ult + 1))[c - 1] if r - 1 < len(ff) else ""
            gr = []
            for t_, s_ in zip(range(4, ult + 1), ordem):
                if t_ == s_ or not (alvo_ini <= t_ <= alvo_fim):
                    continue
                for r in range(6, ate + 1):
                    if str(cel(r, t_)).startswith("="):
                        continue                       # fórmula fica onde está
                    origem = cel(r, s_)
                    gr.append((f"'{ae}'!{_col(t_)}{r}", "" if str(origem).startswith("=") else origem))
            _gravar(api, pid, gr)
            c_ap = ordem.index(c_ap) + 4
            log.append(f"coluna reorganizada na aba {ae}")
        if curto and c_ap:
            _gravar(api, pid, [(f"'{ae}'!{_col(c_ap)}7", curto)])
    # 4) assinatura da mensagem de pagamentos
    if d.get("assina"):
        pa = _aba_real(api, pid, "Pagamentos")
        if pa:
            vp = _ler(api, pid, f"'{pa}'!C1:C80")
            for i_, linha in enumerate(vp, start=1):
                txt = str((list(linha) + [""])[0])
                m = re.match(r"^\*\s*(.+?)\s+-\s+(.+?)\s+(Div\b.*?)\*\s*$", txt)
                if m:
                    _gravar(api, pid, [(f"'{pa}'!C{i_}", f"*{ap} - {cargo or m.group(2)} {m.group(3)}*")])
                    log.append(f"assinatura dos pagamentos: {ap}")
                    break
    return f"{ap}: " + "; ".join(log)

def aplicar_integrante(api, pid, d):
    ap, acao = str(d.get("apelido", "")).strip(), d.get("acao")
    if not ap or acao not in ("novo", "afastar", "reativar", "excluir", "cargo"):
        return "pedido incompleto"
    aba, C, linhas, fim_tab, ult_col = _di_mapa(api, pid)
    sid = _ids(api, pid)[aba]
    r = linhas.get(_sa(ap))
    log = []
    if acao == "cargo":
        return _mudar_cargo(api, pid, d, aba, C, linhas, fim_tab, ult_col)
    if acao == "novo":
        if r:
            return f"{ap} já está na Dados Integrantes (nada feito)"
        dd = d.get("dados", {}) or {}
        r_ap = linhas.get(_sa(d.get("apos", ""))) if d.get("apos") else None
        r = _inserir_linha_di(api, pid, aba, sid, r_ap or fim_tab, ult_col)
        vals = {"num": dd.get("num"), "ap": ap, "cargo": dd.get("cargo"), "grau": dd.get("grau"), "tel": dd.get("tel"),
                "nome": dd.get("nome"), "colete": dd.get("colete"), "tam": dd.get("tam"), "email": dd.get("email"), "sit": "Ativo"}
        gr = [(f"'{aba}'!{_col(C[k])}{r}", v) for k, v in vals.items() if C.get(k) and v not in (None, "")]   # CPF não é gravado
        if C.get("nasc") and dd.get("nasc"):
            try:
                gr.append((f"'{aba}'!{_col(C['nasc'])}{r}", datetime.datetime.strptime(dd["nasc"], "%Y-%m-%d").strftime("%d/%m/%Y")))
            except ValueError:
                pass
        _gravar(api, pid, gr)
        _colunas_nas_abas(api, pid, ap, dd.get("cargo", ""), d.get("apos", ""), log)
        _linha_pgto(api, pid, ap, r, log)
        return f"{ap}: novo integrante na linha {r} da Dados Integrantes; " + "; ".join(log)
    if r is None:
        return f"{ap} não está na Dados Integrantes (nada feito)"
    if acao == "afastar":
        gr = [(f"'{aba}'!{_col(C['sit'])}{r}", "Afastado")]
        if C.get("desde") and d.get("desde"):
            gr.append((f"'{aba}'!{_col(C['desde'])}{r}", datetime.datetime.strptime(d["desde"], "%Y-%m-%d").strftime("%d/%m/%Y")))
        _gravar(api, pid, gr)
        return f"{ap}: afastado desde {d.get('desde', '')}"
    if acao == "excluir":
        pg = _aba_real(api, pid, "PGTO")
        if pg:
            vv, ff = _ler(api, pid, f"'{pg}'!A1:B400"), _ler_f(api, pid, f"'{pg}'!A1:B400")
            for rp in range(2, len(vv) + 1):
                txt = str((list(vv[rp - 1]) + ["", ""])[1]); frm = str((list(ff[rp - 1]) + ["", ""])[1]).replace("$", "") if rp <= len(ff) else ""
                if _sa(txt).endswith(_sa(ap)) or ("dados integrantes" in frm.lower() and re.search(rf"!B{r}(?!\d)", frm)):
                    _estrutura(api, pid, [{"deleteDimension": {"range": {"sheetId": _ids(api, pid)[pg], "dimension": "ROWS", "startIndex": rp - 1, "endIndex": rp}}}])
                    log.append("linha apagada da PGTO"); break
        for ae in _abas_eventos(api, pid):
            v = _ler(api, pid, f"'{ae}'!A1:CZ60")
            cols, ult = _cols_integrantes(v)
            alvo = cols.get(_sa(ap))
            if not alvo:
                continue
            _estrutura(api, pid, [{"deleteDimension": {"range": {"sheetId": _ids(api, pid)[ae], "dimension": "COLUMNS", "startIndex": alvo - 1, "endIndex": alvo}}}])
            log.append(f"coluna apagada da aba {ae}")
        aba, C, linhas, fim_tab, ult_col = _di_mapa(api, pid)
        r = linhas.get(_sa(ap))
        if r:
            _estrutura(api, pid, [{"deleteDimension": {"range": {"sheetId": sid, "dimension": "ROWS", "startIndex": r - 1, "endIndex": r}}}])
        return f"{ap}: excluído de toda a planilha; " + "; ".join(log)
    if acao == "reativar":
        apos = d.get("apos", "")
        r_ap = linhas.get(_sa(apos)) if apos else None
        if r_ap and r != r_ap + 1:      # leva a linha para logo depois de 'apos', dentro do bloco de integrantes
            f_old = (_ler_f(api, pid, f"'{aba}'!A{r}:{_col(ult_col)}{r}") or [[]])[0]
            nova = _inserir_linha_di(api, pid, aba, sid, r_ap, ult_col)
            r_old = r + 1 if r > r_ap else r
            gr = [(f"'{aba}'!{_col(c)}{nova}", (list(f_old) + [""] * ult_col)[c - 1])
                  for c in range(1, ult_col + 1) if not str((list(f_old) + [""] * ult_col)[c - 1]).startswith("=")]
            _gravar(api, pid, gr)
            _estrutura(api, pid, [{"deleteDimension": {"range": {"sheetId": sid, "dimension": "ROWS", "startIndex": r_old - 1, "endIndex": r_old}}}])
            aba, C, linhas, fim_tab, ult_col = _di_mapa(api, pid)
            r = linhas.get(_sa(ap))
        gr = [(f"'{aba}'!{_col(C['sit'])}{r}", "Ativo")]
        if C.get("desde"):
            gr.append((f"'{aba}'!{_col(C['desde'])}{r}", ""))
        _gravar(api, pid, gr)
        cargo = (_ler(api, pid, f"'{aba}'!C{r}:C{r}") or [[""]])[0]
        _colunas_nas_abas(api, pid, ap, (list(cargo) + [""])[0], apos, log)
        _linha_pgto(api, pid, ap, r, log)
        return f"{ap}: reativado; " + "; ".join(log)

def renumerar_google(api, pid, mudancas):
    """Numero Lista (coluna D da Dados Integrantes): ativos numerados, afastados/excluídos em branco."""
    aba = _aba_real(api, pid, "Dados Integrantes")
    _gravar(api, pid, [(f"'{aba}'!D{linha}", num if num else "") for linha, num, ap in mudancas])
    print("Numero Lista acertado: " + ", ".join(f"{ap} = {num if num else '(afastado/excluído)'}" for _, num, ap in mudancas))

def _tipo_aviso(d):
    """Emoji e imagem da notificação conforme o que foi lançado."""
    tp, nome = d.get("tipo"), _sa(str(d.get("evento", "")))
    if tp in ("pagamento_manual", "boletos_sistema"):
        return "💵", "notif-pag.png"
    if tp == "integrante":
        return "🪪", "notif-cadastro.png"
    if tp == "lancamento_presenca" and _sa(str(d.get("aba", ""))).startswith("insight"):
        return "💬", "notif-insight.png"
    if "pub" in nome.split() or nome.startswith("pub"):
        return "🍻", "notif-pub.png"
    if "social" in nome:
        return "🤝", "notif-social.png"
    return "🎉", "notif-evento.png"

MESES_PT = ["janeiro", "fevereiro", "março", "abril", "maio", "junho", "julho", "agosto", "setembro", "outubro", "novembro", "dezembro"]

def _mes_k(v):
    """Mês dos boletos guardados na coluna K (cabeçalho "Boleto 11/2026"), ou None."""
    import re as _re
    m = _re.search(r"(\d{1,2})/(\d{4})", str((list(v[0]) + [""] * 11)[10] if v else ""))
    return (int(m.group(2)), int(m.group(1))) if m else None

BOLETOS_DISPONIVEIS = []   # meses cujos boletos acabaram de entrar na coluna J (nesta rodada)
AVISO_BOLETO_MANUAL = []   # botão "📲 Avisar quem não pagou" (tela Pagamentos): aviso na hora
ENQ_NOVA = []              # enquete nova do comando: avisa os irmãos depois de publicar o painel (a enquete já aparece no app)
VERIFICAR_AGORA = []       # botão "🔄 Verificar agora" (tela Pagamentos): verifica os boletos nesta rodada

def aplicar_boletos(api, pid, d):
    """Boletos lidos no Sistema Insanos (botão no navegador): grava o link de cada irmão na PGTO.
    Mês atual (ou anterior): coluna J. Mês seguinte (boleto pago antes): coluna K, sem mexer no mês atual."""
    aba = _aba_real(api, pid, "PGTO")
    if not aba:
        return "aba PGTO não encontrada"
    v = _ler(api, pid, f"'{aba}'!A1:K400")
    try:
        ano, mes = int(str(d.get("mes"))[:4]), int(str(d.get("mes"))[5:7])
    except Exception:
        hoje = datetime.date.today(); ano, mes = hoje.year, hoje.month
    hoje = datetime.date.today()
    futuro = (ano, mes) > (hoje.year, hoje.month)
    col, letra = (10, "K") if futuro else (9, "J")
    gr, feitos, iguais, sem_linha = [], [], 0, []
    if futuro:
        if _mes_k(v) != (ano, mes):        # outro mês estava na K: começa do zero
            gr += [(f"'{aba}'!K{i}", "") for i in range(2, len(v) + 1) if str((list(v[i - 1]) + [""] * 11)[10]).strip()]
        gr.append((f"'{aba}'!K1", f"Boleto {mes:02d}/{ano}"))
    for it in (d.get("itens") or []):
        ap, link = str(it.get("ap") or "").strip(), str(it.get("link") or "").strip()
        if not ap or not link.startswith("http"):
            continue
        for i, row in enumerate(v[1:], start=2):
            row = list(row) + [""] * 11
            if _sa(str(row[1])).endswith(_sa(ap)):
                if str(row[col]).strip() == link and not futuro:
                    iguais += 1
                else:
                    gr.append((f"'{aba}'!{letra}{i}", link)); feitos.append(ap)
                break
        else:
            sem_linha.append(ap)
    _gravar(api, pid, gr)
    if (feitos or iguais) and not futuro:                 # (o aviso sai uma vez só por mês, mesmo clicando no favorito de novo)
        BOLETOS_DISPONIVEIS.append(f"{ano}-{mes:02d}")       # avisa quem não pagou: "o boleto já está disponível"
    res = f"boletos de {d.get('mes', '')}: {len(feitos)} link(s) gravado(s)"
    if futuro:
        res += f" na coluna K (mês seguinte: entram na J no dia 01/{mes:02d})"
    if iguais:
        res += f", {iguais} já estavam certos"
    if sem_linha:
        res += f" · sem linha na PGTO: {', '.join(sem_linha)}"
    return res

def promover_antecipados(api, pid):
    """Virada do mês: os boletos guardados na K (mês seguinte) passam para a J quando o mês chega."""
    aba = _aba_real(api, pid, "PGTO")
    if not aba:
        return 0
    v = _ler(api, pid, f"'{aba}'!A1:K400")
    mk, hoje = _mes_k(v), datetime.date.today()
    if not mk or mk > (hoje.year, hoje.month):
        return 0
    gr, n = [], 0
    for i, row in enumerate(v[1:], start=2):
        row = list(row) + [""] * 11
        link = str(row[10]).strip()
        if link.startswith("http"):
            gr += [(f"'{aba}'!J{i}", link), (f"'{aba}'!K{i}", "")]; n += 1
    gr.append((f"'{aba}'!K1", "Boleto do próximo mês"))
    _gravar(api, pid, gr)
    if n:
        BOLETOS_DISPONIVEIS.append(f"{mk[0]}-{mk[1]:02d}")
    print(f"Virada do mês: {n} boleto(s) do mês {mk[1]:02d}/{mk[0]} passaram da coluna K para a J.")
    return n

def verificar_antecipados(api, pid, g):
    """Boletos do mês seguinte (coluna K): avisa quem pagou antes do dia 1º."""
    aba = _aba_real(api, pid, "PGTO")
    if not aba:
        return False
    v = _ler(api, pid, f"'{aba}'!A1:K400")
    mk, hoje = _mes_k(v), datetime.date.today()
    if not mk or mk <= (hoje.year, hoje.month):
        return False
    arq = os.path.join(PASTA, "pagamentos_antecipados.json")
    try:
        with open(arq, encoding="utf-8") as f:
            ant = json.load(f)
    except Exception:
        ant = {}
    chave_mes = f"{mk[0]}-{mk[1]:02d}"
    if ant.get("mes") != chave_mes:
        ant = {"mes": chave_mes, "pagos": {}}
    inativos = g.inativos_planilha()
    novos = []
    print(f"Verificando pagamentos antecipados ({MESES_PT[mk[1] - 1]}/{mk[0]})...")
    for row in v[1:]:
        row = list(row) + [""] * 11
        link, rotulo = str(row[10]).strip(), str(row[1]).strip()
        if not link.startswith("http") or rotulo in ant["pagos"] or any(rotulo.lower().endswith(ap) for ap in inativos):
            continue
        try:
            selo, vb = g._ler_boleto(link)
        except Exception as e:
            print(f"  {rotulo}: não abriu ({e})"); continue
        if selo and "pag" in selo.lower() and (not vb or (vb.year, vb.month) == mk):
            ant["pagos"][rotulo] = datetime.datetime.now().strftime("%d/%m às %H:%M")
            novos.append(rotulo.split(" - ")[-1].strip())
            AGRADECER.append((rotulo, MESES_PT[mk[1] - 1], True))
            print(f"  {rotulo}: pago antecipado")
    with open(arq, "w", encoding="utf-8") as f:
        json.dump(ant, f, ensure_ascii=False, indent=1)
    if novos:
        nome_mes = MESES_PT[mk[1] - 1]
        AVISOS_PAG.append(("💵 Pagamento antecipado" if len(novos) == 1 else f"💵 {len(novos)} pagamentos antecipados",
                           f"{', '.join(novos)} {'pagou' if len(novos) == 1 else 'pagaram'} a mensalidade de {nome_mes} antes do dia 1º. "
                           f"Já são {len(ant['pagos'])} pago(s) antecipado(s).", None, False))
    return bool(novos)

def _json_ler(nome, padrao):
    try:
        with open(os.path.join(PASTA, nome), encoding="utf-8") as f:
            return json.load(f)
    except Exception:
        return padrao

def _json_gravar(nome, dados):
    with open(os.path.join(PASTA, nome), "w", encoding="utf-8") as f:
        json.dump(dados, f, ensure_ascii=False, indent=1)

def aplicar_adm_extra(d):
    """Irmãos liberados como ADM pela tela Acessos (além dos fixos do divisao.json): adm_extra.json."""
    ap, acao = str(d.get("apelido", "")).strip(), d.get("acao")
    lista = [x for x in _json_ler("adm_extra.json", []) if isinstance(x, str)]
    if not ap:
        return "pedido incompleto"
    if acao == "liberar" and ap.lower() not in [x.lower() for x in lista]:
        lista.append(ap)
    elif acao == "retirar":
        lista = [x for x in lista if x.lower() != ap.lower()]
    _json_gravar("adm_extra.json", lista)
    return f"acesso de ADM {'liberado para' if acao == 'liberar' else 'retirado de'} {ap}"

def aplicar_visitante(d):
    """Acesso de visitante: uma entrada só, com validade (visitantes.json)."""
    lista = [v for v in _json_ler("visitantes.json", []) if isinstance(v, dict)]
    email = str(d.get("email", "")).strip().lower()
    if d.get("acao") == "cancelar":
        lista = [v for v in lista if v.get("email") != email]
        _json_gravar("visitantes.json", lista)
        return f"visitante cancelado: {email}"
    if "@" not in email:
        return "e-mail inválido"
    horas = max(1, min(int(d.get("horas") or 24), 24 * 30))
    agora = datetime.datetime.now()
    lista = [v for v in lista if v.get("email") != email] + [{
        "email": email, "nome": str(d.get("nome", "")).strip()[:60], "criado": agora.strftime("%Y-%m-%d %H:%M"),
        "expira": (agora + datetime.timedelta(hours=horas)).strftime("%Y-%m-%d %H:%M"), "usado": ""}]
    _json_gravar("visitantes.json", lista)
    return f"visitante liberado: {email} (uma entrada, até {lista[-1]['expira'][8:10]}/{lista[-1]['expira'][5:7]} às {lista[-1]['expira'][11:]})"

RESPOSTAS = []          # respostas "Vou / Não vou" aplicadas nesta rodada (para avisar o comando)

def aplicar_respostas(api, pid, g):
    """Respostas "✅ Vou / ❌ Não vou" dos irmãos (registradas na aba Acessos com a assinatura):
    confere a assinatura e marca SIM/NÃO na aba Presença. Até o dia do evento, vale a resposta do irmão (mesmo onde o comando
    já marcou, exceto o staff); nos insights, só onde o comando não marcou (ou onde a marca veio do próprio irmão)."""
    regs = (g.buscar_acessos() or [])[1:]
    if not regs:
        return []
    est = _json_ler("rsvp.json", {"vistos": [], "celulas": {}})
    vistos, celulas = set(est.get("vistos", [])), est.get("celulas", {})
    hoje = datetime.date.today().isoformat()
    novas = {}
    for r in regs:
        r = list(r) + [""] * 5
        pn = str(r[3])
        if not pn.startswith("✋|"):
            continue
        partes = pn.split("|", 4)
        if len(partes) < 5:
            continue
        _, resp, d, tk, nome = partes
        email = str(r[1]).strip().lower()
        chave = f"{r[0]}|{email}|{d}|{nome}|{resp}"
        lim = (datetime.date.today() - datetime.timedelta(days=10)).isoformat() if resp == "resp" else hoje
        if chave in vistos or d < lim or resp not in ("vou", "nao", "resp"):
            continue
        vistos.add(chave)
        if tk != g.token_rsvp(email, d, nome):
            print(f"Resposta ignorada (assinatura inválida): {email} · {nome}"); continue
        k = (email, d, nome)
        if k not in novas or str(r[0]) >= novas[k][0]:
            novas[k] = (str(r[0]), resp)
    feitas = []
    if novas:
        aba_di, C, linhas, fim_tab, _ = _di_mapa(api, pid)
        vdi = _ler(api, pid, f"'{aba_di}'!A1:AN400")
        ap_de = {}
        if C.get("email"):
            for row in vdi[1:]:
                row = list(row) + [""] * 40
                if "@" in str(row[C["email"] - 1]):
                    ap_de[str(row[C["email"] - 1]).strip().lower()] = str(row[1]).strip()
        gr = []
        # Vou / Não vou -> aba Presenca; Já respondi (insight) -> aba Insights
        for aba_nome, so_resp in (("Presenca", False), ("Insights", True)):
            grupo = {k: x for k, x in novas.items() if (x[1] == "resp") == so_resp}
            if not grupo:
                continue
            aba = _aba_real(api, pid, aba_nome)
            if not aba:
                continue
            v = _ler(api, pid, f"'{aba}'!A1:BZ400")
            lin = lambda r: list(v[r - 1]) + [""] * 80 if len(v) >= r else [""] * 80
            cols, c = {}, 4
            while True:
                nm = str(lin(6)[c - 1]).strip()
                if not nm or nm.lower().startswith("presen") or nm.lower().startswith("insig") or nm.lower().startswith("%"):
                    break
                cols[_sa(nm)] = c; c += 1
            def data_da_linha(r):
                x = lin(r)[1]
                if isinstance(x, (int, float)) and x > 20000:
                    return (datetime.datetime(1899, 12, 30) + datetime.timedelta(days=float(x))).date().isoformat()
                try:
                    return datetime.datetime.strptime(str(x)[:10], "%d/%m/%Y").date().isoformat()
                except ValueError:
                    return ""
            for (email, d, nome), (_, resp) in grupo.items():
                ap = ap_de.get(email)
                cc = cols.get(_sa(ap)) if ap else None
                linha = next((r for r in range(8, len(v) + 1) if data_da_linha(r) == d and _sa(lin(r)[2]) == _sa(nome)), None)
                if not (ap and cc and linha):
                    continue
                atual = str(lin(linha)[cc - 1]).strip().upper()
                cel = f"{aba_nome}|{d}|{nome}|{ap}"
                novo = "SIM" if resp in ("vou", "resp") else "NÃO"
                livre = ("", "-") + (("NÃO", "NAO") if resp == "resp" else ())     # no insight, o NÃO padrão vira SIM quando o irmão marca
                futuro = aba_nome == "Presenca" and d >= hoje   # até o dia do evento, vale a resposta do próprio irmão
                if futuro and atual.startswith("STAFF") and novo == "SIM":
                    continue                                   # o comando marcou staff: "Vou" não tira o staff
                if futuro or atual in livre or celulas.get(cel) == atual or celulas.get(f"{d}|{nome}|{ap}") == atual:   # vale a marcação do comando
                    if atual != novo:
                        gr.append((f"'{aba}'!{_col(cc)}{linha}", novo))
                    celulas[cel] = novo
                    feitas.append((ap, nome, resp))
        _gravar(api, pid, gr)
    _json_gravar("rsvp.json", {"vistos": sorted(vistos)[-800:], "celulas": celulas})
    if feitas:
        print("Respostas dos irmãos aplicadas: " + ", ".join(f"{a} ({'vai' if r == 'vou' else 'respondeu' if r == 'resp' else 'não vai'})" for a, _, r in feitas))
    RESPOSTAS.extend(feitas)
    return feitas

def _num(x):
    """Número vindo da planilha em português: "-23,58" ou "-23.58" -> -23.58 (vazio -> None)."""
    t = str(x if x is not None else "").strip()
    if not t:
        return None
    if "," in t:
        t = t.replace(".", "").replace(",", ".")
    try:
        return float(t)
    except ValueError:
        return None

NOVOS_SOS = []          # SOS enviados nesta rodada (a rotina rápida repete o alerta se ninguém confirmar)

SOS_TIPOS = {"pane": ("🛵", "Pane/acidente de moto"), "saude": ("🚑", "Saúde"), "seguranca": ("⚠️", "Segurança"), "outro": ("🆘", "Outro")}

def processar_sos(api, pid, g, repeticao=None):
    """SOS: alerta para todos os irmãos (com a localização) e aviso de quem está indo (para quem pediu e para o comando).
    repeticao={"id": ..., "n": 2 ou 3}: manda de novo o alerta daquele SOS (ninguém confirmou que está indo)."""
    if "SOS" not in _ids(api, pid):
        return 0
    v = _ler(api, pid, "'SOS'!A1:K600")
    linhas = [(i, list(r) + [""] * 11) for i, r in enumerate(v[1:], start=2)]
    def ficha(r):   # a ficha de emergência (coluna K): vai na notificação e na tela do SOS
        try:
            f = json.loads(str(r[10] or "") or "{}")
            return f if isinstance(f, dict) else {}
        except Exception:
            return {}
    def ficha_txt(f):
        p = []
        if f.get("sangue"): p.append(f"🩸 {f['sangue']}")
        if f.get("alerg"): p.append(f"⚠️ Alergia: {f['alerg']}")
        if f.get("rem"): p.append(f"💊 {f['rem']}")
        c = (f.get("c") or [{}])[0]
        if c.get("t"): p.append(f"📞 {c.get('n') or 'Contato'}{' (' + c['p'] + ')' if c.get('p') else ''}: {c['t']}")
        return " · ".join(p)
    sos = {r[2]: r for _, r in linhas if str(r[1]) == "SOS"}
    if repeticao:      # repetição do alerta: não muda nada na aba, só manda de novo
        r = sos.get(repeticao["id"])
        if r and not any(str(x[1]) == "Estou indo" and x[2] == repeticao["id"] for _, x in linhas):
            emo, nome_tipo = SOS_TIPOS.get(str(r[5]), SOS_TIPOS["outro"])
            lat, lng = _num(r[6]), _num(r[7])
            mh = re.search(r"(\d{1,2}):(\d{2})", str(r[0]))
            hora = f"{int(mh.group(1)):02d}:{mh.group(2)}" if mh else ""
            dados = {"id": repeticao["id"], "lat": lat, "lng": lng, "n": str(r[4]), "t": nome_tipo, "e": emo, "q": hora, **({"f": ficha(r)} if ficha(r) else {})}
            url = "adm.html#sos=" + base64.urlsafe_b64encode(json.dumps(dados, ensure_ascii=False).encode()).decode().rstrip("=")
            g.notificar(f"🆘 SOS ({repeticao['n']}º aviso) · {r[4]} ainda precisa de ajuda", f"{emo} {nome_tipo} · desde as {hora} · ninguém confirmou ainda" + (("\n" + ficha_txt(ficha(r))) if ficha(r) else ""),
                        f"sos-{repeticao['id']}", icone="notif-aviso.png", todos=True, excluir=str(r[3]).strip().lower(), extra={"sos": dados, "url": url})
            print(f"SOS {repeticao['id']}: {repeticao['n']}º aviso enviado (ninguém confirmou ainda).")
            return 1
        return 0
    gr, n = [], 0
    for i, r in linhas:
        if str(r[9]).strip().lower() != "novo":
            continue
        reg, sid, email, ap, cat = str(r[1]), str(r[2]), str(r[3]).strip().lower(), str(r[4]), str(r[5])
        try:
            if reg == "SOS":
                emo, nome_tipo = SOS_TIPOS.get(cat, SOS_TIPOS["outro"])
                lat, lng = _num(r[6]), _num(r[7])      # a planilha em português guarda com vírgula: o mapa precisa de ponto
                mh = re.search(r"(\d{1,2}):(\d{2})", str(r[0]))           # a hora do pedido (registrada na aba SOS)
                hora = f"{int(mh.group(1)):02d}:{mh.group(2)}" if mh else datetime.datetime.now().strftime("%H:%M")
                fi = ficha(r)
                dados = {"id": sid, "lat": lat, "lng": lng, "n": ap, "t": nome_tipo, "e": emo, "q": hora, **({"f": fi} if fi else {})}
                url = "adm.html#sos=" + base64.urlsafe_b64encode(json.dumps(dados, ensure_ascii=False).encode()).decode().rstrip("=")
                corpo = f"{emo} {nome_tipo} · {dados['q']}" + ("" if lat is not None else " · sem localização") + (("\n" + ficha_txt(fi)) if fi else "")
                g.notificar(f"🆘 SOS · {ap} precisa de ajuda", corpo, f"sos-{sid}", icone="notif-aviso.png",
                            todos=True, excluir=email, extra={"sos": dados, "url": url})
                print(f"SOS de {ap} enviado para todos os irmãos.")
                NOVOS_SOS.append(sid)
            elif reg == "Estou indo":
                orig = sos.get(sid)
                km = _num(r[8])
                dist = f" · a {str(round(km, 1)).replace('.', ',')} km" if km is not None else ""
                if orig:
                    g.notificar(f"🏍️ {ap} está indo te ajudar", f"Aguenta firme, irmão{dist}. 🤝", f"sos-indo-{sid}", icone="notif-aviso.png",
                                so_para=str(orig[3]).strip().lower())
                    g.notificar(f"🏍️ {ap} está indo ao SOS do {orig[4]}", f"Apoio a caminho{dist}.", f"sos-indo-{sid}", icone="notif-aviso.png",
                                todos=True, excluir=str(orig[3]).strip().lower())     # todos os irmãos sabem que já tem apoio a caminho
                print(f"{ap} está indo ao SOS {sid}.")
            gr.append((f"'SOS'!J{i}", "avisado")); n += 1
        except Exception as e:
            print(f"Aviso: SOS não processado ({type(e).__name__}: {e}).")
    _gravar(api, pid, gr)
    return n

MSG_LIMITE_IRMAO = 5       # mensagens por dia de quem não é do comando
MSG_RESPOSTAS = [("r_ok", "✅ Recebido"), ("r_zap", "📲 Te respondo no WhatsApp")]

def _msg_token(id_msg, email):
    import hmac, hashlib
    return hmac.new(os.environ.get("OTP_SEGREDO", "").encode(), f"{id_msg}|{email}".encode(), hashlib.sha256).hexdigest()[:16]

def processar_mensagens(api, pid, g):
    """Mensagens pelo app (aba Mensagens, preenchida pelo projeto do código por e-mail):
    comando -> qualquer irmão; irmão -> só o comando (até 5 por dia). Respostas rápidas voltam para quem mandou."""
    if "Mensagens" not in _ids(api, pid):
        return 0
    v = _ler(api, pid, "'Mensagens'!A1:L3000")
    linhas = [(i, list(r) + [""] * 12) for i, r in enumerate(v[1:], start=2)]
    from zoneinfo import ZoneInfo
    agora = datetime.datetime.now(ZoneInfo("America/Sao_Paulo")).replace(tzinfo=None)
    def quando(x):
        if isinstance(x, (int, float)):
            return datetime.datetime(1899, 12, 30) + datetime.timedelta(days=float(x))
        for f in ("%d/%m/%Y %H:%M:%S", "%d/%m/%Y %H:%M", "%Y-%m-%dT%H:%M:%S", "%Y-%m-%d %H:%M:%S"):
            try:
                return datetime.datetime.strptime(str(x)[:19], f)
            except ValueError:
                pass
        return None
    imp = lambda r: str(r[5]).startswith("📌")            # mensagem importante (o comando marcou 📌 ao escrever)
    respondeu = {}
    for _, r in linhas:
        if str(r[1]) == "Resposta":
            respondeu.setdefault(str(r[2]), set()).add(str(r[3]).strip().lower())
    def pendentes(r):                                      # receberam, mas não leram nem responderam
        leram = {x.strip().lower() for x in str(r[10]).split(";") if x.strip()} | respondeu.get(str(r[2]), set())
        return [x.strip().lower() for x in str(r[9]).split(";") if x.strip() and x.strip().lower() not in leram]
    lembrar = []
    for i, r in linhas:
        if str(r[1]) == "Mensagem" and imp(r) and str(r[8]).startswith("enviada") and str(r[11]).strip() in ("", "1"):
            t0 = quando(r[0])
            if t0:
                passou = (agora - t0).total_seconds() / 60
                if (str(r[11]).strip() == "" and passou >= 30) or (str(r[11]).strip() == "1" and passou >= 120):
                    lembrar.append((i, r))
    if not lembrar and not any(str(r[8]).strip() == "novo" for _, r in linhas):
        return 0
    di = _ler(api, pid, "'Dados Integrantes'!A1:BZ300")
    cab = [str(x).strip().lower() for x in (di[0] if di else [])]
    c_mail = next((k for k, x in enumerate(cab) if "mail" in x), None)
    por_ap, por_email = {}, {}
    for r in di[1:]:
        r = list(r) + [""] * 80
        ap, em = str(r[1]).strip(), str(r[c_mail] if c_mail is not None else "").strip().lower()
        if ap:
            por_ap[_sa(ap)] = (em, ap, str(r[2]).strip())
            if "@" in em:
                por_email[em] = (ap, str(r[2]).strip())
    comando = {_sa(a) for a in getattr(g, "AUTORIZADOS", [])} | {_sa(a) for a in (_json_ler("adm_extra.json", []) or []) if isinstance(a, str)}
    hoje = datetime.date.today()
    def dia(x):
        try:
            return datetime.datetime.fromisoformat(str(x)[:19]).date() if not isinstance(x, (int, float)) else (datetime.date(1899, 12, 30) + datetime.timedelta(days=int(x)))
        except ValueError:
            try:
                return datetime.datetime.strptime(str(x)[:10], "%d/%m/%Y").date()
            except ValueError:
                return None
    enviadas_hoje = {}
    for _, r in linhas:
        if str(r[1]) == "Mensagem" and str(r[8]).startswith("enviada") and dia(r[0]) == hoje:
            enviadas_hoje[str(r[3]).strip().lower()] = enviadas_hoje.get(str(r[3]).strip().lower(), 0) + 1
    ou = os.environ.get("OTP_URL", "")
    gr, n = [], 0
    msgs = {str(r[2]): r for _, r in linhas if str(r[1]) == "Mensagem"}
    for i, r in linhas:
        if str(r[8]).strip() != "novo":
            continue
        try:
            de = str(r[3]).strip().lower()
            if de not in por_email:
                gr.append((f"'Mensagens'!I{i}", "recusada: remetente não encontrado na planilha")); continue
            ap_de, cargo_de = por_email[de]
            if str(r[1]) == "Parabens":            # parabéns de aniversário (Meu painel): vai para o aniversariante
                alvo = por_ap.get(_sa(str(r[6]).strip()))
                em = alvo[0] if alvo else ""
                if "@" not in em or em == de:
                    gr.append((f"'Mensagens'!I{i}", "recusada: aniversariante sem e-mail no painel")); continue
                g.notificar(f"🎂 {ap_de} te mandou os parabéns!", str(r[7]), f"parabens-{r[2]}", icone="notif-social.png", so_para=em)
                gr.append((f"'Mensagens'!I{i}", "enviada"))
                gr.append((f"'Mensagens'!J{i}", em))
                print(f"Parabéns de {ap_de} para {r[6]}.")
                n += 1
                continue
            if str(r[1]) == "Mensagem":
                eh_cmd = _sa(ap_de) in comando
                dest = [x.strip() for x in str(r[6]).split(";") if x.strip()]
                if not eh_cmd and any(_sa(x) not in comando for x in dest):
                    gr.append((f"'Mensagens'!I{i}", "recusada: irmão só pode mandar para o comando")); continue
                if not eh_cmd and enviadas_hoje.get(de, 0) >= MSG_LIMITE_IRMAO:
                    gr.append((f"'Mensagens'!I{i}", f"recusada: limite de {MSG_LIMITE_IRMAO} mensagens por dia")); continue
                emails = [por_ap[_sa(x)][0] for x in dest if _sa(x) in por_ap and "@" in por_ap[_sa(x)][0] and por_ap[_sa(x)][0] != de]
                importante = eh_cmd and imp(r)
                for em in emails:
                    g.notificar(("📌 IMPORTANTE · " if importante else "✉️ ") + f"Mensagem de {ap_de}" + (f" ({cargo_de})" if cargo_de else ""), str(r[7]),
                                f"msg-{r[2]}", icone="notif-social.png", so_para=em,
                                extra={"msg": {"id": str(r[2]), "tk": _msg_token(r[2], em), "ou": ou, "rs": MSG_RESPOSTAS, "imp": 1 if importante else 0}})
                gr.append((f"'Mensagens'!I{i}", f"enviada para {len(emails)} irmão(s)"))
                gr.append((f"'Mensagens'!J{i}", "; ".join(emails)))
                if not eh_cmd:
                    enviadas_hoje[de] = enviadas_hoje.get(de, 0) + 1
                print(f"Mensagem de {ap_de} para {len(emails)} irmão(s) ({str(r[5]) or 'individual'}).")
                n += 1
            else:                                   # resposta rápida: volta para quem mandou a mensagem
                orig = msgs.get(str(r[2]))
                if not orig:
                    gr.append((f"'Mensagens'!I{i}", "recusada: mensagem original não encontrada")); continue
                recebeu = [x.strip().lower() for x in str(orig[9]).split(";") if x.strip()]
                alvo = str(orig[3]).strip().lower()
                if de not in recebeu and de != alvo:
                    gr.append((f"'Mensagens'!I{i}", "recusada: não recebeu essa mensagem")); continue
                para = alvo if de != alvo else None
                if para:
                    trecho = str(orig[7])[:60] + ("…" if len(str(orig[7])) > 60 else "")
                    g.notificar(f"↩️ {ap_de} respondeu: {r[7]}", f"Sobre: {trecho}", f"msg-r-{r[2]}", icone="notif-social.png", so_para=para)
                gr.append((f"'Mensagens'!I{i}", "enviada"))
                n += 1
        except Exception as e:
            print(f"Aviso: mensagem não processada ({type(e).__name__}: {e}).")
    # importantes: avisa de novo (30 min e 2 h depois) quem ainda não leu nem respondeu
    for i, r in lembrar:
        try:
            ap_de, cargo_de = por_email.get(str(r[3]).strip().lower(), (str(r[4]), ""))
            pend = pendentes(r)
            for em in pend:
                g.notificar(f"📌 Lembrete · Mensagem de {ap_de}" + (f" ({cargo_de})" if cargo_de else ""), str(r[7]), f"msg-{r[2]}",
                            icone="notif-social.png", so_para=em,
                            extra={"msg": {"id": str(r[2]), "tk": _msg_token(r[2], em), "ou": ou, "rs": MSG_RESPOSTAS, "imp": 1}})
            gr.append((f"'Mensagens'!L{i}", "1" if str(r[11]).strip() == "" else "2"))
            print(f"Mensagem importante {r[2]}: lembrete para {len(pend)} irmão(s) que ainda não leram.")
        except Exception as e:
            print(f"Aviso: lembrete não enviado ({type(e).__name__}: {e}).")
    _gravar(api, pid, gr)
    return n

def acompanhar_sos(api, pid, g, dormir=None, relogio=None):
    """Durante ~3 minutos: avisa na hora quem está indo e repete o alerta (com 1 e 3 minutos) se ninguém confirmar."""
    import time
    dormir, relogio = dormir or time.sleep, relogio or time.time
    inicio, pendentes = relogio(), {sid: [60, 180] for sid in NOVOS_SOS}
    while pendentes and relogio() - inicio < 200:
        dormir(15)
        try:
            processar_sos(api, pid, g)                   # "Estou indo" que chegaram enquanto isso
            passou = relogio() - inicio
            for sid in list(pendentes):
                if pendentes[sid] and passou >= pendentes[sid][0]:
                    n = 4 - len(pendentes[sid])          # 2º aviso, depois 3º
                    if not processar_sos(api, pid, g, repeticao={"id": sid, "n": n}):
                        pendentes.pop(sid); continue  # alguém já confirmou: para de repetir
                    pendentes[sid].pop(0)
                if sid in pendentes and not pendentes[sid]:
                    pendentes.pop(sid)
        except Exception as e:
            print(f"Aviso: acompanhamento do SOS falhou ({type(e).__name__}: {e}).")

def escrever_otp(api, pid, g):
    """Código por e-mail: escreve na aba escondida "_otp" os e-mails dos visitantes ativos
    (os dos irmãos o projeto do código lê direto da aba Dados Integrantes)."""
    if not g.otp_ativo():
        return
    emails = sorted({it[0] for it in getattr(g, "PESSOAIS", []) if len(it) > 2 and it[2] == "v"})
    if "_otp" not in _ids(api, pid):
        _estrutura(api, pid, [{"addSheet": {"properties": {"title": "_otp", "hidden": True}}}])
    api.spreadsheets().values().clear(spreadsheetId=pid, range="'_otp'!A1:A200").execute()
    if emails:
        _gravar(api, pid, [(f"'_otp'!A{i}", e) for i, e in enumerate(emails, start=1)])

def enviar_aviso_insight(g, ins_novos):
    """Assim que um insight é lançado: aviso para quem ainda não consta como respondido (só o insight mais recente)."""
    L = getattr(g, "LEMBRETES", None) or {}
    env = _json_ler("lembretes_enviados.json", {})
    hoje = datetime.date.today().isoformat()
    n = 0
    for d, nome in ins_novos:
        for p in L.get("pessoas", []):
            if not any(x["d"] == d and x["nome"] == nome for x in p.get("ins", [])):
                continue
            k = f"{hoje}|lanc|ins|{d}|{nome}|{p['email']}"
            if k in env:
                continue
            try:
                g.notificar(f"💬 {nome} no ar", f"Fala, {p['n']}! Saiu o {nome}. Responda no grupo do WhatsApp e toque em Já respondi. 🤝",
                            "lembrete-ins", icone="notif-insight.png", so_para=p["email"], rsvp={"d": d, "e": nome, "m": "ins"})
                env[k] = 1; n += 1
            except Exception as e:
                print(f"Aviso: aviso do insight não enviado ({e}).")
    _json_gravar("lembretes_enviados.json", env)
    if n:
        print(f"Aviso do insight enviado: {n} irmão(s).")

def enviar_convites(g, eventos_novos):
    """Assim que um evento é lançado: convite com os botões ✅ Vou / ❌ Não vou para quem ainda não respondeu."""
    L = getattr(g, "LEMBRETES", None) or {}
    env = _json_ler("lembretes_enviados.json", {})
    hoje = datetime.date.today().isoformat()
    n = 0
    for d, nome in eventos_novos:
        for p in L.get("pessoas", []):
            if not any(ev["d"] == d and ev["nome"] == nome for ev in p.get("ev", [])):
                continue
            k = f"{hoje}|lanc|ev|{d}|{nome}|{p['email']}"
            if k in env:
                continue
            emoji, icone = _tipo_aviso({"tipo": "lancamento_presenca", "aba": "Presenca", "evento": nome})
            try:
                g.notificar(f"{emoji} {nome} · {d[8:10]}/{d[5:7]}", f"Fala, {p['n']}! Você vai na {nome}? Responde aqui embaixo. 🤝",
                            "convite", icone=icone, so_para=p["email"], rsvp={"d": d, "e": nome})
                env[k] = 1; n += 1
            except Exception as e:
                print(f"Aviso: convite não enviado ({e}).")
    _json_gravar("lembretes_enviados.json", env)
    if n:
        print(f"Convites com Vou / Não vou enviados: {n}.")

def enviar_boleto_disponivel(g):
    """Assim que os boletos do mês entram na planilha: aviso na hora para quem ainda não pagou, e mais um no dia seguinte (9h).
    Depois disso, só os lembretes de 3, 2, 1 dia antes e do dia do vencimento (enviar_lembretes)."""
    L = getattr(g, "LEMBRETES", None) or {}
    try:
        venc = datetime.date.fromisoformat(str(L.get("venc"))[:10])
    except Exception:
        return
    mes_k = f"{venc.year}-{venc.month:02d}"
    env = _json_ler("lembretes_enviados.json", {})
    agora = datetime.datetime.now(); hoje = agora.date()
    chave_dia = f"9999|disp|{mes_k}"                     # dia em que os boletos ficaram disponíveis (o 9999 evita a limpeza dos 3 dias)
    manual = bool(AVISO_BOLETO_MANUAL)
    if (mes_k in BOLETOS_DISPONIVEIS or manual) and chave_dia not in env:
        env[chave_dia] = hoje.isoformat()
    dia = env.get(chave_dia)
    if not dia:
        return
    try:
        dia = datetime.date.fromisoformat(dia)
    except ValueError:
        return
    if manual:
        vez = 1                                          # botão da tela Pagamentos: manda agora (no máximo um aviso de boleto por dia)
    elif hoje == dia:
        vez = 1
    elif hoje == dia + datetime.timedelta(days=1) and agora.hour >= 9:
        vez = 2
    else:
        _json_gravar("lembretes_enviados.json", env); return
    if hoje > venc:
        _json_gravar("lembretes_enviados.json", env); return
    mes, dv, n = MESES_PT[venc.month - 1], venc.strftime("%d/%m"), 0
    for p in L.get("pessoas", []):
        if p.get("pago"):
            continue
        k = f"9999|disp{vez}|{mes_k}|{p['email']}"
        if (k in env and not manual) or f"{hoje}|pag|{p['email']}" in env:     # já avisado (ou já recebeu um aviso de boleto hoje)
            continue
        ap = p["n"]
        corpo = (f"Fala, {ap}! O boleto da mensalidade de {mes} já está disponível (vence em {dv}). Está a um toque, aqui embaixo. 🤝" if vez == 1 else
                 f"{ap}, passando para lembrar que o boleto de {mes} está disponível, com vencimento em {dv}. Se já pagou, obrigado e pode desconsiderar. 💀")
        link = p.get("boleto") if str(p.get("boleto", "")).startswith("http") else None
        try:
            g.notificar("📅 Agenda da divisão", corpo, "lembrete-pag", icone="notif-pag.png", so_para=p["email"], link=link, link_rot="Abrir meu boleto")
            env[k] = 1; env[f"{hoje}|pag|{p['email']}"] = 1; n += 1
        except Exception as e:
            print(f"Aviso: aviso de boleto disponível não enviado ({e}).")
    _json_gravar("lembretes_enviados.json", env)
    AVISO_BOLETO_MANUAL.clear()
    if n:
        print(f"Boleto disponível ({vez}º aviso{', pelo botão' if manual else ''}): {n} irmão(s) avisado(s).")

def enviar_lembretes(g):
    """Lembretes individuais (só para o irmão): insights sem resposta (13h e 19h), eventos sem confirmação
    (9h, 13h e 19h, até o dia do evento) e a agenda da mensalidade (9h, de 3 dias antes até o vencimento)."""
    L = getattr(g, "LEMBRETES", None) or {}
    if not L.get("pessoas"):
        return
    agora = datetime.datetime.now()
    slot = {9: "m", 13: "t", 19: "n"}.get(agora.hour)
    if not slot:
        return
    hoje = agora.date()
    env = _json_ler("lembretes_enviados.json", {})
    corte = (hoje - datetime.timedelta(days=3)).isoformat()
    env = {k: v for k, v in env.items() if k[:10] >= corte or k.startswith("9999|")}
    br = lambda d: f"{d[8:10]}/{d[5:7]}"
    n_env = 0
    def manda(chave, *a, **k):
        nonlocal n_env
        if chave in env:
            return
        try:
            g.notificar(*a, **k); env[chave] = 1; n_env += 1
        except Exception as e:
            print(f"Aviso: lembrete não enviado ({e}).")
    venc = None
    try:
        venc = datetime.date.fromisoformat(str(L.get("venc"))[:10])
    except Exception:
        pass
    for p in L["pessoas"]:
        ap, email = p["n"], p["email"]
        # insights sem resposta: 13h e 19h, até constar que respondeu
        if slot in ("t", "n") and p.get("ins"):
            xs = p["ins"]; nomes = [x["nome"] for x in xs]
            if len(xs) > 1:
                tit = "💬 Insights te esperando"
                corpo = f"Fala, {ap}! Ainda faltam as suas respostas no {', '.join(nomes[:-1])} e no {nomes[-1]}. Quando puder, dá uma passada no grupo. 🤝"
            elif xs[0]["d"] == hoje.isoformat():
                tit = f"💬 {nomes[0]} no ar"
                corpo = f"Fala, {ap}! Ainda falta a sua resposta no {nomes[0]}. Quando puder, dá uma passada no grupo. 🤝"
            else:
                tit = "💬 Só falta você, irmão"
                corpo = f"{ap}, o {nomes[0]} já tem {xs[0]['resp']} respostas. A sua opinião também conta: responde lá quando der. 💀"
            manda(f"{hoje}|{slot}|ins|{email}", tit, corpo + " Já respondeu? Toque em Já respondi.", "lembrete-ins", icone="notif-insight.png", so_para=email,
                  rsvp={"d": xs[0]["d"], "e": xs[0]["nome"], "m": "ins"} if len(xs) == 1 else None)
        # eventos sem confirmação: 9h, 13h e 19h, até o dia do evento
        for ev in p.get("ev", []):
            emoji, icone = _tipo_aviso({"tipo": "lancamento_presenca", "aba": "Presenca", "evento": ev["nome"]})
            try:   # previsão do tempo (até 15 dias antes) e o endereço para o Como chegar
                end_, hr_ = g.lugar_evento(ev["d"], ev["nome"])
                clima_ = g.clima_evento(ev["d"], hr_, end_)
            except Exception:
                end_, hr_, clima_ = "", "", ""
            prev = f" Previsão: {clima_}." if clima_ else ""
            if ev["d"] == hoje.isoformat():
                tit, corpo = f"{emoji} {ev['nome']} é hoje", f"Fala, {ap}! Hoje tem {ev['nome']}{' às ' + hr_ if hr_ else ''}. Se ainda não confirmou, dá um toque no grupo. Contamos com você! 🤝{prev}"
                mapa = ("https://www.google.com/maps/search/?api=1&query=" + urllib.parse.quote(end_)) if end_ else None
                manda(f"{hoje}|{slot}|ev|{ev['d']}|{ev['nome']}|{email}", tit, corpo, "lembrete-ev", icone=icone, so_para=email,
                      **({"link": mapa, "link_rot": "🗺️ Como chegar"} if mapa else {"rsvp": {"d": ev["d"], "e": ev["nome"]}}))
            else:
                tit, corpo = f"{emoji} {ev['nome']} · {br(ev['d'])}", f"Fala, {ap}! A lista da {ev['nome']} está aberta. Você vai? Responde aqui embaixo. Contamos com você! 🤝{prev}"
                manda(f"{hoje}|{slot}|ev|{ev['d']}|{ev['nome']}|{email}", tit, corpo, "lembrete-ev", icone=icone, so_para=email, rsvp={"d": ev["d"], "e": ev["nome"]})
        # mensalidade: 9h, de 3 dias antes até o dia do vencimento (só quem ainda não aparece como pago)
        if slot == "m" and venc and not p.get("pago"):
            k = (venc - hoje).days
            if 0 <= k <= 3:
                mes, dv = MESES_PT[venc.month - 1], venc.strftime("%d/%m")
                corpo = {3: f"Fala, {ap}! Passando só para avisar que a mensalidade de {mes} vence no dia {dv}. Se já estiver em dia, obrigado e desconsidere. 🤝",
                         2: f"{ap}, a mensalidade de {mes} vence em {dv}. O seu boleto está a um toque, aqui embaixo, se precisar. 💀",
                         1: f"Fala, {ap}! Amanhã ({dv}) vence a mensalidade de {mes}. Se já pagou, valeu: às vezes o banco leva uns dias para confirmar, então é só desconsiderar. 🤝",
                         0: f"{ap}, hoje é o vencimento da mensalidade de {mes}. Se já está tudo certo, obrigado pelo compromisso com o Insanos! Qualquer coisa, é só chamar o ADM no privado. 🤝"}[k]
                link = p.get("boleto") if str(p.get("boleto", "")).startswith("http") else None
                manda(f"{hoje}|pag|{email}", "📅 Agenda da divisão", corpo, "lembrete-pag", icone="notif-pag.png", so_para=email,
                      link=link, link_rot="Abrir meu boleto")
    _json_gravar("lembretes_enviados.json", env)
    if n_env:
        print(f"Lembretes individuais enviados: {n_env}.")

def fila_recente(api, pid, n=15):
    """Resumo das últimas linhas da aba Fila, sem dados pessoais: o tipo do pedido, o evento e a situação."""
    v = _ler(api, pid, "'Fila'!A1:F600")
    nomes = {"pagamentos": "Pagamentos", "pagamento": "Pagamentos", "cadastro": "Cadastro", "boletos_sistema": "Boletos do Sistema Insanos",
             "boletos_aviso": "Aviso de boleto disponível", "mudar_cargo": "Mudança de cargo", "visitante": "Acesso de visitante",
             "adm_extra": "Acesso de ADM", "situacao": "Situação de integrante", "teste_pagamento": "Teste de aviso"}
    out = []
    for r in reversed(v[1:]):
        r = list(r) + [""] * 6
        if not str(r[2]).strip():
            continue
        tipo = str(r[2]).strip()
        try:
            d = json.loads(r[3]) if str(r[3]).strip().startswith("{") else {}
        except Exception:
            d = {}
        if tipo == "lancamento_presenca":
            ins = str(d.get("aba") or "").lower().startswith("insig")
            dt = str(d.get("data", ""))
            txt = f"{'Insight' if ins else 'Lista'} · {d.get('evento', '')}" + (f" ({dt[8:10]}/{dt[5:7]})" if len(dt) >= 10 else "")
        else:
            txt = nomes.get(tipo, tipo.replace("_", " ").capitalize())
        q = str(r[0])
        try:
            if isinstance(r[0], (int, float)) or q.replace(".", "", 1).isdigit():
                q = (datetime.datetime(1899, 12, 30) + datetime.timedelta(days=float(r[0]))).strftime("%d/%m %H:%M")
        except Exception:
            pass
        out.append({"q": q[:16], "quem": str(r[1]).split("@")[0][:30], "txt": txt[:80], "sit": str(r[4])[:20], "res": str(r[5])[:140]})
        if len(out) >= n:
            break
    return out

def processar_fila(api, pid):
    aba = _aba_real(api, pid, "Fila")
    if not aba:
        return 0
    v = _ler(api, pid, f"'{aba}'!A1:F2000")
    n, resumo = 0, []
    for i, row in enumerate(v[1:], start=2):
        row = row + [""] * 6
        if _sa(row[4]) not in ("pendente", "aguardando"):
            continue
        try:
            d = json.loads(row[3]); tp = d.get("tipo")
            if tp == "lancamento_presenca":
                res, sit = aplicar_evento(api, pid, d), "feito"
                try:
                    gravar_justificativas(api, pid, d, row[1])
                except Exception as e:
                    print(f"Aviso: justificativas não gravadas ({type(e).__name__}: {e}).")
                try:
                    gravar_info_evento(api, pid, d, row[1])
                except Exception as e:
                    print(f"Aviso: informações do evento não gravadas ({type(e).__name__}: {e}).")
            elif tp == "pagamento_manual":
                res, sit = aplicar_pagamento(api, pid, d), "feito"
            elif tp == "integrante":
                res, sit = aplicar_integrante(api, pid, d), "feito"
            elif tp == "boletos_sistema":
                res, sit = aplicar_boletos(api, pid, d), "feito"
            elif tp == "verificar_pagamentos":
                VERIFICAR_AGORA.append(str(row[1]).strip().lower() or "?")   # guarda quem pediu (para o aviso ao comando)
                res, sit = "verificação dos boletos pedida pela tela Pagamentos", "feito"
            elif tp in ("enquete_nova", "enquete_fim"):
                res, sit = aplicar_enquete(api, pid, d, row[1]), "feito"
            elif tp == "boletos_aviso":
                AVISO_BOLETO_MANUAL.append(1)
                res, sit = "aviso de boleto disponível: enviado depois de publicar o painel", "feito"
            elif tp == "adm_extra":
                res, sit = aplicar_adm_extra(d), "feito"
            elif tp == "visitante":
                res, sit = aplicar_visitante(d), "feito"
            elif tp == "teste_pagamento":     # botão de teste: não mexe em nada; manda o aviso só para quem pediu
                AVISOS_PAG.append(("💵 Pagamento novo (teste)", "Teste do aviso de pagamento: toque em 📲 Enviar no WhatsApp para ver a lista.",
                                   str(row[1]).strip().lower()))
                res, sit = "aviso de teste enviado só para quem pediu", "feito"
            else:
                res, sit = f"tipo desconhecido: {tp}", "erro"
        except Exception as e:
            res, sit = f"{type(e).__name__}: {e}", "erro"
        _gravar(api, pid, [(f"'{aba}'!E{i}", sit), (f"'{aba}'!F{i}", f"{datetime.datetime.now():%d/%m/%Y %H:%M} · {res}")])
        print(f"Fila linha {i}: {sit} · {res}"); n += 1
        if sit == "feito" and tp not in ("teste_pagamento", "adm_extra", "visitante", "boletos_aviso", "verificar_pagamentos"):   # estes não avisam o comando
            resumo.append((_resumo_item(d), str(row[1]).strip().lower(), d))
    return resumo

def _resumo_publico(itens_d):
    """Para a tela de entrada (que fica aberta): eventos com o nome; pagamento e cadastro sem dizer de quem."""
    eventos, comp, cad = [], 0, 0
    for d in itens_d:
        tp = d.get("tipo")
        if tp == "lancamento_presenca":
            try:
                eventos.append(f"{d.get('aba') or 'Presenca'}: {d.get('evento', '')} ({datetime.datetime.strptime(d['data'], '%Y-%m-%d'):%d/%m})")
            except Exception:
                eventos.append(f"{d.get('aba') or 'Presenca'}: {d.get('evento', '')}")
        elif tp == "pagamento_manual":
            comp += len([1 for v in (d.get("status") or {}).values() if v])
        elif tp == "integrante":
            cad += 1
        elif tp == "boletos_sistema":
            eventos.append("boletos do mês atualizados")
    partes = eventos[:2] + ([f"e mais {len(eventos) - 2} evento(s)"] if len(eventos) > 2 else [])
    if comp:
        partes.append(f"compensação de pagamento ({comp} irmão{'s' if comp > 1 else ''})")
    if cad:
        partes.append(f"cadastro ({cad} alteraç{'ões' if cad > 1 else 'ão'})")
    return " · ".join(partes)

def _quem(emails, xlsx, g):
    """Apelido (como está escrito na planilha) de quem fez os lançamentos."""
    ap_de = {str(k).strip().lower(): v for k, v in g.EMAIL_APELIDO.items()}
    try:
        import openpyxl
        _wb = openpyxl.load_workbook(xlsx, read_only=True, data_only=True)
        bonito = {str(r[1]).strip().lower(): str(r[1]).strip() for r in _wb["Dados Integrantes"].iter_rows(min_row=2, max_col=2, values_only=True) if r[1]}
        _wb.close()
    except Exception:
        bonito = {}
    if not ap_de:      # a lista e-mail -> apelido ainda não foi montada: lê da planilha
        try:
            import openpyxl
            _wb = openpyxl.load_workbook(xlsx, read_only=True, data_only=True)
            linhas = list(_wb["Dados Integrantes"].iter_rows(min_row=1, max_col=60, values_only=True)); _wb.close()
            ce = next((k for k, v in enumerate(linhas[0]) if _sa(v).replace("-", "").startswith("email")), None)
            if ce is not None:
                ap_de = {str(r[ce]).strip().lower(): str(r[1]).strip().lower() for r in linhas[1:] if r[1] and ce < len(r) and r[ce]}
        except Exception:
            pass
    return sorted({bonito.get(ap_de.get(e, ""), ap_de.get(e, "")) for e in emails if e} - {""})

def _resumo_item(d):
    """Texto curto do lançamento, para a notificação."""
    tp = d.get("tipo")
    try:
        if tp == "lancamento_presenca":
            dt = datetime.datetime.strptime(d["data"], "%Y-%m-%d").strftime("%d/%m")
            # adesão no momento: SIM (e Staff) ÷ todos os integrantes marcados no lançamento (igual ao painel)
            marc = [str(v or "").strip().lower() for v in (d.get("marcacoes") or {}).values()]
            sim = sum(1 for v in marc if v.startswith("sim") or v.startswith("staff"))
            ins = str(d.get("aba", "")).lower().startswith("insight")
            pc = f" · {round(100 * sim / len(marc))}% de {'resposta' if ins else 'adesão'} ({sim} de {len(marc)})" if marc else ""
            falta = [k for k, v in (d.get("marcacoes") or {}).items()
                     if str(v or "").strip().lower() in ("", "-") or (ins and str(v or "").strip().lower().startswith("n"))]
            ft = (" · Faltam: " + ", ".join(falta[:8]) + (f" e mais {len(falta) - 8}" if len(falta) > 8 else "")) if falta and len(falta) < len(marc) else ""
            return f"{d.get('aba') or 'Presenca'}: {d.get('evento', '')} ({dt}){pc}{ft}"
        if tp == "pagamento_manual":
            st = d.get("status") or {}
            ades = [k for k, v in st.items() if "ades" in _sa(str(v))]
            comp = [k for k, v in st.items() if v and "ades" not in _sa(str(v))]
            return " · ".join(([f"Adesão paga: {', '.join(ades)}"] if ades else []) + ([f"Compensação: {', '.join(comp)}"] if comp else [])) or "Pagamentos"
        if tp == "boletos_sistema":
            m = str(d.get("mes", ""))
            return f"Boletos do Sistema Insanos ({m[5:7]}/{m[:4]})" if len(m) >= 7 else "Boletos do Sistema Insanos"
        if tp == "integrante":
            if d.get("acao") == "cargo":
                return f"Cadastro: {d.get('apelido', '')} agora é {d.get('cargo', '')}"
            return f"Cadastro: {({'novo': 'novo integrante', 'afastar': 'afastado', 'reativar': 'reativado', 'excluir': 'excluído'}).get(d.get('acao'), d.get('acao'))} {d.get('apelido', '')}"
    except Exception:
        pass
    return str(tp or "alteração")

# ---------------------------------------------------------------------
#  Memória entre execuções (na nuvem cada execução começa do zero):
#  os arquivos de estado do painel ficam guardados na aba escondida "_estado" da planilha
# ---------------------------------------------------------------------
ESTADO = ["geo_cache.json", "pagamentos_verificacao.json", "pagamentos_manual.json", "pagamentos_antecipados.json", "adm_extra.json", "visitantes.json", "lembretes_enviados.json", "rsvp.json", "Lancamentos/notificacoes_enviadas.json", "Lancamentos/ultima_fila.json",
          "Lancamentos/ultima_alteracao.json"]

def _aba_estado(api, pid):
    aba = _aba_real(api, pid, "_estado")
    if not aba:
        api.spreadsheets().batchUpdate(spreadsheetId=pid, body={"requests": [
            {"addSheet": {"properties": {"title": "_estado", "hidden": True}}}]}).execute()
        aba = "_estado"
    return aba

def estado_carregar(api, pid):
    aba = _aba_estado(api, pid)
    for row in _ler(api, pid, f"'{aba}'!A1:B20"):
        row = list(row) + ["", ""]
        if row[0] in ESTADO and str(row[1]).strip():
            destino = os.path.join(PASTA, row[0]); os.makedirs(os.path.dirname(destino), exist_ok=True)
            with open(destino, "w", encoding="utf-8") as f:
                f.write(str(row[1]))
    return aba

def estado_salvar(api, pid, aba):
    linhas = []
    for nome in ESTADO:
        caminho = os.path.join(PASTA, nome)
        if os.path.exists(caminho):
            with open(caminho, encoding="utf-8") as f:
                linhas.append([nome, f.read()[:49000]])
    if linhas:
        api.spreadsheets().values().update(spreadsheetId=pid, range=f"'{aba}'!A1:B{len(linhas)}",
            valueInputOption="RAW", body={"values": linhas}).execute()

# ---------------------------------------------------------------------
#  Verificação dos boletos da Iugu (mesma regra do PC), gravando na aba PGTO da planilha do Google
# ---------------------------------------------------------------------
AVISOS_PAG = []          # avisos de pagamento novo, enviados depois de publicar o painel
AGRADECER = []           # (rótulo do irmão, mês, antecipado): agradecimento individual depois de publicar

def verificar_boletos(api, pid, g):
    hoje = datetime.date.today()
    ciclo = g.vencimento(hoje.year, hoje.month)
    if hoje > ciclo:
        print(f"Verificação de pagamentos: fora do período (vencimento foi {ciclo:%d/%m}). Nada a fazer.")
        return False
    aba = _aba_real(api, pid, "PGTO")
    if not aba:
        print("Verificação de pagamentos: aba PGTO não encontrada."); return False
    print(f"Verificando pagamentos (ciclo com vencimento em {ciclo:%d/%m/%Y})...")
    diag_mostrados = 0
    v = _ler(api, pid, f"'{aba}'!A1:J400")
    manual = g._manual_carregar(ciclo)
    inativos = g.inativos_planilha()
    itens, grav, novos_pagos = [], [], 0
    for r, row in enumerate(v[1:], start=2):
        row = list(row) + [""] * 10
        link = str(row[9]).strip()
        if not link.startswith("http"):
            continue
        rotulo, atual = str(row[1]).strip(), str(row[2]).strip()
        if any(rotulo.lower().endswith(ap) for ap in inativos):
            print(f"  {rotulo}: afastado/excluído (não verificado)"); continue
        if "ades" in _sa(atual) and f"{ciclo:%m/%Y}" in atual:     # pagou a adesão neste mês: não cobra a mensalidade
            itens.append({"rotulo": rotulo, "selo": None, "venc_boleto": None, "status": atual, "obs": "adesão paga no mês em que entrou"})
            print(f"  {rotulo}: adesão paga ({atual})"); continue
        item = {"rotulo": rotulo, "selo": None, "venc_boleto": None, "status": atual, "obs": ""}
        try:
            selo, vb = g._ler_boleto(link)
        except Exception as e:
            item["obs"] = "não consegui abrir o boleto"; itens.append(item); print(f"  {rotulo}: não abriu ({e})"); continue
        item["selo"], item["venc_boleto"] = selo, (vb.isoformat() if vb else None)
        print(f"  {rotulo}: lido no boleto -> situação: {selo or 'NÃO ENCONTRADA'} · vencimento: {vb:%d/%m/%Y}" if vb else f"  {rotulo}: lido no boleto -> situação: {selo or 'NÃO ENCONTRADA'} · vencimento: não encontrado")
        if not selo and diag_mostrados < 2:      # mostra o trecho da página para acertar a leitura (sem dados pessoais: só o que vem depois do vencimento)
            diag_mostrados += 1
            pg = getattr(g, "ULTIMA_PAGINA_BOLETO", "") or ""
            i0 = pg.find("Vencimento")
            print("    trecho da página depois do vencimento: " + re.sub(r"\s+", " ", pg[i0:i0 + 700] if i0 >= 0 else pg[:300]))
        boleto_do_mes = not vb or (vb.year, vb.month) == (ciclo.year, ciclo.month)
        if boleto_do_mes and selo and "pag" in selo.lower():
            novo = "Pago em Dia"
        else:
            novo = "Em aberto"
            if not boleto_do_mes:
                item["obs"] = f"boleto com vencimento {vb:%d/%m/%Y}: o boleto deste mês ainda não foi cadastrado"
            elif selo:
                item["obs"] = f"situação no boleto: {selo}"
        ap_man = g._manual_de(rotulo, manual)
        if not ap_man and novo != "Pago em Dia" and atual.lower().startswith("aguard"):
            ap_man = rotulo.split()[-1] if rotulo else rotulo
            manual[ap_man] = "Aguardando Compensação"
        if novo == "Pago em Dia":
            if ap_man:
                manual.pop(ap_man, None)
            item["boleto_pago"] = True
        elif ap_man:
            novo = manual[ap_man]
            item["obs"] = (item.get("obs") + " · " if item.get("obs") else "") + "marcado à mão"
        item["status"] = novo
        if novo != atual:
            grav.append((f"'{aba}'!C{r}", novo)); print(f"  {rotulo}: {atual or '(vazio)'} -> {novo}")
            if novo == "Pago em Dia":
                novos_pagos += 1; item["novo_pago"] = True
        itens.append(item)
    _gravar(api, pid, grav)
    g._manual_salvar(ciclo, manual)
    quando = datetime.datetime.now().strftime("%d/%m/%Y às %H:%M")
    ultimo, pagos_em = None, {}
    try:
        with open(g.ARQ_VERIF, encoding="utf-8") as f:
            ant = json.load(f)
        ultimo = ant.get("ultimo_pagamento")
        if str(ant.get("ciclo") or "")[:7] == ciclo.isoformat()[:7]:
            pagos_em = dict(ant.get("pagos_em") or {})
    except Exception:
        pass
    agora_ts = datetime.datetime.now().strftime("%Y-%m-%dT%H:%M:%S")
    for it in itens:
        rot = it.get("rotulo", "")
        if it.get("status") == "Pago em Dia":
            pagos_em.setdefault(rot, agora_ts if it.get("novo_pago") else "")
        else:
            pagos_em.pop(rot, None)
    if novos_pagos:
        ultimo = {"ts": agora_ts, "hora": datetime.datetime.now().strftime("%d/%m às %H:%M")}
    with open(g.ARQ_VERIF, "w", encoding="utf-8") as f:
        json.dump({"quando": quando, "ciclo": ciclo.isoformat(), "novos_pagos": novos_pagos, "ultimo_pagamento": ultimo,
                   "pagos_em": pagos_em, "itens": itens}, f, ensure_ascii=False, indent=1)
    print(f"Verificação concluída: {len(itens)} boletos lidos, {len(grav)} status alterado(s).")
    try:
        with open(os.path.join(PASTA, "pagamentos_antecipados.json"), encoding="utf-8") as f:
            ja_avisados = set((json.load(f).get("pagos") or {}).keys())
    except Exception:
        ja_avisados = set()
    novos = [str(it.get("rotulo", "")).split(" - ")[-1].strip() for it in itens if it.get("novo_pago") and it.get("rotulo") not in ja_avisados]
    AGRADECER.extend((it.get("rotulo", ""), MESES_PT[ciclo.month - 1], False) for it in itens if it.get("novo_pago") and it.get("rotulo") not in ja_avisados)
    if novos:
        hora = datetime.datetime.now().strftime("%d/%m às %H:%M")
        # o aviso sai DEPOIS de publicar o painel: o botão "📲 Enviar no WhatsApp" monta a lista com os dados novos
        pagos = sum(1 for it in itens if str(it.get("status", "")).lower().startswith("pago"))
        situacao = f" {round(100 * pagos / len(itens))}% pagos ({pagos} de {len(itens)})." if itens else ""
        if len(novos) == 1:
            AVISOS_PAG.append(("💵 Pagamento novo", f"{novos[0]} pagou a mensalidade (identificado {hora})." + situacao))
        else:
            AVISOS_PAG.append((f"💵 {len(novos)} pagamentos novos", ", ".join(novos) + f" (identificados {hora})." + situacao))
    return bool(grav)

def preparar_ambiente():
    """Chaves do cofre -> arquivos temporários que o painel espera."""
    lanc = os.path.join(PASTA, "Lancamentos"); os.makedirs(lanc, exist_ok=True)
    pem, pub = os.environ.get("PUSH_VAPID_PEM", "").strip(), os.environ.get("PUSH_VAPID_PUB", "").strip()
    if pem and pub:
        with open(os.path.join(lanc, "push_vapid_privada.pem"), "w", encoding="utf-8") as f:
            f.write(pem + "\n")
        with open(os.path.join(lanc, "push_vapid.json"), "w", encoding="utf-8") as f:
            f.write(pub)
    # publicação: a chave automática do GitHub, neste mesmo repositório
    dono, repo = (os.environ.get("GITHUB_REPOSITORY", "/") + "/").split("/")[:2]
    tok = os.path.join(PASTA, ".token_github.json")
    with open(tok, "w", encoding="utf-8") as f:
        json.dump({"token": os.environ.get("GITHUB_TOKEN", ""), "usuario": dono, "repo": repo}, f)
    return tok, dono, repo

def main():
    planilha_id, cred = os.environ.get("PLANILHA_ID", "").strip(), os.environ.get("GOOGLE_CREDENTIALS", "")
    if not planilha_id or not cred:
        print("[ERRO] Faltam os segredos PLANILHA_ID e/ou GOOGLE_CREDENTIALS no GitHub."); sys.exit(1)
    evento = os.environ.get("MODO", "")   # schedule (horário) · repository_dispatch (lançamento/edição) · workflow_dispatch (botão)
    acao = ""   # no repository_dispatch: "lancamento" (planilha) ou "rsvp" (um irmão respondeu Vou / Não vou / Já respondi)
    try:
        if os.environ.get("GITHUB_EVENT_PATH"):
            acao = str(json.load(open(os.environ["GITHUB_EVENT_PATH"], encoding="utf-8")).get("action", ""))
    except Exception:
        pass
    if evento == "repository_dispatch" and acao == "agenda":
        evento = "schedule"   # o despertador do Apps Script vale como uma rodada agendada (boletos, lembretes e avisos dos horários)
    # horários automáticos: na época de pagamentos (dia 1 ao vencimento) de meia em meia hora, das 7h às 18h;
    # fora dela, só a rodada das 9h (avisos do dia). Modificações (painel ou planilha) atualizam na hora.
    sys.argv = [sys.argv[0], "--publicar"]
    import painel as g
    agora = datetime.datetime.now(); hoje = agora.date()
    ciclo = g.vencimento(hoje.year, hoje.month)
    epoca_pag = hoje <= ciclo
    if evento == "schedule":
        h, mi = agora.hour, agora.minute
        rodada_avisos = h in (9, 13, 19) and mi < 30          # avisos de aniversário: 9h, 13h e 19h
        if not ((epoca_pag and 7 <= h <= 18) or rodada_avisos):
            print(f"Nada a fazer agora ({agora:%H:%M}): fora da época de pagamentos e dos horários de aviso (9h, 13h e 19h)."); return
    api = api_google(cred)
    if os.environ.get("SOMENTE_SOS"):       # rotina rápida do SOS (sos.yml): só o alerta, sem gerar o painel
        preparar_ambiente()
        print(f"SOS: {processar_sos(api, planilha_id, g)} registro(s) processado(s).")
        try:
            print(f"Mensagens: {processar_mensagens(api, planilha_id, g)} processada(s).")
        except Exception as e:
            print(f"Aviso: mensagens não processadas ({type(e).__name__}: {e}).")
        acompanhar_sos(api, planilha_id, g)
        return
    aba_estado = estado_carregar(api, planilha_id)
    migrar_entradas(api, planilha_id)
    resumo_fila = processar_fila(api, planilha_id)
    try:
        g.FILA_RECENTE[:] = fila_recente(api, planilha_id)     # "Meus envios" do protótipo do novo layout
    except Exception as e:
        print(f"Aviso: resumo da Fila não montado ({type(e).__name__}: {e}).")
    arq_fila = os.path.join(PASTA, "Lancamentos", "ultima_fila.json")
    lanc_recente = False          # houve lançamento pelo painel nos últimos 10 minutos? (aí não é "edição à mão")
    try:
        with open(arq_fila, encoding="utf-8") as f:
            ult = datetime.datetime.fromisoformat(json.load(f)["ts"])
        lanc_recente = (datetime.datetime.now() - ult).total_seconds() < 600
    except Exception:
        pass
    if resumo_fila:
        print(f"Fila: {len(resumo_fila)} lançamento(s) aplicado(s) na planilha.")
        os.makedirs(os.path.dirname(arq_fila), exist_ok=True)
        with open(arq_fila, "w", encoding="utf-8") as f:
            json.dump({"ts": datetime.datetime.now().isoformat(timespec="seconds")}, f)
    tok, dono, repo = preparar_ambiente()
    try:                                     # reforço: se a rotina rápida do SOS falhou, o alerta sai aqui
        processar_sos(api, planilha_id, g)
    except Exception as e:
        print(f"Aviso: SOS não processado ({type(e).__name__}: {e}).")
    try:                                     # reforço também para as mensagens
        processar_mensagens(api, planilha_id, g)
    except Exception as e:
        print(f"Aviso: mensagens não processadas ({type(e).__name__}: {e}).")
    xlsx = os.path.join(PASTA, "Insanos.xlsx")
    print("Lendo a planilha do Google...")
    montar_planilha(ler_google(planilha_id, cred, api), xlsx)
    g.PLANILHA = xlsx
    g.ARQ_TOKEN = tok
    g.GITHUB_USUARIO, g.GITHUB_REPO = dono, repo
    g.SITE_URL = f"https://{dono}.github.io/{repo}/"
    g.limpar_historico_github = lambda: True       # na nuvem o histórico fica (sem limpeza forçada)
    g.PLANILHA_URL = f"https://docs.google.com/spreadsheets/d/{planilha_id}/edit"   # atalho 📗 Planilha (só para quem vê a tela Acessos)
    arq_alt = os.path.join(PASTA, "Lancamentos", "ultima_alteracao.json")
    agora_txt = datetime.datetime.now().strftime("%d/%m às %H:%M")
    quem_lancou = _quem([e for _, e, _d in resumo_fila], xlsx, g) if resumo_fila else []
    alt = None
    if resumo_fila:
        alt = _resumo_publico([d for _, _, d in resumo_fila]) + (f" · por {', '.join(quem_lancou)}" if quem_lancou else "") + f" · {agora_txt}"
    elif evento == "repository_dispatch" and not lanc_recente:
        alt = f"planilha editada diretamente · {agora_txt}"
    if alt:
        os.makedirs(os.path.dirname(arq_alt), exist_ok=True)
        with open(arq_alt, "w", encoding="utf-8") as f:
            json.dump({"txt": alt}, f, ensure_ascii=False)
    try:
        with open(arq_alt, encoding="utf-8") as f:
            g.ULTIMA_ALTERACAO_TXT = json.load(f).get("txt", "")
    except Exception:
        pass
    if os.environ.get("LANC_URL") and os.environ.get("LANC_CHAVE"):
        g.NUVEM = {"url": os.environ["LANC_URL"], "chave": os.environ["LANC_CHAVE"]}   # o painel manda os lançamentos direto
    # partes que só existem no PC (arquivos de lançamento, Excel pelo Windows, tarefa agendada): desligadas na nuvem
    g.renumerar_lista = lambda mudancas: renumerar_google(api, planilha_id, mudancas)
    for nome in ("migrar_staff", "importar_lancamentos", "importar_integrantes", "importar_pagamentos_manuais",
                 "importar_acessos", "garantir_tarefa_avisos"):
        setattr(g, nome, lambda *a, **k: None)
    # notificações no celular: só com o segredo NOTIFICAR = sim (ligado na virada, para não chegarem em dobro com o PC)
    if os.environ.get("NOTIFICAR", "").strip().lower() not in ("sim", "s", "1", "true"):
        g.notificar = lambda titulo, corpo, tag="aviso", **k: print(f"(notificações desligadas na nuvem) teria enviado: {titulo} · {corpo}")
        g.avisos_diarios = lambda *a, **k: None
        print("Notificações: desligadas na nuvem (segredo NOTIFICAR não está como 'sim').")
    else:
        print("Notificações: ligadas.")
    # boletos: nas rodadas de hora em hora (8h às 17h) e no botão "Run workflow"; nos lançamentos, não (para ser rápido)
    try:
        if aplicar_respostas(api, planilha_id, g):
            montar_planilha(ler_google(planilha_id, cred, api), xlsx)
    except Exception as e:
        print(f"Aviso: respostas dos irmãos não aplicadas ({type(e).__name__}: {e}).")
    try:
        remover_abas_sem_uso(api, planilha_id)
    except Exception as e:
        print(f"Aviso: não consegui apagar as abas sem uso ({type(e).__name__}: {e}).")
    try:
        limpar_totais(api, planilha_id)
    except Exception as e:
        print(f"Aviso: limpeza das contas de Sim/Não/% falhou ({type(e).__name__}: {e}).")
    try:                                         # virada do mês: boletos do mês seguinte (coluna K) passam para a J
        if promover_antecipados(api, planilha_id):
            montar_planilha(ler_google(planilha_id, cred, api), xlsx)
    except Exception as e:
        print(f"Aviso: virada dos boletos antecipados falhou ({type(e).__name__}: {e}).")
    if evento in ("workflow_dispatch", "schedule"):  # pagamentos antecipados: em todas as rodadas automáticas
        try:
            verificar_antecipados(api, planilha_id, g)
        except Exception as e:
            print(f"Aviso: verificação dos antecipados falhou ({type(e).__name__}: {e}).")
    if evento == "workflow_dispatch" or VERIFICAR_AGORA or (evento == "schedule" and epoca_pag and 7 <= agora.hour <= 18):
        try:
            if verificar_boletos(api, planilha_id, g):
                print("Relendo a planilha com os pagamentos atualizados...")
                montar_planilha(ler_google(planilha_id, cred, api), xlsx)
        except Exception as e:
            print(f"Aviso: verificação de pagamentos falhou ({type(e).__name__}: {e}). O painel é gerado mesmo assim.")
    try:
        g.main()
        try:
            escrever_otp(api, planilha_id, g)
        except Exception as e:
            print(f"Aviso: lista do código por e-mail não atualizada ({type(e).__name__}: {e}).")
        # verificação feita à mão (Ver. Agora ou Run workflow): avisa o comando quem pediu, por onde e a hora
        if VERIFICAR_AGORA or evento == "workflow_dispatch":
            try:
                hora = datetime.datetime.now().strftime("%H:%M")
                if VERIFICAR_AGORA:
                    em = VERIFICAR_AGORA[0]
                    ap = (getattr(g, "EMAIL_APELIDO", {}) or {}).get(em) or em.split("@")[0]
                    nomes = list(((getattr(g, "PAG_MEU", {}) or {}).get("st") or {}).keys())   # os apelidos como estão na planilha
                    nome = next((x for x in nomes if x.lower() == str(ap).lower()), str(ap).title())
                    onde = "pelo botão Ver. Agora"
                else:
                    ator = os.environ.get("GITHUB_ACTOR", "")
                    dono = os.environ.get("GITHUB_REPOSITORY_OWNER", "")
                    resp = list(getattr(g, "ACESSOS_SOMENTE", []) or [])
                    nome = resp[0] if (ator and ator == dono and resp) else (ator or "alguém")
                    onde = "pelo Run workflow do GitHub"
                resp_ = {str(x).strip().lower() for x in (getattr(g, "ACESSOS_SOMENTE", []) or [])}
                teste = " (Possível teste ou melhoria)" if str(nome).strip().lower() in resp_ else ""   # o responsável pelo painel: pode ser teste
                g.notificar("🔄 Verificação manual dos pagamentos", f"Pedida por {nome}, às {hora}.{teste}", "verif-manual", icone="notif-pag.png")
            except Exception as e:
                print(f"Aviso: aviso da verificação manual não enviado ({type(e).__name__}: {e}).")
        for aviso in AVISOS_PAG:
            titulo, corpo, so_para, com_zap = (list(aviso) + [None, True])[:4]
            try:
                zap = None
                if com_zap:   # a mensagem pronta vai dentro do aviso (abre direto no WhatsApp); se for grande demais, o app monta na hora
                    zap = getattr(g, "ZAP_PAG", "") or "__STATUS_PAGAMENTOS__"
                    if len(zap.encode("utf-8")) > 2800:
                        zap = "__STATUS_PAGAMENTOS__"
                g.notificar(titulo, corpo, "pagamento", zap=zap, icone="notif-pag.png", so_para=so_para)
            except Exception as e:
                print(f"Aviso: notificação de pagamento não enviada ({e}).")
        AVISOS_PAG.clear()
        # agradecimento individual: só para o irmão que pagou
        for rotulo, mes, antec in AGRADECER:
            try:
                rl = str(rotulo).lower()
                ap = max((k for k in g.EMAIL_APELIDO_INV if k and rl.endswith(k)), key=len, default="")
                email = g.EMAIL_APELIDO_INV.get(ap, "")
                if "@" not in email:
                    continue
                nome = next((m for m in [str(rotulo).split(" - ")[-1].strip()] if m), ap)
                nome = nome[-len(ap):] if len(nome) >= len(ap) else nome
                g.notificar(f"🙏 Valeu, {nome}!",
                            f"Recebemos o pagamento da mensalidade de {mes}{' antes do vencimento' if antec else ''}. Obrigado pelo compromisso com o Insanos M.C.! 💀🏦",
                            "agradecimento", icone="notif-pag.png", so_para=email)
            except Exception as e:
                print(f"Aviso: agradecimento não enviado ({e}).")
        AGRADECER.clear()
        # notificação para o comando: o site já foi publicado com a alteração
        try:
            if resumo_fila:
                quem = _quem([e for _, e, _d in resumo_fila], xlsx, g)
                itens = [r for r, _, _d in resumo_fila]
                corpo = " · ".join(itens[:3]) + (f" · e mais {len(itens) - 3}" if len(itens) > 3 else "")
                zaps = [d.get("zap") for _, _, d in resumo_fila if d.get("zap") and d.get("tipo") == "pagamento_manual"]   # pagamentos lançados
                for _, em_, d_ in resumo_fila:   # evento lançado ou corrigido: a lista atualizada do evento, pronta para o WhatsApp
                    if d_.get("tipo") == "lancamento_presenca":
                        try:
                            z = str(d_.get("zap") or "") or g.zap_evento(d_, em_)   # a mensagem do Lançar (igual à da Lista); senão, a da nuvem
                            if z and len(z.encode("utf-8")) <= 2800:
                                zaps.append(z)
                        except Exception as e:
                            print(f"Aviso: lista do evento para o WhatsApp não montada ({type(e).__name__}: {e}).")
                # emoji e imagem: pagamento > PUB > ação social > evento > insight > cadastro (o mais importante da rodada)
                ordem = ["notif-pag.png", "notif-pub.png", "notif-social.png", "notif-evento.png", "notif-insight.png", "notif-cadastro.png"]
                tipos = [_tipo_aviso(d) for _, _, d in resumo_fila]
                emoji, icone = min(tipos, key=lambda x: ordem.index(x[1])) if tipos else ("✅", None)
                resp_ = {str(x).strip().lower() for x in (getattr(g, "ACESSOS_SOMENTE", []) or [])}
                teste = " (Possível teste ou melhoria)" if any(str(q).strip().lower() in resp_ for q in quem) else ""   # lançado pelo responsável pelo painel
                g.notificar(f"{emoji} Painel atualizado", corpo + (f" (lançado por {', '.join(quem)})" if quem else "") + teste, "atualizacao",
                            zap=zaps[-1] if zaps else None, icone=icone)      # a mensagem pronta do último lançamento com lista
            elif evento == "repository_dispatch" and not lanc_recente and acao != "rsvp":
                # a planilha avisou e a Fila estava vazia: foi edição feita à mão na planilha
                g.notificar("📝 Painel atualizado", "A planilha foi editada à mão. O painel já está com a alteração.", "atualizacao")
        except Exception as e:
            print(f"Aviso: notificação de atualização não enviada ({e}).")
        try:
            novos = [(d.get("data"), d.get("evento")) for _, _, d in (resumo_fila or [])
                     if d.get("tipo") == "lancamento_presenca" and str(d.get("aba") or "Presenca").lower().startswith("presen")
                     and str(d.get("data", "")) >= datetime.date.today().isoformat()]
            if novos:
                enviar_convites(g, novos)
            ins_novos = [(d.get("data"), d.get("evento")) for _, _, d in (resumo_fila or [])
                         if d.get("tipo") == "lancamento_presenca" and str(d.get("aba") or "").lower().startswith("insig")]
            if ins_novos:
                enviar_aviso_insight(g, ins_novos)
        except Exception as e:
            print(f"Aviso: convites não enviados ({type(e).__name__}: {e}).")
        if RESPOSTAS:
            try:
                por = {}
                for ap, nome, r in RESPOSTAS:
                    por.setdefault(nome, []).append(f"{ap} {'✅' if r in ('vou', 'resp') else '❌'}")
                g.notificar("✋ Respostas dos irmãos", " · ".join(f"{n}: {', '.join(x)}" for n, x in por.items()), "atualizacao", icone="notif-evento.png")
            except Exception as e:
                print(f"Aviso: aviso das respostas não enviado ({e}).")
            RESPOSTAS.clear()
        try:
            enviar_boleto_disponivel(g)
        except Exception as e:
            print(f"Aviso: aviso de boleto disponível falhou ({type(e).__name__}: {e}).")
        if evento == "schedule" and datetime.datetime.now().hour >= 9:   # ☀️ bom-dia: uma vez por dia, a partir das 9h, para todos
            try:
                env_bd = _json_ler("lembretes_enviados.json", {})
                k_bd = f"{datetime.date.today().isoformat()}|bomdia"
                if k_bd not in env_bd:
                    fora = []   # quem desligou o bom-dia (aba Preferências, gravada pelo Apps Script)
                    try:
                        if _aba_real(api, planilha_id, "Preferencias"):
                            fora = [str(r[0]).strip().lower() for r in _ler(api, planilha_id, "'Preferencias'!A2:B1000") if len(r) > 1 and str(r[1]).strip().upper() == "SIM"]
                    except Exception:
                        pass
                    tit_bd, corpo_bd = g.bom_dia(getattr(g, "HOJE_EVS", []))
                    g.notificar(tit_bd, corpo_bd, "bomdia", icone="notif-evento.png", todos=True, excluir=fora)
                    env_bd[k_bd] = datetime.datetime.now().isoformat(timespec="minutes")
                    _json_gravar("lembretes_enviados.json", env_bd)
                    print(f"Bom-dia enviado (fora: {len(fora)}).")
            except Exception as e:
                print(f"Aviso: bom-dia não enviado ({type(e).__name__}: {e}).")
        if ENQ_NOVA:   # enquete nova do comando: todos os celulares cadastrados (a enquete já está no app)
            try:
                g.notificar("📊 Enquete do comando", ENQ_NOVA[-1] + " Toque para votar.", "enquete", icone="notif-importante.png", todos=True)
                print("Aviso da enquete nova enviado.")
            except Exception as e:
                print(f"Aviso: enquete não avisada ({type(e).__name__}: {e}).")
            ENQ_NOVA.clear()
        if evento == "schedule":
            try:
                enviar_lembretes(g)
            except Exception as e:
                print(f"Aviso: lembretes individuais não enviados ({type(e).__name__}: {e}).")
    finally:
        try:
            estado_salvar(api, planilha_id, aba_estado)
        except Exception as e:
            print(f"Aviso: não consegui guardar a memória na aba _estado ({e}).")

if __name__ == "__main__":
    main()
