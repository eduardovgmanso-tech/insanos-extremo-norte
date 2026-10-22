"""
Painel Insanos na nuvem (GitHub Actions).
1) Lê a planilha do Google (conta de serviço) e monta uma planilha temporária só com as abas que o painel usa.
2) Roda o mesmo programa do painel (painel.py) e publica neste repositório.
Nada de dados pessoais fica no repositório: a planilha temporária existe só durante a execução.
"""
import os, sys, json, datetime, unicodedata, re

PASTA = os.path.dirname(os.path.abspath(__file__))
ABAS = ["Presenca", "Insights", "Dados Integrantes", "PGTO", "Pagamentos"]
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
    for a in ABAS:
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
    res = api.spreadsheets().values().batchGet(spreadsheetId=planilha_id, ranges=[f"'{real[a]}'!A1:BZ500" for a in nomes],
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
                if "pago" in _sa(atual) and "aguard" not in _sa(atual):
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
        gr = []
        for r in range(8, ((fim - 2) if fim else 42) + 1):
            gr += _totais_linha(aba, v, r, ult)
        _gravar(api, pid, gr)
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

def aplicar_integrante(api, pid, d):
    ap, acao = str(d.get("apelido", "")).strip(), d.get("acao")
    if not ap or acao not in ("novo", "afastar", "reativar", "excluir"):
        return "pedido incompleto"
    aba, C, linhas, fim_tab, ult_col = _di_mapa(api, pid)
    sid = _ids(api, pid)[aba]
    r = linhas.get(_sa(ap))
    log = []
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
            v = _ler(api, pid, f"'{ae}'!A1:CZ60"); ult -= 1
            fim = _linha_adesao(v); gr = []
            for rr in range(8, ((fim - 2) if fim else 42) + 1):
                gr += _totais_linha(ae, v, rr, ult)
            _gravar(api, pid, gr)
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
            elif tp == "pagamento_manual":
                res, sit = aplicar_pagamento(api, pid, d), "feito"
            elif tp == "integrante":
                res, sit = aplicar_integrante(api, pid, d), "feito"
            else:
                res, sit = f"tipo desconhecido: {tp}", "erro"
        except Exception as e:
            res, sit = f"{type(e).__name__}: {e}", "erro"
        _gravar(api, pid, [(f"'{aba}'!E{i}", sit), (f"'{aba}'!F{i}", f"{datetime.datetime.now():%d/%m/%Y %H:%M} · {res}")])
        print(f"Fila linha {i}: {sit} · {res}"); n += 1
        if sit == "feito":
            resumo.append((_resumo_item(d), str(row[1]).strip().lower()))
    return resumo

def _resumo_item(d):
    """Texto curto do lançamento, para a notificação."""
    tp = d.get("tipo")
    try:
        if tp == "lancamento_presenca":
            dt = datetime.datetime.strptime(d["data"], "%Y-%m-%d").strftime("%d/%m")
            return f"{d.get('aba') or 'Presenca'}: {d.get('evento', '')} ({dt})"
        if tp == "pagamento_manual":
            return "Compensação: " + ", ".join(k for k, v in (d.get("status") or {}).items() if v)
        if tp == "integrante":
            return f"Cadastro: {({'novo': 'novo integrante', 'afastar': 'afastado', 'reativar': 'reativado', 'excluir': 'excluído'}).get(d.get('acao'), d.get('acao'))} {d.get('apelido', '')}"
    except Exception:
        pass
    return str(tp or "alteração")

# ---------------------------------------------------------------------
#  Memória entre execuções (na nuvem cada execução começa do zero):
#  os arquivos de estado do painel ficam guardados na aba escondida "_estado" da planilha
# ---------------------------------------------------------------------
ESTADO = ["pagamentos_verificacao.json", "pagamentos_manual.json", "Lancamentos/notificacoes_enviadas.json", "Lancamentos/ultima_fila.json"]

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
        item = {"rotulo": rotulo, "selo": None, "venc_boleto": None, "status": atual, "obs": ""}
        try:
            selo, vb = g._ler_boleto(link)
        except Exception as e:
            item["obs"] = "não consegui abrir o boleto"; itens.append(item); print(f"  {rotulo}: não abriu ({e})"); continue
        item["selo"], item["venc_boleto"] = selo, (vb.isoformat() if vb else None)
        boleto_do_mes = not vb or (vb.year, vb.month) == (ciclo.year, ciclo.month)
        if boleto_do_mes and selo and "pag" in selo.lower():
            novo = "Pago em Dia"
        else:
            novo = "Em aberto"
            if not boleto_do_mes:
                item["obs"] = f"boleto com vencimento {vb:%d/%m/%Y}: o boleto deste mês ainda não foi cadastrado"
            elif selo:
                item["obs"] = f"selo da Iugu: {selo}"
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
    novos = [str(it.get("rotulo", "")).split(" - ")[-1].strip() for it in itens if it.get("novo_pago")]
    if novos:
        hora = datetime.datetime.now().strftime("%d/%m às %H:%M")
        try:
            if len(novos) == 1:
                g.notificar("💰 Pagamento novo", f"{novos[0]} pagou a mensalidade (identificado {hora}).", "pagamento")
            else:
                g.notificar(f"💰 {len(novos)} pagamentos novos", ", ".join(novos) + f" (identificados {hora}).", "pagamento")
        except Exception as e:
            print(f"Aviso: notificação não enviada ({e}).")
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
    aba_estado = estado_carregar(api, planilha_id)
    resumo_fila = processar_fila(api, planilha_id)
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
    xlsx = os.path.join(PASTA, "Insanos.xlsx")
    print("Lendo a planilha do Google...")
    montar_planilha(ler_google(planilha_id, cred, api), xlsx)
    g.PLANILHA = xlsx
    g.ARQ_TOKEN = tok
    g.GITHUB_USUARIO, g.GITHUB_REPO = dono, repo
    g.SITE_URL = f"https://{dono}.github.io/{repo}/"
    g.limpar_historico_github = lambda: True       # na nuvem o histórico fica (sem limpeza forçada)
    if os.environ.get("LANC_URL") and os.environ.get("LANC_CHAVE"):
        g.NUVEM = {"url": os.environ["LANC_URL"], "chave": os.environ["LANC_CHAVE"]}   # o painel manda os lançamentos direto
    # partes que só existem no PC (arquivos de lançamento, Excel pelo Windows, tarefa agendada): desligadas na nuvem
    g.renumerar_lista = lambda mudancas: renumerar_google(api, planilha_id, mudancas)
    for nome in ("migrar_staff", "importar_lancamentos", "importar_integrantes", "importar_pagamentos_manuais",
                 "importar_acessos", "garantir_tarefa_avisos"):
        setattr(g, nome, lambda *a, **k: None)
    # notificações no celular: só com o segredo NOTIFICAR = sim (ligado na virada, para não chegarem em dobro com o PC)
    if os.environ.get("NOTIFICAR", "").strip().lower() not in ("sim", "s", "1", "true"):
        g.notificar = lambda titulo, corpo, tag="aviso": print(f"(notificações desligadas na nuvem) teria enviado: {titulo} · {corpo}")
        g.avisos_diarios = lambda *a, **k: None
        print("Notificações: desligadas na nuvem (segredo NOTIFICAR não está como 'sim').")
    else:
        print("Notificações: ligadas.")
    # boletos: nas rodadas de hora em hora (8h às 17h) e no botão "Run workflow"; nos lançamentos, não (para ser rápido)
    if evento == "workflow_dispatch" or (evento == "schedule" and epoca_pag and 7 <= agora.hour <= 18):
        try:
            if verificar_boletos(api, planilha_id, g):
                print("Relendo a planilha com os pagamentos atualizados...")
                montar_planilha(ler_google(planilha_id, cred, api), xlsx)
        except Exception as e:
            print(f"Aviso: verificação de pagamentos falhou ({type(e).__name__}: {e}). O painel é gerado mesmo assim.")
    try:
        g.main()
        # notificação para o comando: o site já foi publicado com a alteração
        try:
            if resumo_fila:
                ap_de = {str(k).strip().lower(): v for k, v in g.EMAIL_APELIDO.items()}
                try:      # apelido como está escrito na planilha ("Edu", e não "edu")
                    import openpyxl
                    _wb = openpyxl.load_workbook(xlsx, read_only=True, data_only=True)
                    bonito = {str(r[1]).strip().lower(): str(r[1]).strip() for r in _wb["Dados Integrantes"].iter_rows(min_row=2, max_col=2, values_only=True) if r[1]}
                    _wb.close()
                except Exception:
                    bonito = {}
                quem = sorted({bonito.get(ap_de.get(e, ""), ap_de.get(e, "")) for _, e in resumo_fila if e} - {""})
                itens = [r for r, _ in resumo_fila]
                corpo = " · ".join(itens[:3]) + (f" · e mais {len(itens) - 3}" if len(itens) > 3 else "")
                g.notificar("✅ Painel atualizado", corpo + (f" (lançado por {', '.join(quem)})" if quem else ""), "atualizacao")
            elif evento == "repository_dispatch" and not lanc_recente:
                # a planilha avisou e a Fila estava vazia: foi edição feita à mão na planilha
                g.notificar("📝 Painel atualizado", "A planilha foi editada à mão. O painel já está com a alteração.", "atualizacao")
        except Exception as e:
            print(f"Aviso: notificação de atualização não enviada ({e}).")
    finally:
        try:
            estado_salvar(api, planilha_id, aba_estado)
        except Exception as e:
            print(f"Aviso: não consegui guardar a memória na aba _estado ({e}).")

if __name__ == "__main__":
    main()
