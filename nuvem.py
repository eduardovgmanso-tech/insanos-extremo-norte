"""
Painel Insanos na nuvem (GitHub Actions).
1) Lê a planilha do Google (conta de serviço) e monta uma planilha temporária só com as abas que o painel usa.
2) Roda o mesmo programa do painel (painel.py) e publica neste repositório.
Nada de dados pessoais fica no repositório: a planilha temporária existe só durante a execução.
"""
import os, sys, json, datetime, unicodedata

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

def ler_google(planilha_id, cred_json):
    from google.oauth2 import service_account
    from googleapiclient.discovery import build
    cred = service_account.Credentials.from_service_account_info(json.loads(cred_json),
           scopes=["https://www.googleapis.com/auth/spreadsheets.readonly"])
    api = build("sheets", "v4", credentials=cred, cache_discovery=False)
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
    print("Lendo a planilha do Google...")
    dados = ler_google(planilha_id, cred)
    print("Abas lidas:", ", ".join(f"{a} ({len(v)} linhas)" for a, v in dados.items()))
    xlsx = os.path.join(PASTA, "Insanos.xlsx")
    montar_planilha(dados, xlsx)
    tok, dono, repo = preparar_ambiente()
    sys.argv = [sys.argv[0], "--publicar"]
    import painel as g
    g.PLANILHA = xlsx
    g.ARQ_TOKEN = tok
    g.GITHUB_USUARIO, g.GITHUB_REPO = dono, repo
    g.SITE_URL = f"https://{dono}.github.io/{repo}/"
    g.limpar_historico_github = lambda: True       # na nuvem o histórico fica (sem limpeza forçada)
    g.main()

if __name__ == "__main__":
    main()
