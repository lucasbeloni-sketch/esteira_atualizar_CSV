# med_parcial_CSV.py — gera/substitui MED PARCIAL.csv no Google Drive
# Mantém o mesmo tratamento do código original:
# - Lê MED PARCIAIS GERAL!A1:P
# - Trata valores numéricos nas colunas F e J da origem
# - Cria a coluna A "PROJETO CORRIGIDO" com os 9 primeiros caracteres da coluna B da origem
# - Gera arquivo CSV no Google Drive

import os
import re
import time
import random
import json
import pathlib
import csv
import io
import base64
import pandas as pd
import gspread

from datetime import datetime
from typing import Optional

from google.oauth2.service_account import Credentials
from gspread.exceptions import APIError, WorksheetNotFound
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError
from googleapiclient.http import MediaIoBaseUpload


__VERSION__ = "med_parcial_CSV.py v1 CSV Drive"

print(f">>> {__VERSION__} — caminho: {__file__}", flush=True)


# ================== FUSO ==================
os.environ.setdefault("TZ", "America/Sao_Paulo")
try:
    import time as _t
    _t.tzset()
except Exception:
    pass


# ================== CONFIG =================
ID_PLANILHA_ORIGEM = "19xV_P6KIoZB9U03yMcdRb2oF_Q7gVdaukjAvE4xOvl8"
ABA_ORIGEM = "MED PARCIAIS GERAL"
RANGE_ORIGEM = "A1:P"

DRIVE_FOLDER_ID = "1weGikVXLxPdNeDNT0gLfjYViYXy6YHIV"
CSV_NAME = "MED PARCIAL.csv"

CAMINHO_CREDENCIAIS = "credenciais.json"

CSV_DELIMITER = ";"
CSV_ENCODING = "utf-8-sig"

MAX_RETRIES = 6
BASE_SLEEP = 1.0
TRANSIENT = {429, 500, 502, 503, 504}

SCOPES = [
    "https://www.googleapis.com/auth/spreadsheets.readonly",
    "https://www.googleapis.com/auth/drive",
]


# ================== LOG ====================
def now():
    return datetime.now().strftime("%H:%M:%S")


def now_full():
    return datetime.now().strftime("%d/%m/%Y %H:%M:%S")


def log(msg):
    print(f"[{now()}] {msg}", flush=True)


# ================== CREDENCIAIS ====================
def make_creds():
    env_json = os.environ.get("GOOGLE_CREDENTIALS")

    if env_json:
        # Aceita GOOGLE_CREDENTIALS como JSON puro
        try:
            return Credentials.from_service_account_info(json.loads(env_json), scopes=SCOPES)
        except Exception:
            pass

        # Aceita GOOGLE_CREDENTIALS como Base64 do JSON
        try:
            decoded = base64.b64decode(env_json).decode("utf-8")
            return Credentials.from_service_account_info(json.loads(decoded), scopes=SCOPES)
        except Exception as e:
            raise RuntimeError(
                "GOOGLE_CREDENTIALS inválido. Esperado JSON puro ou Base64 do JSON. "
                f"Erro: {e}"
            )

    env_path = os.environ.get("GOOGLE_APPLICATION_CREDENTIALS")

    if env_path and os.path.isfile(env_path):
        return Credentials.from_service_account_file(env_path, scopes=SCOPES)

    script_dir = pathlib.Path(__file__).resolve().parent

    for p in (script_dir / CAMINHO_CREDENCIAIS, pathlib.Path.cwd() / CAMINHO_CREDENCIAIS):
        if p.is_file():
            return Credentials.from_service_account_file(str(p), scopes=SCOPES)

    raise FileNotFoundError(
        "Credenciais não encontradas. Use GOOGLE_CREDENTIALS, "
        "GOOGLE_APPLICATION_CREDENTIALS ou credenciais.json."
    )


# ================== RETRY ====================
def _status_code_api(e: APIError) -> Optional[int]:
    m = re.search(r"\[(\d+)\]", str(e))
    return int(m.group(1)) if m else None


def _status_code_http(e: HttpError) -> Optional[int]:
    try:
        return int(e.resp.status)
    except Exception:
        return None


def with_retry(func, *args, retries=MAX_RETRIES, base=BASE_SLEEP, desc="", **kwargs):
    tent = 0

    while True:
        try:
            return func(*args, **kwargs)

        except APIError as e:
            tent += 1
            code = _status_code_api(e)

            if tent >= retries or (code is not None and code not in TRANSIENT):
                log(f"❌ Falhou: {desc or func.__name__} | {e}")
                raise

            sleep_s = min(60, base * (2 ** (tent - 1)) + random.uniform(0, 0.75))

            log(
                f"⚠️ HTTP {code} — retry {tent}/{retries - 1} "
                f"em {sleep_s:.1f}s — passo: {desc or func.__name__}"
            )

            time.sleep(sleep_s)


def api_retry(callable_execute, desc="", retries=MAX_RETRIES, base=BASE_SLEEP):
    tent = 0

    while True:
        try:
            return callable_execute().execute()

        except HttpError as e:
            tent += 1
            code = _status_code_http(e)

            if tent >= retries or (code is not None and code not in TRANSIENT):
                log(f"❌ Falhou: {desc} | {e}")
                raise

            sleep_s = min(60, base * (2 ** (tent - 1)) + random.uniform(0, 0.75))

            log(
                f"⚠️ HTTP {code} — retry {tent}/{retries - 1} "
                f"em {sleep_s:.1f}s — passo: {desc}"
            )

            time.sleep(sleep_s)


# ================== TRATAMENTOS ====================
def limpar_valor(valor):
    if valor is None:
        return ""

    try:
        if pd.isna(valor):
            return ""
    except Exception:
        pass

    s = str(valor).strip()

    if not s:
        return ""

    s = re.sub(r"[^\d,.\-]", "", s)
    s = s.replace(".", "").replace(",", ".")

    try:
        return float(s)
    except Exception:
        return ""


def formatar_numero_ptbr(valor):
    """
    Para CSV em padrão PT-BR:
    1926.25 -> 1926,25
    """
    if valor is None:
        return ""

    try:
        if pd.isna(valor):
            return ""
    except Exception:
        pass

    s = str(valor).strip()

    if not s:
        return ""

    try:
        num = float(s)
        return f"{num:.2f}".replace(".", ",")
    except Exception:
        return s.replace(".", ",")


def pad_row(row, width):
    row = list(row) if row else []
    row = row[:width]
    row += [""] * (width - len(row))
    return row


def preparar_dados_csv(dados_origem):
    """
    Mantém a lógica original:
    - Cabeçalho da coluna A = PROJETO CORRIGIDO
    - Coluna A dos dados = 9 primeiros caracteres da coluna B da origem
    - Colunas B:P do CSV = colunas B:P da origem
    - Colunas F e J tratadas como número
    """
    cabecalho = pad_row(dados_origem[0], 16)
    dados = [pad_row(linha, 16) for linha in dados_origem[1:]]

    log(f"🔎 Linhas carregadas sem cabeçalho: {len(dados)}")

    log("🧽 Limpando valores numéricos das colunas F e J…")

    for linha in dados:
        # Coluna F da origem
        linha[5] = limpar_valor(linha[5])

        # Coluna J da origem
        linha[9] = limpar_valor(linha[9])

        # Para CSV PT-BR, grava com vírgula decimal
        linha[5] = formatar_numero_ptbr(linha[5])
        linha[9] = formatar_numero_ptbr(linha[9])

    log("🧮 Montando coluna A: PROJETO CORRIGIDO…")

    linhas_csv = []

    # Cabeçalho:
    # A = PROJETO CORRIGIDO
    # B:P = cabeçalho original B:P
    linhas_csv.append(["PROJETO CORRIGIDO"] + cabecalho[1:16])

    # Dados:
    # A = primeiros 9 caracteres da coluna B da origem
    # B:P = dados originais B:P
    for linha in dados:
        projeto_corrigido = str(linha[1] or "")[:9]
        linhas_csv.append([projeto_corrigido] + linha[1:16])

    return linhas_csv


# ================== CSV ====================
def gerar_csv_bytes(linhas_csv):
    output = io.StringIO(newline="")

    writer = csv.writer(
        output,
        delimiter=CSV_DELIMITER,
        lineterminator="\n",
        quoting=csv.QUOTE_MINIMAL,
    )

    for linha in linhas_csv:
        writer.writerow(linha)

    return output.getvalue().encode(CSV_ENCODING)


# ================== GOOGLE DRIVE ====================
def buscar_arquivo_drive(drive_service, folder_id, file_name):
    safe_name = file_name.replace("'", "\\'")

    query = (
        f"name = '{safe_name}' "
        f"and '{folder_id}' in parents "
        f"and trashed = false"
    )

    result = api_retry(
        lambda: drive_service.files().list(
            q=query,
            spaces="drive",
            fields="files(id, name, modifiedTime, webViewLink)",
            orderBy="modifiedTime desc",
            pageSize=10,
            supportsAllDrives=True,
            includeItemsFromAllDrives=True,
        ),
        desc=f"buscar arquivo {file_name} no Drive",
    )

    files = result.get("files", [])
    return files[0] if files else None


def salvar_ou_substituir_csv_drive(drive_service, folder_id, file_name, csv_bytes):
    media = MediaIoBaseUpload(
        io.BytesIO(csv_bytes),
        mimetype="text/csv",
        resumable=False,
    )

    arquivo_existente = buscar_arquivo_drive(drive_service, folder_id, file_name)

    if arquivo_existente:
        file_id = arquivo_existente["id"]

        atualizado = api_retry(
            lambda: drive_service.files().update(
                fileId=file_id,
                media_body=media,
                body={
                    "name": file_name,
                    "mimeType": "text/csv",
                },
                fields="id, name, modifiedTime, webViewLink",
                supportsAllDrives=True,
            ),
            desc=f"substituir {file_name}",
        )

        log(f"✅ Arquivo substituído: {atualizado.get('name')} | ID: {atualizado.get('id')}")
        return atualizado

    criado = api_retry(
        lambda: drive_service.files().create(
            body={
                "name": file_name,
                "parents": [folder_id],
                "mimeType": "text/csv",
            },
            media_body=media,
            fields="id, name, modifiedTime, webViewLink",
            supportsAllDrives=True,
        ),
        desc=f"criar {file_name}",
    )

    log(f"✅ Arquivo criado: {criado.get('name')} | ID: {criado.get('id')}")
    return criado


# ================== INÍCIO =================
inicio = time.time()

log("🚀 Iniciando MED PARCIAL CSV")

# ---- Autenticação
log("🔐 Autenticando no Google…")
cred = make_creds()

gc = gspread.authorize(cred)
drive_service = build("drive", "v3", credentials=cred, cache_discovery=False)

# ---- Abertura origem
log("📂 Abrindo planilha origem…")
planilha_origem = with_retry(
    gc.open_by_key,
    ID_PLANILHA_ORIGEM,
    desc="open_by_key origem",
)

try:
    aba_origem = with_retry(
        planilha_origem.worksheet,
        ABA_ORIGEM,
        desc="worksheet origem",
    )
except WorksheetNotFound:
    log("❌ Aba de origem não encontrada.")
    raise

# ---- Leitura origem
log(f"📥 Lendo dados da origem ({ABA_ORIGEM}!{RANGE_ORIGEM})…")

dados_origem = with_retry(
    aba_origem.get,
    RANGE_ORIGEM,
    desc=f"get {ABA_ORIGEM}!{RANGE_ORIGEM}",
)

if not dados_origem:
    log("⚠️ Sem dados na origem. Será gerado CSV vazio apenas com cabeçalho.")
    dados_origem = [["PROJETO CORRIGIDO"] + [""] * 15]

# ---- Prepara dados
linhas_csv = preparar_dados_csv(dados_origem)

log(f"📏 Tamanho final do CSV: {len(linhas_csv)} linhas × 16 colunas")

# ---- Gera CSV
log(f"📄 Gerando CSV em memória: {CSV_NAME}")
csv_bytes = gerar_csv_bytes(linhas_csv)

log(
    f"📄 CSV gerado: {CSV_NAME} | "
    f"Linhas: {len(linhas_csv)} | "
    f"Atualizado em {now_full()}"
)

# ---- Salva/substitui no Drive
log(f"☁️ Salvando/substituindo {CSV_NAME} no Drive…")

arquivo = salvar_ou_substituir_csv_drive(
    drive_service=drive_service,
    folder_id=DRIVE_FOLDER_ID,
    file_name=CSV_NAME,
    csv_bytes=csv_bytes,
)

log(
    f"🏁 Concluído em {time.time() - inicio:.1f}s — "
    f"MED PARCIAL CSV OK | "
    f"Arquivo: {arquivo.get('name')} | "
    f"ID: {arquivo.get('id')} | "
    f"Modificado em: {arquivo.get('modifiedTime')}"
)
