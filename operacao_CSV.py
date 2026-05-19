# operacao_CSV.py — gera/substitui OPERACAO.csv no Google Drive
# Mantém o mesmo tratamento do código original:
# - Lê Quadro Geral!B17:M
# - Trata coluna D do CSV como número
# - Trata coluna E do CSV como data
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

from datetime import datetime, date
from typing import Optional

from google.oauth2.service_account import Credentials
from gspread.exceptions import APIError, WorksheetNotFound
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError
from googleapiclient.http import MediaIoBaseUpload


__VERSION__ = "operacao_CSV.py v1 CSV Drive"

print(f">>> {__VERSION__} — caminho: {__file__}", flush=True)


# ================== FUSO ==================
os.environ.setdefault("TZ", "America/Sao_Paulo")
try:
    import time as _t
    _t.tzset()
except Exception:
    pass


# ================== CONFIG =================
ID_ORIGEM = "18-AoLupeaUIOdkW89o6SLK6Z9d8X0dKXgdjft_daMBk"
ABA_ORIGEM = "Quadro Geral"
RANGE_ORIGEM = "B17:M"  # 12 colunas

DRIVE_FOLDER_ID = "1weGikVXLxPdNeDNT0gLfjYViYXy6YHIV"
CSV_NAME = "OPERACAO.csv"

CAM_CRED = "credenciais.json"

CSV_DELIMITER = ";"
CSV_ENCODING = "utf-8-sig"

SRC_WIDTH = 12

MAX_RETRIES = 6
BASE_SLEEP = 1.0
TRANSIENT_CODES = {429, 500, 502, 503, 504}

SCOPES = [
    "https://www.googleapis.com/auth/spreadsheets.readonly",
    "https://www.googleapis.com/auth/drive",
]


# ================== LOG ====================
def now():
    return datetime.now().strftime("%d/%m/%Y %H:%M:%S")


def log(msg):
    print(f"[{now()}] {msg}", flush=True)


# ================== CREDENCIAIS ==================
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

    for p in (script_dir / CAM_CRED, pathlib.Path.cwd() / CAM_CRED):
        if p.is_file():
            return Credentials.from_service_account_file(str(p), scopes=SCOPES)

    raise FileNotFoundError(
        "Credenciais não encontradas. Use GOOGLE_CREDENTIALS, "
        "GOOGLE_APPLICATION_CREDENTIALS ou credenciais.json."
    )


# ================== RETRY ==================
def _status_code_api(e: APIError) -> Optional[int]:
    m = re.search(r"\[(\d+)\]", str(e))
    return int(m.group(1)) if m else None


def _status_code_http(e: HttpError) -> Optional[int]:
    try:
        return int(e.resp.status)
    except Exception:
        return None


def with_retries(fn, *args, retries=MAX_RETRIES, base_sleep=BASE_SLEEP, desc="", **kwargs):
    tent = 0

    while True:
        try:
            return fn(*args, **kwargs)

        except APIError as e:
            tent += 1
            code = _status_code_api(e)

            if tent >= retries or (code is not None and code not in TRANSIENT_CODES):
                log(f"❌ Falhou: {desc or fn.__name__} | {e}")
                raise

            slp = min(60, base_sleep * (2 ** (tent - 1)) + random.uniform(0, 0.75))

            log(
                f"⚠️ HTTP {code} — retry {tent}/{retries - 1} "
                f"em {slp:.1f}s — passo: {desc or fn.__name__}"
            )

            time.sleep(slp)


def api_retry(callable_execute, desc="", retries=MAX_RETRIES, base_sleep=BASE_SLEEP):
    tent = 0

    while True:
        try:
            return callable_execute().execute()

        except HttpError as e:
            tent += 1
            code = _status_code_http(e)

            if tent >= retries or (code is not None and code not in TRANSIENT_CODES):
                log(f"❌ Falhou: {desc} | {e}")
                raise

            slp = min(60, base_sleep * (2 ** (tent - 1)) + random.uniform(0, 0.75))

            log(
                f"⚠️ HTTP {code} — retry {tent}/{retries - 1} "
                f"em {slp:.1f}s — passo: {desc}"
            )

            time.sleep(slp)


# ================== HELPERS ==================
def pad_row(row, width=SRC_WIDTH):
    row = list(row) if row else []
    row = row[:width]
    row += [""] * (width - len(row))
    return row


def is_empty_value(v):
    if v is None:
        return True

    try:
        if pd.isna(v):
            return True
    except Exception:
        pass

    return str(v).strip() == ""


def limpar_numero(valor):
    """
    Mantém a lógica do código original para a coluna D:
    - Remove aspas especiais
    - Remove caracteres que não são número, vírgula, ponto ou sinal
    - Converte para float
    """
    if is_empty_value(valor):
        return ""

    bruto = str(valor).strip()
    bruto = bruto.replace("’", "").replace("‘", "").replace("'", "")
    bruto = re.sub(r"[^\d.,-]", "", bruto)

    if bruto in ("", ".", "-", ","):
        return ""

    if "," in bruto and "." in bruto:
        bruto = bruto.replace(".", "").replace(",", ".")
    elif "," in bruto:
        bruto = bruto.replace(",", ".")

    try:
        return float(bruto)
    except Exception:
        return ""


def formatar_numero_ptbr(valor):
    """
    Como CSV não possui formatação de célula,
    grava o número já no padrão PT-BR.

    Exemplo:
    1926.25 -> 1926,25
    """
    if is_empty_value(valor):
        return ""

    try:
        num = float(valor)
        return f"{num:.2f}".replace(".", ",")
    except Exception:
        return str(valor).replace(".", ",")


def limpar_data(valor):
    """
    Mantém a lógica do código original para a coluna E:
    aceita:
    - dd/mm/yyyy
    - dd/mm/yy
    - yyyy-mm-dd

    No CSV, grava como dd/mm/yyyy para ficar legível no padrão BR.
    """
    if is_empty_value(valor):
        return ""

    s = str(valor).strip()
    s = s.replace("’", "").replace("‘", "").replace("'", "")
    s = re.sub(r"[^\d/:-]", "", s)

    if not s:
        return ""

    for fmt in ("%d/%m/%Y", "%d/%m/%y", "%Y-%m-%d"):
        try:
            d = datetime.strptime(s, fmt)
            return d.strftime("%d/%m/%Y")
        except Exception:
            pass

    return ""


def normalize_cell(v):
    if is_empty_value(v):
        return ""

    if isinstance(v, (pd.Timestamp, datetime, date)):
        return v.strftime("%d/%m/%Y")

    return v


def preparar_dados_csv(dados):
    """
    Origem: Quadro Geral!B17:M
    Destino CSV:
    A:L = dados da origem B:M

    Tratamentos:
    - Coluna D do CSV, índice 3, tratada como número
    - Coluna E do CSV, índice 4, tratada como data
    - Primeira linha é considerada cabeçalho e não passa pelos tratamentos
    """
    linhas = [pad_row(linha, SRC_WIDTH) for linha in dados]

    if not linhas:
        return []

    log("🧽 Tratando colunas D número e E data — ignorando cabeçalho…")

    for i in range(1, len(linhas)):
        linha = linhas[i]

        # Coluna D do CSV
        linha[3] = formatar_numero_ptbr(limpar_numero(linha[3]))

        # Coluna E do CSV
        linha[4] = limpar_data(linha[4])

    linhas_normalizadas = []

    for linha in linhas:
        linhas_normalizadas.append([normalize_cell(c) for c in linha])

    return linhas_normalizadas


# ================== CSV ==================
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


# ================== GOOGLE DRIVE ==================
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
t0 = time.time()

log("🚀 Iniciando OPERACAO CSV")

# ---- Autenticação
log("🔐 Autenticando…")
cred = make_creds()

gc = gspread.authorize(cred)
drive_service = build("drive", "v3", credentials=cred, cache_discovery=False)

# ---- Abertura origem
log("📂 Abrindo planilha origem…")

plan_origem = with_retries(
    gc.open_by_key,
    ID_ORIGEM,
    desc="open_by_key origem",
)

try:
    aba_origem = with_retries(
        plan_origem.worksheet,
        ABA_ORIGEM,
        desc="worksheet origem",
    )
except WorksheetNotFound:
    log("❌ Aba de origem não encontrada.")
    raise

# ---- Leitura origem
log(f"📥 Lendo origem ({ABA_ORIGEM}!{RANGE_ORIGEM})…")

dados = with_retries(
    aba_origem.get,
    RANGE_ORIGEM,
    desc=f"get {ABA_ORIGEM}!{RANGE_ORIGEM}",
)

log(f"🔎 Linhas lidas, incluindo cabeçalho: {len(dados)}")

if not dados:
    log("⚠️ Origem vazia. Será gerado CSV vazio.")
    linhas_csv = []
else:
    linhas_csv = preparar_dados_csv(dados)

log(f"📏 Tamanho final do CSV: {len(linhas_csv)} linhas × {SRC_WIDTH} colunas")

# ---- Gerar CSV
log(f"📄 Gerando CSV em memória: {CSV_NAME}")

csv_bytes = gerar_csv_bytes(linhas_csv)

log(
    f"📄 CSV gerado: {CSV_NAME} | "
    f"Linhas: {len(linhas_csv)} | "
    f"Atualizado em {now()}"
)

# ---- Salvar/substituir no Drive
log(f"☁️ Salvando/substituindo {CSV_NAME} no Drive…")

arquivo = salvar_ou_substituir_csv_drive(
    drive_service=drive_service,
    folder_id=DRIVE_FOLDER_ID,
    file_name=CSV_NAME,
    csv_bytes=csv_bytes,
)

log(
    f"🎉 OPERACAO CSV concluído em {time.time() - t0:.1f}s | "
    f"Arquivo: {arquivo.get('name')} | "
    f"ID: {arquivo.get('id')} | "
    f"Modificado em: {arquivo.get('modifiedTime')}"
)
