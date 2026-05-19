# lv_CSV.py — gera/substitui LV CICLO.csv no Google Drive
# Mantém o mesmo tratamento de dados do lv.py original.
#
# Origem:
#   Planilha: 19xV_P6KIoZB9U03yMcdRb2oF_Q7gVdaukjAvE4xOvl8
#   Aba: LV GERAL
#   Intervalo: A:Y
#
# Destino:
#   Google Drive pasta: 1weGikVXLxPdNeDNT0gLfjYViYXy6YHIV
#   Arquivo: LV CICLO.csv

import os
import re
import time
import random
import json
import pathlib
import csv
import io
import base64
from datetime import datetime
from typing import Optional

import pandas as pd
import gspread
from gspread.exceptions import APIError, WorksheetNotFound
from google.oauth2.service_account import Credentials as SACreds
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError
from googleapiclient.http import MediaIoBaseUpload

__VERSION__ = "lv_CSV.py v1 CSV Drive"

print(f">>> {__VERSION__} — caminho: {__file__}", flush=True)

# ====== FUSO ======
os.environ.setdefault("TZ", "America/Sao_Paulo")
try:
    import time as _t
    _t.tzset()
except Exception:
    pass

# ====== CONFIGURAÇÕES ======
ID_ORIGEM = "19xV_P6KIoZB9U03yMcdRb2oF_Q7gVdaukjAvE4xOvl8"
ABA_ORIGEM = "LV GERAL"
RANGE_ORIGEM = "A:Y"  # 25 colunas, inclui cabeçalho

DRIVE_FOLDER_ID = "1weGikVXLxPdNeDNT0gLfjYViYXy6YHIV"
CSV_NAME = "LV CICLO.csv"

CAM_CRED = "credenciais.json"

CSV_DELIMITER = ";"
CSV_ENCODING = "utf-8-sig"

MAX_RETRIES = 6
BASE_SLEEP = 1.1
RETRYABLE = {429, 500, 502, 503, 504}

SCOPES = [
    "https://www.googleapis.com/auth/spreadsheets.readonly",
    "https://www.googleapis.com/auth/drive",
]


# ====== LOG ======
def now_str():
    return datetime.now().strftime("%d/%m/%Y %H:%M:%S")


def log(msg):
    print(f"[{now_str()}] {msg}", flush=True)


def _status_from_apierror(e: APIError) -> Optional[int]:
    m = re.search(r"\[(\d+)\]", str(e))
    return int(m.group(1)) if m else None


def _status_from_httperror(e: HttpError) -> Optional[int]:
    try:
        return int(e.resp.status)
    except Exception:
        return None


# ====== CREDENCIAIS FLEXÍVEIS ======
def make_creds():
    env_json = os.environ.get("GOOGLE_CREDENTIALS")

    if env_json:
        # Aceita GOOGLE_CREDENTIALS como JSON puro
        try:
            return SACreds.from_service_account_info(json.loads(env_json), scopes=SCOPES)
        except Exception:
            pass

        # Aceita GOOGLE_CREDENTIALS como Base64 do JSON
        try:
            decoded = base64.b64decode(env_json).decode("utf-8")
            return SACreds.from_service_account_info(json.loads(decoded), scopes=SCOPES)
        except Exception as e:
            raise RuntimeError(
                "GOOGLE_CREDENTIALS inválido. Esperado JSON puro ou Base64 do JSON. "
                f"Erro: {e}"
            )

    env_path = os.environ.get("GOOGLE_APPLICATION_CREDENTIALS")

    if env_path and os.path.isfile(env_path):
        return SACreds.from_service_account_file(env_path, scopes=SCOPES)

    if os.path.isfile(CAM_CRED):
        return SACreds.from_service_account_file(CAM_CRED, scopes=SCOPES)

    script_dir = pathlib.Path(__file__).resolve().parent

    for p in (script_dir / CAM_CRED, pathlib.Path.cwd() / CAM_CRED):
        if p.is_file():
            return SACreds.from_service_account_file(str(p), scopes=SCOPES)

    raise FileNotFoundError(
        "Credenciais não encontradas. Use GOOGLE_CREDENTIALS, "
        "GOOGLE_APPLICATION_CREDENTIALS ou credenciais.json."
    )


# ====== RETRY GOOGLE SHEETS ======
def with_retry(fn, *args, max_retries=MAX_RETRIES, base_sleep=BASE_SLEEP, desc="", **kwargs):
    tent = 0

    while True:
        try:
            return fn(*args, **kwargs)

        except APIError as e:
            tent += 1
            code = _status_from_apierror(e)

            if tent >= max_retries or (code is not None and code not in RETRYABLE):
                log(f"❌ Falhou: {desc or fn.__name__} | {e}")
                raise

            sleep_s = min(
                60.0,
                (base_sleep * (2 ** (tent - 1))) + random.uniform(0, 0.75)
            )

            log(
                f"⚠️ {e} — retry {tent}/{max_retries - 1} "
                f"em {sleep_s:.1f}s — {desc or fn.__name__}"
            )

            time.sleep(sleep_s)


# ====== RETRY GOOGLE DRIVE ======
def api_retry(callable_execute, desc="", max_retries=MAX_RETRIES, base_sleep=BASE_SLEEP):
    tent = 0

    while True:
        try:
            return callable_execute().execute()

        except HttpError as e:
            tent += 1
            code = _status_from_httperror(e)

            if tent >= max_retries or (code is not None and code not in RETRYABLE):
                log(f"❌ Falhou: {desc} | {e}")
                raise

            sleep_s = min(
                60.0,
                (base_sleep * (2 ** (tent - 1))) + random.uniform(0, 0.75)
            )

            log(
                f"⚠️ {e} — retry {tent}/{max_retries - 1} "
                f"em {sleep_s:.1f}s — {desc}"
            )

            time.sleep(sleep_s)


# ====== DRIVE ======
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


# ====== CSV ======
def gerar_csv_bytes(df):
    output = io.StringIO(newline="")

    writer = csv.writer(
        output,
        delimiter=CSV_DELIMITER,
        lineterminator="\n",
        quoting=csv.QUOTE_MINIMAL,
    )

    values = df.values.tolist()

    for row in values:
        writer.writerow(row)

    return output.getvalue().encode(CSV_ENCODING)


# ====== INÍCIO ======
log("🟢 INÍCIO LV CICLO CSV")
t0_total = time.time()

# Autenticação
log("🔐 Autenticando…")
cred = make_creds()

gc = gspread.authorize(cred)
drive_service = build("drive", "v3", credentials=cred, cache_discovery=False)

# Abertura origem
log("📂 Abrindo planilha origem…")
book_src = with_retry(gc.open_by_key, ID_ORIGEM, desc="open_by_key origem")

try:
    ws_src = with_retry(book_src.worksheet, ABA_ORIGEM, desc="abrir ws origem")
except WorksheetNotFound:
    log("❌ Aba de origem não encontrada.")
    raise

# Leitura origem
log(f"📥 Lendo dados da origem ({ABA_ORIGEM}!{RANGE_ORIGEM})…")
dados = with_retry(ws_src.get, RANGE_ORIGEM, desc=f"get {ABA_ORIGEM}!{RANGE_ORIGEM}")

# Força DF como object, igual ao código original
df = pd.DataFrame(dados, dtype=object)
df = df.astype(object)

log(f"🔎 Linhas lidas, incluindo cabeçalho: {len(df)}")

# Garante 25 colunas A:Y
if df.shape[1] < 25:
    add = 25 - df.shape[1]
    log(f"➕ Normalizando colunas: adicionando {add} colunas vazias até Y")

    for _ in range(add):
        df[df.shape[1]] = ""

# Se vier com mais de 25 colunas por algum motivo, limita em A:Y
if df.shape[1] > 25:
    log(f"✂️ Limitando colunas para A:Y. Colunas atuais: {df.shape[1]}")
    df = df.iloc[:, :25]

df = df.astype(object)

# ====== TRATAMENTOS IGUAIS AO CÓDIGO ORIGINAL ======
log("🧽 Tratando colunas numéricas e data…")

# Colunas numéricas:
# F, K, T, V, W
# Índices 0-based: 5, 10, 19, 21, 22
num_cols = [5, 10, 19, 21, 22]

# Coluna de data:
# H
# Índice 0-based: 7
date_col = 7

# Números, a partir da linha 2 para preservar cabeçalho
for c in num_cols:
    if c < df.shape[1]:
        s = (
            df.iloc[1:, c]
            .astype(str)
            .str.replace("’", "", regex=False)
            .str.replace("‘", "", regex=False)
            .str.replace("'", "", regex=False)
            .str.replace(r"[^\d,.\-]", "", regex=True)
            .str.replace(".", "", regex=False)
            .str.replace(",", ".", regex=False)
        )

        df.iloc[1:, c] = pd.to_numeric(s, errors="coerce")

# Data, a partir da linha 2 para preservar cabeçalho
if date_col < df.shape[1]:
    serie = (
        df.iloc[1:, date_col]
        .astype(str)
        .str.replace("’", "", regex=False)
        .str.replace("‘", "", regex=False)
        .str.replace("'", "", regex=False)
        .str.replace(r"[^\d/:\-]", "", regex=True)
    )

    dt = pd.to_datetime(serie, dayfirst=True, errors="coerce")
    df.iloc[1:, date_col] = dt.dt.strftime("%d/%m/%Y")

# Troca NaN/NaT por vazio
df = df.where(pd.notnull(df), "")

n_rows, n_cols = df.shape

log(f"📏 Tamanho final do CSV: {n_rows} linhas × {n_cols} colunas")

# ====== GERAR CSV EM MEMÓRIA ======
log(f"📄 Gerando CSV em memória: {CSV_NAME}")
csv_bytes = gerar_csv_bytes(df)

log(
    f"📄 CSV gerado: {CSV_NAME} | "
    f"Linhas: {n_rows} | "
    f"Colunas: {n_cols} | "
    f"Atualizado em {now_str()}"
)

# ====== SALVAR / SUBSTITUIR NO DRIVE ======
log(f"☁️ Salvando/substituindo {CSV_NAME} no Drive…")

arquivo = salvar_ou_substituir_csv_drive(
    drive_service=drive_service,
    folder_id=DRIVE_FOLDER_ID,
    file_name=CSV_NAME,
    csv_bytes=csv_bytes,
)

log(
    f"🎉 LV CICLO CSV concluído em {time.time() - t0_total:.1f}s | "
    f"Arquivo: {arquivo.get('name')} | "
    f"ID: {arquivo.get('id')} | "
    f"Modificado em: {arquivo.get('modifiedTime')}"
)
