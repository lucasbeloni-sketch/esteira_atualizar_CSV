# ciclo_CSV.py — v4 CSV Drive com dados iniciando na coluna D — 2026-04-30 BRT
# Lê OBRAS GERAL!A1:T, normaliza dados e salva/substitui CICLO.csv no Google Drive.
# No CSV, cria 3 colunas vazias antes dos dados para simular início na coluna D.

from datetime import datetime
import os
import time
import re
import random
import json
import pathlib
import csv
import io
from typing import Optional

import gspread
from gspread.exceptions import APIError
from google.oauth2.service_account import Credentials as SACreds
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError
from googleapiclient.http import MediaIoBaseUpload

__VERSION__ = "ciclo_CSV.py v4 CSV Drive coluna D"

print(f">>> {__VERSION__} — caminho: {__file__}", flush=True)

# ========= FUSO =========
os.environ.setdefault("TZ", "America/Sao_Paulo")
try:
    import time as _t
    _t.tzset()
except Exception:
    pass

# =========================
# CONFIG
# =========================
ID_ORIGEM = "19xV_P6KIoZB9U03yMcdRb2oF_Q7gVdaukjAvE4xOvl8"
ABA_ORIGEM = "OBRAS GERAL"
INTERVALO_ORIGEM = "A1:T"  # 20 colunas

DRIVE_FOLDER_ID = "1weGikVXLxPdNeDNT0gLfjYViYXy6YHIV"
CSV_NAME = "CICLO.csv"

SRC_WIDTH = 20

# Para simular que os dados começam na coluna D do CSV
# A, B e C ficam vazias; dados entram de D até W.
CSV_COLUNAS_VAZIAS_ANTES = 3

# Separador compatível com Excel/Sheets em PT-BR
CSV_DELIMITER = ";"
CSV_ENCODING = "utf-8-sig"

# Credenciais
CREDENTIALS_PATH = "credenciais.json"
SCOPES = [
    "https://www.googleapis.com/auth/spreadsheets.readonly",
    "https://www.googleapis.com/auth/drive",
]

# Retry
MAX_RETRIES = 6
BASE_SLEEP = 1.0
RETRYABLE_CODES = {429, 500, 502, 503, 504}


# =========================
# UTILS
# =========================
def agora_str():
    return datetime.now().strftime("%d/%m/%Y %H:%M:%S")


def _status_from_apierror(e: APIError) -> Optional[int]:
    m = re.search(r"\[(\d+)\]", str(e))
    return int(m.group(1)) if m else None


def _status_from_httperror(e: HttpError) -> Optional[int]:
    try:
        return int(e.resp.status)
    except Exception:
        return None


def gs_retry(fn, *args, desc="", max_tries=MAX_RETRIES, base=BASE_SLEEP, **kw):
    tent = 0

    while True:
        try:
            return fn(*args, **kw)

        except APIError as e:
            tent += 1
            code = _status_from_apierror(e)

            if tent >= max_tries or (code is not None and code not in RETRYABLE_CODES):
                print(f"❌ {desc or fn.__name__}: {e}", flush=True)
                raise

            slp = min(30.0, base * (2 ** (tent - 1)) + random.uniform(0, 0.6))

            print(
                f"[retry] ⚠️ {desc or fn.__name__}: {e} — retry {tent}/{max_tries - 1} em {slp:.1f}s",
                flush=True,
            )

            time.sleep(slp)


def api_retry(callable_execute, desc="", max_tries=MAX_RETRIES, base=BASE_SLEEP):
    tent = 0

    while True:
        try:
            return callable_execute().execute()

        except HttpError as e:
            tent += 1
            code = _status_from_httperror(e)

            if tent >= max_tries or (code is not None and code not in RETRYABLE_CODES):
                print(f"❌ {desc}: {e}", flush=True)
                raise

            slp = min(30.0, base * (2 ** (tent - 1)) + random.uniform(0, 0.6))

            print(
                f"[retry] ⚠️ {desc}: {e} — retry {tent}/{max_tries - 1} em {slp:.1f}s",
                flush=True,
            )

            time.sleep(slp)


def make_creds():
    env_json = os.environ.get("GOOGLE_CREDENTIALS")

    if env_json:
        return SACreds.from_service_account_info(json.loads(env_json), scopes=SCOPES)

    env_path = os.environ.get("GOOGLE_APPLICATION_CREDENTIALS")

    if env_path and os.path.isfile(env_path):
        return SACreds.from_service_account_file(env_path, scopes=SCOPES)

    script_dir = pathlib.Path(__file__).resolve().parent

    for p in (script_dir / CREDENTIALS_PATH, pathlib.Path.cwd() / CREDENTIALS_PATH):
        if p.is_file():
            return SACreds.from_service_account_file(str(p), scopes=SCOPES)

    raise FileNotFoundError(
        "Credenciais não encontradas. Use GOOGLE_CREDENTIALS, "
        "GOOGLE_APPLICATION_CREDENTIALS ou credenciais.json."
    )


def normalizar_data(txt):
    if not txt:
        return ""

    s = str(txt).strip().lstrip("'").strip()

    m = re.match(r"^(\d{4})-(\d{2})-(\d{2})$", s)
    if m:
        return f"{m.group(3)}/{m.group(2)}/{m.group(1)}"

    if re.match(r"^\d{2}/\d{2}/\d{4}$", s):
        return s

    m = re.match(r"^(\d{2})/(\d{2})/(\d{2})$", s)
    if m:
        return f"{m.group(1)}/{m.group(2)}/20{m.group(3)}"

    return s


def tratar_linhas(linhas):
    for r in linhas:
        # Colunas numéricas na origem:
        # índice 10 = K, índice 11 = L, índice 15 = P
        for idx in (10, 11, 15):
            if idx < len(r):
                bruto = str(r[idx]).replace("R$", "").replace(".", "").replace(",", ".")
                bruto = re.sub(r"[^\d.\-]", "", bruto)

                try:
                    r[idx] = float(bruto) if bruto not in ("", ".", "-") else ""
                except Exception:
                    r[idx] = ""

        # Colunas de data na origem:
        # índice 9 = J, índice 12 = M, índice 14 = O
        for idx in (9, 12, 14):
            if idx < len(r):
                r[idx] = normalizar_data(r[idx])

    return linhas


def pad_row(row, width=SRC_WIDTH):
    row = list(row) if row else []
    row = row[:width]
    row += [""] * (width - len(row))
    return row


def aplicar_offset_coluna_d(row):
    """
    Adiciona 3 colunas vazias no início da linha.
    Assim, quando abrir o CSV em planilha:
    A, B e C ficam vazias;
    os dados começam na coluna D.
    """
    return [""] * CSV_COLUNAS_VAZIAS_ANTES + pad_row(row)


def gerar_csv_bytes(hdr, linhas):
    output = io.StringIO(newline="")

    writer = csv.writer(
        output,
        delimiter=CSV_DELIMITER,
        lineterminator="\n",
        quoting=csv.QUOTE_MINIMAL,
    )

    # Cabeçalho começando na coluna D
    writer.writerow(aplicar_offset_coluna_d(hdr))

    # Dados começando na coluna D
    for linha in linhas:
        writer.writerow(aplicar_offset_coluna_d(linha))

    return output.getvalue().encode(CSV_ENCODING)


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

        print(
            f"✅ Arquivo substituído: {atualizado.get('name')} | ID: {atualizado.get('id')}",
            flush=True,
        )

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

    print(
        f"✅ Arquivo criado: {criado.get('name')} | ID: {criado.get('id')}",
        flush=True,
    )

    return criado


# =========================
# AUTH
# =========================
creds = make_creds()

gc = gspread.authorize(creds)
drive_service = build("drive", "v3", credentials=creds, cache_discovery=False)

# =========================
# LEITURA ORIGEM
# =========================
b_src = gs_retry(gc.open_by_key, ID_ORIGEM, desc="open origem")
w_src = gs_retry(b_src.worksheet, ABA_ORIGEM, desc="ws origem")

dados = gs_retry(
    w_src.get,
    INTERVALO_ORIGEM,
    desc=f"get {ABA_ORIGEM}!{INTERVALO_ORIGEM}",
)

if not dados:
    hdr = []
    linhas = []
    print("⚠️ Origem sem dados. Será gerado um CSV vazio.", flush=True)
else:
    hdr = dados[0]
    linhas = dados[1:]

# =========================
# NORMALIZAÇÕES
# =========================
linhas = tratar_linhas(linhas)

# =========================
# GERAR CSV
# =========================
csv_bytes = gerar_csv_bytes(hdr, linhas)

total_linhas_csv = len(linhas) + 1

print(
    f"📄 CSV gerado em memória: {CSV_NAME} | "
    f"Linhas: {total_linhas_csv} | "
    f"Dados iniciando na coluna D | "
    f"Atualizado em {agora_str()}",
    flush=True,
)

# =========================
# SALVAR / SUBSTITUIR NO DRIVE
# =========================
arquivo = salvar_ou_substituir_csv_drive(
    drive_service=drive_service,
    folder_id=DRIVE_FOLDER_ID,
    file_name=CSV_NAME,
    csv_bytes=csv_bytes,
)

print(
    f"✅ CICLO.csv finalizado no Drive — ID: {arquivo.get('id')} — modificado em {arquivo.get('modifiedTime')}",
    flush=True,
)
