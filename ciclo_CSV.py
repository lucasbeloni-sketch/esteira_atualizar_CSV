# ciclo_CSV.py — v9 CSV Drive com DATA FATURAMENTO corrigida — 2026-04-30 BRT
# Lê OBRAS GERAL!A1:T, normaliza dados e salva/substitui CICLO.csv no Google Drive.
# Também lê BD_Config!A:B para preencher a coluna B do CSV.
#
# No CSV:
# A = vazio
# B = Unidade tratada => PROCX(D:D; BD_Config!A:A; BD_Config!B:B; "-")
# C = Tipo => Fora carteira [valor da coluna H]
# D até W = dados da origem
#
# Ajuste importante:
# A coluna O do CSV = coluna L da origem.
# Coluna L da origem agora é tratada como DATA FATURAMENTO, não como número.

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

__VERSION__ = "ciclo_CSV.py v9 CSV Drive DATA FATURAMENTO corrigida"

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

# Planilha onde está a aba BD_Config
ID_CONFIG = "1gDktQhF0WIjfAX76J2yxQqEeeBsSfMUPGs5svbf9xGM"
ABA_CONFIG = "BD_Config"
INTERVALO_CONFIG = "A:B"

DRIVE_FOLDER_ID = "1weGikVXLxPdNeDNT0gLfjYViYXy6YHIV"
CSV_NAME = "CICLO.csv"

SRC_WIDTH = 20

# No CSV final:
# A = vazio
# B = Unidade tratada
# C = Tipo
# D:W = dados da origem
CSV_PREFIX_HEADER = ["", "Unidade tratada", "Tipo"]

# Como os dados da origem começam na coluna D do CSV:
# D = origem A
# E = origem B
# F = origem C
# G = origem D
# H = origem E
IDX_ORIGEM_PARA_COLUNA_D_CSV = 0
IDX_ORIGEM_PARA_COLUNA_H_CSV = 4

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


def normalizar_chave(valor):
    if valor is None:
        return ""

    s = str(valor).strip()

    if s.endswith(".0"):
        try:
            return str(int(float(s)))
        except Exception:
            return s

    return s


def normalizar_data(txt):
    if not txt:
        return ""

    s = str(txt).strip().lstrip("'").strip()

    # Trata datas zeradas do Google Sheets/Excel como vazio
    if s in ("30/12/1899", "31/12/1899", "00/01/1900"):
        return ""

    # Formato 2026-03-12 -> 12/03/2026
    m = re.match(r"^(\d{4})-(\d{2})-(\d{2})$", s)
    if m:
        data = f"{m.group(3)}/{m.group(2)}/{m.group(1)}"
        return "" if data in ("30/12/1899", "31/12/1899", "00/01/1900") else data

    # Formato 12/03/2026
    if re.match(r"^\d{2}/\d{2}/\d{4}$", s):
        return s

    # Formato 12/03/26 -> 12/03/2026
    m = re.match(r"^(\d{2})/(\d{2})/(\d{2})$", s)
    if m:
        return f"{m.group(1)}/{m.group(2)}/20{m.group(3)}"

    return s


def tratar_linhas(linhas):
    for r in linhas:
        # Colunas numéricas na origem:
        # índice 10 = K
        # índice 15 = P
        #
        # IMPORTANTE:
        # Removido o índice 11 daqui, pois índice 11 = coluna L da origem.
        # A coluna L da origem vira a coluna O no CSV e representa DATA FATURAMENTO.
        for idx in (10, 15):
            if idx < len(r):
                bruto = str(r[idx]).replace("R$", "").replace(".", "").replace(",", ".")
                bruto = re.sub(r"[^\d.\-]", "", bruto)

                try:
                    r[idx] = float(bruto) if bruto not in ("", ".", "-") else ""
                except Exception:
                    r[idx] = ""

        # Colunas de data na origem:
        # índice 9  = J
        # índice 11 = L -> DATA FATURAMENTO, coluna O no CSV
        # índice 12 = M
        # índice 14 = O
        for idx in (9, 11, 12, 14):
            if idx < len(r):
                r[idx] = normalizar_data(r[idx])

    return linhas


def pad_row(row, width=SRC_WIDTH):
    row = list(row) if row else []
    row = row[:width]
    row += [""] * (width - len(row))
    return row


def montar_mapa_bd_config(linhas_config):
    """
    Simula o comportamento do PROCX:
    - procura na coluna A da BD_Config
    - retorna a coluna B da BD_Config
    - se não encontrar, retorna "-"
    - se houver chave repetida, mantém a primeira ocorrência
    """
    mapa = {}

    for linha in linhas_config:
        if not linha:
            continue

        chave = normalizar_chave(linha[0]) if len(linha) >= 1 else ""
        valor = linha[1] if len(linha) >= 2 else ""

        if chave and chave not in mapa:
            mapa[chave] = valor

    return mapa


def obter_valor_coluna_d_csv(linha_origem):
    """
    Como no CSV final os dados da origem começam na coluna D:
    D = origem A

    Portanto, para preencher a coluna B com base na coluna D do CSV,
    usamos o índice 0 da linha de origem.
    """
    if IDX_ORIGEM_PARA_COLUNA_D_CSV < len(linha_origem):
        return normalizar_chave(linha_origem[IDX_ORIGEM_PARA_COLUNA_D_CSV])

    return ""


def gerar_unidade_tratada(linha_origem, mapa_config):
    valor_d = obter_valor_coluna_d_csv(linha_origem)

    if not valor_d:
        return ""

    return mapa_config.get(valor_d, "-")


def obter_valor_coluna_h_csv(linha_origem):
    """
    Como no CSV final os dados da origem começam na coluna D:
    H = origem E

    Portanto, para preencher a coluna C com base na coluna H do CSV,
    usamos o índice 4 da linha de origem.
    """
    if IDX_ORIGEM_PARA_COLUNA_H_CSV < len(linha_origem):
        return str(linha_origem[IDX_ORIGEM_PARA_COLUNA_H_CSV]).strip()

    return ""


def gerar_tipo(linha_origem):
    valor_h = obter_valor_coluna_h_csv(linha_origem)

    if valor_h:
        return f"Fora carteira [{valor_h}]"

    return ""


def gerar_csv_bytes(hdr, linhas, mapa_config):
    output = io.StringIO(newline="")

    writer = csv.writer(
        output,
        delimiter=CSV_DELIMITER,
        lineterminator="\n",
        quoting=csv.QUOTE_MINIMAL,
    )

    # Cabeçalho:
    # A vazio | B Unidade tratada | C Tipo | D:W cabeçalho da origem
    writer.writerow(CSV_PREFIX_HEADER + pad_row(hdr))

    # Dados:
    # A vazio | B calculado pelo BD_Config | C calculada pela coluna H | D:W dados da origem
    for linha in linhas:
        unidade_tratada = gerar_unidade_tratada(linha, mapa_config)
        tipo = gerar_tipo(linha)

        prefixo_dados = [
            "",
            unidade_tratada,
            tipo,
        ]

        writer.writerow(prefixo_dados + pad_row(linha))

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
    print("⚠️ Origem sem dados. Será gerado CSV somente com cabeçalho.", flush=True)
else:
    hdr = dados[0]
    linhas = dados[1:]

# =========================
# LEITURA BD_CONFIG
# =========================
b_config = gs_retry(gc.open_by_key, ID_CONFIG, desc="open config")
w_config = gs_retry(b_config.worksheet, ABA_CONFIG, desc="ws BD_Config")

dados_config = gs_retry(
    w_config.get,
    INTERVALO_CONFIG,
    desc=f"get {ABA_CONFIG}!{INTERVALO_CONFIG}",
)

mapa_config = montar_mapa_bd_config(dados_config)

print(
    f"🔎 BD_Config carregada: {len(mapa_config)} chaves encontradas para preenchimento da coluna B.",
    flush=True,
)

# =========================
# NORMALIZAÇÕES
# =========================
linhas = tratar_linhas(linhas)

# =========================
# GERAR CSV
# =========================
csv_bytes = gerar_csv_bytes(hdr, linhas, mapa_config)

total_linhas_csv = len(linhas) + 1

print(
    f"📄 CSV gerado em memória: {CSV_NAME} | "
    f"Linhas: {total_linhas_csv} | "
    f"Coluna B via BD_Config | "
    f"Coluna C calculada por: Fora carteira [Coluna H] | "
    f"DATA FATURAMENTO corrigida na coluna O | "
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
