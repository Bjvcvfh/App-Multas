import os
import re
import sys
from pathlib import Path
from datetime import datetime

def resource_path(relative_path: str) -> str:
    base = getattr(sys, "_MEIPASS", os.path.abspath(os.path.dirname(__file__)))
    # __file__ aqui é ...\utils\helpers.py, então precisamos subir 1 nível no modo .py
    if not getattr(sys, "_MEIPASS", None):
        base = os.path.abspath(os.path.join(os.path.dirname(__file__), ".."))
    return os.path.join(base, relative_path)

def ensure_dirs(paths):
    for p in paths:
        os.makedirs(p, exist_ok=True)

def sanitize_filename(name: str) -> str:
    name = re.sub(r"[^\w\-. ]", "", name, flags=re.UNICODE).strip()
    name = name.replace(" ", "_")
    return name[:120] if len(name) > 120 else name

def parse_money_to_float(x) -> float:
    try:
        import pandas as pd
        if pd.isna(x):
            return 0.0
    except Exception:
        pass

    s = str(x).strip().replace("R$", "").replace(" ", "")
    s = re.sub(r"[^0-9,\.\-]", "", s)
    if s.count(",") > 0 and s.count(".") > 0:
        s = s.replace(".", "").replace(",", ".")
    elif s.count(",") > 0 and s.count(".") == 0:
        s = s.replace(",", ".")
    try:
        return float(s)
    except:
        return 0.0

def format_brl(v: float) -> str:
    s = f"{v:,.2f}"
    s = s.replace(",", "X").replace(".", ",").replace("X", ".")
    return f"R$ {s}"

def data_por_extenso_ptbr(dt: datetime) -> str:
    meses = [
        "janeiro", "fevereiro", "março", "abril", "maio", "junho",
        "julho", "agosto", "setembro", "outubro", "novembro", "dezembro"
    ]
    return f"{dt.day} de {meses[dt.month - 1].capitalize()} de {dt.year}"

def get_persistent_app_dir(app_name="AppMultas"):
    """
    Retorna diretório persistente em AppData\\Local
    Ex:
    C:\\Users\\Renan\\AppData\\Local\\TransVarzea\\AppMultas
    """
    base = (
        Path(os.environ.get("LOCALAPPDATA", Path.home() / "AppData" / "Local"))
        / app_name
    )
    base.mkdir(parents=True, exist_ok=True)
    return base