# app/gui/wbs_helpers.py
import re, unicodedata
import pandas as pd
from collections import OrderedDict

def _casefold(s: str) -> str:
    if not isinstance(s, str):
        return ""
    s = unicodedata.normalize("NFKD", s)
    s = "".join(ch for ch in s if not unicodedata.combining(ch))
    s = re.sub(r"\s+", " ", s.strip().lower())
    return s

def normalize(text):
    if not isinstance(text, str):
        return ""
    return unicodedata.normalize('NFKD', text).encode('ASCII', 'ignore').decode().lower().strip()

def find_wbs_columns(df: pd.DataFrame):
    col_wbs = col_desc = col_nivel = None

    for col in df.columns:
        norm = normalize(col)
        if norm == "wbs":
            col_wbs = col
        elif norm in {"descricao", "desc"}:
            col_desc = col
        elif norm == "nivel":
            col_nivel = col

    return col_wbs, col_desc, col_nivel


def split_levels(df: pd.DataFrame, col_nivel: str):
    def to_int(v):
        try:
            if pd.isna(v): return None
            return int(float(str(v).strip().replace(",", ".")))
        except Exception:
            return None
    return df[col_nivel].apply(to_int)

def children_at_level(df, col_nivel, col_wbs, col_desc, prefix, level):
    if level == 1 or prefix is None:
        sub = df[df[col_nivel] == 1][[col_wbs, col_desc]].dropna(subset=[col_wbs])
    else:
        pfx = str(prefix) + "."
        sub = df[(df[col_nivel] == level) & (df[col_wbs].astype(str).str.startswith(pfx))][[col_wbs, col_desc]]
    return sub.sort_values(by=col_wbs)

def branch_has_children(df, col_nivel, col_wbs, parent_code, child_level):
    if not parent_code:
        return False
    pfx = str(parent_code) + "."
    sub = df[(df[col_nivel] == child_level) & (df[col_wbs].astype(str).str.startswith(pfx))]
    return not sub.empty

def ensure_level10_row(df, col_nivel, col_wbs, col_desc, leaf_code):
    mask_leaf = (df[col_wbs].astype(str) == str(leaf_code)) & (df[col_nivel].fillna(0) < 10)
    leaf_idx = df.index[mask_leaf]
    if len(leaf_idx) == 0:
        raise ValueError(f"Código WBS não encontrado: {leaf_code}")
    i = leaf_idx[0]
    j = i + 1
    while j in df.index:
        nivel = df.at[j, col_nivel]
        wbs   = df.at[j, col_wbs]
        if pd.notna(nivel) and int(float(nivel)) == 10:
            return j
        if isinstance(wbs, str) and wbs.strip():
            break
        j += 1
    new_row = {c: None for c in df.columns}
    new_row[col_nivel] = 10
    new_row[col_wbs]   = None
    new_row[col_desc]  = ""
    upper = df.loc[:i]
    lower = df.loc[i+1:]
    new_df = pd.concat([upper, pd.DataFrame([new_row]), lower], ignore_index=True)
    df.drop(df.index, inplace=True)
    for c in new_df.columns:
        df[c] = new_df[c]
    return i + 1

def find_level10_text(df, col_nivel, col_wbs, col_desc, leaf_code):
    mask_leaf = (df[col_wbs].astype(str) == str(leaf_code)) & (df[col_nivel].fillna(0) < 10)
    leaf_idx = df.index[mask_leaf]
    if len(leaf_idx) == 0:
        return ""
    i = leaf_idx[0]
    j = i + 1
    while j in df.index:
        nivel = df.at[j, col_nivel]
        wbs   = df.at[j, col_wbs]
        if pd.notna(nivel) and int(float(nivel)) == 10:
            val = df.at[j, col_desc]
            return "" if pd.isna(val) else str(val)
        if isinstance(wbs, str) and wbs.strip():
            break
        j += 1
    return ""

def list_ancestors(wbs_code: str):
    parts = [p for p in str(wbs_code).split(".") if p]
    return [".".join(parts[:k]) for k in range(1, len(parts))]
    
def detect_relevant_leaves(df, col_nivel, col_wbs, col_desc, baseline_series):
    """
    Devolve uma lista de tuplos (wbs_code, leaf_desc) para cada 'folha' cujo nível 10
    logo abaixo tem descrição DIFERENTE do baseline (ou seja, o utilizador editou).
    Mantém a ordem de aparecimento no ficheiro e não duplica códigos.
    """
    if df is None or col_nivel is None: 
        return []

    niv = split_levels(df, col_nivel)

    def norm(s):
        if pd.isna(s): return ""
        return str(s).strip()

    baseline = baseline_series if baseline_series is not None else df[col_desc].copy()

    last_code = None
    last_desc = ""
    seen = OrderedDict()

    for i in range(len(df)):
        lvl = niv.iat[i]
        code = df.at[i, col_wbs] if col_wbs in df.columns else None

        if lvl is not None and lvl < 10 and isinstance(code, str) and code.strip():
            last_code = str(code).strip()
            leaf_desc = df.at[i, col_desc] if col_desc in df.columns else ""
            last_desc = "" if pd.isna(leaf_desc) else str(leaf_desc)
            continue

        if lvl == 10:
            new_txt = norm(df.at[i, col_desc])
            old_txt = norm(baseline.iat[i])
            if new_txt and (new_txt != old_txt) and last_code:
                seen[last_code] = last_desc

    return list(seen.items())