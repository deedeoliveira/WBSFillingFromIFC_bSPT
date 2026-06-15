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


_COLUMN_ALIASES = {
    "wbs":            "col_wbs",
    "nivel":          "col_nivel",
    "descricao":      "col_desc",
    "desc":           "col_desc",
    "unidades":       "col_unidades",
    "unid.":          "col_unidades",
    "unid":           "col_unidades",
    "ifc class":      "col_ifc_class",
    "ifcclass":       "col_ifc_class",
    "predefinedtype": "col_predef",
    "objecttype":     "col_objtype",
    "ifc property":   "col_ifc_prop",
    "ifcproperty":    "col_ifc_prop",
}


def find_wbs_columns(df: pd.DataFrame) -> dict:
    result = {k: None for k in set(_COLUMN_ALIASES.values())}
    for col in df.columns:
        norm = normalize(str(col))
        canonical = _COLUMN_ALIASES.get(norm)
        if canonical and result[canonical] is None:
            result[canonical] = col
    return result


def unpack_core_columns(cols: dict):
    return cols["col_wbs"], cols["col_desc"], cols["col_nivel"]


def split_levels(df: pd.DataFrame, col_nivel: str):
    def to_int(v):
        try:
            if pd.isna(v):
                return None
            return int(float(str(v).strip().replace(",", ".")))
        except Exception:
            return None
    return df[col_nivel].apply(to_int)


def children_at_level(df, col_nivel, col_wbs, col_desc, prefix, level):
    if level == 1 or prefix is None:
        sub = df[df[col_nivel] == 1][[col_wbs, col_desc]].dropna(subset=[col_wbs])
    else:
        pfx = str(prefix) + "."
        sub = df[
            (df[col_nivel] == level) &
            (df[col_wbs].astype(str).str.startswith(pfx))
        ][[col_wbs, col_desc]]
    return sub.sort_values(by=col_wbs)


def branch_has_children(df, col_nivel, col_wbs, parent_code, child_level):
    if not parent_code:
        return False
    pfx = str(parent_code) + "."
    sub = df[
        (df[col_nivel] == child_level) &
        (df[col_wbs].astype(str).str.startswith(pfx))
    ]
    return not sub.empty


def ensure_level10_row(df, col_nivel, col_wbs, col_desc, leaf_code):
    mask_leaf = (
        (df[col_wbs].astype(str) == str(leaf_code)) &
        (df[col_nivel].fillna(0) < 10)
    )
    leaf_idx = df.index[mask_leaf]
    if len(leaf_idx) == 0:
        raise ValueError(f"WBS code not found: {leaf_code}")
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
    lower = df.loc[i + 1:]
    new_df = pd.concat([upper, pd.DataFrame([new_row]), lower], ignore_index=True)
    df.drop(df.index, inplace=True)
    for c in new_df.columns:
        df[c] = new_df[c]
    return i + 1


def find_level10_text(df, col_nivel, col_wbs, col_desc, leaf_code):
    mask_leaf = (
        (df[col_wbs].astype(str) == str(leaf_code)) &
        (df[col_nivel].fillna(0) < 10)
    )
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
    if df is None or col_nivel is None:
        return []

    niv = split_levels(df, col_nivel)

    def norm(s):
        if pd.isna(s):
            return ""
        return str(s).strip()

    baseline = baseline_series if baseline_series is not None else df[col_desc].copy()

    last_code = None
    last_desc = ""
    seen = OrderedDict()

    for i in range(len(df)):
        lvl  = niv.iat[i]
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


def _safe_str(val) -> str:
    if val is None:
        return ""
    try:
        if pd.isna(val):
            return ""
    except Exception:
        pass
    return str(val).strip()


def _split_cell(value: str, separator: str = "/") -> list[str]:
    return [t.strip() for t in value.split(separator) if t.strip()]


def _build_filter_specs(raw_class: str, raw_predef: str, raw_objtype: str) -> list[dict]:
    objtype = raw_objtype.strip()

    class_tokens_raw = _split_cell(raw_class)
    if any("." in t for t in class_tokens_raw):
        specs = []
        for token in class_tokens_raw:
            if "." in token:
                parts = token.split(".", 1)
                ifc_cls = parts[0].strip()
                predef  = parts[1].strip()
            else:
                ifc_cls = token.strip()
                predef  = ""
            if ifc_cls:
                specs.append({
                    "ifc_class":   ifc_cls,
                    "predefined":  predef,
                    "object_type": objtype if predef.upper() == "USERDEFINED" else "",
                    "props": [],
                })
        return specs

    class_tokens  = _split_cell(raw_class)
    predef_tokens = _split_cell(raw_predef)

    if not class_tokens:
        return []

    specs = []

    if len(class_tokens) > 1:
        for cls, prd in zip(class_tokens, predef_tokens or [""]):
            specs.append({
                "ifc_class":   cls,
                "predefined":  prd,
                "object_type": objtype if prd.upper() == "USERDEFINED" else "",
                "props": [],
            })
    else:
        single_cls = class_tokens[0]
        if not predef_tokens:
            predef_tokens = [""]
        for prd in predef_tokens:
            specs.append({
                "ifc_class":   single_cls,
                "predefined":  prd,
                "object_type": objtype if prd.upper() == "USERDEFINED" else "",
                "props": [],
            })

    return specs


def extract_partial_mapping(df: pd.DataFrame, cols: dict) -> dict:
    col_wbs       = cols.get("col_wbs")
    col_nivel     = cols.get("col_nivel")
    col_ifc_class = cols.get("col_ifc_class")
    col_predef    = cols.get("col_predef")
    col_objtype   = cols.get("col_objtype")
    col_ifc_prop  = cols.get("col_ifc_prop")

    if not all([col_wbs, col_nivel, col_ifc_class]):
        return {}

    niv = split_levels(df, col_nivel)
    rules = {}

    for i in range(len(df)):
        lvl = niv.iat[i]
        if lvl is None or lvl >= 10:
            continue

        code = _safe_str(df.at[i, col_wbs])
        if not code:
            continue

        raw_class = _safe_str(df.at[i, col_ifc_class])
        if not raw_class:
            continue

        raw_predef  = _safe_str(df.at[i, col_predef])  if col_predef  else ""
        raw_objtype = _safe_str(df.at[i, col_objtype]) if col_objtype else ""
        ifc_prop    = _safe_str(df.at[i, col_ifc_prop]) if col_ifc_prop else ""

        filter_specs = _build_filter_specs(raw_class, raw_predef, raw_objtype)
        if not filter_specs:
            continue

        qty_pset = ""
        qty_prop = ""
        if ifc_prop:
            if "." in ifc_prop:
                qty_pset, qty_prop = ifc_prop.split(".", 1)
                qty_pset = qty_pset.strip()
                qty_prop = qty_prop.strip()
            else:
                qty_prop = ifc_prop

        mappings = []
        for fspec in filter_specs:
            entry = {"filter": fspec}
            if qty_prop:
                entry["quantity_detail"] = {
                    "pset": qty_pset,
                    "prop": qty_prop,
                }
            mappings.append(entry)

        rules[code] = {
            "mappings": mappings,
            "material": "",
            "quantity": {
                "type": "prop" if ifc_prop else "count",
            },
            "agrupamento": {
                "pset": "",
                "prop": "",
            },
        }

    return rules
