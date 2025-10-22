# app/gui/app.py
import os
from pathlib import Path
import threading
import tkinter as tk
from tkinter import ttk, filedialog, messagebox

import pandas as pd

from openpyxl.styles import Font, PatternFill, Alignment, Border, Side

from app.gui.wbs_helpers import find_wbs_columns, split_levels, detect_relevant_leaves, list_ancestors
from app.gui.views.home import HomePage
from app.gui.views.wbs_editor import WBSPage
from app.gui.views.qty import QtyPage
from app.gui.views.report import ReportPage
from app.core.structural_engine import IFCInvestigator

from ifcopenshell.util.element import get_psets

class WBSApp(tk.Tk):

    def __init__(self):
        super().__init__()

        try:
            from app import __version__
            version_str = f" | v{__version__}"
        except ImportError:
            version_str = ""

        self.title(f"Extração de Quantidades IFC → WBS{version_str}")

        self.minsize(1200, 800)
        try:
            self.state("zoomed")
        except Exception:
            pass
        self._fullscreen = False
        self.bind("<F11>", self._toggle_fullscreen)
        self.bind("<Escape>", self._exit_fullscreen)

        self.df_raw: pd.DataFrame | None = None
        self.df_desc0: pd.Series | None = None
        self.col_nivel: str | None = "Nível"
        self.col_wbs: str | None = "WBS"
        self.col_desc: str | None = "DESCRIÇÃO"

        self.wbs_xlsx_var = tk.StringVar(value="")
        self.ifc_var = tk.StringVar(value="")
        self.out_var = tk.StringVar(value="")
        self.map_var = tk.StringVar(value="")

        self.rules = {}
        self.wbs_finalized = False

        self.ifc_file = None
        self.ifc_path_loaded = None
        self.inv = IFCInvestigator()

        self._previous_tab_index = None

        self._build_ui()

    def pick_file(self, var: tk.StringVar, title: str, patterns):
        path = filedialog.askopenfilename(title=title, filetypes=patterns)
        if path:
            var.set(path)

    def pick_dir(self, var: tk.StringVar, title: str):
        path = filedialog.askdirectory(title=title)
        if path:
            var.set(path)

    def _build_ui(self):
        self.notebook = ttk.Notebook(self)
        self.notebook.pack(fill="both", expand=True, padx=8, pady=8)

        print("\n" + "="*60)
        print("INICIANDO CONSTRUÇÃO DA UI")
        print("="*60)

        try:
            print("[1/4] Criando HomePage...")
            self.page_home = HomePage(self.notebook, self)
            print(f"      ✓ HomePage criada: {type(self.page_home)}")
        except Exception as e:
            print(f"      ✗ ERRO ao criar HomePage: {e}")
            raise

        try:
            print("[2/4] Criando WBSPage...")
            self.page_wbs = WBSPage(self.notebook, self)
            print(f"      ✓ WBSPage criada: {type(self.page_wbs)}")
        except Exception as e:
            print(f"      ✗ ERRO ao criar WBSPage: {e}")
            raise

        try:
            print("[3/4] Criando QtyPage...")
            self.page_qty = QtyPage(self.notebook, self)
            print(f"      ✓ QtyPage criada: {type(self.page_qty)}")
        except Exception as e:
            print(f"      ✗ ERRO ao criar QtyPage: {e}")
            raise

        try:
            print("[4/4] Criando ReportPage...")
            self.page_report = ReportPage(self.notebook, self)
            print(f"      ✓ ReportPage criada: {type(self.page_report)}")
        except Exception as e:
            print(f"      ✗ ERRO ao criar ReportPage: {e}")
            print(f"      AVISO: Continuando sem a aba Report")
            self.page_report = None

        print("\nAdicionando abas ao notebook...")

        try:
            self.notebook.add(self.page_home, text="Home")
            print("  [0] Home adicionada")
        except Exception as e:
            print(f"  [0] ERRO ao adicionar Home: {e}")

        try:
            self.notebook.add(self.page_wbs, text="WBS e descrição")
            print("  [1] WBS e descrição adicionada")
        except Exception as e:
            print(f"  [1] ERRO ao adicionar WBS: {e}")

        try:
            self.notebook.add(self.page_qty, text="Mapeamento IFC")
            print("  [2] Mapeamento IFC adicionada")
        except Exception as e:
            print(f"  [2] ERRO ao adicionar Mapeamento: {e}")

        if self.page_report is not None:
            try:
                self.notebook.add(self.page_report, text="Extrair quantidades e Gerar WBS preenchido")
                print("  [3] Extrair quantidades adicionada")
            except Exception as e:
                print(f"  [3] ERRO ao adicionar Extrair: {e}")
        else:
            print("  [3] PULANDO: ReportPage não foi criada")

        total_abas = self.notebook.index('end')
        print(f"\n{'='*60}")
        print(f"UI CONSTRUÍDA: {total_abas} abas no notebook")
        print(f"{'='*60}\n")

        self.notebook.bind("<<NotebookTabChanged>>", self._on_tab_changed)

    def _get_tab_index(self, page_widget):
        try:
            for i in range(self.notebook.index('end')):
                if self.notebook.nametowidget(self.notebook.tabs()[i]) == page_widget:
                    return i
            return None
        except Exception as e:
            print(f"[ERRO] Falha ao encontrar índice da aba: {e}")
            return None

    def go_home(self):
        try:
            idx = self._get_tab_index(self.page_home)
            if idx is not None:
                self.notebook.select(idx)
            else:
                print("[ERRO] Aba Home não encontrada")
        except Exception as e:
            print(f"[ERRO] Falha ao ir para Home: {e}")

    def go_wbs(self):
        try:
            idx = self._get_tab_index(self.page_wbs)
            if idx is not None:
                self.notebook.select(idx)
            else:
                print("[ERRO] Aba WBS não encontrada")
        except Exception as e:
            print(f"[ERRO] Falha ao ir para WBS: {e}")

    def open_mapping(self, source: str = "home"):
        try:
            if hasattr(self.page_qty, "set_mode"):
                self.page_qty.set_mode(source)
            idx = self._get_tab_index(self.page_qty)
            if idx is not None:
                self.notebook.select(idx)
            else:
                print("[ERRO] Aba Mapeamento não encontrada")
        except Exception as e:
            print(f"[ERRO] Falha ao ir para Mapeamento: {e}")

    def open_extract(self, source: str = "home"):
        try:
            if hasattr(self.page_report, "set_mode"):
                self.page_report.set_mode(source)
            idx = self._get_tab_index(self.page_report)
            if idx is not None:
                self.notebook.select(idx)
            else:
                print("[ERRO] Aba Extrair não encontrada")
        except Exception as e:
            print(f"[ERRO] Falha ao ir para Extrair: {e}")

    def go_mapping(self):
        self.open_mapping("home")

    def go_extract(self):
        self.open_extract("home")

    def has_user_descriptions(self) -> bool:
        if self.df_raw is None or self.col_nivel is None:
            return False
        try:
            pairs = detect_relevant_leaves(
                self.df_raw, self.col_nivel, self.col_wbs, self.col_desc, self.df_desc0
            )
            return bool(pairs)
        except Exception:
            return False

    def has_ifc_mapping(self) -> bool:
        return any(self.rules.values())

    def ensure_ifc_loaded(self, path: str):
        if not path:
            raise ValueError("Caminho IFC vazio.")
        if self.ifc_path_loaded == path and self.ifc_file is not None:
            return self.ifc_file
        self.ifc_file = self.inv.open_ifc(path)
        self.ifc_path_loaded = path
        return self.ifc_file

    def _toggle_fullscreen(self, event=None):
        self._fullscreen = not getattr(self, "_fullscreen", False)
        self.attributes("-fullscreen", self._fullscreen)

    def _exit_fullscreen(self, event=None):
        self._fullscreen = False
        self.attributes("-fullscreen", False)
    
    def _clear_memory_for_extraction(self):
        """
        Limpa memória quando entra na aba de Extração.
        Mantém apenas: rules (mapeamento) e os caminhos dos arquivos.
        Força recarregamento de WBS e IFC dos arquivos.
        """
        print("\n" + "="*60)
        print("LIMPEZA DE MEMORIA PARA EXTRACAO")
        print("="*60)
        
        print("  - Limpando df_raw e df_desc0...")
        self.df_raw = None
        self.df_desc0 = None
        
        print("  - Limpando cache do IFC...")
        self.ifc_file = None
        self.ifc_path_loaded = None
        
        num_rules = len(self.rules) if self.rules else 0
        print(f"  - Mantendo rules: {num_rules} codigos mapeados")
        
        print("  - Mantendo caminhos:")
        print(f"      WBS: {self.wbs_xlsx_var.get()}")
        print(f"      IFC: {self.ifc_var.get()}")
        print(f"      MAP: {self.map_var.get()}")
        
        print("Limpeza concluida")
        print("="*60 + "\n")

    def _on_tab_changed(self, event):
        try:
            current_tab = self.notebook.select()
            current_tab_index = self.notebook.index(current_tab)
            
            tab_names = {
                0: "Home",
                1: "WBS",
                2: "Mapeamento",
                3: "Extração"
            }
            
            previous_tab = tab_names.get(self._previous_tab_index, "Desconhecida")
            current_tab_name = tab_names.get(current_tab_index, "Desconhecida")
            
            print(f"\n🔄 Mudança de aba: {previous_tab} → {current_tab_name}")
            
            if current_tab_index == 3 and self._previous_tab_index in [1, 2]:
                print("⚠️  Entrando em Extração vindo de outra aba - executando limpeza")
                self._clear_memory_for_extraction()
            
            if current_tab_index == 0:
                if hasattr(self.page_home, 'refresh_on_show'):
                    self.page_home.refresh_on_show()

            self._previous_tab_index = current_tab_index
            
        except Exception as e:
            print(f"[ERRO] Falha ao processar mudança de aba: {e}")

    def run_generate_report(self, log_widget, on_finish=None):
        def log(msg):
            log_widget.configure(state="normal")
            log_widget.insert("end", msg + ("\n" if not msg.endswith("\n") else ""))
            log_widget.see("end")
            log_widget.configure(state="disabled")

        def worker():
            try:
                
                print("\n" + "="*60)
                print("🚀 INICIANDO GERAÇÃO DE RELATÓRIO")
                print("="*60)
                
                if not self.rules:
                    raise RuntimeError("Carregue um mapeamento (JSON).")
                print(f"✓ Mapeamento carregado: {len(self.rules)} códigos")
                
                wbs_path = self.wbs_xlsx_var.get().strip()
                if not wbs_path or not Path(wbs_path).is_file():
                    raise RuntimeError("Selecione um ficheiro WBS válido.")
                
                if getattr(self, "wbs_finalized", False) and self.df_raw is not None:
                    print("📋 Usando WBS da memória (já editado na aba WBS)")
                    log("Usando WBS editado (com descrições do utilizador)")
                    
                    if not all([self.col_wbs, self.col_desc, self.col_nivel]):
                        raise RuntimeError("Colunas WBS não detectadas. Recarregue o WBS.")
                    
                else:
                    print(f"📂 Recarregando WBS de: {wbs_path}")
                    log(f"Carregando WBS: {Path(wbs_path).name}")
                    
                    self.df_raw = pd.read_excel(wbs_path, header=1)
                    print(f"   ✓ {len(self.df_raw)} linhas carregadas")
                    
                    print(f"   📋 Colunas no DataFrame: {list(self.df_raw.columns)}")
                    print(f"   🔍 'UNID.' na colunas? {'UNID.' in self.df_raw.columns}")
                    
                    self.col_wbs, self.col_desc, self.col_nivel = find_wbs_columns(self.df_raw)
                    if not all([self.col_wbs, self.col_desc, self.col_nivel]):
                        raise RuntimeError("Não foi possível detectar as colunas WBS/Descrição/Nível")
                    print(f"   ✓ Colunas detectadas: WBS={self.col_wbs}, DESC={self.col_desc}, NIVEL={self.col_nivel}")
                    
                    self.df_desc0 = self.df_raw[self.col_desc].copy()
                
                ifc_path = self.ifc_var.get().strip()
                if not ifc_path or not Path(ifc_path).is_file():
                    raise RuntimeError("Selecione um ficheiro IFC válido.")
                
                print(f"📂 Recarregando IFC de: {ifc_path}")
                log(f"Carregando IFC: {Path(ifc_path).name}")
                
                self.ifc_file = None
                self.ifc_path_loaded = None
                self.inv.open_ifc(ifc_path)
                print(f"   ✓ IFC carregado")

                out_dir = Path(self.out_var.get().strip() or Path.home())
                out_dir.mkdir(parents=True, exist_ok=True)

                out_path = out_dir / "WBS_Preenchido.xlsx"
                if out_path.exists():
                    n = 1
                    while out_path.exists():
                        out_path = out_dir / f"WBS_Preenchido({n}).xlsx"
                        n += 1
                
                print(f"📁 Saída: {out_path}")

                self.last_code_extensions = {}

                code_to_qty: dict[str, float] = {}
                code_to_groupvals: dict[str, list[str]] = {}
                code_to_groupqtys: dict[str, dict[str, float]] = {}
                code_to_desc_index: dict[str, int] = {}
                code_to_unit: dict[str, str] = {}

                project_info = self.inv.get_project_info()
                ifc_project = project_info.get("project", "n/a")
                ifc_site = project_info.get("site", "n/a")
                ifc_building = project_info.get("building", "n/a")

                all_details = []

                def _key(code: str):
                    return tuple(int(x) if x.isdigit() else x for x in code.split("."))

                log("\nExtraindo quantidades por código WBS:")
                
                for code in sorted(self.rules.keys(), key=_key):
                    rule = self.rules[code]
                    try:
                        elems = self.inv.filter_elements(rule)

                        q_pset = rule["quantity"]["pset"]
                        q_prop = rule["quantity"]["prop"]
                        q, details = self.inv.sum_quantity(elems, q_pset, q_prop)
                        
                        grp_vals = []
                        grp_sums = {}
                        agr = rule.get("agrupamento") or rule.get("agrupamento".encode("utf-8").decode("utf-8"))
                        g_pset = g_prop = None
                        if agr and isinstance(agr, dict):
                            g_pset = agr.get("pset")
                            g_prop = agr.get("prop")

                        if g_pset and g_prop:
                            seen = set()
                            for det in (details or []):
                                e = det.get("element")
                                elem_val = det.get("valor", 0.0)
                                try:
                                    psets = get_psets(e) or {}
                                    gval = psets.get(g_pset, {}).get(g_prop, None)
                                    
                                    if gval is None:
                                        continue
                                    sval = str(gval).strip()
                                    
                                    if not sval:
                                        continue

                                    if sval not in seen:
                                        seen.add(sval)
                                        grp_vals.append(sval)

                                    grp_sums[sval] = grp_sums.get(sval, 0.0) + float(elem_val or 0.0)
                                except Exception:
                                    pass
                        else:
                            grp_vals = []
                            grp_sums = {}

                        code_to_groupvals[str(code).strip()] = grp_vals
                        code_to_groupqtys[str(code).strip()] = grp_sums

                        
                        code_to_qty[str(code).strip()] = q or 0.0
                        log(f" - {code}: {q}")

                        if details:
                            f = rule.get("filter", {})
                            for det in details:
                                e = det.get("element")
                                guid = det.get("guid", "n/a")
                                val = det.get("valor", None)
                                material = self.inv.get_element_material(e) if e else "n/a"
                                classification_code = self.inv.get_classification_code(e) if e else "n/a"
                                buildingstorey = self.inv.get_building_storey(e) if e else "n/a"

                                desc, unidade = "", "n/a"
                                try:
                                    df_desc = self.df_raw.copy()
                                    df_desc[self.col_nivel] = split_levels(df_desc, self.col_nivel)
                                    parent_code = str(code).strip()
                                    leaf_rows = df_desc[
                                        (df_desc[self.col_wbs].astype(str).str.strip() == parent_code) &
                                        (df_desc[self.col_nivel] < 10)
                                    ]
                                    if not leaf_rows.empty:
                                        leaf_idx = leaf_rows.index[0]
                                        next_idx = leaf_idx + 1
                                        if next_idx < len(df_desc) and df_desc.at[next_idx, self.col_nivel] == 10:
                                            desc = str(df_desc.at[next_idx, self.col_desc]) if self.col_desc in df_desc.columns else ""
                                            if "UNID." in df_desc.columns:
                                                unidade = str(df_desc.at[next_idx, "UNID."])
                                            code_to_desc_index.setdefault(parent_code, next_idx)
                                        code_to_unit.setdefault(parent_code, unidade)
                                
                                except Exception as ex:
                                    print(f"[WARN] Erro ao obter descrição/unidade: {ex}")

                                group_value = None
                                try:
                                    agr = rule.get("agrupamento")
                                    if agr and isinstance(agr, dict):
                                        g_pset = agr.get("pset")
                                        g_prop = agr.get("prop")
                                        if g_pset and g_prop and e is not None:
                                            psets = get_psets(e) or {}
                                            gv = psets.get(g_pset, {}).get(g_prop)
                                            if gv is not None:
                                                sgv = str(gv).strip()
                                                group_value = sgv if sgv else None
                                except Exception as ex:
                                    print(f"[WARN] Erro ao obter valor de agrupamento: {ex}")

                                all_details.append({
                                    "ifc_project": ifc_project,
                                    "ifc_site": ifc_site,
                                    "ifc_building": ifc_building,
                                    "wbs_codigo": str(code).strip(),
                                    "parent_code": str(code).strip(),
                                    "group_value": group_value,
                                    "descricao": desc,
                                    "ifc_class": f.get("ifc_class", "n/a"),
                                    "predefinedtype": f.get("predefined", "n/a"),
                                    "objecttype": f.get("object_type", "n/a"),
                                    "wbs_grouping": "n/a",
                                    "material": material,
                                    "classification_code": classification_code,
                                    "buildingstorey": buildingstorey,
                                    "ifc_guid": guid,
                                    "ifc_valor": val,
                                    "unidade": unidade
                                })

                    except Exception as e:
                        log(f" - {code}: erro ({e})")

                df = self.df_raw.copy()
                lvl = split_levels(df, self.col_nivel)

                col_qty = "QDTE."
                if col_qty not in df.columns:
                    df[col_qty] = ""

                columns_to_keep = [self.col_wbs, self.col_desc, col_qty, "UNID."]
                df_export = df[[c for c in columns_to_keep if c in df.columns]]

                out_rows = []
                row_kinds = []

                idx10_to_parent = {}
                for parent_code, idx10 in code_to_desc_index.items():
                    idx10_to_parent[idx10] = parent_code

                self.last_code_extensions = {}

                for idx, row in df_export.iterrows():
                    is_desc10 = (idx < len(lvl) and lvl.iat[idx] == 10)
                    kind = "desc10" if is_desc10 else "wbs"

                    new_row = row.copy()

                    if is_desc10:
                        if col_qty in new_row.index:
                            new_row[col_qty] = ""
                        if "UNID." in new_row.index:
                            new_row["UNID."] = ""

                    if is_desc10:
                        parent_code = idx10_to_parent.get(idx)
                        if parent_code:
                            desc_ext = f"{parent_code}.01"
                            new_row[self.col_wbs] = desc_ext
                            
                            group_vals = code_to_groupvals.get(parent_code) or []
                            groups_map = {}
                            for k, gval in enumerate(group_vals, start=1):
                                g_ext = f"{desc_ext}.{k:02d}"
                                groups_map[str(gval)] = g_ext

                            self.last_code_extensions[parent_code] = {
                                "desc": desc_ext,
                                "groups": groups_map
                            }

                    out_rows.append(new_row)
                    row_kinds.append(kind)

                    if is_desc10:
                        parent_code = idx10_to_parent.get(idx)
                        if parent_code:
                            group_vals = code_to_groupvals.get(parent_code) or []
                            groups_map = self.last_code_extensions.get(parent_code, {}).get("groups", {})
                            
                            for gval in group_vals:

                                ins_data = {}
                                
                                for col in df_export.columns:
                                    if col == self.col_wbs:
                                        ins_data[col] = groups_map.get(str(gval), "")
                                    elif col == self.col_desc:
                                        ins_data[col] = str(gval)
                                    elif col == col_qty or col == "QDTE.":
                                        qty_value = code_to_groupqtys.get(parent_code, {}).get(str(gval), 0.0)
                                        ins_data[col] = qty_value
                                        print(f"[DEBUG INS] {parent_code} -> {gval}: qty={qty_value}, col={col}")
                                    elif col == "UNID.":
                                        ins_data[col] = code_to_unit.get(parent_code, "")
                                    else:
                                        ins_data[col] = ""

                                new_series = pd.Series(ins_data, index=df_export.columns)
                                out_rows.append(new_series)
                                row_kinds.append("insert")
                                
                                print(f"[DEBUG SERIES] Criada linha: {new_series[col_qty] if col_qty in new_series.index else 'N/A'}")

                df_export2 = pd.DataFrame(out_rows, columns=df_export.columns)

                with pd.ExcelWriter(out_path, engine="openpyxl") as writer:
                    df_export2.to_excel(writer, index=False, sheet_name="WBS Preenchido")
                    ws = writer.sheets["WBS Preenchido"]

                    header_fill = PatternFill(start_color="00B050", end_color="00B050", fill_type="solid")
                    header_font = Font(bold=True, color="FFFFFF", size=12)
                    header_alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)

                    gray_fill = PatternFill(start_color="F2F2F2", end_color="F2F2F2", fill_type="solid")
                    white_fill = PatternFill(start_color="FFFFFF", end_color="FFFFFF", fill_type="solid")

                    border = Border(
                        left=Side(style="thin"),
                        right=Side(style="thin"),
                        top=Side(style="thin"),
                        bottom=Side(style="thin"),
                    )
                    data_alignment = Alignment(horizontal="left", vertical="top", wrap_text=True)

                    for cell in ws[1]:
                        if cell.value:
                            cell.fill = header_fill
                            cell.font = header_font
                            cell.alignment = header_alignment
                            cell.border = border

                    for i, kind in enumerate(row_kinds, start=2):
                        fill = gray_fill if kind == "wbs" else white_fill
                        for cell in ws[i]:
                            cell.alignment = data_alignment
                            cell.border = border
                            cell.fill = fill

                    ws.column_dimensions["A"].width = 20
                    ws.column_dimensions["B"].width = 50
                    ws.column_dimensions["C"].width = 15
                    ws.column_dimensions["D"].width = 12

                log(f"\n✓ Exportado: {out_path}")

                headers = [
                    "ifc_filename", 
                    "wbs_codigo", "descricao",
                    "ifc_class", "predefinedtype", "objecttype",
                    "wbs_group", "material",
                    "ifc_guid", "buildingstorey", "classification_code",
                    "ifc_project", "ifc_site", "ifc_building",
                    "ifc_valor", "unidade"
                ]

                wbs_rows = []
                df_w = self.df_raw.copy()
                lvl_w = split_levels(df_w, self.col_nivel)
                for i in range(len(df_w)):
                    if lvl_w.iat[i] is not None and lvl_w.iat[i] < 10:
                        code = str(df_w.at[i, self.col_wbs]).strip() if pd.notna(df_w.at[i, self.col_wbs]) else ""
                        desc = str(df_w.at[i, self.col_desc]).strip() if pd.notna(df_w.at[i, self.col_desc]) else ""
                        if code:
                            wbs_rows.append({"wbs_codigo": code, "descricao": desc})

                per_code = {}
                for det in (all_details or []):
                    code = det.get("wbs_codigo")
                    if not code:
                        continue
                    blob = per_code.setdefault(code, {"meta": {}, "groups": {}})
                    blob["meta"] = {
                        "ifc_project": det.get("ifc_project", "n/a"),
                        "ifc_site":    det.get("ifc_site", "n/a"),
                        "ifc_building":det.get("ifc_building", "n/a"),
                    }
                    gval = det.get("group_value") or "n/a"
                    lst = blob["groups"].setdefault(gval, [])
                    lst.append({
                        "ifc_class": det.get("ifc_class", "n/a"),
                        "predefined": det.get("predefinedtype", "n/a"),
                        "objecttype": det.get("objecttype", "n/a"),
                        "material": det.get("material", "n/a"),
                        "guid": det.get("ifc_guid", "n/a"),
                        "buildingstorey": det.get("buildingstorey", "n/a"),
                        "classification_code": det.get("classification_code", "n/a"),
                        "value": det.get("ifc_valor", ""),
                    })
                    
                self._last_csv_cache = {
                    "headers": headers,
                    "wbs_rows": wbs_rows,
                    "per_code": per_code,
                    "code_to_unit": code_to_unit,
                }

                self.last_detailed_rows = all_details
                self.page_report.btn_export_csv.config(state="normal")
                
                if not hasattr(self, "last_detailed_rows"):
                    self.last_detailed_rows = []
                if not hasattr(self, "last_code_extensions"):
                    self.last_code_extensions = {}
                
                log(f"Detalhes recolhidos: {len(self.last_detailed_rows)} elementos")
                log(f"Extensões de WBS: {len(self.last_code_extensions)} pais")

                print("\n" + "="*60)
                print("✅ RELATÓRIO GERADO COM SUCESSO")
                print("="*60 + "\n")

                if on_finish:
                    on_finish("WBS preenchido gerado com sucesso.")

            except Exception as e:
                print(f"\n❌ ERRO NA GERAÇÃO: {e}")
                if on_finish:
                    on_finish(f"Erro: {e}")

        threading.Thread(target=worker, daemon=True).start()