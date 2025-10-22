# app/gui/views/report.py
import json
from pathlib import Path
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from tkinter.scrolledtext import ScrolledText

import pandas as pd

from app.gui.wbs_helpers import find_wbs_columns, split_levels

import os
import ifcopenshell


class ReportPage(tk.Frame):

    def __init__(self, notebook, app):
        super().__init__(notebook)
        self.app = app
        self.pack_propagate(False)

        def _s(v):
            try:
                return v.get().strip()
            except Exception:
                return ""
            
        wbs_guess = getattr(app, "last_exported_wbs", "") or _s(getattr(app, "wbs_xlsx_var", "")) or getattr(app, "wbs_path_loaded", "") or ""
        map_guess = getattr(app, "map_path", "") or _s(getattr(app, "map_var", "")) or ""
        ifc_guess = getattr(app, "ifc_file", "") or _s(getattr(app, "ifc_var", "")) or ""
        out_guess = _s(getattr(app, "out_var", "")) or (str(Path(wbs_guess).parent) if wbs_guess else str(Path.home()))

        self.wbs_var = tk.StringVar(value=wbs_guess)
        self.map_var = tk.StringVar(value=map_guess)
        self.ifc_var = tk.StringVar(value=ifc_guess)
        self.out_var = tk.StringVar(value=out_guess)

        self._build_ui()
        
    def _build_ui(self):
        self.grid_columnconfigure(0, weight=0)
        self.grid_columnconfigure(1, weight=1)
        self.grid_columnconfigure(2, weight=0)
        self.grid_columnconfigure(3, weight=0)
        pad = {"padx": (10, 8), "pady": (6, 4)}

        self.banner = tk.Label(self, anchor="w",
                               text="Carregue o WBS (com descrições), o mapeamento (JSON) e o IFC.")
        self.banner.grid(row=0, column=0, columnspan=4, sticky="we", padx=10, pady=(8, 4))

        tk.Label(self, text="WBS com descrições (Excel):").grid(row=1, column=0, sticky="e", **pad)
        self.e_wbs = tk.Entry(self, textvariable=self.wbs_var)
        self.e_wbs.grid(row=1, column=1, sticky="we", **pad)
        self.b_wbs_pick = tk.Button(self, text="Procurar…", command=self.browse_wbs)
        self.b_wbs_pick.grid(row=1, column=2, sticky="w", **pad)
        self.b_wbs_load = tk.Button(self, text="Carregar WBS", command=self.on_load_wbs)
        self.b_wbs_load.grid(row=1, column=3, sticky="w", **pad)

        tk.Label(self, text="Mapeamento (JSON):").grid(row=2, column=0, sticky="e", **pad)
        self.e_map = tk.Entry(self, textvariable=self.map_var)
        self.e_map.grid(row=2, column=1, sticky="we", **pad)
        self.b_map_pick = tk.Button(self, text="Procurar…", command=self.browse_map)
        self.b_map_pick.grid(row=2, column=2, sticky="w", **pad)
        self.b_map_load = tk.Button(self, text="Carregar mapeamento", command=self.on_load_mapping)
        self.b_map_load.grid(row=2, column=3, sticky="w", **pad)

        tk.Label(self, text="Modelo IFC:").grid(row=3, column=0, sticky="e", **pad)
        self.e_ifc = tk.Entry(self, textvariable=self.ifc_var)
        self.e_ifc.grid(row=3, column=1, sticky="we", **pad)
        self.b_ifc_pick = tk.Button(self, text="Procurar…", command=self.browse_ifc)
        self.b_ifc_pick.grid(row=3, column=2, sticky="w", **pad)
        self.b_ifc_load = tk.Button(self, text="Carregar IFC", command=self.on_load_ifc)
        self.b_ifc_load.grid(row=3, column=3, sticky="w", **pad)

        tk.Label(self, text="Pasta de saída:").grid(row=4, column=0, sticky="e", **pad)
        self.e_out = tk.Entry(self, textvariable=self.out_var)
        self.e_out.grid(row=4, column=1, sticky="we", **pad)
        self.b_out_pick = tk.Button(self, text="Procurar…", command=self.browse_outdir)
        self.b_out_pick.grid(row=4, column=2, sticky="w", **pad)

        self.run_btn = tk.Button(self, text="Gerar WBS preenchido", command=self.on_run)
        self.run_btn.grid(row=5, column=0, columnspan=2, sticky="we", padx=(10, 4), pady=(8, 6))
             
        self.btn_export_csv = tk.Button(
            self,
            text="Exportar CSV detalhado",
            state="disabled",
            command=self.on_export_csv
        )
        self.btn_export_csv.grid(row=5, column=2, columnspan=2, sticky="we", padx=(4, 10), pady=(8, 6))

        self.log = ScrolledText(self, height=14, state="disabled")
        self.log.grid(row=6, column=0, columnspan=4, sticky="nsew", padx=10, pady=(4, 10))
        self.grid_rowconfigure(6, weight=1)

    def set_mode(self, source: str = "home"):

        self._mode = source or "home"

        from_previous = source in ("wbs", "mapping", "from_mapping")
        
        if from_previous:

            self.banner.config(text="A utilizar o WBS, mapeamento e IFC definidos na aba anterior.")
            self._set_inputs_state("disabled")
        else:

            self.banner.config(text="Carregue o WBS (com descrições), o mapeamento (JSON) e o IFC.")
            self._set_inputs_state("normal")

        try:
            self.after(100, self._autoload_from_previous_tab)
        except Exception:
            self._autoload_from_previous_tab()

    def _set_inputs_state(self, state: str):

        for w in (self.e_wbs, self.b_wbs_pick, self.b_wbs_load,
                self.e_map, self.b_map_pick, self.b_map_load,
                self.e_ifc, self.b_ifc_pick, self.b_ifc_load):
            try:
                w.configure(state=state)
            except Exception:
                pass

        try:
            self.e_out.configure(state="normal")
            self.b_out_pick.configure(state="normal")
        except Exception:
            pass

    def _autoload_from_previous_tab(self):

        app = self.app
        loaded = []
        warnings = []

        try:
            wbs_path = getattr(app, "wbs_xlsx_var", None)
            wbs_path_str = wbs_path.get() if wbs_path else ""
            
            if wbs_path_str:
                self.wbs_var.set(wbs_path_str)

            if getattr(app, "df_raw", None) is None and wbs_path_str and Path(wbs_path_str).is_file():
                self._log("Carregando WBS automaticamente do arquivo...\n")
                try:

                    df = pd.read_excel(wbs_path_str, header=1)

                    col_wbs, col_desc, col_nivel = find_wbs_columns(df)
                    
                    if col_wbs and col_desc and col_nivel:
                        app.df_raw = df.copy()
                        app.col_wbs = col_wbs
                        app.col_desc = col_desc
                        app.col_nivel = col_nivel
                        app.df_desc0 = app.df_raw[app.col_desc].copy()
                        loaded.append(f"[OK] WBS: {len(df)} linhas carregadas")
                        self._log(f"    Colunas: {col_wbs}, {col_desc}, {col_nivel}\n")
                    else:
                        warnings.append("[X] WBS: colunas nao encontradas")
                        self._log(f"    ERRO: Colunas disponiveis: {list(df.columns)}\n")
                except Exception as e:
                    warnings.append(f"[X] Erro ao carregar WBS: {e}")
                    self._log(f"    ERRO: {e}\n")
            elif getattr(app, "df_raw", None) is not None:
                loaded.append("[OK] WBS (ja carregado em memoria)")
            else:
                warnings.append("[X] WBS nao carregado")
        except Exception as e:
            warnings.append(f"[X] Erro ao verificar WBS: {e}")

        try:
            if getattr(app, "rules", None):
                loaded.append(f"[OK] Mapeamento: {len(app.rules)} codigos")

                if getattr(app, "map_var", None) and app.map_var.get():
                    self.map_var.set(app.map_var.get())
                else:
                    self.map_var.set("[carregado da aba anterior]")
            else:
                warnings.append("[X] Mapeamento nao carregado")
        except Exception as e:
            warnings.append(f"[X] Erro ao verificar mapeamento: {e}")

        try:
            ifc_path = getattr(app, "ifc_var", None)
            ifc_path_str = ifc_path.get() if ifc_path else ""
            
            if ifc_path_str:
                self.ifc_var.set(ifc_path_str)

            if getattr(app, "ifc_file", None) is None and ifc_path_str and Path(ifc_path_str).is_file():
                self._log("Carregando IFC automaticamente...\n")
                try:
                    app.inv.open_ifc(ifc_path_str)
                    app.ifc_path_loaded = ifc_path_str
                    app.ifc_file = app.inv.ifc_file
                    loaded.append(f"[OK] IFC: {Path(ifc_path_str).name}")
                except Exception as e:
                    warnings.append(f"[X] Erro ao carregar IFC: {e}")
                    self._log(f"    ERRO: {e}\n")
            elif getattr(app, "ifc_file", None) is not None:
                loaded.append("[OK] IFC (ja carregado em memoria)")
            else:
                warnings.append("[X] IFC nao carregado")
        except Exception as e:
            warnings.append(f"[X] Erro ao verificar IFC: {e}")

        try:
            if not self.out_var.get():
                if self.wbs_var.get():
                    self.out_var.set(str(Path(self.wbs_var.get()).parent))
                else:
                    self.out_var.set(str(Path.home()))
        except Exception:
            pass

        lines = ["=== Dados carregados automaticamente ===\n"]
        
        for msg in loaded:
            lines.append(msg)
        
        if warnings:
            lines.append("\nAvisos:")
            for msg in warnings:
                lines.append(msg)
            lines.append("\nCarregue os dados em falta antes de gerar o WBS preenchido.")
        else:
            lines.append("\n[OK] Todos os dados estao prontos!")
            lines.append("Pode clicar em 'Gerar WBS preenchido' quando estiver pronto.\n")

        try:
            self.log.configure(state="normal")
            self.log.delete("1.0", "end")
            self.log.insert("end", "\n".join(lines))
            self.log.configure(state="disabled")
            self.log.see("end")
        except Exception as e:
            print(f"[DEBUG] Erro ao atualizar log: {e}")


    def browse_wbs(self):
        p = filedialog.askopenfilename(title="Escolher WBS com descrições (Excel)",
                                       filetypes=[("Excel (*.xlsx *.xls)", "*.xlsx *.xls")])
        if p:
            self.wbs_var.set(p)

    def browse_map(self):
        p = filedialog.askopenfilename(title="Escolher mapeamento (JSON)",
                                       filetypes=[("JSON", "*.json")])
        if p:
            self.map_var.set(p)

    def browse_ifc(self):
        p = filedialog.askopenfilename(title="Escolher IFC",
                                       filetypes=[("IFC", "*.ifc")])
        if p:
            self.ifc_var.set(p)

    def browse_outdir(self):
        p = filedialog.askdirectory(title="Escolher pasta de saída")
        if p:
            self.out_var.set(p)

    def on_load_wbs(self):
        path = self.wbs_var.get().strip()
        if not path:
            messagebox.showinfo("Extrair quantidades e Gerar WBS", "Selecione um ficheiro Excel.")
            return
        if not Path(path).is_file():
            messagebox.showerror("Extrair quantidades e Gerar WBS", "Caminho invalido.")
            return
        try:

            df = pd.read_excel(path, header=1)
            
            col_wbs, col_desc, col_nivel = find_wbs_columns(df)
            
            if col_wbs and col_desc and col_nivel:
                self.app.df_raw = df.copy()
                self.app.col_wbs = col_wbs
                self.app.col_desc = col_desc
                self.app.col_nivel = col_nivel
                self.app.df_desc0 = self.app.df_raw[self.app.col_desc].copy()
                messagebox.showinfo("Extrair quantidades e Gerar", 
                                f"WBS carregado com sucesso.\n{len(df)} linhas")
                return
            
            raise RuntimeError(f"Colunas nao encontradas. Disponiveis: {list(df.columns)}")
        except Exception as e:
            messagebox.showerror("Extrair quantidades e Gerar WBS", f"Falha a ler o WBS:\n{e}")

    def on_load_mapping(self):
            path = self.map_var.get().strip()
            if not path:
                messagebox.showinfo("Extrair quantidades e Gerar WBS", "Selecione um ficheiro JSON.")
                return
            if not Path(path).is_file():
                messagebox.showerror("Extrair quantidades e Gerar WBS", "Caminho inválido.")
                return
            try:
                with open(path, "r", encoding="utf-8") as f:
                    data = json.load(f)
                rules = data.get("rules", data)
                if not isinstance(rules, dict) or not rules:
                    raise RuntimeError("Estrutura inválida ou vazia.")

                missing_agr = []
                for code, rule in rules.items():
                    agr = rule.get("agrupamento")
                    if not agr or not agr.get("pset") or not agr.get("prop"):
                        missing_agr.append(code)
                
                if missing_agr:
                    messagebox.showerror("Mapeamento IFC", 
                                    f"Agrupamento é obrigatório!\n\n"
                                    f"Os seguintes códigos não têm agrupamento:\n"
                                    f"{', '.join(missing_agr)}\n\n"
                                    f"Por favor, corrija o mapeamento.")
                    return
                
                self.app.rules = rules
                messagebox.showinfo("Extrair quantidades e Gerar WBS", "Mapeamento carregado com sucesso.")
            except Exception as e:
                messagebox.showerror("Extrair quantidades e Gerar WBS", f"Falha a carregar o mapeamento:\n{e}")

    def on_load_ifc(self):
        path = self.ifc_var.get().strip()
        if not path or not os.path.isfile(path):
            messagebox.showerror("Extrair quantidades e Gerar WBS", "Selecione primeiro um ficheiro IFC válido.")
            return
        try:
            self.app.inv.open_ifc(path)
            self.app.ifc_path_loaded = path
            self.app.ifc_file = self.app.inv.ifc_file
            messagebox.showinfo("Extrair quantidades e Gerar WBS", "IFC carregado com sucesso.")
        except Exception as e:
            messagebox.showerror("Extrair quantidades e Gerar WBS", f"Falha a abrir o ficheiro IFC:\n{e}")

    def on_run(self):
        self.run_btn.configure(state="disabled", text="A gerar…")
        self.log.configure(state="normal")
        self.log.delete("1.0", "end")
        self.log.insert("end", "A processar regras...\n")
        self.log.configure(state="disabled")

        self.app.wbs_xlsx_var.set(self.wbs_var.get().strip())
        self.app.map_var.set(self.map_var.get().strip())
        self.app.ifc_var.set(self.ifc_var.get().strip())
        self.app.out_var.set(self.out_var.get().strip())

        missing = []
        if not self.app.wbs_xlsx_var.get().strip(): missing.append("WBS (Excel)")
        if not self.app.map_var.get().strip():      missing.append("Mapeamento (JSON)")
        if not self.app.ifc_var.get().strip():      missing.append("IFC")
        if not self.app.out_var.get().strip():      missing.append("Pasta de saída")
        if missing:
            messagebox.showinfo("Dados em falta", "Indique: " + ", ".join(missing) + ".")
            self.run_btn.configure(state="normal", text="Gerar WBS preenchido")
            return

        def done(msg):
            messagebox.showinfo("Extrair quantidades e Gerar WBS", msg)
            self.run_btn.configure(state="normal", text="Gerar WBS preenchido")

        if hasattr(self.app, "run_generate_report"):
            self.after(10, lambda: self.app.run_generate_report(self.log, on_finish=done))
            return

        self.log.configure(state="normal")
        self.log.insert("end", "Pipeline de geração não encontrado na app. (fallback)\n")
        self.log.insert("end", "Nada foi gerado — implementa app.run_generate_report.\n")
        self.log.configure(state="disabled")
        done("Fluxo de UI testado com sucesso (sem geração).")

    def _log(self, text: str):
        self.log.configure(state="normal")
        self.log.insert("end", text + ("\n" if not text.endswith("\n") else ""))
        self.log.see("end")
        self.log.configure(state="disabled")      
        
    @staticmethod
    def _parse_wbs_code(code: str) -> tuple:
        if not code or not isinstance(code, str):
            return (float('inf'),)
        
        parts = []
        for part in code.split("."):
            part = part.strip()
            if part.isdigit():
                parts.append(int(part))
            else:
                parts.append(part)
        
        return tuple(parts)


    def on_export_csv(self):
        import csv
        from pathlib import Path
        import pandas as pd

        if self.app.df_raw is None:
            messagebox.showinfo("Extrair quantidades e Gerar WBS", "Carregue o WBS (com descrições) e gere primeiro o WBS preenchido.")
            return

        cache = getattr(self.app, "_last_csv_cache", None)
        if not isinstance(cache, dict):
            messagebox.showinfo("Extrair quantidades e Gerar WBS", "Gere primeiro o WBS preenchido (nenhum detalhe em cache).")
            return

        headers   = cache["headers"]
        headers   = [h for h in headers if h != "wbs_group"]
        
        per_code  = cache["per_code"]
        units_by_parent = cache.get("code_to_unit", {})

        if not units_by_parent:
            df_raw = self.app.df_raw.copy()
            lvl_raw = split_levels(df_raw, self.app.col_nivel)

            def pick_in_df(df, *cands):
                for c in cands:
                    if c in df.columns:
                        return c
                return None

            col_unit_raw = pick_in_df(
                df_raw,
                "UNID.", "Unid.", "UNID", "Unid",
                "UNIDADE", "Unidade", "Unidª", "UNIDª"
            )

            units_by_parent = {}
            if col_unit_raw:
                for i in range(len(df_raw)):
                    try:
                        if lvl_raw.iat[i] == 10:
                            parent_code = str(df_raw.at[i, self.app.col_wbs]).strip() if pd.notna(df_raw.at[i, self.app.col_wbs]) else ""
                            unit_val    = str(df_raw.at[i, col_unit_raw]).strip() if pd.notna(df_raw.at[i, col_unit_raw]) else ""
                            if parent_code:
                                units_by_parent[parent_code] = unit_val
                    except Exception:
                        pass

        out_dir = Path(self.app.out_var.get().strip() or Path.home())
        excel_path = out_dir / "WBS_Preenchido.xlsx"
        
        ifc_path_str = (self.ifc_var.get().strip()
                        or getattr(self.app, "ifc_path_loaded", "")
                        or "")
        ifc_filename = Path(ifc_path_str).stem if ifc_path_str else "n/a"
        
        if not excel_path.exists():
            messagebox.showerror("Erro", f"Excel não encontrado: {excel_path}")
            return

        try:
            df_excel = pd.read_excel(excel_path, sheet_name="WBS Preenchido")
        except Exception as e:
            messagebox.showerror("Erro", f"Falha ao ler o Excel:\n{e}")
            return

        col_wbs_excel  = self.app.col_wbs  if self.app.col_wbs  in df_excel.columns else "WBS"
        col_desc_excel = self.app.col_desc if self.app.col_desc in df_excel.columns else "DESCRIÇÃO"

        parent_order          = []
        desc_code_by_parent   = {}
        group_codes_by_parent = {}

        for _, row in df_excel.iterrows():
            code = str(row.get(col_wbs_excel, "")).strip()
            if not code:
                continue

            if code in per_code and code not in parent_order:
                parent_order.append(code)
                continue

            if code.endswith(".01"):
                parent = code[:-3]
                desc_code_by_parent[parent] = code
                continue

            if ".01." in code:
                parent = code.split(".01.")[0]
                group_codes_by_parent.setdefault(parent, []).append(code)

        code_ext = getattr(self.app, "last_code_extensions", {})

        out_csv = out_dir / "WBS_InputsParaDashboard.csv"
        n = 1
        while out_csv.exists():
            out_csv = out_dir / f"WBS_InputsParaDashboard({n}).csv"
            n += 1

        df_raw = self.app.df_raw.copy()
        lvl    = split_levels(df_raw, self.app.col_nivel)
        col_wbs_raw  = self.app.col_wbs
        col_desc_raw = self.app.col_desc

        try:
            all_rows = []

            for i in range(len(df_raw)):
                if lvl.iat[i] is not None and lvl.iat[i] < 10:
                    code = str(df_raw.at[i, col_wbs_raw]).strip() if pd.notna(df_raw.at[i, col_wbs_raw]) else ""
                    desc = str(df_raw.at[i, col_desc_raw]).strip() if pd.notna(df_raw.at[i, col_desc_raw]) else ""
                    if code:
                        row = {h: "" for h in headers}
                        row["wbs_codigo"] = code
                        row["descricao"]  = desc
                        all_rows.append(row)

            for parent in parent_order:
                blob = per_code.get(parent, {})
                meta = blob.get("meta", {})

                desc_code = desc_code_by_parent.get(parent) or code_ext.get(parent, {}).get("desc")
                if desc_code:
                    user_desc = ""
                    parent_idx = None
                    
                    for i in range(len(df_raw)):
                        if (lvl.iat[i] is not None and lvl.iat[i] < 10 and 
                            str(df_raw.at[i, col_wbs_raw]).strip() == parent):
                            parent_idx = i
                            break
                    
                    if parent_idx is not None:
                        next_idx = parent_idx + 1
                        if next_idx < len(df_raw) and lvl.iat[next_idx] == 10:
                            user_desc = str(df_raw.at[next_idx, col_desc_raw]).strip() if pd.notna(df_raw.at[next_idx, col_desc_raw]) else ""
                    
                    row = {h: "" for h in headers}
                    row["ifc_filename"] = ifc_filename
                    row["wbs_codigo"] = desc_code
                    row["descricao"]  = user_desc
                    all_rows.append(row)

                groups = blob.get("groups", {})
                group_map = code_ext.get(parent, {}).get("groups", {})
                
                if group_map:
                    for gval, elems in sorted(groups.items(),
                                              key=lambda kv: tuple(int(x) if x.isdigit() else x
                                                                   for x in group_map.get(kv[0], "zzz").split("."))):
                        element_code = group_map.get(gval, "")
                        for el in elems:
                            row = {h: "" for h in headers}
                            row["ifc_filename"]     = ifc_filename
                            row["wbs_codigo"]       = element_code
                            row["descricao"]        = (gval or "n/a")
                            row["ifc_class"]        = el.get("ifc_class", "n/a")
                            row["predefinedtype"]   = el.get("predefined", "n/a")
                            row["objecttype"]       = el.get("objecttype", "n/a")
                            row["material"]         = el.get("material", "n/a")
                            row["ifc_guid"]         = el.get("guid", "n/a")
                            row["buildingstorey"]   = el.get("buildingstorey", "n/a")
                            row["classification_code"] = el.get("classification_code", "n/a")
                            row["ifc_project"]      = meta.get("ifc_project", "n/a")
                            row["ifc_site"]         = meta.get("ifc_site", "n/a")
                            row["ifc_building"]     = meta.get("ifc_building", "n/a")
                            row["ifc_valor"]        = el.get("value", "")
                            row["unidade"]          = units_by_parent.get(parent, "")
                            all_rows.append(row)
                else:
                    for i, (gval, elems) in enumerate(groups.items(), start=1):
                        element_code = group_codes_by_parent.get(parent, [])
                        element_code = element_code[i-1] if i-1 < len(element_code) else f"{parent}.01.{i:02d}"
                        for el in elems:
                            row = {h: "" for h in headers}
                            row["ifc_filename"]     = ifc_filename
                            row["wbs_codigo"]       = element_code
                            row["descricao"]        = (gval or "n/a")
                            row["ifc_class"]        = el.get("ifc_class", "n/a")
                            row["predefinedtype"]   = el.get("predefined", "n/a")
                            row["objecttype"]       = el.get("objecttype", "n/a")
                            row["material"]         = el.get("material", "n/a")
                            row["ifc_guid"]         = el.get("guid", "n/a")
                            row["buildingstorey"]   = el.get("buildingstorey", "n/a")
                            row["classification_code"] = el.get("classification_code", "n/a")
                            row["ifc_project"]      = meta.get("ifc_project", "n/a")
                            row["ifc_site"]         = meta.get("ifc_site", "n/a")
                            row["ifc_building"]     = meta.get("ifc_building", "n/a")
                            row["ifc_valor"]        = el.get("value", "")
                            row["unidade"]          = units_by_parent.get(parent, "")
                            all_rows.append(row)

            all_rows.sort(key=lambda row: ReportPage._parse_wbs_code(row.get("wbs_codigo", "")))

            with open(out_csv, "w", newline="", encoding="utf-8-sig") as f:
                w = csv.DictWriter(f, fieldnames=headers)
                w.writeheader()
                for row in all_rows:
                    w.writerow(row)

            messagebox.showinfo("Extrair quantidades e Gerar WBS", f"Ficheiro CSV exportado com sucesso em:\n{out_csv}")

        except Exception as e:
            messagebox.showerror("Erro", f"Falha ao exportar CSV:\n{e}")