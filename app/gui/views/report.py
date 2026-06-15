import json
import os
from pathlib import Path
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from tkinter.scrolledtext import ScrolledText

import pandas as pd

from app.gui.wbs_helpers import find_wbs_columns, unpack_core_columns, split_levels
from app.core.structural_engine import load_and_migrate_rules, migrate_rule_v1_to_v2

import ifcopenshell


class ReportPage(tk.Frame):

    def __init__(self, notebook, app):
        super().__init__(notebook)
        self.app = app
        self.pack_propagate(False)

        def _s(v):
            try:    return v.get().strip()
            except: return ""

        wbs_guess = (getattr(app, "last_exported_wbs", "")
                     or _s(getattr(app, "wbs_xlsx_var", ""))
                     or getattr(app, "wbs_path_loaded", "") or "")
        map_guess = _s(getattr(app, "map_var", ""))
        ifc_guess = _s(getattr(app, "ifc_var", ""))
        out_guess = (_s(getattr(app, "out_var", ""))
                     or (str(Path(wbs_guess).parent) if wbs_guess else str(Path.home())))

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

        self.banner = tk.Label(
            self, anchor="w",
            text="Carregue o WBS (com descrições), o mapeamento (JSON) e o IFC.",
        )
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
            self, text="Exportar CSV detalhado",
            state="disabled", command=self.on_export_csv,
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
            try: w.configure(state=state)
            except Exception: pass
        try:
            self.e_out.configure(state="normal")
            self.b_out_pick.configure(state="normal")
        except Exception:
            pass

    def _autoload_from_previous_tab(self):
        app    = self.app
        loaded = []
        warns  = []

        try:
            wbs_path_str = (getattr(app, "wbs_xlsx_var", None) or tk.StringVar()).get()
            if wbs_path_str:
                self.wbs_var.set(wbs_path_str)
            if getattr(app, "df_raw", None) is None and wbs_path_str and Path(wbs_path_str).is_file():
                self._log("Carregando WBS automaticamente...\n")
                df = pd.read_excel(wbs_path_str, header=1)
                cols = find_wbs_columns(df)
                col_wbs, col_desc, col_nivel = unpack_core_columns(cols)
                if col_wbs and col_desc and col_nivel:
                    app.df_raw    = df.copy()
                    app.col_wbs   = col_wbs
                    app.col_desc  = col_desc
                    app.col_nivel = col_nivel
                    app.wbs_cols  = cols
                    app.df_desc0  = app.df_raw[app.col_desc].copy()
                    loaded.append(f"[OK] WBS: {len(df)} linhas carregadas")
                else:
                    warns.append("[X] WBS: colunas não encontradas")
            elif getattr(app, "df_raw", None) is not None:
                loaded.append("[OK] WBS (já em memória)")
            else:
                warns.append("[X] WBS não carregado")
        except Exception as e:
            warns.append(f"[X] Erro WBS: {e}")

        try:
            if getattr(app, "rules", None):
                loaded.append(f"[OK] Mapeamento: {len(app.rules)} códigos")
                map_path = (getattr(app, "map_var", None) or tk.StringVar()).get()
                self.map_var.set(map_path or "[carregado da aba anterior]")
            else:
                warns.append("[X] Mapeamento não carregado")
        except Exception as e:
            warns.append(f"[X] Erro mapeamento: {e}")

        try:
            ifc_path_str = (getattr(app, "ifc_var", None) or tk.StringVar()).get()
            if ifc_path_str:
                self.ifc_var.set(ifc_path_str)
            if getattr(app, "ifc_file", None) is None and ifc_path_str and Path(ifc_path_str).is_file():
                self._log("Carregando IFC automaticamente...\n")
                app.inv.open_ifc(ifc_path_str)
                app.ifc_path_loaded = ifc_path_str
                app.ifc_file        = app.inv.ifc_file
                loaded.append(f"[OK] IFC: {Path(ifc_path_str).name}")
            elif getattr(app, "ifc_file", None) is not None:
                loaded.append("[OK] IFC (já em memória)")
            else:
                warns.append("[X] IFC não carregado")
        except Exception as e:
            warns.append(f"[X] Erro IFC: {e}")

        try:
            if not self.out_var.get():
                self.out_var.set(
                    str(Path(self.wbs_var.get()).parent) if self.wbs_var.get() else str(Path.home())
                )
        except Exception:
            pass

        lines = ["=== Dados carregados automaticamente ===\n"] + loaded
        if warns:
            lines += ["\nAvisos:"] + warns + ["\nCarregue os dados em falta antes de gerar o WBS."]
        else:
            lines += ["\n[OK] Todos os dados estão prontos!",
                      "Pode clicar em 'Gerar WBS preenchido' quando estiver pronto.\n"]

        try:
            self.log.configure(state="normal")
            self.log.delete("1.0", "end")
            self.log.insert("end", "\n".join(lines))
            self.log.configure(state="disabled")
            self.log.see("end")
        except Exception:
            pass

    def browse_wbs(self):
        p = filedialog.askopenfilename(
            title="Escolher WBS (Excel)",
            filetypes=[("Excel (*.xlsx *.xls)", "*.xlsx *.xls")],
        )
        if p: self.wbs_var.set(p)

    def browse_map(self):
        p = filedialog.askopenfilename(
            title="Escolher mapeamento (JSON)",
            filetypes=[("JSON", "*.json")],
        )
        if p: self.map_var.set(p)

    def browse_ifc(self):
        p = filedialog.askopenfilename(
            title="Escolher IFC",
            filetypes=[("IFC", "*.ifc")],
        )
        if p: self.ifc_var.set(p)

    def browse_outdir(self):
        p = filedialog.askdirectory(title="Escolher pasta de saída")
        if p: self.out_var.set(p)

    def on_load_wbs(self):
        path = self.wbs_var.get().strip()
        if not path:
            messagebox.showinfo("Extrair quantidades", "Selecione um ficheiro Excel.")
            return
        if not Path(path).is_file():
            messagebox.showerror("Extrair quantidades", "Caminho inválido.")
            return
        try:
            df   = pd.read_excel(path, header=1)
            cols = find_wbs_columns(df)
            col_wbs, col_desc, col_nivel = unpack_core_columns(cols)
            if col_wbs and col_desc and col_nivel:
                self.app.df_raw    = df.copy()
                self.app.col_wbs   = col_wbs
                self.app.col_desc  = col_desc
                self.app.col_nivel = col_nivel
                self.app.wbs_cols  = cols
                self.app.df_desc0  = self.app.df_raw[self.app.col_desc].copy()
                messagebox.showinfo("Extrair quantidades", f"WBS carregado: {len(df)} linhas.")
                return
            raise RuntimeError(f"Colunas não encontradas. Disponíveis: {list(df.columns)}")
        except Exception as e:
            messagebox.showerror("Extrair quantidades", f"Falha a ler o WBS:\n{e}")

    def on_load_mapping(self):
        path = self.map_var.get().strip()
        if not path:
            messagebox.showinfo("Extrair quantidades", "Selecione um ficheiro JSON.")
            return
        if not Path(path).is_file():
            messagebox.showerror("Extrair quantidades", "Caminho inválido.")
            return
        try:
            with open(path, "r", encoding="utf-8") as f:
                data = json.load(f)

            rules = load_and_migrate_rules(data)
            if not rules:
                raise RuntimeError("Estrutura inválida ou vazia.")

            is_partial = data.get("partial", False)

            if is_partial:
                incomplete = []
                for code, rule in rules.items():
                    for m in rule.get("mappings", []):
                        qd = m.get("quantity_detail", {})
                        if not qd.get("pset"):
                            incomplete.append(code)
                            break
                if incomplete:
                    messagebox.showwarning(
                        "Mapeamento parcial",
                        f"Este é um mapeamento parcial.\n"
                        f"{len(incomplete)} código(s) têm pset de quantidade em falta.\n\n"
                        f"Esses códigos serão ignorados na extração ou gerarão avisos.\n\n"
                        f"Recomenda-se completar o mapeamento na aba 'Mapeamento IFC' primeiro.",
                    )

            self.app.rules = rules
            partial_note = " (parcial)" if is_partial else ""
            messagebox.showinfo(
                "Extrair quantidades",
                f"Mapeamento{partial_note} carregado: {len(rules)} código(s) WBS.",
            )
        except Exception as e:
            messagebox.showerror("Extrair quantidades", f"Falha a carregar o mapeamento:\n{e}")

    def on_load_ifc(self):
        path = self.ifc_var.get().strip()
        if not path or not os.path.isfile(path):
            messagebox.showerror("Extrair quantidades", "Selecione primeiro um ficheiro IFC válido.")
            return
        try:
            self.app.inv.open_ifc(path)
            self.app.ifc_path_loaded = path
            self.app.ifc_file        = self.app.inv.ifc_file
            messagebox.showinfo("Extrair quantidades", "IFC carregado com sucesso.")
        except Exception as e:
            messagebox.showerror("Extrair quantidades", f"Falha a abrir o IFC:\n{e}")

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
        if not self.app.ifc_var.get().strip():      missing.append("IFC")
        if not self.app.out_var.get().strip():      missing.append("Pasta de saída")
        if not self.app.rules:                      missing.append("Mapeamento (JSON)")
        if missing:
            messagebox.showinfo("Dados em falta", "Indique: " + ", ".join(missing) + ".")
            self.run_btn.configure(state="normal", text="Gerar WBS preenchido")
            return

        def done(msg):
            messagebox.showinfo("Extrair quantidades", msg)
            self.run_btn.configure(state="normal", text="Gerar WBS preenchido")

        if hasattr(self.app, "run_generate_report"):
            self.after(10, lambda: self.app.run_generate_report(self.log, on_finish=done))
            return

        self._log("Pipeline não encontrado. (fallback)\n")
        done("Fluxo testado.")

    def on_export_csv(self):
        import csv as csv_mod

        cache = getattr(self.app, "_last_csv_cache", None)
        if not isinstance(cache, dict):
            messagebox.showinfo("Extrair quantidades", "Gere primeiro os ficheiros de output.")
            return

        headers         = [h for h in cache["headers"] if h != "wbs_group"]
        per_code        = cache["per_code"]
        units_by_parent = cache.get("code_to_unit", {})
        wbs_rows        = cache.get("wbs_rows", [])
        code_ext        = getattr(self.app, "last_code_extensions", {})

        out_dir      = Path(self.app.out_var.get().strip() or Path.home())
        ifc_path_str = self.ifc_var.get().strip() or getattr(self.app, "ifc_path_loaded", "") or ""
        ifc_stem     = Path(ifc_path_str).stem if ifc_path_str else "output"
        ifc_filename = ifc_stem

        out_csv = out_dir / f"ElementosQuantificados_{ifc_stem}.csv"
        n = 1
        while out_csv.exists():
            out_csv = out_dir / f"ElementosQuantificados_{ifc_stem}({n}).csv"
            n += 1

        try:
            all_rows = []

            found_codes = {
                code for code, blob in per_code.items()
                if not blob.get("no_elements")
            }

            ancestor_codes = set()
            for code in found_codes:
                parts = code.split(".")
                for k in range(1, len(parts)):
                    ancestor_codes.add(".".join(parts[:k]))

            for entry in wbs_rows:
                code = entry.get("wbs_codigo", "")
                if code in ancestor_codes:
                    row = {h: "" for h in headers}
                    row["wbs_codigo"] = code
                    row["descricao"]  = entry.get("descricao", "")
                    all_rows.append(row)

            parent_order = sorted(
                found_codes,
                key=lambda c: ReportPage._parse_wbs_code(c)
            )

            for parent in parent_order:
                blob      = per_code.get(parent, {})
                meta      = blob.get("meta", {})
                desc_ext  = code_ext.get(parent, {}).get("desc", f"{parent}.01")
                group_map = code_ext.get(parent, {}).get("groups", {})

                if blob.get("no_elements"):
                    continue

                user_desc = next(
                    (e.get("descricao", "") for e in wbs_rows
                     if e.get("wbs_codigo") == parent),
                    ""
                )
                row = {h: "" for h in headers}
                row["ifc_filename"] = ifc_filename
                row["wbs_codigo"]   = desc_ext
                row["descricao"]    = user_desc
                all_rows.append(row)

                groups = blob.get("groups", {})
                if group_map:
                    sorted_groups = sorted(
                        groups.items(),
                        key=lambda kv: ReportPage._parse_wbs_code(
                            group_map.get(kv[0], "zzz"))
                    )
                else:
                    sorted_groups = list(groups.items())

                for gi, (gval, elems) in enumerate(sorted_groups, start=1):
                    element_code = group_map.get(gval) or f"{desc_ext}.{gi:02d}"
                    for el in elems:
                        row = {h: "" for h in headers}
                        row["ifc_filename"]        = ifc_filename
                        row["wbs_codigo"]          = element_code
                        row["descricao"]           = gval or "n/a"
                        row["ifc_class"]           = el.get("ifc_class",           "n/a")
                        row["predefinedtype"]       = el.get("predefined",          "n/a")
                        row["objecttype"]          = el.get("objecttype",           "n/a")
                        row["material"]            = el.get("material",             "n/a")
                        row["ifc_guid"]            = el.get("guid",                 "n/a")
                        row["buildingstorey"]      = el.get("buildingstorey",       "n/a")
                        row["classification_code"] = el.get("classification_code",  "n/a")
                        row["ifc_project"]         = meta.get("ifc_project",        "n/a")
                        row["ifc_site"]            = meta.get("ifc_site",           "n/a")
                        row["ifc_building"]        = meta.get("ifc_building",       "n/a")
                        row["ifc_valor"]           = el.get("value",                "")
                        row["unidade"]             = units_by_parent.get(parent,    "")
                        row["qty_type"]            = el.get("qty_type",             "prop")
                        all_rows.append(row)

            all_rows.sort(key=lambda r: ReportPage._parse_wbs_code(r.get("wbs_codigo", "")))

            with open(out_csv, "w", newline="", encoding="utf-8-sig") as f:
                w = csv_mod.DictWriter(f, fieldnames=headers)
                w.writeheader()
                for row in all_rows:
                    w.writerow(row)

            messagebox.showinfo("Extrair quantidades", "Um ficheiro exportado com sucesso.")

        except Exception as e:
            messagebox.showerror("Erro", f"Falha ao exportar CSV:\n{e}")

    def _log(self, text: str):
        self.log.configure(state="normal")
        self.log.insert("end", text + ("\n" if not text.endswith("\n") else ""))
        self.log.see("end")
        self.log.configure(state="disabled")

    @staticmethod
    def _parse_wbs_code(code: str) -> tuple:
        if not code or not isinstance(code, str):
            return (float("inf"),)
        parts = []
        for part in code.split("."):
            part = part.strip()
            parts.append(int(part) if part.isdigit() else part)
        return tuple(parts)
