# app/gui/views/qty.py
import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import re
import pandas as pd
import json
from pathlib import Path

from app.gui.wbs_helpers import (
    detect_relevant_leaves,
    find_level10_text,
)

from app.core.structural_engine import IFCInvestigator


class QtyPage(tk.Frame):

    def __init__(self, master, app):
        super().__init__(master)
        self.app = app

        self.relevant_codes = []
        self.relevant_set = set()
        self.code_to_desc = {}
        self.selected_code = None
        self.prop_rows = []
        self.path_stack = []
        self.ifc_model = None
        self.predefs_by_class = {}

        self._build_ui()
        self.after_idle(lambda: self.refresh_items(silent=True))

    def _build_ui(self):
        self.grid_columnconfigure(0, weight=2, minsize=600)
        self.grid_columnconfigure(1, weight=1, minsize=480)
        self.grid_rowconfigure(0, weight=1)

        left = tk.Frame(self, bd=1, relief="groove")
        left.grid(row=0, column=0, sticky="nsew", padx=8, pady=8)
        left.grid_propagate(False)

        top_left = tk.Frame(left)
        top_left.pack(fill="x", padx=8, pady=(8, 4))

        self.mode = "home"
        self.src_frame = tk.LabelFrame(left, text="WBS de entrada (com descrições)", padx=8, pady=8)
        self.src_frame.pack(fill="x", padx=8, pady=(0, 6))

        row_src = tk.Frame(self.src_frame)
        row_src.pack(fill="x")

        tk.Label(row_src, text="WBS (Excel):").pack(side="left")

        self.wbs_entry = tk.Entry(row_src, textvariable=self.app.wbs_xlsx_var, width=40)
        self.wbs_entry.pack(side="left", padx=6, fill="x", expand=True)

        self.wbs_browse_btn = tk.Button(
            row_src, text="Procurar…",
            command=lambda: self.app.pick_file(
                self.app.wbs_xlsx_var, "Escolher WBS (Excel)",
                [("Excel (*.xlsx *.xls)", "*.xlsx *.xls")]
            )
        )
        self.wbs_browse_btn.pack(side="left", padx=(0, 4))
        
        self.wbs_load_btn = tk.Button(row_src, text="Carregar", command=self._load_wbs_from_file)
        self.wbs_load_btn.pack(side="left")

        nav = tk.Frame(left); nav.pack(fill="x", padx=8)
        self.path_label = tk.Label(nav, text="Caminho: – (Nível 1)")
        self.path_label.pack(side="left", anchor="w")

        nav_btns = tk.Frame(left); nav_btns.pack(fill="x", padx=8, pady=(4, 4))
        self.btn_up   = tk.Button(nav_btns, text="Nível acima",  command=self.on_back, state="disabled")
        self.btn_down = tk.Button(nav_btns, text="Nível abaixo", command=self.on_next, state="disabled")
        self.btn_up.pack(side="left"); self.btn_down.pack(side="left", padx=(6, 0))

        list_wrap = tk.Frame(left); list_wrap.pack(fill="both", expand=True, padx=8, pady=(0, 8))
        self.listbox = tk.Listbox(list_wrap, height=22, width=56, exportselection=False)
        self.listbox.pack(side="left", fill="both", expand=True)
        yscroll = tk.Scrollbar(list_wrap, orient="vertical", command=self.listbox.yview)
        yscroll.pack(side="right", fill="y")
        self.listbox.configure(yscrollcommand=yscroll.set)
        self.listbox.bind("<<ListboxSelect>>", self.on_select_list)

        actions_left = tk.Frame(left)
        actions_left.pack(fill="x", padx=8, pady=(0, 8))

        tk.Button(
            actions_left,
            text="Limpar regra do item",
            command=self.clear_rule_current
        ).pack(side="right")

        right_wrapper = tk.Frame(self)
        right_wrapper.grid(row=0, column=1, sticky="nsew", padx=(0, 8), pady=8)
        right_wrapper.grid_columnconfigure(0, weight=1)
        right_wrapper.grid_rowconfigure(0, weight=1)

        canvas = tk.Canvas(right_wrapper, bg="white", highlightthickness=0, bd=0)
        canvas.grid(row=0, column=0, sticky="nsew")
        scrollbar = tk.Scrollbar(right_wrapper, orient="vertical", command=canvas.yview)
        scrollbar.grid(row=0, column=1, sticky="ns")
        canvas.configure(yscrollcommand=scrollbar.set)

        right = tk.Frame(canvas, bg="white")
        canvas_window = canvas.create_window(0, 0, window=right, anchor="nw")

        def _on_mousewheel(event):
            canvas.yview_scroll(int(-1*(event.delta/120)), "units")
        canvas.bind("<MouseWheel>", _on_mousewheel)
        canvas.bind("<Button-4>", lambda e: canvas.yview_scroll(-1, "units"))
        canvas.bind("<Button-5>", lambda e: canvas.yview_scroll(1, "units"))

        def on_frame_configure(event=None):
            canvas.configure(scrollregion=canvas.bbox("all"))
            canvas_width = canvas.winfo_width()
            if canvas_width > 1:
                canvas.itemconfig(canvas_window, width=canvas_width)
            bbox = canvas.bbox("all")
            if bbox:
                content_height = bbox[3] - bbox[1]
                canvas_height = canvas.winfo_height()
                if content_height > canvas_height:
                    scrollbar.grid()
                else:
                    scrollbar.grid_remove()

        right.bind("<Configure>", on_frame_configure)
        canvas.bind("<Configure>", on_frame_configure)
        right.grid_columnconfigure(0, weight=1)

        ifc = tk.LabelFrame(right, text="IFC de entrada", padx=8, pady=8, bg="white")
        ifc.grid(row=0, column=0, sticky="we", padx=8, pady=(8, 4))
        ifc.grid_columnconfigure(1, weight=1)

        tk.Label(ifc, text="IFC:", bg="white").grid(row=0, column=0, sticky="w", padx=(0, 6), pady=4)
        tk.Entry(ifc, textvariable=self.app.ifc_var, width=40)\
            .grid(row=0, column=1, sticky="we", padx=(0, 6), pady=4)
        tk.Button(
            ifc, text="Procurar…",
            command=lambda: self.app.pick_file(self.app.ifc_var, "Selecionar IFC", [("IFC (*.ifc)", "*.ifc")])
        ).grid(row=0, column=2, sticky="w", padx=(0, 4), pady=4)
        tk.Button(ifc, text="Carregar", command=self._load_ifc)\
            .grid(row=0, column=3, sticky="w", pady=4)

        self.title_var = tk.StringVar(value="Selecione um item (folha) à esquerda…")
        tk.Label(right, textvariable=self.title_var, font=("Segoe UI", 11, "bold"), bg="white")\
            .grid(row=1, column=0, sticky="w", padx=8, pady=(8, 4))

        desc_box = tk.LabelFrame(right, text="Descrição do utilizador (WBS)", padx=8, pady=8, bg="white")
        desc_box.grid(row=2, column=0, sticky="we", padx=8, pady=(4, 6))
        self.user_desc_lbl = tk.Label(desc_box, text="—", fg="#555", wraplength=400, justify="left", bg="white")
        self.user_desc_lbl.pack(fill="x")

        filt = tk.LabelFrame(right, text="Filtro (obrigatório)", padx=8, pady=8, bg="white")
        filt.grid(row=3, column=0, sticky="we", padx=8, pady=(4, 4))
        filt.grid_columnconfigure(1, weight=1)

        tk.Label(filt, text="IfcClass:", bg="white").grid(row=0, column=0, sticky="e", padx=(0,6), pady=4)
        self.ifc_class = ttk.Combobox(filt, state="readonly", values=[])
        self.ifc_class.grid(row=0, column=1, sticky="we", pady=4)
        self.ifc_class.bind("<<ComboboxSelected>>", self._on_class_selected)

        tk.Label(filt, text="PredefinedType:", bg="white").grid(row=1, column=0, sticky="e", padx=(0,6), pady=4)
        self.predef = ttk.Combobox(filt, state="readonly", values=[])
        self.predef.grid(row=1, column=1, sticky="we", pady=4)
        self.predef.bind("<<ComboboxSelected>>", self._toggle_object_type)
        self.predef.bind("<KeyRelease>", self._toggle_object_type)

        tk.Label(filt, text="ObjectType (se USERDEFINED):", bg="white").grid(row=2, column=0, sticky="e", padx=(0,6), pady=4)
        self.objtype = tk.Entry(filt)
        self.objtype.grid(row=2, column=1, sticky="we", pady=4)

        mat = tk.LabelFrame(right, text="Material (opcional)", padx=8, pady=8, bg="white")
        mat.grid(row=4, column=0, sticky="we", padx=8, pady=(4, 4))
        mat.grid_columnconfigure(1, weight=1)
        tk.Label(mat, text="Material:", bg="white").grid(row=0, column=0, sticky="e", padx=(0,6), pady=4)
        self.material = ttk.Combobox(mat, state="readonly", values=[])
        self.material.grid(row=0, column=1, sticky="we", pady=4)

        extra = tk.LabelFrame(right, text="Filtros adicionais (opcional)", padx=8, pady=8, bg="white")
        extra.grid(row=5, column=0, sticky="we", padx=8, pady=(4, 4))
        extra.grid_columnconfigure(0, weight=1)

        hdr = tk.Frame(extra, bg="white")
        hdr.pack(fill="x", padx=0, pady=(0, 6))
        hdr.grid_columnconfigure(0, weight=2, uniform="flt")
        hdr.grid_columnconfigure(1, weight=2, uniform="flt")
        hdr.grid_columnconfigure(2, weight=1, uniform="flt")
        hdr.grid_columnconfigure(3, weight=0)
        tk.Label(hdr, text="Grupo (Pset)", font=("Segoe UI", 9, "bold"), bg="white")\
            .grid(row=0, column=0, sticky="w", padx=(0, 6))
        tk.Label(hdr, text="Propriedade",  font=("Segoe UI", 9, "bold"), bg="white")\
            .grid(row=0, column=1, sticky="w", padx=(0, 6))
        tk.Label(hdr, text="Valor",        font=("Segoe UI", 9, "bold"), bg="white")\
            .grid(row=0, column=2, sticky="w", padx=(0, 6))
        tk.Label(hdr, text="", bg="white").grid(row=0, column=3, sticky="e")

        self.extra_container = tk.Frame(extra, bg="white")
        self.extra_container.pack(fill="x", expand=True)
        self.prop_rows = []
        self.btn_add_filter = tk.Button(extra, text="+ Adicionar filtro", command=self.add_prop_row)
        self.btn_add_filter.pack(anchor="e", pady=(6,0))

        qty = tk.LabelFrame(right, text="Quantidade (obrigatório)", padx=8, pady=8, bg="white")
        qty.grid(row=6, column=0, sticky="we", padx=8, pady=(4, 4))
        qty.grid_columnconfigure(1, weight=1)
        tk.Label(qty, text="Pset:", bg="white").grid(row=0, column=0, sticky="e", padx=(0,6), pady=4)
        self.q_pset = tk.Entry(qty); self.q_pset.grid(row=0, column=1, sticky="we", pady=4)
        tk.Label(qty, text="Propriedade:", bg="white").grid(row=1, column=0, sticky="e", padx=(0,6), pady=4)
        self.q_prop = tk.Entry(qty); self.q_prop.grid(row=1, column=1, sticky="we", pady=4)

        agr = tk.LabelFrame(right, text="Propriedade de Agrupamento (obrigatório)", padx=8, pady=8, bg="white")
        agr.grid(row=7, column=0, sticky="we", padx=8, pady=(4, 4))
        agr.grid_columnconfigure(1, weight=1)
        tk.Label(agr, text="Pset:", bg="white").grid(row=0, column=0, sticky="e", padx=(0,6), pady=4)
        self.agr_pset = tk.Entry(agr); self.agr_pset.grid(row=0, column=1, sticky="we", pady=4)
        tk.Label(agr, text="Propriedade:", bg="white").grid(row=1, column=0, sticky="e", padx=(0,6), pady=4)
        self.agr_prop = tk.Entry(agr); self.agr_prop.grid(row=1, column=1, sticky="we", pady=4)

        actions = tk.Frame(right, bg="white")
        actions.grid(row=8, column=0, sticky="we", padx=8, pady=(6, 8))
        actions.grid_columnconfigure(0, weight=1)
        actions.grid_columnconfigure(1, weight=1)
        self.btn_save_rule = tk.Button(actions, text="Guardar regra", command=self.save_rule_current)
        self.btn_save_rule.grid(row=0, column=0, sticky="we", padx=(0, 4))
        tk.Button(actions, text="Salvar e exportar", command=self.save_rules_dialog)\
            .grid(row=0, column=1, sticky="we", padx=(4, 0))

        self._toggle_object_type()
        self._apply_mode()
        self._set_edit_enabled(False)

    def _load_ifc(self):
            path = (self.app.ifc_var.get() or "").strip()
            if not path:
                messagebox.showwarning("Mapeamento IFC", "Selecione primeiro o ficheiro do modelo IFC (.ifc).")
                return

            if not path.lower().endswith(".ifc"):
                messagebox.showerror("Mapeamento IFC", "Apenas ficheiros .ifc são aceites.")
                return

            try:
                import ifcopenshell
            except Exception:
                messagebox.showerror("Mapeamento IFC", "Biblioteca 'ifcopenshell' em falta. Instale com: pip install ifcopenshell")
                return

            try:
                model = ifcopenshell.open(path)
            except Exception as e:
                messagebox.showerror("Mapeamento IFC", f"Falha a abrir o IFC:\n{e}")
                return

            from collections import defaultdict
            class_to_predefs = defaultdict(set)

            try:
                products = model.by_type("IfcProduct")
            except Exception:
                products = []

            for el in products:
                cls = el.is_a()
                pre = getattr(el, "PredefinedType", None)
                if isinstance(pre, str):
                    pre = pre.upper().strip()
                if pre is not None and pre != "":
                    class_to_predefs[cls].add(pre)
                else:
                    class_to_predefs[cls].add("NOTDEFINED")

            self.ifc_model = model
            self.predefs_by_class = {k: sorted(v) for k, v in class_to_predefs.items()}

            self.app.inv.ifc_file = model
            materials_dict = self.app.inv.extract_all_materials()
            material_categories = sorted(materials_dict.keys())

            classes = sorted(self.predefs_by_class.keys())
            self.ifc_class.configure(values=classes)
            self.ifc_class.set("")
            self.predef.configure(values=[]); self.predef.set("")

            self.material.configure(values=material_categories)
            self.material.set("")
            
            self._toggle_object_type()

            messagebox.showinfo("Mapeamento IFC", f"IFC carregado com sucesso.")

    def _on_class_selected(self, *_):
        cls = (self.ifc_class.get() or "").strip()
        opts = self.predefs_by_class.get(cls, [])
        self.predef.configure(values=opts)
        self.predef.set("" if not opts else opts[0])
        self._toggle_object_type()

    @staticmethod
    def _tokens(code: str):
        return [p for p in str(code).split(".") if p]

    def _current_level(self):
        return 1 if not self.path_stack else len(self._tokens(self.path_stack[-1])) + 1

    def _desc_for(self, code: str):
        return self.code_to_desc.get(code, "")

    def _is_leaf(self, code: str):
        return code in self.relevant_set

    def _has_child_in_relevant(self, code: str):
        pref = f"{code}."
        return any(c.startswith(pref) and c != code for c in self.relevant_codes)

    def _candidates_for_level(self, level: int, prefix: str | None):
        cands = set()
        for leaf in self.relevant_codes:
            toks = self._tokens(leaf)
            if level <= len(toks):
                cand = ".".join(toks[:level])
                if prefix is None:
                    cands.add(cand)
                else:
                    if cand.startswith(prefix + "."):
                        cands.add(cand)
        return sorted(cands)

    def _render_list(self):
        lvl = self._current_level()
        prefix = self.path_stack[-1] if self.path_stack else None
        cands = self._candidates_for_level(lvl, prefix)
        self.listbox.delete(0, "end")
        for code in cands:
            desc = self._desc_for(code)
            label = f"{code} — {desc}" if desc else code
            self.listbox.insert("end", label)

        path_txt = " > ".join(self.path_stack) if self.path_stack else "—"
        self.path_label.configure(text=f"Caminho: {path_txt} (Nível {lvl})")

        self.btn_up.configure(state=("normal" if self.path_stack else "disabled"))
        sel = self.listbox.curselection()
        if sel:
            code = self.listbox.get(sel[0]).split(" — ")[0]
            can_down = self._has_child_in_relevant(code) and not self._is_leaf(code)
        else:
            can_down = False
        self.btn_down.configure(state=("normal" if can_down else "disabled"))

        self.on_select_list()
        self._set_edit_enabled(self._is_leaf_selected())
    
    def set_mode(self, source: str):
            self.mode = "wbs" if source == "wbs" else "home"
            self._apply_mode()

            if self.mode == "wbs":
                self._disable_wbs_upload()
                self.after(100, lambda: self.refresh_items(silent=True))
            else:
                self._enable_wbs_upload()

    def _apply_mode(self):

        try:
            if not self.src_frame.winfo_ismapped():
                self.src_frame.pack(fill="x", padx=8, pady=(0, 6))
        except Exception:
            pass

        self.clear_form()
        self._set_edit_enabled(False)
        self._render_list()
    
    def _disable_wbs_upload(self):

        try:
            self.wbs_entry.configure(state="disabled", bg="#F0F0F0")
            self.wbs_browse_btn.configure(state="disabled")
            self.wbs_load_btn.configure(state="disabled")
            self.src_frame.configure(text="WBS de entrada (carregado da aba anterior)")
        except Exception as e:
            print(f"[WARN] Erro ao desabilitar upload WBS: {e}")
    
    def _enable_wbs_upload(self):
        """Habilita a zona de upload do WBS (modo normal)"""
        try:
            self.wbs_entry.configure(state="normal", bg="white")
            self.wbs_browse_btn.configure(state="normal")
            self.wbs_load_btn.configure(state="normal")
            self.src_frame.configure(text="WBS de entrada (com descrições)")
        except Exception as e:
            print(f"[WARN] Erro ao habilitar upload WBS: {e}")


    def _load_wbs_from_file(self):
        import pandas as pd
        from app.gui.wbs_helpers import find_wbs_columns, split_levels

        path = (self.app.wbs_xlsx_var.get() or "").strip()
        if not path:
            messagebox.showwarning("Mapeamento IFC", "Selecione primeiro o ficheiro WBS com descrições (.xls or .xlsx).")
            return

        try:

            df = pd.read_excel(path, header=1)
            col_wbs, col_desc, col_nivel = find_wbs_columns(df)

            if not all([col_wbs, col_desc, col_nivel]):
                df = pd.read_excel(path, header=0)
                col_wbs, col_desc, col_nivel = find_wbs_columns(df)

            df[col_nivel] = split_levels(df, col_nivel)

            self.app.df_raw    = df
            self.app.df_desc0  = pd.Series([None]*len(df))
            self.app.col_nivel = col_nivel
            self.app.col_wbs   = col_wbs
            self.app.col_desc  = col_desc

            self.app.wbs_finalized = True

            self.refresh_items()
            messagebox.showinfo("Mapeamento IFC", "WBS com descrições carregado com sucesso.")

        except Exception as e:
            messagebox.showerror("Erro", f"Falha a ler o WBS:\n{e}")

    def refresh_items(self, silent: bool = False):
        import pandas as pd
        from app.gui.wbs_helpers import detect_relevant_leaves

        if self.mode == "wbs":
            if not getattr(self.app, "wbs_finalized", False):
                if not silent:
                    messagebox.showinfo("Mapeamento IFC", "Finalize o WBS na aba 'WBS e descrição' antes de prosseguir.")
                return
            if self.app.df_raw is None:
                if not silent:
                    messagebox.showinfo("Mapeamento IFC", "Carregue e edite um WBS na aba 'WBS e descrição' antes.")
                return
        else:
            if self.app.df_raw is None:
                if not silent:
                    messagebox.showinfo("Mapeamento IFC", "Carregue um WBS (Excel) acima.")
                return

        pairs = detect_relevant_leaves(
            self.app.df_raw,
            self.app.col_nivel,
            self.app.col_wbs,
            self.app.col_desc,
            self.app.df_desc0,
        )
        self.relevant_codes = [code for code, _ in pairs]
        self.relevant_set = set(self.relevant_codes)

        self.code_to_desc.clear()
        df = self.app.df_raw
        col_wbs, col_desc = self.app.col_wbs, self.app.col_desc
        for _, row in df.iterrows():
            code = row[col_wbs]
            if isinstance(code, str) and code.strip():
                val = row[col_desc]
                self.code_to_desc[str(code).strip()] = "" if pd.isna(val) else str(val).strip()

        self.path_stack = []
        self.selected_code = None
        self.title_var.set("Selecione um item (folha) à esquerda…")
        self.user_desc_lbl.configure(text="—")
        self.clear_form()
        self._set_edit_enabled(False)
        self._render_list()

    def on_select_list(self, *_):
        sel = self.listbox.curselection()
        if not sel:
            self.btn_down.configure(state="disabled")
            self.selected_code = None
            self.title_var.set("Selecione um item (folha) à esquerda…")
            self.user_desc_lbl.configure(text="—")
            self.clear_form()
            self._set_edit_enabled(False)
            return

        code = self.listbox.get(sel[0]).split(" — ")[0]
        can_down = self._has_child_in_relevant(code) and not self._is_leaf(code)
        self.btn_down.configure(state=("normal" if can_down else "disabled"))

        if self._is_leaf(code):
            self.selected_code = code
            base = self._desc_for(code)
            user_txt = find_level10_text(
                self.app.df_raw, self.app.col_nivel, self.app.col_wbs, self.app.col_desc, code
            )
            self.title_var.set(f"Item selecionado: {code} — {base}")
            self.user_desc_lbl.configure(text=(user_txt or "—"))

            self._set_edit_enabled(True)
            self.load_rule_into_form(self.app.rules.get(code))

        else:
            self.selected_code = None
            self.title_var.set(f"Secção: {code} — {self._desc_for(code)}")
            self.user_desc_lbl.configure(text="—")
            self.clear_form()
            self._set_edit_enabled(False)


    def on_next(self):
        sel = self.listbox.curselection()
        if not sel:
            messagebox.showwarning("Mapeamento IFC", "Selecione um item na lista para avançar.")
            return
        code = self.listbox.get(sel[0]).split(" — ")[0]
        if self._is_leaf(code) or not self._has_child_in_relevant(code):
            messagebox.showinfo("Mapeamento IFC", "Já está no último nível desta ramificação.")
            return
        self.path_stack.append(code)
        self._render_list()

    def on_back(self):
        if not self.path_stack:
            return
        self.path_stack.pop()
        self._render_list()
             
    def add_prop_row(self, preset=None):
        row = tk.Frame(self.extra_container)
        row.pack(fill="x", pady=2)

        row.grid_columnconfigure(0, weight=2, uniform="flt")
        row.grid_columnconfigure(1, weight=2, uniform="flt")
        row.grid_columnconfigure(2, weight=1, uniform="flt")
        row.grid_columnconfigure(3, weight=0)

        e_pset = tk.Entry(row)
        e_prop = tk.Entry(row)
        e_val  = tk.Entry(row)
        btn_rm = tk.Button(row, text="Remover", command=lambda r=row: self._remove_prop_row(r))

        e_pset.grid(row=0, column=0, sticky="we", padx=(0, 6))
        e_prop.grid(row=0, column=1, sticky="we", padx=(0, 6))
        e_val .grid(row=0, column=2, sticky="we", padx=(0, 6))
        btn_rm.grid(row=0, column=3, sticky="e")

        if preset:
            e_pset.insert(0, preset.get("pset", ""))
            e_prop.insert(0, preset.get("prop", ""))
            e_val .insert(0, preset.get("value", ""))

        self.prop_rows.append((row, e_pset, e_prop, e_val, btn_rm))

    def _remove_prop_row(self, row_widget):
        for i, (r, *_rest) in enumerate(self.prop_rows):
            if r is row_widget:
                r.destroy()
                del self.prop_rows[i]
                break

    def _toggle_object_type(self, *_):
        use = (self.predef.get() or "").strip().upper() == "USERDEFINED"
        try:
            self.objtype.configure(state=("normal" if use and self.btn_save_rule["state"] != "disabled" else "disabled"))
        except Exception:
            pass
        if not use:
            try:
                self.objtype.delete(0, "end")
            except Exception:
                pass
    
    def _letters_only(self, s: str) -> str:
            return re.sub(r"[^A-Za-z]", "", s or "")

    def _filter_letters(self, widget, *, pascal=False, upper=False):
        s = widget.get()
        s2 = self._letters_only(s)
        if upper:
            s2 = s2.upper()
        elif pascal and s2:
            s2 = s2[0].upper() + s2[1:]
        if s2 != s:
            pos = widget.index("insert")
            widget.delete(0, "end")
            widget.insert(0, s2)
            try:
                widget.icursor(min(pos, len(s2)))
            except Exception:
                pass

    def clear_form(self):
            try: self.ifc_class.set("")
            except Exception: pass
            try: self.predef.set("")
            except Exception: pass
            try: self.material.set("")
            except Exception: pass

            for name in ("objtype", "q_pset", "q_prop", "agr_pset", "agr_prop"):
                w = getattr(self, name, None)
                if w is not None:
                    try: w.delete(0, "end")
                    except Exception: pass

            if not hasattr(self, "prop_rows"):
                self.prop_rows = []
            for row, *_ in list(self.prop_rows):
                try: row.destroy()
                except Exception: pass
            self.prop_rows.clear()

            self._toggle_object_type()


    def load_rule_into_form(self, rule: dict | None):
            self.clear_form()
            if not rule:
                return

            f = rule.get("filter", {}) or {}
            q = rule.get("quantity", {}) or {}
            mat = rule.get("material", "")
            agr = rule.get("agrupamento", {}) or {}

            if f.get("ifc_class"):
                self.ifc_class.set(f["ifc_class"])
                self._on_class_selected()
            if f.get("predefined"):
                self.predef.set(f["predefined"])
            self._toggle_object_type()
            if (f.get("predefined","").strip().upper() == "USERDEFINED") and f.get("object_type"):
                self.objtype.insert(0, f["object_type"])

            for p in (f.get("props") or []):
                self.add_prop_row(preset=p)

            if q.get("pset"): self.q_pset.insert(0, q["pset"])
            if q.get("prop"): self.q_prop.insert(0, q["prop"])

            if mat:
                self.material.set(mat)

            if agr.get("pset"):
                self.agr_pset.insert(0, agr["pset"])
            if agr.get("prop"):
                self.agr_prop.insert(0, agr["prop"])


    def clear_rule_current(self):
        if not self.selected_code or not self._is_leaf_selected():
            messagebox.showinfo("Mapeamento IFC", "Selecione um item (folha) primeiro.")
            return

        self.app.rules.pop(self.selected_code, None)
        self.clear_form()
        self._set_edit_enabled(True)
        try:
            self.ifc_class.focus_set()
        except Exception:
            pass

        messagebox.showinfo("Mapeamento IFC", f"Regra do item {self.selected_code} removida. Pode definir uma nova regra.")

    def save_rule_current(self):
            try:
                if not self._is_leaf_selected():
                    messagebox.showwarning("Mapeamento IFC", "Selecione um item do último nível (folha) na lista à esquerda.")
                    return

                ic = self.ifc_class.get().strip()
                pd = self.predef.get().strip().upper()
                mat = self.material.get().strip()
                
                if not ic or not pd:
                    messagebox.showerror("Mapeamento IFC", "IFC Class e PredefinedType são obrigatórios.")
                    return

                ot = self.objtype.get().strip() if pd == "USERDEFINED" else ""
                if pd == "USERDEFINED" and not ot:
                    messagebox.showerror("Mapeamento IFC", "ObjectType é obrigatório quando PredefinedType = USERDEFINED.")
                    return

                props = []
                for row, e_pset, e_prop, e_val, _btn in getattr(self, "prop_rows", []):
                    pset = (e_pset.get() or "").strip()
                    prop = (e_prop.get() or "").strip()
                    val_txt = (e_val.get() or "").strip()
                    if not pset and not prop and not val_txt:
                        continue
                    if not pset or not prop:
                        messagebox.showerror("Mapeamento IFC", "Complete Pset e Propriedade (ou remova a linha).")
                        return
                    props.append({"pset": pset, "prop": prop, "value": self._parse_value_token(val_txt)})

                qpset = self.q_pset.get().strip()
                qprop = self.q_prop.get().strip()
                if not qpset or not qprop:
                    messagebox.showerror("Mapeamento IFC", "Quantidade: Pset e Propriedade são obrigatórios.")
                    return

                agrpset = self.agr_pset.get().strip()
                agrprop = self.agr_prop.get().strip()
                
                rule = {
                    "filter": {"ifc_class": ic, "predefined": pd},
                    "material": mat,
                    "quantity": {"pset": qpset, "prop": qprop},
                    "agrupamento": {"pset": agrpset, "prop": agrprop}
                }
                
                if pd == "USERDEFINED":
                    rule["filter"]["object_type"] = ot
                
                if props:
                    rule["filter"]["props"] = props

                if not agrpset or not agrprop:
                    messagebox.showerror("Mapeamento IFC", 
                                       "Agrupamento é obrigatório.\n"
                                       "Preencha Pset e Propriedade de agrupamento.")
                    return

                self.app.rules[self.selected_code] = rule
                messagebox.showinfo("Mapeamento IFC", f"Regra guardada para {self.selected_code}.")

            except Exception as e:
                messagebox.showerror("Erro", f"Falha a guardar a regra:\n{e}")
        
    def _normalize_rule(self, rule: dict) -> dict:
            f = rule.get("filter", {})
            return {
                "filter": {
                    "ifc_class": f.get("ifc_class", "").strip(),
                    "predefined": f.get("predefined", "").strip(),
                    "object_type": f.get("object_type", "").strip(),
                    "props": [
                        {
                            "pset": fp.get("pset", "").strip(),
                            "prop": fp.get("prop", "").strip(),
                            "value": fp.get("value")
                        }
                        for fp in f.get("props", [])
                        if fp.get("pset") and fp.get("prop")
                    ]
                },
                "material": rule.get("material", "").strip(),
                "quantity": {
                    "pset": rule.get("quantity", {}).get("pset", "").strip(),
                    "prop": rule.get("quantity", {}).get("prop", "").strip()
                },
                "agrupamento": {
                    "pset": rule.get("agrupamento", {}).get("pset", "").strip(),
                    "prop": rule.get("agrupamento", {}).get("prop", "").strip()
                } if rule.get("agrupamento") else None
            }


    def _validate_rule(self, code: str, r: dict) -> tuple[bool, str]:
        try:
            f = r["filter"]
            ic = f["ifc_class"]
            pd = f["predefined"]
        except Exception:
            return False, f"[{code}] Regra inválida: faltam campos obrigatórios."

        if not ic or not pd:
            return False, f"[{code}] 'ifc_class' e 'predefined' são obrigatórios."

        if pd.upper() == "USERDEFINED":
            ot = (f.get("object_type") or "").strip()
            if not ot:
                return False, f"[{code}] 'object_type' é obrigatório quando predefined=USERDEFINED."

        q = r.get("quantity") or {}
        if (not q.get("pset")) or (not q.get("prop")):
            return False, f"[{code}] quantity.pset e quantity.prop são obrigatórios."
        agr = r.get("agrupamento")
        if not agr or not agr.get("pset") or not agr.get("prop"):
            return False, f"[{code}] Agrupamento é obrigatório. Forneça agrupamento.pset e agrupamento.prop."
        
        return True, ""

    def _validate_loaded_rules(self, data: dict) -> tuple[bool, str]:
        if not isinstance(data, dict):
            return False, "Ficheiro JSON inválido."
        if str(data.get("version")) not in {"1", 1}:
            return False, "Versão do ficheiro não suportada (esperado 'version': 1)."
        rules = data.get("rules")
        if not isinstance(rules, dict) or not rules:
            return False, "Estrutura 'rules' ausente ou vazia."

        for code, r in rules.items():
            ok, msg = self._validate_rule(code, r)
            if not ok:
                return False, msg
        return True, ""


    def save_rules_dialog(self):
            if not self.app.rules:
                messagebox.showinfo("Mapeamento IFC", "Não há regras para guardar.")
                return

            initialdir = (
                Path(self.app.wbs_xlsx_var.get().strip()).parent
                if self.app.wbs_xlsx_var.get().strip() else Path.home()
            )

            path = filedialog.asksaveasfilename(
                title="Guardar mapeamento",
                defaultextension=".json",
                filetypes=[("JSON", "*.json")],
                initialdir=str(initialdir),
                initialfile="mapeamento_ifc.json",
            )
            if not path:
                return

            cleaned = {}
            for code, r in (self.app.rules or {}).items():
                nr = self._normalize_rule(r)
                ok, msg = self._validate_rule(code, nr)
                if not ok:
                    messagebox.showerror("Mapeamento", f"Não foi possível guardar:\n{msg}")
                    return
                cleaned[code] = nr

            payload = {
                "version": 1,
                "ifc_path": self.app.ifc_var.get().strip() or None,
                "wbs_path": self.app.wbs_xlsx_var.get().strip() or None,
                "rules": cleaned,
            }

            try:
                with open(path, "w", encoding="utf-8") as f:
                    json.dump(payload, f, ensure_ascii=False, indent=2)
                messagebox.showinfo("Mapeamento IFC", f"Mapeamento guardado com sucesso em:\n{path}")

                try:
                    self.app.open_extract("from_mapping")
                except Exception:
                    self.app.go_extract()

            except Exception as e:
                messagebox.showerror("Erro", f"Falha a guardar o mapeamento:\n{e}")


    def load_rules_dialog(self):
        
        if self.app.df_raw is None:
            messagebox.showinfo(
                "Mapeamento",
                "Para editar um mapeamento é necessário carregar primeiro um WBS com descrições."
            )
            return
        
        path = filedialog.askopenfilename(
            title="Editar mapeamento existente",
            filetypes=[("JSON", "*.json"), ("Todos", "*.*")],
            initialdir=str(Path.home()),
        )
        if not path:
            return

        try:
            with open(path, "r", encoding="utf-8") as f:
                data = json.load(f)
        except Exception as e:
            messagebox.showerror("Erro", f"Não foi possível ler o ficheiro:\n{e}")
            return

        ok, msg = self._validate_loaded_rules(data)
        if not ok:
            messagebox.showerror("Mapeamento", f"Ficheiro inválido:\n{msg}")
            return

        rules = {code: self._normalize_rule(r) for code, r in data["rules"].items()}

        if self.app.rules:
            if not messagebox.askyesno(
                "Mapeamento",
                "Substituir o mapeamento atualmente em memória pelo ficheiro carregado?"
            ):
                return

        self.app.rules = rules

        if not self.app.ifc_var.get().strip():
            if data.get("ifc_path"):
                self.app.ifc_var.set(data["ifc_path"])
        if not self.app.wbs_xlsx_var.get().strip():
            if data.get("wbs_path"):
                self.app.wbs_xlsx_var.set(data["wbs_path"])

        self._render_list()
        self._set_edit_enabled(False)
        messagebox.showinfo("Mapeamento IFC", f"Mapeamento carregado com sucesso em:\n{path}")
       
        
    def _is_leaf_selected(self) -> bool:
        return bool(getattr(self, "selected_code", None)) and \
               (self.selected_code in getattr(self, "relevant_set", set()))

    def _set_edit_enabled(self, enabled: bool):
        state_combo = "readonly" if enabled else "disabled"
        state_entry = "normal"   if enabled else "disabled"

        try: self.ifc_class.configure(state=state_combo)
        except Exception: pass
        try: self.predef.configure(state=state_combo)
        except Exception: pass
        try: self.objtype.configure(state=state_entry)
        except Exception: pass
        try: self.q_pset.configure(state=state_entry)
        except Exception: pass
        try: self.q_prop.configure(state=state_entry)
        except Exception: pass

        try: self.btn_add_filter.configure(state=("normal" if enabled else "disabled"))
        except Exception: pass

        try: self.material.configure(state=(state_combo if enabled else "disabled"))
        except Exception: pass

        try: self.agr_pset.configure(state=(state_entry if enabled else "disabled"))
        except Exception: pass
        try: self.agr_prop.configure(state=(state_entry if enabled else "disabled"))
        except Exception: pass

        for row_frame in self.extra_container.winfo_children():
            for w in row_frame.winfo_children():
                try:
                    w.configure(state=(state_entry if not isinstance(w, ttk.Combobox) else state_combo))
                except Exception:
                    pass

        try: self.btn_save_rule.configure(state=("normal" if enabled else "disabled"))
        except Exception: pass

        can = bool(enabled and getattr(self.app, "ifc_file", None) and self._is_leaf_selected())
        state = ("normal" if can else "disabled")
        try:
            self.btn_test_filter.configure(state=state)
            self.btn_test_qty.configure(state=state)
        except Exception:
            pass
        
        self._toggle_object_type()
        
    
    def _parse_value_token(self, s: str):
        if s.lower() in {"true", "false"}:
            return s.lower() == "true"
        try:
            return int(s)
        except Exception:
            pass
        try:
            return float(s)
        except Exception:
            pass
        return s
        
    def _build_rule_from_form(self):
        if not self._is_leaf_selected():
            messagebox.showwarning("Mapeamento IFC", "Selecione um item do último nível (folha) na lista à esquerda.")
            return None

        ic = self.ifc_class.get().strip()
        pd = self.predef.get().strip().upper()
        if not ic or not pd:
            messagebox.showerror("Mapeamento IFC", "IFC Class e PredefinedType são obrigatórios.")
            return None

        ot = self.objtype.get().strip() if pd == "USERDEFINED" else ""
        if pd == "USERDEFINED" and not ot:
            messagebox.showerror("Mapeamento IFC", "ObjectType é obrigatório quando PredefinedType = USERDEFINED.")
            return None

        props = []
        for row, e_pset, e_prop, e_val, _btn in getattr(self, "prop_rows", []):
            pset = (e_pset.get() or "").strip()
            prop = (e_prop.get() or "").strip()
            val_txt = (e_val.get() or "").strip()
            if not pset and not prop and not val_txt:
                continue
            if not pset or not prop:
                messagebox.showerror("Mapeamento IFC", "Complete Pset e Propriedade (ou remova a linha vazia).")
                return None
            props.append({"pset": pset, "prop": prop, "value": self._parse_value_token(val_txt)})

        qpset = self.q_pset.get().strip()
        qprop = self.q_prop.get().strip()
        if not qpset or not qprop:
            messagebox.showerror("Mapeamento IFC", "Quantidade: Pset e Propriedade são obrigatórios.")
            return None

        rule = {"filter": {"ifc_class": ic, "predefined": pd}, "quantity": {"pset": qpset, "prop": qprop}}
        if pd == "USERDEFINED":
            rule["filter"]["object_type"] = ot
        if props:
            rule["filter"]["props"] = props
        return rule

    def test_filter_current(self):
        if not getattr(self.app, "ifc_file", None):
            messagebox.showinfo("IFC", "Carregue primeiro um ficheiro IFC (parte inferior da página).")
            return
        rule = self._build_rule_from_form()
        if not rule:
            return
        try:
            elems = list(self.app.inv.filter_elements(self.app.ifc_file, rule))
            messagebox.showinfo("Filtro", f"Filtro encontrou {len(elems)} elemento(s) no IFC.")
        except Exception as e:
            messagebox.showerror("Erro", f"Falha ao testar o filtro:\n{e}")

    def test_quantity_current(self):
        if not getattr(self.app, "ifc_file", None):
            messagebox.showinfo("IFC", "Carregue primeiro um ficheiro IFC (parte inferior da página).")
            return
        rule = self._build_rule_from_form()
        if not rule:
            return
        try:
            total, n = self.app.inv.sum_quantity(self.app.ifc_file, rule)
            messagebox.showinfo(
                "Quantidade",
                f"Elementos com valor: {n}\nSoma da quantidade: {total:g}"
            )
        except Exception as e:
            messagebox.showerror("Erro", f"Falha ao testar a quantidade:\n{e}")