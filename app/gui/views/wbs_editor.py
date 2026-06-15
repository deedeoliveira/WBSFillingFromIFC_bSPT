import json
import os
from pathlib import Path

import tkinter as tk
from tkinter import ttk, filedialog, messagebox

import pandas as pd

from app.gui.wbs_helpers import (
    find_wbs_columns, unpack_core_columns, split_levels,
    children_at_level, branch_has_children,
    ensure_level10_row, find_level10_text, list_ancestors,
    extract_partial_mapping,
)


class WBSPage(tk.Frame):
    def __init__(self, master, app):
        super().__init__(master)
        self.app = app

        self.section_var  = tk.StringVar()
        self.level_var    = tk.IntVar(value=1)
        self.path_stack:  list[str] = []
        self.current_leaf: str | None = None
        self._mode_var    = tk.StringVar(value="__none__")
        self._mode_confirmed = False
        self._existing_descs: dict[str, str] = {}
        self._editing_existing = False

        self._build_ui()

    def _build_ui(self):
        self.grid_columnconfigure(0, weight=1)
        self.grid_rowconfigure(3, weight=1)

        self._mode_lf = tk.LabelFrame(self, text="Modo de trabalho", padx=10, pady=8)
        self._mode_lf.grid(row=0, column=0, sticky="we", padx=10, pady=(10, 4))

        tk.Radiobutton(
            self._mode_lf,
            text="Novo WBS com descrições  (começar do zero)",
            variable=self._mode_var, value="new",
        ).pack(anchor="w")
        tk.Radiobutton(
            self._mode_lf,
            text="Continuar WBS existente  (já tenho um WBS com descrições parciais)",
            variable=self._mode_var, value="continue",
        ).pack(anchor="w")
        self._btn_confirm_mode = tk.Button(
            self._mode_lf, text="Confirmar modo →", command=self._confirm_mode,
        )
        self._btn_confirm_mode.pack(anchor="e", pady=(6, 0))

        self._upload_frame = tk.Frame(self)
        self._upload_frame.grid(row=1, column=0, sticky="we", padx=10, pady=(0, 4))
        self._upload_frame.grid_columnconfigure(1, weight=1)
        self._upload_frame.grid_remove()

        tk.Label(self._upload_frame, text="WBS original (bSPT):").grid(
            row=0, column=0, sticky="e", padx=(0, 6), pady=3)
        tk.Entry(self._upload_frame, textvariable=self.app.wbs_xlsx_var, width=80).grid(
            row=0, column=1, sticky="we", pady=3)
        tk.Button(self._upload_frame, text="Procurar…", command=self._pick_wbs_original).grid(
            row=0, column=2, padx=(4, 0), pady=3)
        tk.Button(self._upload_frame, text="Carregar", command=self.on_load_wbs).grid(
            row=0, column=3, padx=(4, 0), pady=3)

        self._partial_lbl = tk.Label(self._upload_frame, text="WBS com descrições (parcial):")
        self._partial_var = tk.StringVar()
        self._partial_entry = tk.Entry(self._upload_frame, textvariable=self._partial_var, width=80)
        self._partial_btn_pick = tk.Button(
            self._upload_frame, text="Procurar…", command=self._pick_wbs_partial)
        self._partial_btn_load = tk.Button(
            self._upload_frame, text="Carregar parcial", command=self._load_wbs_partial)

        self._info_lbl = tk.Label(
            self._upload_frame, text="", fg="#666", wraplength=860, justify="left")
        self._info_lbl.grid(row=2, column=0, columnspan=4, sticky="w", pady=(2, 0))

        self._nav_frame = tk.Frame(self)
        self._nav_frame.grid(row=2, column=0, sticky="we", padx=10, pady=4)
        self._nav_frame.grid_columnconfigure(1, weight=1)
        self._nav_frame.grid_remove()

        tk.Label(self._nav_frame, text="Secção (Nível 1):").grid(row=0, column=0, sticky="w")
        self.section_combo = ttk.Combobox(
            self._nav_frame, textvariable=self.section_var, state="readonly")
        self.section_combo.grid(row=0, column=1, sticky="we", padx=8)
        self.section_combo.bind("<<ComboboxSelected>>", self.on_section_selected)

        nav_btns = tk.Frame(self._nav_frame)
        nav_btns.grid(row=0, column=2)
        self.btn_up   = tk.Button(nav_btns, text="Nível acima",  command=self.on_back,  state="disabled")
        self.btn_down = tk.Button(nav_btns, text="Nível abaixo", command=self.on_next,  state="disabled")
        self.btn_add  = tk.Button(nav_btns, text="Adicionar descrição",
                                  command=self.on_begin_add_desc, state="disabled")
        self.btn_up.pack(side="left", padx=(0, 4))
        self.btn_down.pack(side="left", padx=(0, 4))
        self.btn_add.pack(side="left")

        self.path_label = tk.Label(self._nav_frame, text="Caminho: —", fg="#555")
        self.path_label.grid(row=1, column=0, columnspan=3, sticky="w", pady=(2, 0))

        tbl = tk.Frame(self)
        tbl.grid(row=3, column=0, sticky="nsew", padx=10, pady=4)
        tbl.grid_remove()
        self._tbl_frame = tbl

        self.tree = ttk.Treeview(tbl, columns=("wbs", "desc", "has_desc"),
                                  show="headings", height=14)
        self.tree.heading("wbs",      text="WBS")
        self.tree.heading("desc",     text="Descrição")
        self.tree.heading("has_desc", text="✓")
        self.tree.column("wbs",      width=200, anchor="w")
        self.tree.column("desc",     width=660, anchor="w")
        self.tree.column("has_desc", width=30,  anchor="center")
        self.tree.pack(side="left", fill="both", expand=True)
        yscroll = ttk.Scrollbar(tbl, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscroll=yscroll.set)
        yscroll.pack(side="right", fill="y")
        self.tree.bind("<<TreeviewSelect>>", lambda e: self._update_nav_buttons())

        self._editor_lf = tk.LabelFrame(
            self, text="Descrição do utilizador (disponível no último nível da secção)",
            padx=10, pady=8,
        )
        self._editor_lf.grid(row=4, column=0, sticky="we", padx=10, pady=(0, 4))
        self._editor_lf.grid_remove()

        self.user_desc = tk.Text(self._editor_lf, height=3, width=120, state="disabled")
        self.user_desc.pack(fill="x")

        btn_row = tk.Frame(self._editor_lf)
        btn_row.pack(fill="x", pady=(6, 0))

        self.save_desc_btn = tk.Button(
            btn_row, text="Guardar descrição", state="disabled",
            command=self.on_save_user_desc,
        )
        self.save_desc_btn.pack(side="right")

        self.edit_desc_btn = tk.Button(
            btn_row, text="Editar descrição", state="disabled",
            command=self._on_edit_existing,
        )
        self.edit_desc_btn.pack(side="right", padx=(0, 6))
        self.edit_desc_btn.pack_forget()

        self._bottom_frame = tk.Frame(self)
        self._bottom_frame.grid(row=999, column=0, sticky="e", padx=8, pady=(4, 8))
        self._bottom_frame.grid_remove()
        self.btn_save_export = tk.Button(
            self._bottom_frame, text="Salvar etapa e exportar WBS",
            command=self.on_save_and_export,
        )
        self.btn_save_export.pack(side="right")

    def _confirm_mode(self):
        mode = self._mode_var.get()
        if not mode:
            messagebox.showwarning("WBS e descrição", "Escolha um modo antes de confirmar.")
            return

        self._mode_confirmed = True
        self._btn_confirm_mode.configure(state="disabled", text="Modo confirmado ✓")
        for rb in self._mode_lf.winfo_children():
            if isinstance(rb, tk.Radiobutton):
                rb.configure(state="disabled")

        self._upload_frame.grid()

        if mode == "continue":
            self._partial_lbl.grid(row=1, column=0, sticky="e", padx=(0, 6), pady=3)
            self._partial_entry.grid(row=1, column=1, sticky="we", pady=3)
            self._partial_btn_pick.grid(row=1, column=2, padx=(4, 0), pady=3)
            self._partial_btn_load.grid(row=1, column=3, padx=(4, 0), pady=3)
            self._info_lbl.configure(
                text="Carregue primeiro o WBS original (bSPT) e depois o WBS com descrições parciais."
            )
        else:
            self._info_lbl.configure(
                text="Carregue o WBS original da buildingSMART Portugal."
            )

    def _pick_wbs_original(self):
        self.app.pick_file(
            self.app.wbs_xlsx_var,
            "Selecionar WBS original (buildingSMART Portugal)",
            [("Excel (*.xlsx *.xls)", "*.xlsx *.xls")],
        )

    def _pick_wbs_partial(self):
        path = filedialog.askopenfilename(
            title="Selecionar WBS com descrições parciais",
            filetypes=[("Excel (*.xlsx *.xls)", "*.xlsx *.xls")],
        )
        if path:
            self._partial_var.set(path)

    def on_load_wbs(self):
        path = self.app.wbs_xlsx_var.get().strip()
        if not path or not os.path.isfile(path):
            messagebox.showerror("Erro", "Selecione primeiro o ficheiro WBS original.")
            return
        if not path.lower().endswith((".xls", ".xlsx")):
            messagebox.showerror("WBS e descrição", "Apenas ficheiros .xls ou .xlsx são aceites.")
            return

        try:
            df_full    = pd.read_excel(path, header=None)
            header_row = self._find_header_row(df_full)
            if header_row is None:
                messagebox.showerror(
                    "Erro",
                    "Cabeçalhos não encontrados. O ficheiro deve ter: NÍVEL, WBS e DESCRIÇÃO.",
                )
                return

            self.app.df_raw = pd.read_excel(path, header=header_row)
            cols = find_wbs_columns(self.app.df_raw)
            self.app.col_wbs, self.app.col_desc, self.app.col_nivel = unpack_core_columns(cols)
            self.app.wbs_cols = cols

            if not all([self.app.col_wbs, self.app.col_desc, self.app.col_nivel]):
                messagebox.showerror(
                    "Erro", "Não foi possível detectar colunas obrigatórias (NÍVEL, WBS, DESCRIÇÃO).")
                return

            self.app.df_desc0 = self.app.df_raw[self.app.col_desc].copy()

            if self._mode_var.get() == "continue" and self._existing_descs:
                self._apply_existing_descs()

            self._populate_sections()

            ifc_cols = [k for k in ("col_ifc_class","col_predef","col_objtype","col_ifc_prop")
                        if cols.get(k)]
            extra = (f"\n\nColunas IFC detectadas: {len(ifc_cols)}/4"
                     if ifc_cols else
                     "\n\nNenhuma coluna IFC detectada.")
            messagebox.showinfo("WBS e descrição",
                                f"WBS original carregado com sucesso.{extra}")

        except Exception as e:
            messagebox.showerror("Erro", f"Falha a ler o WBS:\n{e}")
            import traceback; traceback.print_exc()

    def _load_wbs_partial(self):
        path = self._partial_var.get().strip()
        if not path or not os.path.isfile(path):
            messagebox.showerror("Erro", "Selecione o WBS com descrições parciais.")
            return

        try:
            df_p = pd.read_excel(path, header=None)
            hr   = self._find_header_row(df_p)
            if hr is None:
                df_p = pd.read_excel(path, header=1)
            else:
                df_p = pd.read_excel(path, header=hr)

            cols_p = find_wbs_columns(df_p)
            col_wbs_p, col_desc_p, col_nivel_p = unpack_core_columns(cols_p)
            if not all([col_wbs_p, col_desc_p, col_nivel_p]):
                messagebox.showerror("Erro", "Colunas obrigatórias não encontradas no WBS parcial.")
                return

            niv_p = split_levels(df_p, col_nivel_p)
            self._existing_descs.clear()

            for i in range(len(df_p)):
                lvl = niv_p.iat[i]
                if lvl is None:
                    continue
                if lvl == 10:
                    desc_val = df_p.at[i, col_desc_p]
                    if pd.isna(desc_val) or not str(desc_val).strip():
                        continue
                    for j in range(i - 1, -1, -1):
                        parent_lvl = niv_p.iat[j]
                        if parent_lvl is not None and parent_lvl < 10:
                            code = df_p.at[j, col_wbs_p]
                            if isinstance(code, str) and code.strip():
                                self._existing_descs[code.strip()] = str(desc_val).strip()
                            break

            if not self._existing_descs:
                messagebox.showwarning("WBS e descrição",
                                       "Nenhuma descrição encontrada no WBS parcial.")
                return

            if self.app.df_raw is not None:
                self._apply_existing_descs()
                self._populate_sections()

            messagebox.showinfo(
                "WBS e descrição",
                f"WBS parcial carregado: {len(self._existing_descs)} descrição(ões) encontrada(s).\n"
                f"Os itens com descrição aparecem marcados com ✓ na lista.",
            )

        except Exception as e:
            messagebox.showerror("Erro", f"Falha a ler o WBS parcial:\n{e}")
            import traceback; traceback.print_exc()

    def _apply_existing_descs(self):
        if self.app.df_raw is None or not self._existing_descs:
            return
        df  = self.app.df_raw
        niv = split_levels(df, self.app.col_nivel)

        for i in range(len(df)):
            lvl  = niv.iat[i]
            code = df.at[i, self.app.col_wbs]
            if lvl is None or lvl >= 10 or not isinstance(code, str) or not code.strip():
                continue
            code = code.strip()
            if code not in self._existing_descs:
                continue
            try:
                idx = ensure_level10_row(
                    df, self.app.col_nivel, self.app.col_wbs, self.app.col_desc, code)
                df.at[idx, self.app.col_desc] = self._existing_descs[code]
                if self.app.df_desc0 is not None and len(self.app.df_desc0) < len(df):
                    diff = len(df) - len(self.app.df_desc0)
                    self.app.df_desc0 = pd.concat(
                        [self.app.df_desc0, pd.Series([None] * diff)], ignore_index=True)
            except Exception:
                pass

    def _find_header_row(self, df):
        import unicodedata
        def _norm(text):
            if not isinstance(text, str):
                return ""
            return unicodedata.normalize("NFKD", text).encode("ASCII", "ignore").decode().lower().strip()
        target = {"wbs", "nivel"}
        desc_v = {"descricao", "desc"}
        for idx, row in df.iterrows():
            normed = {_norm(str(v)) for v in row if pd.notna(v)}
            if target.issubset(normed) and bool(desc_v & normed):
                return idx
        return None

    def _populate_sections(self):
        df  = self.app.df_raw.copy()
        df[self.app.col_nivel] = split_levels(df, self.app.col_nivel)
        top = (
            df[df[self.app.col_nivel] == 1]
            [[self.app.col_wbs, self.app.col_desc]]
            .dropna(subset=[self.app.col_wbs])
        )
        items = [
            f"{w} – {d}" if isinstance(d, str) and d.strip() else str(w)
            for w, d in zip(top[self.app.col_wbs], top[self.app.col_desc])
        ]
        self.section_combo["values"] = items
        if items:
            self.section_combo.current(0)
            self.on_section_selected()

        self._nav_frame.grid()
        self._tbl_frame.grid()
        self._editor_lf.grid()
        self._bottom_frame.grid()

    def on_section_selected(self, *_):
        txt = self.section_var.get().strip()
        if not txt:
            return
        code = txt.split(" – ")[0].strip() if " – " in txt else txt.split(" - ")[0].strip()
        self.path_stack  = [code]
        self.level_var.set(2)
        self.current_leaf = None
        self._toggle_editor(False)
        self._render_table_for(code, 2)

    def on_next(self):
        sel = self.tree.selection()
        if not sel:
            messagebox.showwarning("WBS e descrição", "Selecione um item.")
            return
        sel_code = self.tree.item(sel[0], "values")[0]
        curr = self.level_var.get()
        has_child = branch_has_children(
            self.app.df_raw, self.app.col_nivel, self.app.col_wbs, sel_code, curr + 1)
        if not has_child:
            self.current_leaf = sel_code
            self._toggle_editor(False)
            messagebox.showinfo("WBS e descrição",
                                "Atingiu o último nível. Use 'Adicionar descrição'.")
            self._update_nav_buttons()
            return
        self.path_stack.append(sel_code)
        self.level_var.set(curr + 1)
        self.current_leaf = None
        self._toggle_editor(False)
        self._render_table_for(sel_code, self.level_var.get())

    def on_back(self):
        if not self.path_stack:
            return
        if len(self.path_stack) == 1:
            code = self.path_stack[0]
            self.level_var.set(2)
            self.current_leaf = None
            self._toggle_editor(False)
            self._render_table_for(code, 2)
        else:
            self.path_stack.pop()
            parent = self.path_stack[-1]
            self.level_var.set(1 + len(self.path_stack))
            self.current_leaf = None
            self._toggle_editor(False)
            self._render_table_for(parent, self.level_var.get())
        self._update_nav_buttons()

    def _render_table_for(self, prefix, level):
        df  = self.app.df_raw.copy()
        df[self.app.col_nivel] = split_levels(df, self.app.col_nivel)
        sub = children_at_level(
            df, self.app.col_nivel, self.app.col_wbs, self.app.col_desc, prefix, level)
        for i in self.tree.get_children():
            self.tree.delete(i)
        for _, row in sub.iterrows():
            code  = str(row[self.app.col_wbs])
            desc  = "" if pd.isna(row[self.app.col_desc]) else str(row[self.app.col_desc])
            has_d = "✓" if code in self._existing_descs else ""
            self.tree.insert("", "end", values=(code, desc, has_d))
        self.path_label.configure(
            text=f"Caminho: {' > '.join(self.path_stack)} (Nível {level})")
        self._update_nav_buttons()

    def _update_nav_buttons(self):
        has_path = bool(self.path_stack)
        sel      = self.tree.selection()
        sel_code = self.tree.item(sel[0], "values")[0] if sel else None
        curr     = self.level_var.get()
        self.btn_up.configure(state=(tk.NORMAL if has_path else tk.DISABLED))
        has_child = (
            branch_has_children(
                self.app.df_raw, self.app.col_nivel, self.app.col_wbs, sel_code, curr + 1)
            if sel_code else False)
        self.btn_down.configure(state=(tk.NORMAL if (sel_code and has_child) else tk.DISABLED))
        can_add = bool(sel_code) and not has_child
        self.btn_add.configure(state=(tk.NORMAL if can_add else tk.DISABLED))
        if not can_add:
            self._toggle_editor(False)

    def _toggle_editor(self, enabled: bool, leaf: str | None = None,
                       existing_text: str | None = None):
        if not enabled:
            self.user_desc.configure(state="disabled", bg="#F0F0F0")
            self.save_desc_btn.configure(state="disabled")
            self.edit_desc_btn.pack_forget()
            self._editing_existing = False
            return

        self.user_desc.configure(state="normal")
        self.user_desc.delete("1.0", "end")

        if existing_text is not None:
            self.user_desc.insert("1.0", existing_text)
            self.user_desc.configure(state="disabled", bg="#F5F5DC")
            self.save_desc_btn.configure(state="disabled")
            self.edit_desc_btn.configure(state="normal", text="Editar descrição")
            self.edit_desc_btn.pack(side="right", padx=(0, 6))
            self._editing_existing = False
        else:
            self.user_desc.configure(state="normal", bg="white")
            self.save_desc_btn.configure(state="normal")
            self.edit_desc_btn.pack_forget()
            self._editing_existing = False

        if leaf:
            self.path_label.configure(
                text=f"Caminho: {' > '.join(self.path_stack)} (Folha: {leaf})")

    def _on_edit_existing(self):
        if self._editing_existing:
            orig = self._existing_descs.get(self.current_leaf, "")
            self.user_desc.configure(state="normal")
            self.user_desc.delete("1.0", "end")
            self.user_desc.insert("1.0", orig)
            self.user_desc.configure(state="disabled", bg="#F5F5DC")
            self.save_desc_btn.configure(state="disabled")
            self.edit_desc_btn.configure(text="Editar descrição")
            self._editing_existing = False
        else:
            self.user_desc.configure(state="normal", bg="white")
            self.save_desc_btn.configure(state="normal")
            self.edit_desc_btn.configure(text="Cancelar edição")
            self._editing_existing = True

    def on_begin_add_desc(self):
        sel = self.tree.selection()
        if not sel:
            messagebox.showwarning("WBS e descrição", "Selecione um item.")
            return
        code = self.tree.item(sel[0], "values")[0]
        curr = self.level_var.get()
        if branch_has_children(
                self.app.df_raw, self.app.col_nivel, self.app.col_wbs, code, curr + 1):
            messagebox.showwarning("WBS e descrição", "Ainda não está no último nível.")
            return

        self.current_leaf = code

        existing = find_level10_text(
            self.app.df_raw, self.app.col_nivel, self.app.col_wbs, self.app.col_desc, code)

        if existing:
            self._toggle_editor(True, leaf=code, existing_text=existing)
        else:
            self._toggle_editor(True, leaf=code, existing_text=None)

    def on_save_user_desc(self):
        if not self.current_leaf:
            messagebox.showwarning("WBS e descrição",
                                   "Selecione um item no último nível.")
            return
        txt = self.user_desc.get("1.0", "end").strip()
        if not txt:
            messagebox.showwarning("WBS e descrição", "A descrição está vazia.")
            return

        idx = ensure_level10_row(
            self.app.df_raw, self.app.col_nivel, self.app.col_wbs, self.app.col_desc,
            self.current_leaf)
        self.app.df_raw.at[idx, self.app.col_desc] = txt

        if self.app.df_desc0 is not None and len(self.app.df_desc0) != len(self.app.df_raw):
            b = self.app.df_desc0
            self.app.df_desc0 = pd.concat(
                [b.iloc[:idx], pd.Series([None]), b.iloc[idx:]], ignore_index=True)

        self._existing_descs[self.current_leaf] = txt

        messagebox.showinfo("WBS e descrição", "Descrição guardada com sucesso.")
        self.user_desc.delete("1.0", "end")
        self._toggle_editor(False)
        self.current_leaf  = None
        self._editing_existing = False
        self._update_nav_buttons()
        if self.path_stack:
            self._render_table_for(self.path_stack[-1], self.level_var.get())

    def on_export_wbs(self) -> bool:
        if self.app.df_raw is None:
            messagebox.showwarning("WBS e descrição",
                                   "Carregue primeiro o WBS da buildingSMART Portugal.")
            return False

        initialdir = (
            self.app.out_var.get().strip()
            or os.path.dirname(self.app.wbs_xlsx_var.get().strip())
            or str(Path.home())
        )
        path = filedialog.asksaveasfilename(
            title="Guardar WBS atualizado",
            defaultextension=".xlsx",
            filetypes=[("Excel", "*.xlsx")],
            initialdir=initialdir,
            initialfile="WBS_ComDescricao.xlsx",
        )
        if not path:
            return False

        try:
            df       = self.app.df_raw
            niv      = split_levels(df, self.app.col_nivel)
            baseline = (self.app.df_desc0
                        if self.app.df_desc0 is not None
                        else df[self.app.col_desc].copy())

            def norm(v):
                return "" if pd.isna(v) else str(v).strip()

            code_to_idx = {}
            for i, code in enumerate(df[self.app.col_wbs]):
                lvl = niv.iat[i]
                if lvl is not None and lvl < 10 and isinstance(code, str) and code.strip():
                    code_to_idx[str(code).strip()] = i

            include_idx = set()
            last_code = None; last_code_idx = None

            for i in range(len(df)):
                lvl  = niv.iat[i]
                code = df.at[i, self.app.col_wbs]
                if lvl is not None and lvl < 10 and isinstance(code, str) and code.strip():
                    last_code     = str(code).strip()
                    last_code_idx = i
                    continue
                if lvl == 10:
                    new_txt = norm(df.at[i, self.app.col_desc])
                    old_txt = norm(baseline.iat[i]) if i < len(baseline) else ""
                    if new_txt and last_code is not None:
                        include_idx.add(i)
                        include_idx.add(last_code_idx)
                        for anc in list_ancestors(last_code):
                            j = code_to_idx.get(anc)
                            if j is not None:
                                include_idx.add(j)

            if not include_idx:
                messagebox.showinfo("WBS e descrição",
                                    "Não há descrições do utilizador para exportar.")
                return False

            mini_df = df.loc[sorted(include_idx)]
            with pd.ExcelWriter(path, engine="openpyxl") as writer:
                mini_df.to_excel(writer, index=False, startrow=1)
                ws = next(iter(writer.sheets.values()))
                ws["A1"] = "WBS com descrições"

            self.app.last_exported_wbs = path
            self.app.wbs_xlsx_var.set(path)
            return True

        except Exception as e:
            messagebox.showerror("Erro", f"Falha a exportar o WBS:\n{e}")
            return False

    def _export_partial_mapping(self, excel_path: str):
        if self.app.df_raw is None:
            return False
        cols = getattr(self.app, "wbs_cols", None) or find_wbs_columns(self.app.df_raw)
        has_ifc = any(cols.get(k) for k in
                      ("col_ifc_class", "col_predef", "col_objtype", "col_ifc_prop"))
        if not has_ifc:
            return False
        rules = extract_partial_mapping(self.app.df_raw, cols)
        if not rules:
            return False
        payload = {"version": 2, "partial": True, "ifc_path": None,
                   "wbs_path": excel_path, "rules": rules}
        json_path = str(Path(excel_path).with_name(
            Path(excel_path).stem + "_mapeamento_parcial.json"))
        try:
            with open(json_path, "w", encoding="utf-8") as f:
                json.dump(payload, f, ensure_ascii=False, indent=2)
            return json_path
        except Exception as e:
            messagebox.showerror("Erro", f"Falha a exportar mapeamento parcial:\n{e}")
            return False

    def on_save_and_export(self):
        self.app.wbs_finalized = True
        ok = self.on_export_wbs()
        if not ok:
            return
        exported_wbs_path = self.app.wbs_xlsx_var.get().strip()
        json_result = self._export_partial_mapping(exported_wbs_path)
        if json_result:
            self.app.map_var.set(json_result)
            n_rules = 0
            try:
                with open(json_result, encoding="utf-8") as f:
                    n_rules = len(json.load(f).get("rules", {}))
            except Exception:
                pass
            messagebox.showinfo(
                "WBS e descrição",
                f"WBS exportado com sucesso.\n\n"
                f"Mapeamento parcial gerado com {n_rules} código(s) WBS:\n{json_result}\n\n"
                f"Este mapeamento será pré-carregado na aba 'Mapeamento IFC'.",
            )
        else:
            messagebox.showinfo("WBS e descrição",
                                f"WBS exportado com sucesso.\n{exported_wbs_path}")
        self.app.open_mapping(source="wbs")
