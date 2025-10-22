# app/gui/views/wbs_editor.py
import os
from pathlib import Path

import tkinter as tk
from tkinter import ttk, filedialog, messagebox

import pandas as pd

from app.gui.wbs_helpers import (
    find_wbs_columns, split_levels, children_at_level, branch_has_children,
    ensure_level10_row, find_level10_text, list_ancestors
)

class WBSPage(tk.Frame):
    def __init__(self, master, app):
        super().__init__(master)
        self.app = app

        self.section_var = tk.StringVar()
        self.level_var   = tk.IntVar(value=1)
        self.path_stack: list[str] = []
        self.current_leaf: str | None = None

        self._build_ui()

    def _build_ui(self):
        self.grid_columnconfigure(1, weight=1)

        top = tk.Frame(self); top.grid(row=0, column=0, columnspan=3, sticky="we", padx=10, pady=(10,4))
        tk.Label(top, text="WBS (Excel, bSPT):").pack(side="left")
        tk.Entry(top, textvariable=self.app.wbs_xlsx_var, width=90).pack(side="left", padx=6, fill="x", expand=True)
        tk.Button(top, text="Procurar…", command=self.pick_wbs).pack(side="left", padx=4)
        tk.Button(top, text="Carregar", command=self.on_load_wbs).pack(side="left", padx=8)

        tk.Label(self, text="Faça upload do WBS da buildingSMART Portugal. "
                            "As descrições do utilizador são adicionadas no nível 10 (linha abaixo do item).",
                 fg="#666", wraplength=900, justify="left").grid(row=1, column=0, columnspan=3, sticky="w", padx=12)

        sect = tk.Frame(self); sect.grid(row=2, column=0, columnspan=3, sticky="we", padx=10, pady=6)
        tk.Label(sect, text="Secção (Nível 1):").grid(row=0, column=0, sticky="w")
        self.section_combo = ttk.Combobox(sect, textvariable=self.section_var, state="readonly")
        self.section_combo.grid(row=0, column=1, sticky="we", padx=8); sect.grid_columnconfigure(1, weight=1)
        self.section_combo.bind("<<ComboboxSelected>>", self.on_section_selected)

        nav = tk.Frame(self); nav.grid(row=3, column=0, columnspan=3, sticky="we", padx=10, pady=4)
        nav.grid_columnconfigure(0, weight=1)
        self.path_label = tk.Label(nav, text="Caminho: —"); self.path_label.grid(row=0, column=0, sticky="w")
        
        self.btn_up   = tk.Button(nav, text="Nível acima",  command=self.on_back,  state="disabled")
        self.btn_down = tk.Button(nav, text="Nível abaixo", command=self.on_next, state="disabled")
        
        self.btn_add  = tk.Button(nav, text="Adicionar descrição", command=self.on_begin_add_desc, state="disabled")
        
        self.btn_up.grid(row=0, column=1, sticky="e", padx=(0,6))
        self.btn_down.grid(row=0, column=2, sticky="e", padx=(0,6))
        self.btn_add.grid(row=0, column=3, sticky="e")

        tbl = tk.Frame(self); tbl.grid(row=4, column=0, columnspan=3, sticky="nsew", padx=10, pady=6)
        self.tree = ttk.Treeview(tbl, columns=("wbs","desc"), show="headings", height=16)
        self.tree.heading("wbs", text="WBS"); self.tree.column("wbs", width=220, anchor="w")
        self.tree.heading("desc", text="Descrição"); self.tree.column("desc", width=720, anchor="w")
        self.tree.pack(side="left", fill="both", expand=True)
        yscroll = ttk.Scrollbar(tbl, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscroll=yscroll.set); yscroll.pack(side="right", fill="y")
        self.tree.bind("<<TreeviewSelect>>", lambda e: self._update_nav_buttons())

        editor = tk.LabelFrame(self, text="Descrição do utilizador (disponível no último nível da secção)", padx=10, pady=10)
        editor.grid(row=5, column=0, columnspan=3, sticky="we", padx=10, pady=(4,10))
        self.user_desc = tk.Text(editor, height=3, width=120, state="disabled")
        self.user_desc.pack(fill="x")
        self.save_desc_btn = tk.Button(editor, text="Guardar descrição", state="disabled", command=self.on_save_user_desc)
        self.save_desc_btn.pack(anchor="e", pady=(6,0))

        bottom = tk.Frame(self)
        bottom.grid(row=999, column=0, columnspan=4, sticky="e", padx=8, pady=(6, 8))

        self.btn_save_export = tk.Button(
            bottom,
            text="Salvar etapa e exportar WBS",
            command=self.on_save_and_export
        )
        self.btn_save_export.pack(side="right")

    def pick_wbs(self):
        self.app.pick_file(
            self.app.wbs_xlsx_var,
            "Selecionar WBS (buildingSMART Portugal)",
            [("Excel (*.xlsx *.xls)", "*.xlsx *.xls")]
        )

    def on_load_wbs(self):
            path = self.app.wbs_xlsx_var.get().strip()
            if not path or not os.path.isfile(path):
                messagebox.showerror("Erro", "Selecione primeiro o ficheiro WBS (Excel).")
                return
            
            if not path.lower().endswith((".xls", ".xlsx")):
                messagebox.showerror("WBS e descrição", "Apenas ficheiros .xls ou .xlsx são aceites.")
                return

            try:
                self.app.df_raw = pd.read_excel(path, header=1)

                self.app.col_wbs, self.app.col_desc, self.app.col_nivel = find_wbs_columns(self.app.df_raw)
                
                self.app.df_desc0 = self.app.df_raw[self.app.col_desc].copy()

                df = self.app.df_raw.copy()
                df[self.app.col_nivel] = split_levels(df, self.app.col_nivel)
                top = df[df[self.app.col_nivel] == 1][[self.app.col_wbs, self.app.col_desc]].dropna(subset=[self.app.col_wbs])
                items = [f"{w} — {d}" if isinstance(d, str) and d.strip() else f"{w}"
                        for w, d in zip(top[self.app.col_wbs], top[self.app.col_desc])]
                self.section_combo["values"] = items
                if items:
                    self.section_combo.current(0)
                    self.on_section_selected()
                else:
                    self.section_var.set("")
                messagebox.showinfo("WBS e descrição", "WBS carregado com sucesso. Selecione uma secção (Nível 1).")
            except Exception as e:
                messagebox.showerror("Erro", f"Falha a ler o WBS:\n{e}")

    def on_section_selected(self, *_):
        txt = self.section_var.get().strip()
        if not txt: return
        code = txt.split(" — ")[0].strip()
        self.path_stack = [code]
        self.level_var.set(2)
        self.current_leaf = None
        self._toggle_editor(False)
        self._render_table_for(code, 2)

    def on_next(self):
        if not self.path_stack:
            messagebox.showwarning("WBS e descrição", "Selecione primeiro uma secção (Nível 1)."); return
        sel = self.tree.selection()
        if not sel:
            messagebox.showwarning("WBS e descrição", "Selecione um item na lista para avançar."); return
        sel_code = self.tree.item(sel[0], "values")[0]
        curr = self.level_var.get()
        has_child = branch_has_children(self.app.df_raw, self.app.col_nivel, self.app.col_wbs, sel_code, curr + 1)
        if not has_child:
            self.current_leaf = sel_code
            self._toggle_editor(False)
            messagebox.showinfo("WBS e descrição", "Atingiu o último nível desta secção. Use o botão 'Adicionar descrição'.")
            self._update_nav_buttons(); return
        self.path_stack.append(sel_code)
        self.level_var.set(curr + 1)
        self.current_leaf = None
        self._toggle_editor(False)
        self._render_table_for(sel_code, self.level_var.get())

    def on_back(self):
        if not self.path_stack: return
        if len(self.path_stack) == 1:
            code = self.path_stack[0]; self.level_var.set(2); self.current_leaf=None; self._toggle_editor(False)
            self._render_table_for(code, 2)
        else:
            self.path_stack.pop()
            parent = self.path_stack[-1]
            self.level_var.set(1 + len(self.path_stack))
            self.current_leaf=None; self._toggle_editor(False)
            self._render_table_for(parent, self.level_var.get())
        self._update_nav_buttons()

    def _render_table_for(self, prefix, level):
        df = self.app.df_raw.copy()
        df[self.app.col_nivel] = split_levels(df, self.app.col_nivel)
        sub = children_at_level(df, self.app.col_nivel, self.app.col_wbs, self.app.col_desc, prefix, level) \
              if df is not None else pd.DataFrame(columns=[self.app.col_wbs or "WBS", self.app.col_desc or "Descrição"])
        for i in self.tree.get_children():
            self.tree.delete(i)
        for _, row in sub.iterrows():
            self.tree.insert("", "end", values=(str(row[self.app.col_wbs]),
                                                "" if pd.isna(row[self.app.col_desc]) else str(row[self.app.col_desc])))
        self.path_label.configure(text=f"Caminho: {' > '.join(self.path_stack) if self.path_stack else '—'} (Nível {level})")
        self._update_nav_buttons()

    def _update_nav_buttons(self):
        has_path = bool(self.path_stack)
        sel = self.tree.selection()
        sel_code = self.tree.item(sel[0], "values")[0] if sel else None
        curr = self.level_var.get()
        self.btn_up.configure(state=(tk.NORMAL if has_path else tk.DISABLED))
        has_child = branch_has_children(self.app.df_raw, self.app.col_nivel, self.app.col_wbs, sel_code, curr+1) if sel_code else False
        self.btn_down.configure(state=(tk.NORMAL if (sel_code and has_child) else tk.DISABLED))
        can_add = bool(sel_code) and not has_child
        self.btn_add.configure(state=(tk.NORMAL if can_add else tk.DISABLED))
        if not can_add:
            self._toggle_editor(False)

    def _toggle_editor(self, enabled, leaf=None):
            if enabled:
                self.user_desc.configure(state="normal", bg="white")
                self.save_desc_btn.configure(state="normal")
                if leaf:
                    self.path_label.configure(text=f"Caminho: {' > '.join(self.path_stack)} (Folha: {leaf})")
            else:
                self.user_desc.configure(state="disabled", bg="#F0F0F0")
                self.save_desc_btn.configure(state="disabled")

    def on_begin_add_desc(self):
        sel = self.tree.selection()
        if not sel:
            messagebox.showwarning("WBS e descrição", "Selecione um item."); return
        code = self.tree.item(sel[0], "values")[0]
        curr = self.level_var.get()
        if branch_has_children(self.app.df_raw, self.app.col_nivel, self.app.col_wbs, code, curr+1):
            messagebox.showwarning("WBS e descrição", "Ainda não está no último nível."); return
        self.current_leaf = code
        existing = find_level10_text(self.app.df_raw, self.app.col_nivel, self.app.col_wbs, self.app.col_desc, code)
        self.user_desc.configure(state="normal"); self.user_desc.delete("1.0","end"); self.user_desc.insert("1.0", existing)
        self._toggle_editor(True, leaf=code)

    def on_save_user_desc(self):
        if not self.current_leaf:
            messagebox.showwarning("WBS e descrição", "Selecione um item no último nível para guardar a descrição."); return
        txt = self.user_desc.get("1.0","end").strip()
        if not txt:
            messagebox.showwarning("WBS e descrição", "A descrição está vazia."); return
        idx = ensure_level10_row(self.app.df_raw, self.app.col_nivel, self.app.col_wbs, self.app.col_desc, self.current_leaf)
        self.app.df_raw.at[idx, self.app.col_desc] = txt
       
        if self.app.df_desc0 is not None and len(self.app.df_desc0) != len(self.app.df_raw):
            b = self.app.df_desc0
            self.app.df_desc0 = pd.concat([b.iloc[:idx], pd.Series([None]), b.iloc[idx:]], ignore_index=True)

        
        messagebox.showinfo("WBS e descrição", "Descrição guardada com sucesso.")

        self.user_desc.delete("1.0","end")
        self._toggle_editor(False)
        self.current_leaf = None
        self._update_nav_buttons()

    def on_export_wbs(self) -> bool:

        if self.app.df_raw is None:
            messagebox.showwarning("WBS e descrição", "Carregue primeiro o WBS da buildingSMART Portugal (Excel).")
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
            df = self.app.df_raw
            niv = split_levels(df, self.app.col_nivel)

            baseline = (
                self.app.df_desc0
                if self.app.df_desc0 is not None
                else df[self.app.col_desc].copy()
            )

            def norm(v): return "" if pd.isna(v) else str(v).strip()

            code_to_idx = {}
            for i, code in enumerate(df[self.app.col_wbs]):
                lvl = niv.iat[i]
                if lvl is not None and lvl < 10 and isinstance(code, str) and code.strip():
                    code_to_idx[str(code).strip()] = i

            include_idx = set()
            last_code = None
            last_code_idx = None

            for i in range(len(df)):
                lvl = niv.iat[i]
                code = df.at[i, self.app.col_wbs]

                if lvl is not None and lvl < 10 and isinstance(code, str) and code.strip():
                    last_code = str(code).strip()
                    last_code_idx = i
                    continue

                if lvl == 10:
                    new_txt = norm(df.at[i, self.app.col_desc])
                    old_txt = norm(baseline.iat[i]) if i < len(baseline) else ""
                    if new_txt and (new_txt != old_txt) and last_code is not None:
                        include_idx.add(i)
                        include_idx.add(last_code_idx)
                        for anc in list_ancestors(last_code):
                            j = code_to_idx.get(anc)
                            if j is not None:
                                include_idx.add(j)

            if not include_idx:
                messagebox.showinfo("WBS e descrição", "Não há descrições do utilizador para exportar.")
                return False

            mini_df = df.loc[sorted(include_idx)]

            with pd.ExcelWriter(path, engine="openpyxl") as writer:
                mini_df.to_excel(writer, index=False, startrow=1)
                ws = next(iter(writer.sheets.values()))
                ws["A1"] = "WBS com descrições"

            messagebox.showinfo("WBS e descrição", f"WBS exportado com {len(mini_df)} linhas totais. Salvo em:\n{path}.")

            self.app.last_exported_wbs = path
            self.app.wbs_xlsx_var.set(path)
            
            return True

        except Exception as e:
            messagebox.showerror("Erro", f"Falha a exportar o WBS:\n{e}")
            return False


    def on_save_and_export(self):
        self.app.wbs_finalized = True

        ok = self.on_export_wbs()
        if not ok:
            return

        self.app.open_mapping(source="wbs")