import json
import tkinter as tk
from tkinter import ttk, messagebox, filedialog
from pathlib import Path

import pandas as pd

from app.gui.wbs_helpers import (
    detect_relevant_leaves,
    find_level10_text,
    find_wbs_columns,
    unpack_core_columns,
    split_levels,
)
from app.core.structural_engine import migrate_rule_v1_to_v2


def _parse_value_token(s: str):
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


class FilterBlock(tk.LabelFrame):

    def __init__(self, parent, index: int, predefs_by_class: dict,
                 on_remove, qty_type_var: tk.StringVar, free_mode: bool = False):
        super().__init__(parent, padx=6, pady=4)
        self.index           = index
        self.predefs_by_class = predefs_by_class
        self._on_remove      = on_remove
        self.qty_type_var    = qty_type_var
        self.free_mode       = free_mode
        self.prop_rows: list = []

        self._build()
        self._update_qty_visibility()
        self._trace_id = qty_type_var.trace_add("write", lambda *_: self._safe_update_qty_visibility())

    def _build(self):
        self.grid_columnconfigure(1, weight=1)

        hdr = tk.Frame(self)
        hdr.grid(row=0, column=0, columnspan=3, sticky="we", pady=(0, 2))
        hdr.grid_columnconfigure(0, weight=1)
        self._lbl_title = tk.Label(hdr, text=f"Classe IFC {self.index}",
                                   font=("Segoe UI", 9, "bold"))
        self._lbl_title.pack(side="left")
        tk.Button(hdr, text="Remover", fg="red",
                  command=self._do_remove).pack(side="right")

        tk.Label(self, text="IfcClass:").grid(row=1, column=0, sticky="e", padx=(0,6), pady=2)
        if self.free_mode:
            self.ifc_class = tk.Entry(self)
        else:
            self.ifc_class = ttk.Combobox(self, state="readonly",
                                           values=sorted(self.predefs_by_class.keys()))
            self.ifc_class.bind("<<ComboboxSelected>>", self._on_class_selected)
        self.ifc_class.grid(row=1, column=1, sticky="we", pady=2)

        tk.Label(self, text="PredefinedType:").grid(row=2, column=0, sticky="e", padx=(0,6), pady=2)
        if self.free_mode:
            self.predef = tk.Entry(self)
        else:
            self.predef = ttk.Combobox(self, state="readonly", values=[])
            self.predef.bind("<<ComboboxSelected>>", self._toggle_objtype)
        self.predef.grid(row=2, column=1, sticky="we", pady=2)

        tk.Label(self, text="ObjectType (se PredefinedType = USERDEFINED):").grid(
            row=3, column=0, sticky="e", padx=(0,6), pady=2)
        self.objtype = tk.Entry(self,
                                state="normal" if self.free_mode else "disabled")
        self.objtype.grid(row=3, column=1, sticky="we", pady=2)

        xtra = tk.LabelFrame(self, text="Filtros adicionais (opcional)", padx=6, pady=4)
        xtra.grid(row=4, column=0, columnspan=3, sticky="we", pady=(4,2))
        xtra.grid_columnconfigure(0, weight=1)
        hdr2 = tk.Frame(xtra)
        hdr2.pack(fill="x", pady=(0,2))
        for ci, lbl in enumerate(["Grupo (Pset)", "Propriedade", "Valor"]):
            tk.Label(hdr2, text=lbl, font=("Segoe UI", 8, "bold")).grid(
                row=0, column=ci, sticky="w", padx=(0,4))
        hdr2.grid_columnconfigure(0, weight=2, uniform="ef")
        hdr2.grid_columnconfigure(1, weight=2, uniform="ef")
        hdr2.grid_columnconfigure(2, weight=1, uniform="ef")
        self._extra_container = tk.Frame(xtra)
        self._extra_container.pack(fill="x")
        tk.Button(xtra, text="+ Adicionar filtro",
                  command=self._add_prop_row).pack(anchor="e")

        self.qty_lf = tk.LabelFrame(
            self, text="Quantidade – propriedade e grupo (pset) para esta classe",
            padx=6, pady=6)
        self.qty_lf.grid(row=5, column=0, columnspan=3, sticky="we", pady=(4,2))
        self.qty_lf.grid_columnconfigure(1, weight=1)
        tk.Label(self.qty_lf, text="Pset:").grid(row=0, column=0, sticky="e", padx=(0,6), pady=2)
        self.q_pset = tk.Entry(self.qty_lf)
        self.q_pset.grid(row=0, column=1, sticky="we", pady=2)
        tk.Label(self.qty_lf, text="Propriedade:").grid(row=1, column=0, sticky="e", padx=(0,6), pady=2)
        self.q_prop = tk.Entry(self.qty_lf)
        self.q_prop.grid(row=1, column=1, sticky="we", pady=2)

    def _safe_update_qty_visibility(self):
        try:
            self._update_qty_visibility()
        except Exception:
            pass

    def _update_qty_visibility(self):
        if self.qty_type_var.get() == "prop":
            self.qty_lf.grid()
        else:
            self.qty_lf.grid_remove()

    def destroy(self):
        try:
            self.qty_type_var.trace_remove("write", self._trace_id)
        except Exception:
            pass
        super().destroy()

    def _do_remove(self):
        if callable(self._on_remove):
            self._on_remove(self)

    def _on_class_selected(self, *_):
        cls  = (self.ifc_class.get() or "").strip()
        opts = self.predefs_by_class.get(cls, [])
        self.predef.configure(values=opts)
        self.predef.set("" if not opts else opts[0])
        self._toggle_objtype()

    def _toggle_objtype(self, *_):
        if self.free_mode:
            return
        use = (self.predef.get() or "").strip().upper() == "USERDEFINED"
        self.objtype.configure(state=("normal" if use else "disabled"))
        if not use:
            self.objtype.delete(0, "end")

    def _add_prop_row(self, preset=None):
        row = tk.Frame(self._extra_container)
        row.pack(fill="x", pady=1)
        row.grid_columnconfigure(0, weight=2, uniform="ef")
        row.grid_columnconfigure(1, weight=2, uniform="ef")
        row.grid_columnconfigure(2, weight=1, uniform="ef")
        e_pset = tk.Entry(row); e_prop = tk.Entry(row); e_val = tk.Entry(row)
        e_pset.grid(row=0, column=0, sticky="we", padx=(0,4))
        e_prop.grid(row=0, column=1, sticky="we", padx=(0,4))
        e_val .grid(row=0, column=2, sticky="we", padx=(0,4))
        btn = tk.Button(row, text="✕", width=3,
                        command=lambda r=row: self._remove_prop_row(r))
        btn.grid(row=0, column=3)
        if preset:
            e_pset.insert(0, preset.get("pset", ""))
            e_prop.insert(0, preset.get("prop", ""))
            v = preset.get("value", "")
            e_val.insert(0, "" if v is None else str(v))
        self.prop_rows.append((row, e_pset, e_prop, e_val, btn))

    def _remove_prop_row(self, row_widget):
        for i, (r, *_) in enumerate(self.prop_rows):
            if r is row_widget:
                r.destroy()
                del self.prop_rows[i]
                break

    def get_data(self) -> dict:
        ic  = (self.ifc_class.get() if isinstance(self.ifc_class, tk.Entry)
               else self.ifc_class.get() or "").strip()
        prd = (self.predef.get() if isinstance(self.predef, tk.Entry)
               else self.predef.get() or "").strip().upper()
        if not ic or not prd:
            raise ValueError("IfcClass e PredefinedType são obrigatórios em cada classe.")

        ot = ""
        if prd == "USERDEFINED":
            ot = self.objtype.get().strip()
            if not ot:
                raise ValueError("ObjectType obrigatório quando PredefinedType = USERDEFINED.")

        props = []
        for _, ep, epr, ev, _ in self.prop_rows:
            pset = ep.get().strip(); prop = epr.get().strip(); val = ev.get().strip()
            if not pset and not prop and not val:
                continue
            if not pset or not prop:
                raise ValueError("Complete Pset e Propriedade nos filtros adicionais.")
            props.append({"pset": pset, "prop": prop,
                          "value": _parse_value_token(val)})

        entry = {"filter": {"ifc_class": ic, "predefined": prd,
                             "object_type": ot, "props": props}}

        if self.qty_type_var.get() == "prop":
            qp  = self.q_pset.get().strip()
            qpr = self.q_prop.get().strip()
            if not qp or not qpr:
                raise ValueError(
                    f"Pset e Propriedade de quantidade são obrigatórios para '{ic}'.")
            entry["quantity_detail"] = {"pset": qp, "prop": qpr}

        return entry

    def load_from(self, mapping_entry: dict):
        f   = mapping_entry.get("filter", {})
        ic  = f.get("ifc_class", "")
        prd = f.get("predefined", "")
        ot  = f.get("object_type", "")

        if ic:
            if isinstance(self.ifc_class, ttk.Combobox):
                cur = list(self.ifc_class["values"])
                if ic not in cur:
                    self.ifc_class.configure(values=sorted(set(cur) | {ic}))
                self.ifc_class.set(ic)
                opts = self.predefs_by_class.get(ic, [])
                if prd and prd not in opts:
                    opts = sorted(set(opts) | {prd})
                self.predef.configure(values=opts)
            else:
                self.ifc_class.delete(0, "end")
                self.ifc_class.insert(0, ic)

        if prd:
            if isinstance(self.predef, ttk.Combobox):
                self.predef.set(prd)
            else:
                self.predef.delete(0, "end")
                self.predef.insert(0, prd)

        self._toggle_objtype()
        if ot:
            self.objtype.configure(state="normal")
            self.objtype.delete(0, "end")
            self.objtype.insert(0, ot)

        for p in (f.get("props") or []):
            self._add_prop_row(preset=p)

        qd = mapping_entry.get("quantity_detail", {})
        if qd.get("pset"):
            self.q_pset.delete(0, "end"); self.q_pset.insert(0, qd["pset"])
        if qd.get("prop"):
            self.q_prop.delete(0, "end"); self.q_prop.insert(0, qd["prop"])

    def update_title(self, index: int):
        self.index = index
        self._lbl_title.configure(text=f"Classe IFC {index}")


class QtyPage(tk.Frame):

    def __init__(self, master, app):
        super().__init__(master)
        self.app = app

        self.relevant_codes: list[str] = []
        self.relevant_set:   set[str]  = set()
        self.code_to_desc:   dict      = {}
        self._selected_code: str | None = None
        self.path_stack:     list[str]  = []
        self.predefs_by_class: dict     = {}

        self._filter_blocks: list[FilterBlock] = []
        self.qty_type_var  = tk.StringVar(value="prop")
        self._ifc_mode_var = tk.StringVar(value="project")
        self._ifc_loaded   = False
        self._mode_confirmed = False

        self.mode = "home"

        self._build_ui()
        self.after_idle(lambda: self.refresh_items(silent=True))

    @property
    def selected_code(self) -> str | None:
        return self._selected_code

    @selected_code.setter
    def selected_code(self, value: str | None):
        self._selected_code = value

    def _build_ui(self):
        self.grid_columnconfigure(0, weight=2, minsize=560)
        self.grid_columnconfigure(1, weight=1, minsize=480)
        self.grid_rowconfigure(0, weight=1)
        self._build_left()
        self._build_right()

    def _build_left(self):
        left = tk.Frame(self, bd=1, relief="groove")
        left.grid(row=0, column=0, sticky="nsew", padx=8, pady=8)

        self._top_container = tk.Frame(left)
        self._top_container.pack(fill="x")

        self.src_frame = tk.LabelFrame(self._top_container,
                                       text="WBS de entrada (com descrições)",
                                       padx=8, pady=6)
        self.src_frame.pack(fill="x", padx=8, pady=(8,4))
        row_src = tk.Frame(self.src_frame)
        row_src.pack(fill="x")
        tk.Label(row_src, text="WBS (Excel):").pack(side="left")
        self.wbs_entry = tk.Entry(row_src, textvariable=self.app.wbs_xlsx_var, width=36)
        self.wbs_entry.pack(side="left", padx=6, fill="x", expand=True)
        self.wbs_browse_btn = tk.Button(row_src, text="Procurar…",
            command=lambda: self.app.pick_file(
                self.app.wbs_xlsx_var, "Escolher WBS",
                [("Excel (*.xlsx *.xls)", "*.xlsx *.xls")]))
        self.wbs_browse_btn.pack(side="left", padx=(0,4))
        self.wbs_load_btn = tk.Button(row_src, text="Carregar",
                                      command=self._load_wbs_from_file)
        self.wbs_load_btn.pack(side="left")

        self._map_frame = tk.LabelFrame(self._top_container,
                                        text="Mapeamento existente (opcional)",
                                        padx=8, pady=6)
        self._map_frame.pack(fill="x", padx=8, pady=(0,4))
        map_frame = self._map_frame
        row_map = tk.Frame(map_frame)
        row_map.pack(fill="x")
        tk.Label(row_map, text="JSON:").pack(side="left")
        tk.Entry(row_map, textvariable=self.app.map_var, width=36).pack(
            side="left", padx=6, fill="x", expand=True)
        tk.Button(row_map, text="Procurar…",
                  command=self._browse_mapping).pack(side="left", padx=(0,4))
        tk.Button(row_map, text="Carregar",
                  command=self._load_mapping_from_file).pack(side="left")

        self._nav_anchor = tk.Frame(left)
        self._nav_anchor.pack(fill="x")

        self._list_container = tk.Frame(left)
        self._list_container.pack(fill="both", expand=True)

        nav = tk.Frame(self._list_container)
        nav.pack(fill="x", padx=8, pady=(4,2))
        self.path_label = tk.Label(nav, text="Caminho: — (Nível 1)")
        self.path_label.pack(side="left")

        nav_btns = tk.Frame(self._list_container)
        nav_btns.pack(fill="x", padx=8, pady=(0,4))
        self.btn_up = tk.Button(nav_btns, text="Nível acima",
                                command=self.on_back, state="disabled")
        self.btn_down = tk.Button(nav_btns, text="Nível abaixo",
                                  command=self.on_next, state="disabled")
        self.btn_up.pack(side="left")
        self.btn_down.pack(side="left", padx=(6,0))

        list_wrap = tk.Frame(self._list_container)
        list_wrap.pack(fill="both", expand=True, padx=8, pady=(0,4))
        self.listbox = tk.Listbox(list_wrap, height=20, width=52, exportselection=False)
        self.listbox.pack(side="left", fill="both", expand=True)
        yscroll = tk.Scrollbar(list_wrap, orient="vertical", command=self.listbox.yview)
        yscroll.pack(side="right", fill="y")
        self.listbox.configure(yscrollcommand=yscroll.set)
        self.listbox.bind("<<ListboxSelect>>", self.on_select_list)

        bottom = tk.Frame(self._list_container)
        bottom.pack(fill="x", padx=8, pady=(0,8))
        tk.Label(bottom, text="✓ = tem mapeamento associado",
                 fg="#555", font=("Segoe UI", 8)).pack(side="left")
        tk.Button(bottom, text="Limpar regra do item",
                  command=self.clear_rule_current).pack(side="right")

        self._top_container.pack_forget()
        self._list_container.pack_forget()

    def _build_right(self):
        right_wrapper = tk.Frame(self)
        right_wrapper.grid(row=0, column=1, sticky="nsew", padx=(0,8), pady=8)
        right_wrapper.grid_columnconfigure(0, weight=1)
        right_wrapper.grid_rowconfigure(0, weight=1)

        canvas = tk.Canvas(right_wrapper, bg="white", highlightthickness=0, bd=0)
        canvas.grid(row=0, column=0, sticky="nsew")
        sb = tk.Scrollbar(right_wrapper, orient="vertical", command=canvas.yview)
        sb.grid(row=0, column=1, sticky="ns")
        canvas.configure(yscrollcommand=sb.set)

        self._right = tk.Frame(canvas, bg="white")
        cw = canvas.create_window(0, 0, window=self._right, anchor="nw")

        def _mw(e): canvas.yview_scroll(int(-1*(e.delta/120)), "units")
        canvas.bind("<MouseWheel>", _mw)
        canvas.bind("<Button-4>", lambda e: canvas.yview_scroll(-1, "units"))
        canvas.bind("<Button-5>", lambda e: canvas.yview_scroll(1, "units"))

        def _cfg(e=None):
            canvas.configure(scrollregion=canvas.bbox("all"))
            w = canvas.winfo_width()
            if w > 1: canvas.itemconfig(cw, width=w)
        self._right.bind("<Configure>", _cfg)
        canvas.bind("<Configure>", _cfg)
        self._right.grid_columnconfigure(0, weight=1)

        r = self._right

        mode_lf = tk.LabelFrame(r, text="Modo de mapeamento", padx=8, pady=6, bg="white")
        mode_lf.grid(row=0, column=0, sticky="we", padx=8, pady=(8,4))
        self._rb_project = tk.Radiobutton(
            mode_lf, text="Modo projeto  (carregar IFC — dropdowns automáticos)",
            variable=self._ifc_mode_var, value="project", bg="white")
        self._rb_project.pack(anchor="w")
        tk.Label(mode_lf, text="     Só poderá mapear elementos existentes no ficheiro IFC carregado.",
                 fg="#666", font=("Segoe UI", 8), bg="white").pack(anchor="w")
        self._rb_generic = tk.Radiobutton(
            mode_lf, text="Modo genérico  (sem IFC — escrever classes manualmente)",
            variable=self._ifc_mode_var, value="generic", bg="white")
        self._rb_generic.pack(anchor="w")
        self._btn_confirm_mode = tk.Button(
            mode_lf, text="Confirmar modo →", command=self._confirm_mode)
        self._btn_confirm_mode.pack(anchor="e", pady=(6,0))

        self.ifc_lf = tk.LabelFrame(r, text="IFC de entrada", padx=8, pady=6, bg="white")
        self.ifc_lf.grid(row=1, column=0, sticky="we", padx=8, pady=(0,4))
        self.ifc_lf.grid_columnconfigure(1, weight=1)
        tk.Label(self.ifc_lf, text="IFC:", bg="white").grid(
            row=0, column=0, sticky="w", padx=(0,6), pady=4)
        tk.Entry(self.ifc_lf, textvariable=self.app.ifc_var, width=36).grid(
            row=0, column=1, sticky="we", padx=(0,6), pady=4)
        tk.Button(self.ifc_lf, text="Procurar…",
                  command=lambda: self.app.pick_file(
                      self.app.ifc_var, "Selecionar IFC", [("IFC", "*.ifc")])).grid(
            row=0, column=2, padx=(0,4), pady=4)
        tk.Button(self.ifc_lf, text="Carregar",
                  command=self._load_ifc).grid(row=0, column=3, pady=4)

        self.title_var = tk.StringVar(value="Selecione um item (folha) à esquerda…")
        tk.Label(r, textvariable=self.title_var,
                 font=("Segoe UI", 11, "bold"), bg="white").grid(
            row=2, column=0, sticky="w", padx=8, pady=(8,2))

        desc_lf = tk.LabelFrame(r, text="Descrição do utilizador (WBS)",
                                padx=8, pady=6, bg="white")
        desc_lf.grid(row=3, column=0, sticky="we", padx=8, pady=(0,4))
        self.user_desc_lbl = tk.Label(desc_lf, text="—", fg="#555",
                                      wraplength=400, justify="left", bg="white")
        self.user_desc_lbl.pack(fill="x")

        qty_lf = tk.LabelFrame(r, text="Modo de quantificação (obrigatório)",
                               padx=8, pady=6, bg="white")
        qty_lf.grid(row=4, column=0, sticky="we", padx=8, pady=(0,4))
        self._qty_lf = qty_lf
        tk.Radiobutton(qty_lf,
                       text="Leitura de propriedade  (ler valor de uma propriedade IFC)",
                       variable=self.qty_type_var, value="prop",
                       command=self._on_qty_type_changed, bg="white").pack(anchor="w")
        tk.Radiobutton(qty_lf,
                       text="Contagem de elementos  (contar elementos IFC filtrados)",
                       variable=self.qty_type_var, value="count",
                       command=self._on_qty_type_changed, bg="white").pack(anchor="w")

        self.filters_lf = tk.LabelFrame(
            r, text="Elementos IFC a serem pesquisados (obrigatório)",
            padx=8, pady=6, bg="white")
        self.filters_lf.grid(row=5, column=0, sticky="we", padx=8, pady=(0,4))
        self.filters_lf.grid_columnconfigure(0, weight=1)
        self._blocks_container = tk.Frame(self.filters_lf, bg="white")
        self._blocks_container.pack(fill="x", expand=True)
        self._blocks_container.grid_columnconfigure(0, weight=1)
        tk.Button(self.filters_lf, text="+ Adicionar classe IFC",
                  command=lambda: self._add_filter_block()).pack(anchor="e", pady=(4,0))

        mat_lf = tk.LabelFrame(r, text="Material (opcional)", padx=8, pady=6, bg="white")
        mat_lf.grid(row=6, column=0, sticky="we", padx=8, pady=(0,4))
        mat_lf.grid_columnconfigure(1, weight=1)
        tk.Label(mat_lf, text="Material:", bg="white").grid(
            row=0, column=0, sticky="e", padx=(0,6), pady=4)
        self._mat_lf = mat_lf
        self.material = ttk.Combobox(mat_lf, state="readonly", values=[])
        self.material.grid(row=0, column=1, sticky="we", pady=4)

        agr_lf = tk.LabelFrame(r, text="Propriedade de Agrupamento (opcional)",
                               padx=8, pady=6, bg="white")
        agr_lf.grid(row=7, column=0, sticky="we", padx=8, pady=(0,4))
        self._agr_lf = agr_lf
        agr_lf.grid_columnconfigure(1, weight=1)
        tk.Label(agr_lf, text="Pset:", bg="white").grid(
            row=0, column=0, sticky="e", padx=(0,6), pady=4)
        self.agr_pset = tk.Entry(agr_lf)
        self.agr_pset.grid(row=0, column=1, sticky="we", pady=4)
        tk.Label(agr_lf, text="Propriedade:", bg="white").grid(
            row=1, column=0, sticky="e", padx=(0,6), pady=4)
        self.agr_prop = tk.Entry(agr_lf)
        self.agr_prop.grid(row=1, column=1, sticky="we", pady=4)

        act = tk.Frame(r, bg="white")
        act.grid(row=8, column=0, sticky="we", padx=8, pady=(6,8))
        act.grid_columnconfigure(0, weight=1)
        act.grid_columnconfigure(1, weight=1)
        self._act_frame = act
        self.btn_save_rule = tk.Button(act, text="Guardar regra",
                                       command=self.save_rule_current, state="disabled")
        self.btn_save_rule.grid(row=0, column=0, sticky="we", padx=(0,4))
        tk.Button(act, text="Salvar e exportar",
                  command=self.save_rules_dialog).grid(
            row=0, column=1, sticky="we", padx=(4,0))

        for _s in (self.filters_lf, self._qty_lf, self._mat_lf,
                   self._agr_lf, self._act_frame):
            try: _s.grid_remove()
            except Exception: pass

        try: self.ifc_lf.grid_remove()
        except Exception: pass

        self._set_edit_enabled(False)

    def _confirm_mode(self):
        self._mode_confirmed = True
        self._rb_project.configure(state="disabled")
        self._rb_generic.configure(state="disabled")
        self._btn_confirm_mode.configure(state="disabled", text="Modo confirmado ✓")
        self._on_mode_changed()
        if self._ifc_mode_var.get() == "generic":
            self._top_container.pack(fill="x", before=self._nav_anchor)
            self._list_container.pack(fill="both", expand=True)

    def _on_mode_changed(self):
        is_project = self._ifc_mode_var.get() == "project"
        if self._mode_confirmed:
            if is_project:
                self.ifc_lf.grid()
            else:
                self.ifc_lf.grid_remove()

        cur_mat = self.material.get()
        self.material.destroy()
        if is_project:
            self.material = ttk.Combobox(self._mat_lf, state="readonly", values=[])
        else:
            self.material = tk.Entry(self._mat_lf)
            if cur_mat:
                self.material.insert(0, cur_mat)
        self.material.grid(row=0, column=1, sticky="we", pady=4)
        if is_project and cur_mat:
            try: self.material.set(cur_mat)
            except Exception: pass

        current_data = []
        for blk in self._filter_blocks:
            try:
                current_data.append(blk.get_data())
            except Exception:
                current_data.append(None)
        self._clear_filter_blocks()
        if current_data:
            for d in current_data:
                self._add_filter_block(data=d)
        else:
            self._add_filter_block()

    def _add_filter_block(self, data: dict | None = None) -> FilterBlock:
        free = self._ifc_mode_var.get() == "generic"
        idx  = len(self._filter_blocks) + 1
        blk  = FilterBlock(
            self._blocks_container,
            index=idx,
            predefs_by_class=self.predefs_by_class,
            on_remove=self._remove_filter_block,
            qty_type_var=self.qty_type_var,
            free_mode=free,
        )
        blk.grid(row=idx - 1, column=0, sticky="we", pady=(0, 4))
        self._filter_blocks.append(blk)
        if data:
            blk.load_from(data)
        return blk

    def _remove_filter_block(self, block: FilterBlock):
        if len(self._filter_blocks) <= 1:
            messagebox.showwarning("Mapeamento IFC",
                                   "Deve existir pelo menos uma classe IFC.")
            return
        if block in self._filter_blocks:
            self._filter_blocks.remove(block)
            block.destroy()
            self._renumber_blocks()

    def _renumber_blocks(self):
        for i, blk in enumerate(self._filter_blocks, start=1):
            blk.update_title(i)
            blk.grid(row=i - 1, column=0, sticky="we", pady=(0, 4))

    def _clear_filter_blocks(self):
        for blk in self._filter_blocks:
            try: blk.destroy()
            except Exception: pass
        self._filter_blocks.clear()

    def _on_qty_type_changed(self):
        pass

    def _load_ifc(self):
        path = (self.app.ifc_var.get() or "").strip()
        if not path or not path.lower().endswith(".ifc"):
            messagebox.showwarning("Mapeamento IFC",
                                   "Selecione primeiro um ficheiro IFC (.ifc).")
            return
        try:
            import ifcopenshell
        except Exception:
            messagebox.showerror("Mapeamento IFC", "Biblioteca 'ifcopenshell' em falta.")
            return
        try:
            model = ifcopenshell.open(path)
        except Exception as e:
            messagebox.showerror("Mapeamento IFC", f"Falha a abrir o IFC:\n{e}")
            return

        from collections import defaultdict
        c2p = defaultdict(set)
        for el in (model.by_type("IfcProduct") or []):
            pre = getattr(el, "PredefinedType", None)
            c2p[el.is_a()].add(
                pre.upper().strip() if isinstance(pre, str) else "NOTDEFINED")
        self.predefs_by_class = {k: sorted(v) for k, v in c2p.items()}
        self.app.inv.ifc_file = model

        mats = self.app.inv.extract_all_materials()
        self.material.configure(values=sorted(mats.keys()))
        self.material.set("")

        for blk in self._filter_blocks:
            if isinstance(blk.ifc_class, ttk.Combobox):
                cur = blk.ifc_class.get()
                blk.ifc_class.configure(values=sorted(self.predefs_by_class.keys()))
                if cur: blk.ifc_class.set(cur)
            blk.predefs_by_class = self.predefs_by_class

        self._ifc_loaded = True
        self._top_container.pack(fill="x", before=self._nav_anchor)
        self._list_container.pack(fill="both", expand=True)
        if self._is_leaf_selected():
            self._set_edit_enabled(True)
            self.load_rule_into_form(self.app.rules.get(self._selected_code))
            self.title_var.set(
                f"Item selecionado: {self._selected_code}"
                + (f" — {self.code_to_desc.get(self._selected_code, '')}"
                   if self._selected_code else ""))
        messagebox.showinfo("Mapeamento IFC", "IFC carregado com sucesso.")

    def _load_wbs_from_file(self):
        path = (self.app.wbs_xlsx_var.get() or "").strip()
        if not path:
            messagebox.showwarning("Mapeamento IFC", "Selecione primeiro o ficheiro WBS.")
            return
        try:
            df   = pd.read_excel(path, header=1)
            cols = find_wbs_columns(df)
            col_wbs, col_desc, col_nivel = unpack_core_columns(cols)
            if not all([col_wbs, col_desc, col_nivel]):
                df   = pd.read_excel(path, header=0)
                cols = find_wbs_columns(df)
                col_wbs, col_desc, col_nivel = unpack_core_columns(cols)
            df[col_nivel] = split_levels(df, col_nivel)
            self.app.df_raw    = df
            self.app.df_desc0  = pd.Series([None] * len(df))
            self.app.col_nivel = col_nivel
            self.app.col_wbs   = col_wbs
            self.app.col_desc  = col_desc
            self.app.wbs_cols  = cols
            self.app.wbs_finalized = True
            self.refresh_items()
            messagebox.showinfo("Mapeamento IFC", "WBS carregado com sucesso.")
        except Exception as e:
            messagebox.showerror("Erro", f"Falha a ler o WBS:\n{e}")

    def _browse_mapping(self):
        p = filedialog.askopenfilename(
            title="Escolher mapeamento JSON",
            filetypes=[("JSON", "*.json"), ("Todos", "*.*")])
        if p: self.app.map_var.set(p)

    def _load_mapping_from_file(self):
        path = (self.app.map_var.get() or "").strip()
        if not path or not Path(path).is_file():
            messagebox.showerror("Mapeamento IFC", "Caminho inválido ou em falta.")
            return
        try:
            with open(path, "r", encoding="utf-8") as f:
                data = json.load(f)
        except Exception as e:
            messagebox.showerror("Erro", f"Não foi possível ler o ficheiro:\n{e}")
            return

        from app.core.structural_engine import load_and_migrate_rules
        rules = load_and_migrate_rules(data)
        if not rules:
            messagebox.showerror("Mapeamento IFC", "Estrutura inválida ou vazia.")
            return

        is_partial = data.get("partial", False)
        if self.app.rules:
            msg = "Substituir o mapeamento em memória pelo ficheiro carregado?"
            if is_partial:
                msg += "\n\n(Este é um mapeamento parcial — campos por completar.)"
            if not messagebox.askyesno("Mapeamento", msg):
                return

        self.app.rules = rules
        if not self.app.ifc_var.get().strip() and data.get("ifc_path"):
            self.app.ifc_var.set(data["ifc_path"])
        if not self.app.wbs_xlsx_var.get().strip() and data.get("wbs_path"):
            self.app.wbs_xlsx_var.set(data["wbs_path"])

        partial_note = " (parcial)" if is_partial else ""
        messagebox.showinfo(
            "Mapeamento IFC",
            f"Mapeamento{partial_note} carregado: {len(rules)} código(s) WBS.\n"
            f"Selecione um item na lista para ver ou editar.")
        self._render_list()
        self._set_edit_enabled(False)

    def _auto_load_partial_mapping(self):
        path = (self.app.map_var.get() or "").strip()
        if not path or not Path(path).is_file() or self.app.rules:
            return
        try:
            with open(path, "r", encoding="utf-8") as f:
                data = json.load(f)
            if not data.get("partial", False):
                return
            from app.core.structural_engine import load_and_migrate_rules
            rules = load_and_migrate_rules(data)
            if rules:
                self.app.rules = rules
                self._render_list()
        except Exception:
            pass

    @staticmethod
    def _tokens(code: str):
        return [p for p in str(code).split(".") if p]

    def _current_level(self):
        return 1 if not self.path_stack else len(self._tokens(self.path_stack[-1])) + 1

    def _is_leaf(self, code: str):
        return code in self.relevant_set

    def _has_child_in_relevant(self, code: str):
        pref = f"{code}."
        return any(c.startswith(pref) for c in self.relevant_codes)

    def _candidates(self, level: int, prefix: str | None):
        cands = set()
        for leaf in self.relevant_codes:
            toks = self._tokens(leaf)
            if level <= len(toks):
                cand = ".".join(toks[:level])
                if prefix is None or cand.startswith(prefix + "."):
                    cands.add(cand)
        return sorted(cands)

    def _render_list(self):
        lvl    = self._current_level()
        prefix = self.path_stack[-1] if self.path_stack else None
        cands  = self._candidates(lvl, prefix)
        self.listbox.delete(0, "end")
        for code in cands:
            desc  = self.code_to_desc.get(code, "")
            mark  = "✓ " if code in self.app.rules else ""
            label = f"{mark}{code}" + (f" — {desc}" if desc else "")
            self.listbox.insert("end", label)
        path_txt = " > ".join(self.path_stack) if self.path_stack else "—"
        self.path_label.configure(text=f"Caminho: {path_txt} (Nível {lvl})")
        self.btn_up.configure(state="normal" if self.path_stack else "disabled")

    @staticmethod
    def _code_from_label(label: str) -> str:
        return label.lstrip("✓ ").strip().split(" — ")[0].strip()

    def set_mode(self, source: str):
        self.mode = "wbs" if source == "wbs" else "home"
        if self.mode == "wbs":
            self._disable_wbs_upload()
            self.after(100, lambda: self.refresh_items(silent=True))
            self.after(200, self._auto_load_partial_mapping)
        else:
            self._enable_wbs_upload()
        self.clear_form()
        self._set_edit_enabled(False)
        self._render_list()

    def _disable_wbs_upload(self):
        for w in (self.wbs_entry, self.wbs_browse_btn, self.wbs_load_btn):
            try: w.configure(state="disabled", bg="#F0F0F0") if isinstance(w, tk.Entry) \
                 else w.configure(state="disabled")
            except Exception: pass
        self.src_frame.configure(text="WBS de entrada (carregado da aba anterior)")

    def _enable_wbs_upload(self):
        for w in (self.wbs_entry, self.wbs_browse_btn, self.wbs_load_btn):
            try: w.configure(state="normal", bg="white") if isinstance(w, tk.Entry) \
                 else w.configure(state="normal")
            except Exception: pass
        self.src_frame.configure(text="WBS de entrada (com descrições)")

    def refresh_items(self, silent: bool = False):
        if self.mode == "wbs" and not getattr(self.app, "wbs_finalized", False):
            if not silent:
                messagebox.showinfo("Mapeamento IFC", "Finalize o WBS na aba anterior.")
            return
        if self.app.df_raw is None:
            if not silent:
                messagebox.showinfo("Mapeamento IFC", "Carregue um WBS (Excel) acima.")
            return

        pairs = detect_relevant_leaves(
            self.app.df_raw, self.app.col_nivel, self.app.col_wbs,
            self.app.col_desc, self.app.df_desc0)
        self.relevant_codes = [c for c, _ in pairs]
        self.relevant_set   = set(self.relevant_codes)

        self.code_to_desc.clear()
        for _, row in self.app.df_raw.iterrows():
            code = row[self.app.col_wbs]
            if isinstance(code, str) and code.strip():
                val = row[self.app.col_desc]
                self.code_to_desc[code.strip()] = "" if pd.isna(val) else str(val).strip()

        self.path_stack    = []
        self._selected_code = None
        self.title_var.set("Selecione um item (folha) à esquerda…")
        self.user_desc_lbl.configure(text="—")
        self.clear_form()
        self._set_edit_enabled(False)
        self._render_list()

    def on_select_list(self, *_):
        sel = self.listbox.curselection()
        if not sel:
            self._selected_code = None
            self.btn_down.configure(state="disabled")
            self.title_var.set("Selecione um item (folha) à esquerda…")
            self.user_desc_lbl.configure(text="—")
            self.clear_form()
            self._set_edit_enabled(False)
            return

        code = self._code_from_label(self.listbox.get(sel[0]))
        can_down = self._has_child_in_relevant(code) and not self._is_leaf(code)
        self.btn_down.configure(state="normal" if can_down else "disabled")

        if self._is_leaf(code):
            self._selected_code = code
            base     = self.code_to_desc.get(code, "")
            user_txt = find_level10_text(
                self.app.df_raw, self.app.col_nivel,
                self.app.col_wbs, self.app.col_desc, code)
            self.title_var.set(f"Item selecionado: {code} — {base}")
            self.user_desc_lbl.configure(text=user_txt or "—")
            can_edit = self._mapping_enabled()
            self._set_edit_enabled(can_edit)
            if can_edit:
                self.load_rule_into_form(self.app.rules.get(code))
            else:
                if self._ifc_mode_var.get() == "project" and not self._ifc_loaded:
                    self.title_var.set(
                        f"Item selecionado: {code} — {base}  ⚠ Carregue o IFC primeiro")
        else:
            self._selected_code = None
            self.title_var.set(f"Secção: {code} — {self.code_to_desc.get(code, '')}")
            self.user_desc_lbl.configure(text="—")
            self.clear_form()
            self._set_edit_enabled(False)

    def on_next(self):
        sel = self.listbox.curselection()
        if not sel:
            messagebox.showwarning("Mapeamento IFC", "Selecione um item.")
            return
        code = self._code_from_label(self.listbox.get(sel[0]))
        if self._is_leaf(code) or not self._has_child_in_relevant(code):
            messagebox.showinfo("Mapeamento IFC", "Já está no último nível.")
            return
        self.path_stack.append(code)
        self._render_list()

    def on_back(self):
        if not self.path_stack:
            return
        self.path_stack.pop()
        self._render_list()

    def clear_form(self):
        self._clear_filter_blocks()
        try: self.material.set("")
        except Exception: pass
        self.qty_type_var.set("prop")
        for w in (self.agr_pset, self.agr_prop):
            try: w.delete(0, "end")
            except Exception: pass

    def load_rule_into_form(self, rule: dict | None):
        self.clear_form()
        if not rule:
            self._add_filter_block()
            return
        rule = migrate_rule_v1_to_v2(rule)
        qty_type = rule.get("quantity", {}).get("type", "prop")
        self.qty_type_var.set(qty_type if qty_type in ("prop", "count") else "prop")

        mat = rule.get("material", "")
        if mat:
            try: self.material.set(mat)
            except Exception: pass

        for m in rule.get("mappings", []):
            self._add_filter_block(data=m)
        if not self._filter_blocks:
            self._add_filter_block()

        agr = rule.get("agrupamento") or {}
        if agr.get("pset"): self.agr_pset.insert(0, agr["pset"])
        if agr.get("prop"): self.agr_prop.insert(0, agr["prop"])

    def clear_rule_current(self):
        if not self._selected_code or self._selected_code not in self.relevant_set:
            messagebox.showinfo("Mapeamento IFC",
                                "Selecione um item (folha) primeiro.")
            return
        self.app.rules.pop(self._selected_code, None)
        self.clear_form()
        self._add_filter_block()
        self._set_edit_enabled(True)
        self._render_list()
        messagebox.showinfo("Mapeamento IFC",
                            f"Regra do item {self._selected_code} removida.")

    def save_rule_current(self):
        if not self._selected_code or self._selected_code not in self.relevant_set:
            messagebox.showwarning("Mapeamento IFC",
                                   "Selecione um item (folha) à esquerda.")
            return
        if not self._filter_blocks:
            messagebox.showerror("Mapeamento IFC",
                                 "Adicione pelo menos uma classe IFC.")
            return
        try:
            mappings = [blk.get_data() for blk in self._filter_blocks]
        except ValueError as e:
            messagebox.showerror("Mapeamento IFC", str(e))
            return

        rule = {
            "mappings": mappings,
            "material": self.material.get().strip(),
            "quantity": {"type": self.qty_type_var.get()},
            "agrupamento": {
                "pset": self.agr_pset.get().strip(),
                "prop": self.agr_prop.get().strip(),
            },
        }
        self.app.rules[self._selected_code] = rule
        self._render_list()
        messagebox.showinfo("Mapeamento IFC",
                            f"Regra guardada para {self._selected_code}.")

    def save_rules_dialog(self):
        if not self.app.rules:
            messagebox.showinfo("Mapeamento IFC", "Não há regras para guardar.")
            return

        errors = []
        for code, rule in self.app.rules.items():
            ok, msg = self._validate_rule(code, migrate_rule_v1_to_v2(rule))
            if not ok:
                errors.append(msg)
        if errors:
            messagebox.showerror("Mapeamento",
                                 "Erros de validação:\n" + "\n".join(errors))
            return

        initialdir = (Path(self.app.wbs_xlsx_var.get().strip()).parent
                      if self.app.wbs_xlsx_var.get().strip() else Path.home())
        path = filedialog.asksaveasfilename(
            title="Guardar mapeamento",
            defaultextension=".json",
            filetypes=[("JSON", "*.json")],
            initialdir=str(initialdir),
            initialfile="mapeamento_ifc.json")
        if not path:
            return

        payload = {
            "version": 2,
            "partial": False,
            "ifc_path": self.app.ifc_var.get().strip() or None,
            "wbs_path": self.app.wbs_xlsx_var.get().strip() or None,
            "rules": {c: migrate_rule_v1_to_v2(r) for c, r in self.app.rules.items()},
        }
        try:
            with open(path, "w", encoding="utf-8") as f:
                json.dump(payload, f, ensure_ascii=False, indent=2)
            self.app.map_var.set(path)
            messagebox.showinfo("Mapeamento IFC", f"Mapeamento guardado:\n{path}")
            try: self.app.open_extract("from_mapping")
            except Exception:
                try: self.app.go_extract()
                except Exception: pass
        except Exception as e:
            messagebox.showerror("Erro", f"Falha a guardar:\n{e}")

    def _validate_rule(self, code: str, rule: dict) -> tuple[bool, str]:
        mappings = rule.get("mappings", [])
        if not mappings:
            return False, f"[{code}] Pelo menos uma classe IFC é obrigatória."
        qty_type = rule.get("quantity", {}).get("type", "")
        if qty_type not in ("prop", "count"):
            if qty_type == "":
                return True, ""
            return False, f"[{code}] Tipo de quantificação inválido: '{qty_type}'."
        for i, m in enumerate(mappings):
            f = m.get("filter", {})
            if not f.get("ifc_class"):
                continue
            if f.get("ifc_class") and not f.get("predefined"):
                continue
            if f.get("predefined", "").upper() == "USERDEFINED" and not f.get("object_type"):
                return False, f"[{code}] Classe {i+1}: object_type obrigatório para USERDEFINED."
            if qty_type == "prop":
                qd = m.get("quantity_detail", {})
                if not qd.get("pset") and not qd.get("prop"):
                    continue
                if not qd.get("pset") or not qd.get("prop"):
                    return False, f"[{code}] Classe {i+1}: Pset e Propriedade de quantidade são obrigatórios."
        return True, ""

    def _is_leaf_selected(self) -> bool:
        return bool(self._selected_code) and self._selected_code in self.relevant_set

    def _mapping_enabled(self) -> bool:
        if not self._is_leaf_selected():
            return False
        if self._ifc_mode_var.get() == "project" and not self._ifc_loaded:
            return False
        return True

    def _set_edit_enabled(self, enabled: bool):
        for section in (self.filters_lf, self._qty_lf, self._mat_lf,
                        self._agr_lf, self._act_frame):
            try:
                if enabled:
                    section.grid()
                else:
                    section.grid_remove()
            except Exception:
                pass

        state       = "normal"   if enabled else "disabled"
        state_combo = "readonly" if enabled else "disabled"
        try: self.material.configure(state=state_combo if isinstance(self.material, ttk.Combobox) else state)
        except Exception: pass
        for w in (self.agr_pset, self.agr_prop):
            try: w.configure(state=state)
            except Exception: pass
        try: self.btn_save_rule.configure(state=state)
        except Exception: pass
