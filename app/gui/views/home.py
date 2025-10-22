# app/gui/views/home.py
import tkinter as tk
from tkinter import ttk
import webbrowser

class HomePage(tk.Frame):

    def __init__(self, master, app):
        super().__init__(master)
        self.app = app
        self._q2_built = False
        self._build_ui()

    def _build_ui(self):
        self.grid_columnconfigure(0, weight=1)

        title = tk.Label(self, text="Bem-vindo!", font=("Segoe UI", 14, "bold"))
        title.grid(row=0, column=0, sticky="w", padx=16, pady=(16, 6))

        msg = (           
            "A aplicação é dividida em três abas que podem ser utilizadas em sequência ou separadamente:\n"
            "• WBS e descrição: Carregue o WBS da buildingSMARTPortugal (.xls) e preencha as descrições dos itens.\n"
            "• Mapeamento IFC: À partir de um WBS com descrições (.xls), carregue o IFC (.ifc) e crie as regras que relacionam cada código WBS aos elementos IFC.\n"
            "• Extrair quantidades e Gerar WBS preenchido: À partir de um WBS com descrições (.xls) e um mapeamento (.json), investigue o IFC (.ifc) para extração de quantidades.\n\n"
            "Responda às perguntas abaixo para seguir para a etapa certa."
        )
        tk.Label(self, text=msg, justify="left", wraplength=900).grid(
            row=1, column=0, sticky="w", padx=16
        )

        ttk.Separator(self, orient="horizontal").grid(row=2, column=0, sticky="we", padx=16, pady=(12, 8))

        self.status_frame = tk.Frame(self, bg="#f0f0f0", relief="solid", borderwidth=1)
        self.status_frame.grid(row=3, column=0, sticky="we", padx=16, pady=(8, 12))
        tk.Label(
            self.status_frame, text="Estado Atual:", font=("Segoe UI", 10, "bold"), bg="#f0f0f0"
        ).pack(anchor="w", padx=10, pady=(8, 4))
        self.status_label = tk.Label(self.status_frame, text="", justify="left", bg="#f0f0f0", fg="#333")
        self.status_label.pack(anchor="w", padx=10, pady=(0, 8))

        self.q1_frame = tk.Frame(self)
        self.q1_frame.grid(row=4, column=0, sticky="w", padx=16, pady=(8, 4))
        tk.Label(
            self.q1_frame, text="1. Já tem o WBS com descrições dos itens?", font=("Segoe UI", 11, "bold")
        ).pack(anchor="w")
        q1_btns = tk.Frame(self.q1_frame)
        q1_btns.pack(anchor="w", pady=(6, 0))
        tk.Button(q1_btns, text="Não, preciso fazê-lo", command=self._goto_wbs, width=22).pack(side="left")
        tk.Button(q1_btns, text="Sim, já tenho", command=self._show_q2, width=18).pack(side="left", padx=(8, 0))

        self.q2_frame = tk.Frame(self)

        ttk.Separator(self, orient="horizontal").grid(
            row=99, column=0, sticky="we", padx=16, pady=(14, 6)
        )

        footer = tk.Frame(self)
        footer.grid(row=100, column=0, sticky="we", padx=16, pady=(0, 12))
        footer.grid_columnconfigure(0, weight=1)

        def _open(url: str):
            webbrowser.open(url)

        lbl_proj = tk.Label(
            footer,
            text="Projeto da buildingSMART Portugal",
            fg="#666666",
            font=("Segoe UI", 9, "underline"),
            cursor="hand2",
        )
        lbl_proj.grid(row=0, column=0, sticky="w")
        lbl_proj.bind("<Button-1>", lambda _e: _open("https://buildingsmart.pt/"))

        lbl_dev = tk.Label(
            footer,
            text="Desenvolvido por Andressa Oliveira",
            fg="#666666",
            font=("Segoe UI", 9, "underline"),
            cursor="hand2",
        )
        lbl_dev.grid(row=1, column=0, sticky="w")
        lbl_dev.bind("<Button-1>", lambda _e: _open("https://www.linkedin.com/in/andoliveira/"))

        self._refresh_status()

    def refresh_on_show(self):
        self._refresh_status()

    def _refresh_status(self):
        wbs_ok = self.app.has_user_descriptions() or self.app.wbs_finalized
        map_ok = self.app.has_ifc_mapping()

        status_parts = []
        
        if wbs_ok:
            status_parts.append("WBS com descrições: OK")
        else:
            status_parts.append("WBS com descrições: em falta")
        
        if map_ok:
            status_parts.append("Mapeamento IFC: OK")
        else:
            status_parts.append("Mapeamento IFC: em falta")
        
        status_text = "\n".join(status_parts)
        self.status_label.config(text=status_text)

        if wbs_ok and map_ok:
            self._show_shortcut_to_extract()

    def _goto_wbs(self):
        self.app.go_wbs()

    def _show_q2(self):
        if self._q2_built:
            self.q2_frame.grid(row=5, column=0, sticky="w", padx=16, pady=(12, 4))
            return

        self._q2_built = True
        self.q2_frame.grid(row=5, column=0, sticky="w", padx=16, pady=(12, 4))

        tk.Label(
            self.q2_frame, 
            text="2. Já tem o mapeamento entre WBS e IFC?",
            font=("Segoe UI", 11, "bold")
        ).pack(anchor="w")

        q2_btns = tk.Frame(self.q2_frame)
        q2_btns.pack(anchor="w", pady=(6, 0))
        
        tk.Button(
            q2_btns, 
            text="Não, preciso configurar",
            command=lambda: self.app.open_mapping("home"),
            width=25
        ).pack(side="left")

        tk.Button(
            q2_btns, 
            text="Sim, já tenho",
            command=lambda: self.app.open_extract("home"),
            width=18
        ).pack(side="left", padx=(8, 0))

    def _show_shortcut_to_extract(self):
        if hasattr(self, "_shortcut_built"):
            return
        
        self._shortcut_built = True
        
        shortcut_frame = tk.Frame(self, bg="#d4edda", relief="solid", borderwidth=1)
        shortcut_frame.grid(row=6, column=0, sticky="we", padx=16, pady=(16, 8))
        
        tk.Label(
            shortcut_frame,
            text="Tudo pronto! Podes ir diretamente para a extração.",
            font=("Segoe UI", 10, "bold"),
            bg="#d4edda",
            fg="#155724"
        ).pack(anchor="w", padx=10, pady=(8, 4))
        
        tk.Button(
            shortcut_frame,
            text="Ir para Extrair Quantidades",
            command=lambda: self.app.open_extract("home"),
            bg="#28a745",
            fg="white",
            font=("Segoe UI", 10, "bold"),
            relief="flat",
            padx=20,
            pady=8
        ).pack(anchor="w", padx=10, pady=(4, 8))