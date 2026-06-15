import tkinter as tk
from tkinter import ttk


class HomePage(tk.Frame):

    def __init__(self, master, app):
        super().__init__(master)
        self.app = app
        self._q2_frame = None
        self._q3_frame = None
        self._build_ui()

    def _build_ui(self):
        self.grid_columnconfigure(0, weight=1)

        hdr = tk.Frame(self, pady=16)
        hdr.grid(row=0, column=0, sticky="we")
        tk.Label(
            hdr,
            text="Bem-vindo — Extração de Quantidades IFC → WBS",
            font=("Segoe UI", 14, "bold"),
        ).pack()
        tk.Label(
            hdr,
            text="A aplicação está dividida em três etapas que podem ser realizadas em sequência ou de forma independente.",
            font=("Segoe UI", 10),
            fg="#555",
        ).pack(pady=(4, 0))

        steps_lf = tk.LabelFrame(self, text="Como funciona", padx=16, pady=12)
        steps_lf.grid(row=1, column=0, sticky="we", padx=20, pady=(0, 12))

        steps = [
            ("1. WBS e descrição",
             "Carregue o WBS da buildingSMART Portugal e adicione ou edite as descrições "
             "customizadas dos itens. Pode começar do zero ou continuar um WBS já parcialmente preenchido."),
            ("2. Mapeamento IFC",
             "Associe cada item WBS aos elementos IFC correspondentes. Pode usar um IFC de projeto "
             "(com dropdowns automáticos) ou definir o mapeamento manualmente para uso genérico."),
            ("3. Extrair quantidades e gerar MQT",
             "Com o WBS e o mapeamento prontos, extraia as quantidades do modelo IFC "
             "e gere os ficheiros de output."),
        ]
        for title, desc in steps:
            row = tk.Frame(steps_lf)
            row.pack(fill="x", pady=4)
            tk.Label(row, text=title, font=("Segoe UI", 10, "bold"), width=32, anchor="w").pack(side="left")
            tk.Label(row, text=desc, font=("Segoe UI", 10), fg="#333",
                     wraplength=720, justify="left", anchor="w").pack(side="left", fill="x", expand=True)

        q_frame = tk.LabelFrame(self, text="Responda às perguntas para seguir para a etapa certa",
                                padx=16, pady=12)
        q_frame.grid(row=2, column=0, sticky="we", padx=20, pady=(0, 12))

        self._build_question(
            q_frame,
            text="Pretende criar ou editar um WBS com descrições customizadas?",
            yes_label="Sim, quero criar ou editar",
            yes_cmd=lambda: self.app.notebook.select(1),
            no_label="Não, já tenho um WBS pronto",
            no_cmd=self._show_q2,
        )

        self._q2_frame = self._build_question(
            q_frame,
            text="Pretende criar ou editar um mapeamento WBSxIFC?",
            yes_label="Sim, quero criar ou editar o mapeamento",
            yes_cmd=lambda: self.app.notebook.select(2),
            no_label="Não, já tenho um mapeamento pronto",
            no_cmd=self._show_q3,
            hidden=True,
        )

        self._q3_frame = self._build_question(
            q_frame,
            text="Pretende extrair quantidades de um ficheiro IFC e gerar MQT?",
            yes_label="Sim, avançar para extração",
            yes_cmd=lambda: self.app.notebook.select(3),
            no_label=None,
            no_cmd=None,
            hidden=True,
        )

        footer = tk.Frame(self, pady=12)
        footer.grid(row=99, column=0, sticky="we", padx=20)

        def _link(text, url, pady=(0, 0)):
            lnk = tk.Label(
                footer, text=text,
                fg="#888", cursor="hand2",
                font=("Segoe UI", 9, "underline"), anchor="w",
            )
            lnk.pack(anchor="w", pady=pady)
            lnk.bind("<Button-1>", lambda e, u=url: self._open_url(u))

        _link(
            "Aceda ao guida do utilizador desta aplicação",
            "https://github.com/deedeoliveira/WBSFillingFromIFC_bSPT/blob/main/docs/user-guide.md",
        )
        _link(
            "Aceda ao código-fonte no GitHub",
            "https://github.com/deedeoliveira/WBSFillingFromIFC_bSPT",
        )
        _link(
            "Desenvolvido por Andressa Oliveira",
            "https://www.linkedin.com/in/andoliveira/",
            pady=(4, 0),
        )

    def _build_question(self, parent, text, yes_label, yes_cmd,
                        no_label, no_cmd, hidden=False):
        frame = tk.Frame(parent, pady=6)
        frame.pack(fill="x")

        tk.Label(frame, text=text, font=("Segoe UI", 10, "bold"),
                 anchor="w").pack(anchor="w")

        btn_row = tk.Frame(frame)
        btn_row.pack(anchor="w", pady=(4, 0))

        tk.Button(btn_row, text=yes_label, width=38,
                  command=yes_cmd).pack(side="left", padx=(0, 8))

        if no_label and no_cmd:
            tk.Button(btn_row, text=no_label, width=38,
                      command=no_cmd).pack(side="left")

        if hidden:
            frame.pack_forget()

        return frame

    def _show_q2(self):
        self._q2_frame.pack(fill="x")

    def _show_q3(self):
        self._q3_frame.pack(fill="x")

    def _open_url(self, url: str):
        import webbrowser
        try:
            webbrowser.open(url)
        except Exception:
            pass
