# WBS Filling From IFC

![Version](https://img.shields.io/badge/version-0.1.1--beta-orange)
![Python](https://img.shields.io/badge/python-3.9+-blue)
![Platform](https://img.shields.io/badge/platform-Windows-lightgrey)
![License](https://img.shields.io/badge/license-MIT-green)

**Versão atual:** 0.1.1 (Beta)

---

## Sobre

Esta aplicação permite automatizar o processo de extração de quantidades de modelos de informação em IFC e preencher automaticamente a estrutura de Work Breakdown Structure (WBS) desenvolvida pela buildingSMART Portugal.

**Aviso:** Esta é uma versão beta.

---

## Instalação para Usuários

1. Baixe o executável da [página de recursos da buildingSMART Portugal](https://buildingsmart.pt/recursos/)
2. Execute `WBSFillingFromIFC_bSPT_v0_1_1.exe`

---

## Uso

### Arquivos Necessários

| Arquivo | Formato | Descrição |
|---------|---------|-----------|
| WBS Original | `.xlsx` | Estrutura WBS (versão v01) da buildingSMART Portugal |
| Modelo IFC | `.ifc` | Modelo BIM compatível com IFC 2x3 ou 4 |

### Fluxos de Trabalho

#### Fluxo 1: Processo Completo

Use quando começa do zero.

1. **Aba Home**
   - Responder pergunta para ser direcionado a aba "WBS e Descrição"

2. **Aba WBS e Descrição**
   - Carregar WBS original
   - Navegar pela estrutura WBS
   - Adicionar descrições customizadas
   - Salvar WBS editado

3. **Aba Mapeamento IFC**
   - Carregar IFC
   - Selecionar códigos WBS (folhas)
   - Definir filtros IFC (classe, predefined type, propriedades)
   - Definir propriedade de quantidade
   - Definir agrupamento
   - Salvar e exportar mapeamento

4. **Aba Extração**
   - Gerar WBS preenchido
   - Exportar CSV detalhado

#### Fluxo 2: Mapeamento IFC + Extração de informação

Use quando WBS já está customizado, precisa apenas criar mapeamento.

1. **Aba Home**
   - Responder perguntas para ser direcionado a aba "Mapeamento IFC"

2. **Aba Mapeamento IFC**
   - Carregar WBS com descrições customizadas
   - Carregar IFC
   - Selecionar códigos WBS (folhas)
   - Definir filtros IFC (classe, predefined type, propriedades)
   - Definir propriedade de quantidade
   - Definir agrupamento
   - Salvar e exportar mapeamento

3. **Aba Extração**
   - Gerar WBS preenchido
   - Exportar CSV detalhado

#### Fluxo 3: Extração direta de informação

Use quando já tem WBS customizado e mapeamento pronto.

1. **Aba Home**
   - Responder perguntas para ser direcionado a aba "Extrair quantidades e Gerar WBS preenchido"

2. **Aba Extração**
   - Carregar WBS com descrições customizadas
   - Carregar mapeamento JSON
   - Carregar IFC
   - Gerar WBS preenchido
   - Exportar CSV detalhado

---

## Desenvolvimento

### Estrutura do Projeto
```
WBSFillingFromIFC_bSPT/
├── app/
│   ├── __init__.py
│   ├── core/
│   │   └── structural_engine.py
│   └── gui/
│       ├── main.py
│       ├── app.py
│       ├── wbs_helpers.py
│       └── views/
│           ├── home.py
│           ├── wbs_editor.py
│           ├── qty.py
│           └── report.py
├── start.py
├── requirements.txt
└── README.md
```

---

## Autores e Contato

| Papel | Nome | Contato |
|-------|------|---------|
| **Desenvolvedora** | Andressa Oliveira | [LinkedIn](https://www.linkedin.com/in/andoliveira/) • [Email](mailto:soliveira.andressa@gmail.com) |
---

## Contribuições

Contribuições são bem-vindas.

---

*Versão 0.1.1 - Dezembro 2025*