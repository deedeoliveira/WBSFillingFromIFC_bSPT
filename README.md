# WBS Filling From IFC

![Version](https://img.shields.io/badge/version-0.2.0-blue)
![Python](https://img.shields.io/badge/python-3.9+-blue)
![Platform](https://img.shields.io/badge/platform-Windows-lightgrey)
![License](https://img.shields.io/badge/license-MIT-green)
![Tests](https://img.shields.io/badge/tests-41%20passed-brightgreen)
[![DOI](https://zenodo.org/badge/DOI/10.5281/zenodo.18489649.svg)](https://doi.org/10.5281/zenodo.18489649)

---

## Idioma / Language

- 🇧🇷 [Português](#português)
- 🇬🇧 [English](#english)

---

# Português

## Sobre

A extração de quantidades a partir de modelos BIM em IFC é frequentemente manual e pouco padronizada. Esta aplicação automatiza esse processo com base num mapeamento explícito entre itens de uma Work Breakdown Structure (WBS) e elementos IFC, gerando automaticamente um Mapa de Quantidades de Trabalho (MQT).

A WBS de referência utilizada é a desenvolvida pela buildingSMART Portugal (bSPT) para o contexto português, mas a aplicação pode ser adaptada a outras estruturas.

**Para quem:** profissionais da construção civil que trabalham com modelos BIM e precisam de extrair e organizar quantidades a partir de ficheiros IFC.

---

## Instalação

1. Descarregue o executável na [página de recursos da buildingSMART Portugal](https://buildingsmart.pt/recursos/)
2. Execute `WBSFillingFromIFC_bSPT.exe` — não requer instalação

**A partir do código fonte:**

```bash
git clone https://github.com/deedeoliveira/WBSFillingFromIFC_bSPT
cd WBSFillingFromIFC_bSPT
pip install -r requirements.txt
python start.py
```

---

## Ficheiros necessários

| Ficheiro | Formato | Descrição |
|----------|---------|-----------|
| WBS original | `.xlsx` | WBS da buildingSMART Portugal |
| Modelo IFC | `.ifc` | Modelo (IFC 2x3 ou IFC 4) |
| Mapeamento | `.json` | Gerado pela app ou importado |

---

## Fluxos de trabalho

Os três fluxos podem ser executados em sequência ou de forma independente.

### Fluxo 1 — Processo completo
1. **WBS e descrição** — carregue o WBS e adicione descrições customizadas
2. **Mapeamento IFC** — associe itens WBS a classes IFC e defina como extrair quantidades
3. **Extrair quantidades** — carregue o IFC e gere os ficheiros de output

### Fluxo 2 — Mapeamento + Extração
Use quando o WBS com descrições já está pronto.
1. **Mapeamento IFC** → **Extrair quantidades**

### Fluxo 3 — Apenas extração
Use quando WBS e mapeamento já estão prontos.
1. **Extrair quantidades** — carregue os três ficheiros e gere o output

---

## Ficheiros de output

| Ficheiro | Conteúdo |
|----------|----------|
| `MapaQuantidadesTrabalhos_[IFC].xlsx` | MQT preenchido com quantidades — apenas itens com elementos encontrados no modelo |
| `ElementosVerificados_[IFC].xlsx` | Todos os itens mapeados, incluindo os não encontrados (assinalados) |
| `ElementosQuantificados_[IFC].csv` | Detalhe por elemento para dashboards |

---

## Testes

```bash
pip install pytest
pytest tests/ -v
```

---

## Estrutura do projecto

```
WBSFillingFromIFC_bSPT/
├── app/
│   ├── core/
│   │   └── structural_engine.py
│   └── gui/
│       ├── app.py
│       ├── wbs_helpers.py
│       └── views/
│           ├── home.py
│           ├── wbs_editor.py
│           ├── qty.py
│           └── report.py
├── tests/
├── docs/
├── start.py
└── requirements.txt
```

---

## Citação

Oliveira, A. (2026). *WBSFillingFromIFC_bSPT* (Version 0.2.0). Zenodo.
https://doi.org/10.5281/zenodo.18489649

```bibtex
@software{oliveira_2026_wbsfillingfromifc_bspt,
  author    = {Oliveira, Andressa},
  title     = {WBSFillingFromIFC\_bSPT},
  version   = {0.2.0},
  year      = {2026},
  publisher = {Zenodo},
  doi       = {10.5281/zenodo.18489649}
}
```

---

## Autora

| Função | Nome | Contacto |
|--------|------|----------|
| Desenvolvimento | Andressa Oliveira | [LinkedIn](https://www.linkedin.com/in/andoliveira/) · [Email](mailto:soliveira.andressa@gmail.com) |

---

## Contribuições

Contribuições são bem-vindas. Consulte [`docs/contributing.md`](docs/contributing.md) para instruções.

---

*Versão 0.2.0 — Junho 2026*

---

# English

## About

Extracting quantities from BIM/IFC models is often a manual, poorly standardised process. This application automates it through an explicit mapping between Work Breakdown Structure (WBS) items and IFC elements, automatically generating a Bill of Quantities (BoQ / MQT).

The reference WBS used is the one developed by buildingSMART Portugal (bSPT) for the Portuguese context, but the application can be adapted to other structures.

**For:** construction professionals working with BIM models who need to extract and organise quantities from IFC files.

---

## Installation

1. Download the executable from the [buildingSMART Portugal resources page](https://buildingsmart.pt/recursos/)
2. Run `WBSFillingFromIFC_bSPT.exe` — no installation required

**From source:**

```bash
git clone https://github.com/deedeoliveira/WBSFillingFromIFC_bSPT
cd WBSFillingFromIFC_bSPT
pip install -r requirements.txt
python start.py
```

---

## Required files

| File | Format | Description |
|------|--------|-------------|
| Original WBS | `.xlsx` | WBS from buildingSMART Portugal |
| IFC model | `.ifc` | BIM model (IFC 2x3 or IFC 4) |
| Mapping | `.json` | Generated by the app or imported |

---

## Workflows

The three workflows can be run in sequence or independently.

### Workflow 1 — Full process
1. **WBS and descriptions** — load the WBS and add custom descriptions
2. **IFC mapping** — associate WBS items with IFC classes and define quantity extraction
3. **Extract quantities** — load the IFC and generate output files

### Workflow 2 — Mapping + Extraction
Use when the WBS with descriptions is already ready.
1. **IFC mapping** → **Extract quantities**

### Workflow 3 — Extraction only
Use when WBS and mapping are already prepared.
1. **Extract quantities** — load all three files and generate output

---

## Output files

| File | Content |
|------|---------|
| `MapaQuantidadesTrabalhos_[IFC].xlsx` | Filled BoQ with quantities — only items with elements found in the model |
| `ElementosVerificados_[IFC].xlsx` | All mapped items, including not-found ones (flagged) |
| `ElementosQuantificados_[IFC].csv` | Per-element detail for dashboards |

---

## Tests

```bash
pip install pytest
pytest tests/ -v
```

---

## Project structure

```
WBSFillingFromIFC_bSPT/
├── app/
│   ├── core/
│   │   └── structural_engine.py
│   └── gui/
│       ├── app.py
│       ├── wbs_helpers.py
│       └── views/
│           ├── home.py
│           ├── wbs_editor.py
│           ├── qty.py
│           └── report.py
├── tests/
├── docs/
├── start.py
└── requirements.txt
```

---

## Citation

Oliveira, A. (2026). *WBSFillingFromIFC_bSPT* (Version 0.2.0). Zenodo.
https://doi.org/10.5281/zenodo.18489649

```bibtex
@software{oliveira_2026_wbsfillingfromifc_bspt,
  author    = {Oliveira, Andressa},
  title     = {WBSFillingFromIFC\_bSPT},
  version   = {0.2.0},
  year      = {2026},
  publisher = {Zenodo},
  doi       = {10.5281/zenodo.18489649}
}
```

---

## Author

| Role | Name | Contact |
|------|------|---------|
| Development | Andressa Oliveira | [LinkedIn](https://www.linkedin.com/in/andoliveira/) · [Email](mailto:soliveira.andressa@gmail.com) |

---

## Contributions

Contributions are welcome. See [`docs/contributing.md`](docs/contributing.md) for instructions.

---

*Version 0.2.0 — June 2026*
