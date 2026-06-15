# Contribuir — WBSFillingFromIFC_bSPT

## Idioma / Language
- 🇧🇷 [Português](#português)
- 🇬🇧 [English](#english)

---

# Português

## Pré-requisitos

- Python 3.9+
- `pip install -r requirements.txt`
- `pip install pytest`

---

## Correr localmente

```bash
git clone https://github.com/deedeoliveira/WBSFillingFromIFC_bSPT
cd WBSFillingFromIFC_bSPT
pip install -r requirements.txt
python start.py
```

---

## Testes

```bash
pytest tests/ -v
```

Os testes cobrem `wbs_helpers.py` e `structural_engine.py` e não requerem display nem ficheiros IFC — o tkinter é mockado em `tests/conftest.py`.

---

## Estrutura do código

Antes de modificar, lê [`docs/architecture.md`](architecture.md). Em particular:

- Para alterar a lógica de leitura do WBS Excel → `wbs_helpers.py`
- Para alterar a lógica de filtragem ou quantificação IFC → `structural_engine.py`
- Para alterar o formato JSON de mapeamento → `structural_engine.py` (migração) + `qty.py` (UI) + `app.py` (`run_generate_report`)
- Para alterar a UI de uma aba → o ficheiro `views/` correspondente

---

## Convenções

- Português na UI (labels, mensagens, botões)
- Inglês no código (nomes de variáveis, funções, comentários)
- Prints de debug com prefixo `[DEBUG]` ou `[WARN]` — remover antes de fazer commit para produção
- Não fazer hard-code de nomes de colunas — usar sempre `find_wbs_columns()` e as chaves `col_*`

---

## Sugestões e comentários

Este repositório não aceita Pull Requests. Se encontrares um problema ou tiveres uma sugestão, abre uma [Issue](https://github.com/deedeoliveira/WBSFillingFromIFC_bSPT/issues) com descrição do que observaste.

---

## Gerar executável

```bash
pip install pyinstaller
pyinstaller build.spec
```

O executável gerado fica em `dist/`.

---

## Versões

O projecto segue [Semantic Versioning](https://semver.org/):
- `MAJOR` — alterações incompatíveis com versões anteriores
- `MINOR` — nova funcionalidade retrocompatível
- `PATCH` — correcções de bugs

Ao lançar uma nova versão:
1. Actualizar `__version__` em `app/__init__.py`
2. Actualizar `CHANGELOG.md`
3. Fazer commit e tag `vX.Y.Z`

---

# English

## Prerequisites

- Python 3.9+
- `pip install -r requirements.txt`
- `pip install pytest`

---

## Run locally

```bash
git clone https://github.com/deedeoliveira/WBSFillingFromIFC_bSPT
cd WBSFillingFromIFC_bSPT
pip install -r requirements.txt
python start.py
```

---

## Tests

```bash
pytest tests/ -v
```

Tests cover `wbs_helpers.py` and `structural_engine.py` and require no display or IFC files — tkinter is mocked in `tests/conftest.py`.

---

## Code structure

Read [`docs/architecture.md`](architecture.md) before modifying. Key pointers:

- WBS Excel reading logic → `wbs_helpers.py`
- IFC filtering and quantity logic → `structural_engine.py`
- Mapping JSON format → `structural_engine.py` (migration) + `qty.py` (UI) + `app.py` (`run_generate_report`)
- Tab UI → corresponding file in `views/`

---

## Conventions

- UI in Portuguese (labels, messages, buttons)
- Code in English (variable names, functions, comments)
- Debug prints prefixed `[DEBUG]` or `[WARN]` — remove before production commit
- Never hard-code column names — always use `find_wbs_columns()` and `col_*` keys

---

## Suggestions and feedback

This repository does not accept Pull Requests. If you find a bug or have a suggestion, open an [Issue](https://github.com/deedeoliveira/WBSFillingFromIFC_bSPT/issues) with a description of what you observed.

---

## Build executable

```bash
pip install pyinstaller
pyinstaller build.spec
```

Output goes to `dist/`.

---

## Versioning

The project follows [Semantic Versioning](https://semver.org/):
- `MAJOR` — breaking changes
- `MINOR` — new backward-compatible functionality
- `PATCH` — bug fixes

When releasing a new version:
1. Update `__version__` in `app/__init__.py`
2. Update `CHANGELOG.md`
3. Commit and tag `vX.Y.Z`
