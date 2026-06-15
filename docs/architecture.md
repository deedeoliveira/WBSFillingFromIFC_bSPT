# Arquitectura — WBSFillingFromIFC_bSPT

## Idioma / Language
- 🇧🇷 [Português](#português)
- 🇬🇧 [English](#english)

---

# Português

## Visão geral

A aplicação é uma app desktop Windows construída em Python com tkinter. Segue uma arquitectura de pipeline em três etapas, cada uma correspondendo a uma aba da interface:

```
WBS original (Excel)
        │
        ▼
┌─────────────────┐
│  WBSPage        │  → WBS com descrições (Excel) + mapeamento parcial (JSON)
└─────────────────┘
        │
        ▼
┌─────────────────┐
│  QtyPage        │  → mapeamento completo (JSON)
└─────────────────┘
        │
        ▼
┌─────────────────┐
│  ReportPage     │  → MapaQuantidadesTrabalhos (Excel)
└─────────────────┘     ElementosVerificados (Excel)
                        ElementosQuantificados (CSV)
```

---

## Estrutura de ficheiros

```
app/
├── __init__.py                  # versão, metadados
├── core/
│   └── structural_engine.py     # lógica IFC: filtragem, quantificação
└── gui/
    ├── app.py                   # WBSApp (tk.Tk) — orquestra tudo
    ├── wbs_helpers.py           # utilitários de leitura e parsing do WBS Excel
    └── views/
        ├── home.py              # aba Home — navegação e links
        ├── wbs_editor.py        # aba WBS e descrição
        ├── qty.py               # aba Mapeamento IFC
        └── report.py            # aba Extrair quantidades
tests/
├── conftest.py                  # mock tkinter para CI sem display
├── test_wbs_helpers.py
└── test_structural_engine.py
```

---

## Formato JSON de mapeamento (v2)

O ficheiro JSON é o artefacto central que liga a WBSPage à ReportPage.

```json
{
  "version": 2,
  "partial": false,
  "ifc_path": null,
  "wbs_path": null,
  "rules": {
    "08.01.01.03": {
      "mappings": [
        {
          "filter": {
            "ifc_class": "IfcColumn",
            "predefined": "COLUMN",
            "object_type": "",
            "props": [
              { "pset": "Pset_ColumnCommon", "prop": "LoadBearing", "value": true }
            ]
          },
          "quantity_detail": {
            "pset": "Qto_ColumnBaseQuantities",
            "prop": "NetVolume"
          }
        }
      ],
      "material": "Concrete",
      "quantity": { "type": "prop" },
      "agrupamento": {
        "pset": "PTBS_Identification",
        "prop": "WbsGrouping"
      }
    }
  }
}
```

**Campos:**
- `mappings` — lista de classes IFC a pesquisar para este item WBS; cada entrada tem `filter` e `quantity_detail`
- `filter.props` — filtros adicionais por propriedade IFC (com valor esperado)
- `quantity.type` — `"prop"` (ler valor de propriedade) ou `"count"` (contar elementos)
- `quantity_detail` — pset e propriedade onde está o valor de quantidade (só para `type: "prop"`)
- `agrupamento` — propriedade usada para subdividir o resultado (ex: dimensão do elemento)
- `partial: true` — mapeamento incompleto, gerado automaticamente pelo WBSPage

**Retrocompatibilidade:** mapeamentos v1 (com `filter` e `quantity` no nível raiz) são migrados automaticamente para v2 na leitura.

---

## WBSPage — `wbs_editor.py`

**Input esperado:** WBS Excel da buildingSMART Portugal.

O ficheiro tem um cabeçalho nas primeiras linhas (posição variável). A app detecta a linha de cabeçalho procurando a primeira linha que contenha simultaneamente `WBS`, `nivel` e `descricao` (normalizado, sem acentos).

**Colunas lidas:**

| Coluna | Chave interna | Obrigatória |
|--------|--------------|-------------|
| `NÍVEL` | `col_nivel` | ✓ |
| `WBS` | `col_wbs` | ✓ |
| `DESCRIÇÃO` | `col_desc` | ✓ |
| `UNIDADES` | `col_unidades` | — (output incompleto sem ela) |
| `IFC Class` | `col_ifc_class` | — |
| `PredefinedType` | `col_predef` | — |
| `ObjectType` | `col_objtype` | — |
| `IFC Property` | `col_ifc_prop` | — |

A detecção é feita por normalização do nome da coluna (sem acentos, lowercase, sem espaços duplos) — ver `find_wbs_columns()` em `wbs_helpers.py`.

**Nível 10:** cada item folha pode ter uma linha de nível 10 imediatamente abaixo com a descrição do utilizador. Essa linha não tem código WBS.

**Parsing das colunas IFC** (para mapeamento parcial):
- `IFC Class` suporta dot notation (`IfcFooting.PILE_CAP`), múltiplas classes (`IfcDoor / IfcWindow`) e múltiplos predefined types (`IfcCovering + CEILING / FLOORING`)
- `IFC Property` suporta `GRUPO.PROPRIEDADE` ou só `PROPRIEDADE` (pset fica vazio)
- Linha sem `IFC Class` é ignorada no mapeamento parcial

---

## QtyPage — `qty.py`

**Input esperado:** WBS com descrições (Excel) + opcionalmente mapeamento JSON.

### Modos de mapeamento

A aba tem dois modos, escolhidos pelo utilizador antes de iniciar — a escolha fica bloqueada até recarregar:

**Modo projecto** — requer upload de IFC. Os dropdowns de `IfcClass` e `PredefinedType` são populados automaticamente com os valores presentes no modelo. Garante que o mapeamento é válido para o IFC carregado.

**Modo genérico** — sem IFC. Todos os campos são Entry de texto livre. Útil para criar mapeamentos reutilizáveis entre projectos, sem estar dependente de um modelo específico.

### Modos de quantificação

Por regra (por item WBS), o utilizador escolhe um de dois modos:

**Leitura de propriedade (`type: "prop"`)** — lê um valor numérico de uma propriedade IFC específica (pset + prop). O pset e prop podem variar por classe IFC dentro da mesma regra.

**Contagem de elementos (`type: "count"`)** — conta o número de elementos IFC que passam nos filtros. Não requer pset/prop de quantidade.

### Estrutura de uma regra

Cada item folha pode ter uma regra guardada em `app.rules[wbs_code]`. Uma regra pode ter múltiplos `FilterBlock` (multi-classe) — cada bloco representa uma classe IFC com os seus próprios filtros e, em modo `prop`, o seu próprio pset/prop de quantidade. O resultado final é a soma ou contagem de todos os blocos.

O mapeamento pode ser carregado de um JSON existente (parcial ou completo) e editado. Regras parciais (sem `ifc_class` preenchido) são aceites e ignoradas na extração.

**Validação na exportação:**
- `ifc_class` e `predefined` obrigatórios por FilterBlock (regras sem `ifc_class` são ignoradas)
- `quantity_detail.pset` e `prop` obrigatórios quando `type: "prop"`
- `quantity.type` tem de ser `"prop"` ou `"count"`

---

## structural_engine.py — `IFCInvestigator`

Classe principal de interface com o IFC via `ifcopenshell`.

**Indexação:** ao abrir um IFC, todos os `IfcProduct` são indexados por classe em `index_by_class`. Isto evita varrer o modelo inteiro a cada filtragem.

**Pipeline de extração por regra:**

```
extract_quantities(rule)
    │
    ├── para cada mapping entry:
    │       filter_elements_for_mapping()
    │           ├── filtra por ifc_class (index_by_class)
    │           ├── filtra por PredefinedType
    │           ├── filtra por ObjectType (se USERDEFINED)
    │           └── filtra por props extras (_match_props)
    │
    ├── aplica filtro de material
    │
    └── se type="count" → count_elements()
        se type="prop"  → sum_quantity(pset, prop)
```

**Filtro booleano:** `_match_props` suporta todas as representações de booleano IFC: `True`, `"TRUE"`, `".T."`, `"T"`, `"1"`, `"YES"` (e equivalentes false). Tenta também ler propriedades de tipo (`IfcRelDefinesByType`) além das de instância.

---

## ReportPage — `app.py` (`run_generate_report`)

Corre numa thread separada para não bloquear a GUI.

**Sequência:**
1. Valida e migra regras (v1 → v2)
2. Carrega WBS e IFC
3. Para cada código WBS em `rules`: filtra elementos → extrai quantidade → recolhe agrupamento
4. Constrói `_build_rows(include_not_found=True/False)` para os dois Excel
5. Guarda `_last_csv_cache` em memória para exportação CSV posterior

**`_last_csv_cache`:** dict com `headers`, `wbs_rows`, `per_code`, `code_to_unit`. O CSV é gerado exclusivamente a partir deste cache — não lê ficheiros em disco.

---

## Ordem de leitura sugerida

Para **utilizadores** que querem perceber como usar: [`docs/user-guide.md`](user-guide.md)

Para quem quer perceber como o código funciona ou sugerir melhorias: este ficheiro (`architecture.md`).

Sugestões e comentários são bem-vindos via [Issues](https://github.com/deedeoliveira/WBSFillingFromIFC_bSPT/issues).

---

# English

## Overview

Desktop Windows app built in Python with tkinter. Three-stage pipeline, one tab per stage.

## File structure

*(same as Portuguese section above)*

## v2 Mapping JSON format

*(see Portuguese section — structure is language-independent)*

## WBSPage

Reads the buildingSMART Portugal WBS Excel. Header row is auto-detected by looking for a row containing `WBS`, `nivel` and `descricao` (normalised). Columns are matched by normalised name — see `find_wbs_columns()` in `wbs_helpers.py`.

Level-10 rows hold user descriptions (no WBS code, immediately below the leaf item).

IFC columns (`IFC Class`, `PredefinedType`, `ObjectType`, `IFC Property`) are optional and used to auto-generate a partial mapping JSON.

## QtyPage

The tab has two modes, chosen once before starting (locked until reload):

**Project mode** — requires IFC upload. `IfcClass` and `PredefinedType` dropdowns are auto-populated from the loaded model.

**Generic mode** — no IFC needed. All fields are free-text Entry widgets. Useful for reusable mappings across projects.

Each WBS leaf item can have a mapping rule stored in `app.rules[wbs_code]`. A rule supports multiple `FilterBlock` entries (multi-class) — each block has its own filter and, in `prop` mode, its own pset/prop for quantity. Results are summed across all blocks.

Two quantity modes per rule:
- **Property reading (`type: "prop"`)** — reads a numeric value from a specified pset/prop (may vary per class)
- **Element count (`type: "count"`)** — counts matching elements, no pset/prop needed

Partial mappings (empty `ifc_class`) are accepted and silently skipped during extraction.

## structural_engine.py

`IFCInvestigator` wraps ifcopenshell. On open, all `IfcProduct` elements are indexed by class. Filtering applies: class → PredefinedType → ObjectType → extra props → material. Boolean IFC properties are handled in all representations (`.T.`, `TRUE`, `True`, etc.).

## ReportPage

`run_generate_report` runs in a background thread. Builds two Excel files and populates `_last_csv_cache`. CSV export reads from cache only — does not touch disk files.

---

## Suggested reading order

**Users** who want to understand how to use the app: [`docs/user-guide.md`](user-guide.md)

For anyone who wants to understand how the code works or suggest improvements: this file (`architecture.md`).

Suggestions and comments are welcome via [Issues](https://github.com/deedeoliveira/WBSFillingFromIFC_bSPT/issues).
