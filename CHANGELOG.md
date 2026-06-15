# Changelog

---

## Idioma / Language

- 🇧🇷 [Português](#português)
- 🇬🇧 [English](#english)

---

# Português

Todas as mudanças notáveis serão documentadas aqui.
Formato baseado em [Keep a Changelog](https://keepachangelog.com/en/1.0.0/).

---

## [0.2.0] - 2026-06-14

### ✨ Novidades

- Aba WBS: opção de continuar um WBS já parcialmente preenchido
- Aba WBS: geração automática de mapeamento parcial a partir das colunas IFC do WBS original
- Aba Mapeamento: upload de mapeamento existente (parcial ou completo)
- Aba Mapeamento: suporte a múltiplas classes IFC por item WBS
- Aba Mapeamento: modo de quantificação por contagem de elementos
- Aba Mapeamento: modo genérico sem upload de IFC (campos livres)
- Aba Extração: dois ficheiros Excel de output (`MapaQuantidadesTrabalhos` e `ElementosVerificados`)
- Aba Extração: itens sem elementos encontrados assinalados em `ElementosVerificados`
- Testes unitários com pytest (41 testes)

### 🔧 Interno

- Novo formato JSON de mapeamento v2 com suporte a multi-classe e tipo de quantificação
- Migração automática de mapeamentos v1 na leitura
- Novo dispatcher `extract_quantities()` e função `count_elements()` no `structural_engine`

---

## [0.1.1] - 2025-12-01

- Correcções menores na detecção de colunas WBS e no carregamento do IFC

---

## [0.1.0] - 2025-10-22

- Versão beta inicial — 4 abas, 3 fluxos, extração de quantidades IFC, exportação Excel e CSV

---

[0.2.0]: https://github.com/deedeoliveira/WBSFillingFromIFC_bSPT/compare/v0.1.1...v0.2.0
[0.1.1]: https://github.com/deedeoliveira/WBSFillingFromIFC_bSPT/compare/v0.1.0...v0.1.1
[0.1.0]: https://github.com/deedeoliveira/WBSFillingFromIFC_bSPT/releases/tag/v0.1.0

---

# English

All notable changes will be documented here.
Format based on [Keep a Changelog](https://keepachangelog.com/en/1.0.0/).

---

## [0.2.0] - 2026-06-14

### ✨ New features

- WBS tab: option to continue a partially filled WBS
- WBS tab: automatic partial mapping from IFC columns in the original WBS
- Mapping tab: upload existing mapping (partial or complete)
- Mapping tab: multiple IFC classes per WBS item
- Mapping tab: element count as quantity mode
- Mapping tab: generic mode without IFC upload (free-text fields)
- Extraction tab: two Excel output files (`MapaQuantidadesTrabalhos` and `ElementosVerificados`)
- Extraction tab: not-found items flagged in `ElementosVerificados`
- Unit tests with pytest (41 tests)

### 🔧 Internal

- New v2 mapping JSON format with multi-class and quantity type support
- Automatic v1 mapping migration on load
- New `extract_quantities()` dispatcher and `count_elements()` in `structural_engine`

---

## [0.1.1] - 2025-12-01

- Minor fixes to WBS column detection and IFC load error handling

---

## [0.1.0] - 2025-10-22

- Initial beta release — 4 tabs, 3 workflows, IFC quantity extraction, Excel and CSV export

---

[0.2.0]: https://github.com/deedeoliveira/WBSFillingFromIFC_bSPT/compare/v0.1.1...v0.2.0
[0.1.1]: https://github.com/deedeoliveira/WBSFillingFromIFC_bSPT/compare/v0.1.0...v0.1.1
[0.1.0]: https://github.com/deedeoliveira/WBSFillingFromIFC_bSPT/releases/tag/v0.1.0
