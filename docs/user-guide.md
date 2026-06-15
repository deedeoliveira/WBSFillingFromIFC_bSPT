# Guia do Utilizador — WBSFillingFromIFC_bSPT

## Idioma / Language
- 🇧🇷 [Português](#português)
- 🇬🇧 [English](#english)

---

# Português

## Requisitos

- Windows 10 ou superior
- Ficheiro WBS bSPT (`.xlsx`) — disponível em [buildingsmart.pt/recursos](https://buildingsmart.pt/recursos/)
- Modelo IFC (`.ifc`) verificado
- Mapeamento JSON (`.json`) — gerado pela app ou importado

---

## Arrancar a aplicação

**Executável:** clique duas vezes em `WBSFillingFromIFC_bSPT.exe`.

**A partir do código fonte:**
```bash
python start.py
```

---

## Aba Home

Responde às perguntas para seres dirigido à aba correcta. Se já tiveres algum dos ficheiros intermédios prontos (WBS com descrições, mapeamento), podes saltar etapas.

---

## Aba WBS e descrição

### Modo "Novo WBS com descrições"

Use quando começa do zero.

1. Confirma o modo e carrega o WBS original bSPT
2. Selecciona uma secção (Nível 1) no selector
3. Navega pela hierarquia com os botões **Nível abaixo** / **Nível acima**
4. Ao chegar ao último nível, clica **Adicionar descrição** e escreve a descrição do item
5. Clica **Guardar descrição**
6. Repete para todos os itens que precisam de descrição
7. Clica **Salvar etapa e exportar WBS** — gera o Excel com descrições e, se o WBS original tiver colunas IFC, também um JSON de mapeamento parcial

### Modo "Continuar WBS existente"

Use quando já tens um WBS com descrições parcialmente preenchido de uma sessão anterior.

1. Confirma o modo
2. Carrega o WBS original bSPT
3. Carrega o WBS com descrições parciais (ficheiro exportado anteriormente)
4. Os itens já preenchidos aparecem marcados com ✓ e bloqueados — clica **Editar descrição** para alterar
5. Adiciona as descrições em falta e exporta

---

## Aba Mapeamento IFC

### Antes de começar — escolhe o modo

Podes começar um mapeamento do zero ou carregar um mapeamento existente (parcial ou completo) para editar. Um mapeamento parcial é gerado automaticamente pela aba anterior se o WBS original tiver colunas IFC preenchidas — carrega-o aqui para continuar a partir daí.

**Modo projecto:** carrega um IFC específico. Os dropdowns de classe e predefined type são populados automaticamente com os valores do modelo. Usa quando o mapeamento é para um projecto concreto.

**Modo genérico:** sem IFC. Escreves as classes manualmente. Usa quando queres criar um mapeamento reutilizável entre projectos.

Confirma o modo — não pode ser alterado depois.

### Modo projecto

1. Confirma o modo e carrega o IFC
2. Carrega o WBS com descrições (ou vem automaticamente da aba anterior)
3. Carrega um mapeamento existente se tiveres (opcional)
4. Selecciona um item folha na lista (marcados com ✓ se já têm mapeamento)
5. Para cada item:
   - Escolhe o modo de quantificação: **Leitura de propriedade** ou **Contagem de elementos**
   - Em **+ Adicionar classe IFC**: selecciona a classe, predefined type e define os filtros
   - Se **Leitura de propriedade**: preenche o pset e propriedade de quantidade para cada classe
   - Define o material (opcional) e a propriedade de agrupamento (opcional)
   - Clica **Guardar regra**
6. Quando todos os itens estiverem mapeados, clica **Salvar e exportar** → gera o JSON de mapeamento

### Modo genérico

Igual ao modo projecto, mas todos os campos são de texto livre — escreve as classes e propriedades manualmente.

---

## Aba Extrair quantidades e Gerar WBS preenchido

1. Carrega o WBS com descrições, o mapeamento JSON e o IFC
   - Se vieres da aba anterior, os ficheiros são pré-carregados automaticamente
2. Define a pasta de saída
3. Clica **Gerar WBS preenchido**
4. Aguarda — o progresso aparece no log
5. Após concluir, clica **Exportar CSV detalhado** se precisares do ficheiro para Power BI

### Ficheiros gerados

| Ficheiro | Conteúdo |
|----------|----------|
| `MapaQuantidadesTrabalhos_[IFC].xlsx` | Mapa de Quantidades de Trabalho preenchido: hierarquia WBS com descrições, quantidades e unidades por item. Inclui sub-linhas de agrupamento quando definido (ex: dimensões por elemento). Contém apenas itens cujos elementos foram encontrados no modelo. |
| `ElementosVerificados_[IFC].xlsx` | Versão completa do MQT: inclui todos os itens do mapeamento, mesmo os não encontrados. Itens sem elementos no modelo aparecem com a indicação `[ELEMENTOS NÃO ENCONTRADOS]` em fundo amarelo. Útil para verificar a cobertura do mapeamento. |
| `ElementosQuantificados_[IFC].csv` | Um registo por elemento IFC encontrado: código WBS, GUID, classe IFC, piso, material, valor de quantidade e unidade. Estruturado para ligação directa a dashboards (ex: Power BI). |

---

## Perguntas frequentes

**A app não encontra as colunas do WBS.**
O ficheiro Excel tem de ser o WBS bSPT original ou ter as colunas `NÍVEL`, `WBS` e `DESCRIÇÃO`. A app detecta-as automaticamente por nome (sem sensibilidade a acentos ou maiúsculas).

**Posso usar um mapeamento criado para outro projecto?**
Sim. Carrega o JSON na aba Mapeamento IFC. Se os nomes de classes e propriedades IFC forem consistentes entre projectos, o mapeamento funciona sem alterações.

**O `MapaQuantidadesTrabalhos` está vazio / todos os itens estão no `ElementosVerificados` como não encontrados.**
Verifica se o IFC tem as propriedades que o mapeamento exige. Nomes de propriedades são sensíveis a capitalização (`LoadBearing` ≠ `Loadbearing`). Confirma também se o predefined type no IFC corresponde exactamente ao definido no mapeamento.

---

# English

## Requirements

- Windows 10 or later
- bSPT WBS file (`.xlsx`) — available at [buildingsmart.pt/recursos](https://buildingsmart.pt/recursos/)
- Verified IFC model (`.ifc`)
- Mapping JSON (`.json`) — generated by the app or imported

---

## Starting the application

**Executable:** double-click `WBSFillingFromIFC_bSPT.exe`.

**From source:**
```bash
python start.py
```

---

## Home tab

Answer the questions to be directed to the correct tab. If you already have intermediate files ready (WBS with descriptions, mapping), you can skip steps.

---

## WBS and descriptions tab

### Mode "New WBS with descriptions"

Use when starting from scratch.

1. Confirm the mode and load the original bSPT WBS
2. Select a section (Level 1) in the selector
3. Navigate the hierarchy with **Level down** / **Level up** buttons
4. At the last level, click **Add description** and write the item description
5. Click **Save description**
6. Repeat for all items that need a description
7. Click **Save step and export WBS** — generates the Excel with descriptions and, if the original WBS has IFC columns, also a partial mapping JSON

### Mode "Continue existing WBS"

Use when you already have a partially filled WBS from a previous session.

1. Confirm the mode
2. Load the original bSPT WBS
3. Load the partial WBS with descriptions (previously exported file)
4. Already filled items are marked ✓ and locked — click **Edit description** to change them
5. Add missing descriptions and export

---

## IFC Mapping tab

### Before starting — choose a mode

You can start a mapping from scratch or load an existing mapping (partial or complete) to edit. A partial mapping is automatically generated by the previous tab if the original WBS has IFC columns filled in — load it here to continue from where you left off.

**Project mode:** load a specific IFC. Class and predefined type dropdowns are auto-populated from the model. Use when mapping for a specific project.

**Generic mode:** no IFC needed. Write classes manually. Use when creating a reusable mapping across projects.

Confirm the mode — it cannot be changed afterwards.

### Project mode

1. Confirm the mode and load the IFC
2. Load the WBS with descriptions (or it comes automatically from the previous tab)
3. Load an existing mapping if you have one (optional)
4. Select a leaf item in the list (marked ✓ if already mapped)
5. For each item:
   - Choose quantity mode: **Property reading** or **Element count**
   - In **+ Add IFC class**: select class, predefined type and define filters
   - If **Property reading**: fill in the pset and quantity property for each class
   - Set material (optional) and grouping property (optional)
   - Click **Save rule**
6. When all items are mapped, click **Save and export** → generates the mapping JSON

### Generic mode

Same as project mode, but all fields are free text — write classes and properties manually.

---

## Extract quantities tab

1. Load the WBS with descriptions, the mapping JSON and the IFC
   - If coming from the previous tab, files are pre-loaded automatically
2. Set the output folder
3. Click **Generate filled WBS**
4. Wait — progress appears in the log
5. After completion, click **Export detailed CSV** if you need the Power BI file

### Generated files

| File | Content |
|------|---------|
| `MapaQuantidadesTrabalhos_[IFC].xlsx` | Filled Bill of Quantities: WBS hierarchy with descriptions, quantities and units per item. Includes grouping sub-rows when defined (e.g. dimensions per element). Contains only items whose elements were found in the model. |
| `ElementosVerificados_[IFC].xlsx` | Complete BoQ version: includes all mapped items, even those not found. Items with no matching elements are flagged `[ELEMENTOS NÃO ENCONTRADOS]` with a yellow background. Useful for verifying mapping coverage. |
| `ElementosQuantificados_[IFC].csv` | One row per IFC element found: WBS code, GUID, IFC class, floor, material, quantity value and unit. Structured for direct connection to dashboards (e.g. Power BI). |

---

## FAQ

**The app does not find the WBS columns.**
The Excel file must be the original bSPT WBS or have columns named `NÍVEL`, `WBS` and `DESCRIÇÃO`. The app detects them automatically by name (accent and case insensitive).

**Can I use a mapping created for another project?**
Yes. Load the JSON in the IFC Mapping tab. If IFC class and property names are consistent across projects, the mapping works without changes.

**`MapaQuantidadesTrabalhos` is empty / all items appear in `ElementosVerificados` as not found.**
Check that the IFC has the properties the mapping requires. Property names are case-sensitive (`LoadBearing` ≠ `Loadbearing`). Also confirm the predefined type in the IFC matches exactly what is defined in the mapping.
