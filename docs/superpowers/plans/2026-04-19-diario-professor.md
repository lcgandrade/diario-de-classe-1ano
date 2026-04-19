# Diário de Classe — Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Gerar dois arquivos `.xlsx` completos — `diario-de-classe.xlsx` e `planejamento-semanal.xlsx` — com formatação profissional, fórmulas e destaques visuais para apoiar o dia a dia de uma professora de 1º ano do ensino fundamental público.

**Architecture:** Dois scripts Python independentes (`criar_diario.py` e `criar_planejamento.py`) que usam `openpyxl` para gerar os arquivos na raiz do projeto. Cada script pode ser re-executado a qualquer momento para regenerar os arquivos.

**Tech Stack:** Python 3.8+, openpyxl 3.x

---

## File Map

```
school/
├── scripts/
│   ├── criar_diario.py          ← gera diario-de-classe.xlsx
│   └── criar_planejamento.py    ← gera planejamento-semanal.xlsx
├── diario-de-classe.xlsx        ← output (gerado)
├── planejamento-semanal.xlsx    ← output (gerado)
├── .claude/
│   └── ROADMAP.md
└── docs/superpowers/
    ├── specs/2026-04-19-diario-professor-design.md
    └── plans/2026-04-19-diario-professor.md
```

---

## Task 1: Setup — Python e openpyxl

**Files:**
- Create: `scripts/criar_diario.py`
- Create: `scripts/criar_planejamento.py`

- [ ] **Step 1.1: Verificar Python e instalar openpyxl**

```bash
python --version
pip install openpyxl
```

Expected: Python 3.8+ e `Successfully installed openpyxl-...`

- [ ] **Step 1.2: Criar pasta scripts**

```bash
mkdir -p scripts
```

- [ ] **Step 1.3: Criar scripts vazios para validar ambiente**

`scripts/criar_diario.py`:
```python
import openpyxl
print("openpyxl OK:", openpyxl.__version__)
```

`scripts/criar_planejamento.py`:
```python
import openpyxl
print("openpyxl OK:", openpyxl.__version__)
```

- [ ] **Step 1.4: Executar para confirmar ambiente**

```bash
python scripts/criar_diario.py
```

Expected: `openpyxl OK: 3.x.x`

- [ ] **Step 1.5: Commit**

```bash
git add scripts/
git commit -m "chore: setup scripts directory e ambiente openpyxl"
```

---

## Task 2: diario-de-classe.xlsx — Scaffold e Aba Turma

**Files:**
- Modify: `scripts/criar_diario.py`
- Output: `diario-de-classe.xlsx`

- [ ] **Step 2.1: Substituir conteúdo de criar_diario.py com scaffold completo + aba Turma**

```python
# scripts/criar_diario.py
import os
import openpyxl
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
from openpyxl.formatting.rule import CellIsRule
from openpyxl.utils import get_column_letter

N_ALUNOS   = 27
DISCIPLINAS = ["Português", "Matemática", "Ciências", "História", "Geografia", "Artes"]
BIMESTRES  = ["B1", "B2", "B3", "B4"]
MESES      = ["Fev", "Mar", "Abr", "Mai", "Jun", "Jul", "Ago", "Set", "Out", "Nov"]

COR_AZUL   = "1F4E79"
COR_BRANCO = "FFFFFF"
COR_ROXO   = "D9D2E9"   # destaque aluno especial

def _fill(cor):
    return PatternFill(fill_type="solid", fgColor=cor)

def _font(bold=False, cor=COR_BRANCO, size=10):
    return Font(bold=bold, color=cor, size=size)

def _center(wrap=False):
    return Alignment(horizontal="center", vertical="center", wrap_text=wrap)

def _border():
    s = Side(style="thin")
    return Border(left=s, right=s, top=s, bottom=s)

def aplicar_header(cell, valor):
    cell.value = valor
    cell.font   = _font(bold=True)
    cell.fill   = _fill(COR_AZUL)
    cell.alignment = _center(wrap=True)
    cell.border = _border()

def estilizar_dado(cell, centralizar=True):
    cell.border = _border()
    if centralizar:
        cell.alignment = _center()

def add_turma(wb):
    ws = wb.create_sheet("Turma")
    ws.sheet_properties.tabColor = "1F4E79"

    colunas = [
        ("Nº",               5),
        ("Nome Completo",   35),
        ("Dt. Nascimento",  16),
        ("Responsável",     30),
        ("Contato",         18),
        ("Al. Especial",    13),
        ("Observações",     40),
    ]
    for col, (titulo, largura) in enumerate(colunas, 1):
        aplicar_header(ws.cell(row=1, column=col), titulo)
        ws.column_dimensions[get_column_letter(col)].width = largura

    ws.row_dimensions[1].height = 30

    for row in range(2, N_ALUNOS + 2):
        ws.cell(row=row, column=1, value=row - 1)
        ws.cell(row=row, column=6, value="Não")
        for col in range(1, len(colunas) + 1):
            estilizar_dado(ws.cell(row=row, column=col))

    ws.freeze_panes = "B2"
    return ws

def main():
    saida = os.path.join(os.path.dirname(__file__), "..", "diario-de-classe.xlsx")
    wb = openpyxl.Workbook()
    wb.remove(wb.active)

    add_turma(wb)

    wb.save(saida)
    print(f"Gerado: {os.path.abspath(saida)}")

if __name__ == "__main__":
    main()
```

- [ ] **Step 2.2: Executar e verificar**

```bash
python scripts/criar_diario.py
```

Expected: `Gerado: ...diario-de-classe.xlsx`  
Abrir o arquivo e confirmar: aba "Turma" com cabeçalho azul, 27 linhas numeradas, coluna "Al. Especial" com "Não" pré-preenchido.

- [ ] **Step 2.3: Commit**

```bash
git add scripts/criar_diario.py diario-de-classe.xlsx
git commit -m "feat: criar_diario com aba Turma"
```

---

## Task 3: Abas de Presença Mensal

**Files:**
- Modify: `scripts/criar_diario.py`
- Output: `diario-de-classe.xlsx` (atualizado)

- [ ] **Step 3.1: Adicionar função `add_presenca_sheets` antes de `main()`**

```python
def add_presenca_sheets(wb):
    red_fill = _fill("FFC7CE")

    for mes in MESES:
        ws = wb.create_sheet(f"Presença - {mes}")
        ws.sheet_properties.tabColor = "2E75B6"

        # Cabeçalhos fixos
        aplicar_header(ws.cell(row=1, column=1), "Nº")
        aplicar_header(ws.cell(row=1, column=2), "Nome")

        # Dias 1-31
        for dia in range(1, 32):
            aplicar_header(ws.cell(row=1, column=2 + dia), str(dia))
            ws.column_dimensions[get_column_letter(2 + dia)].width = 4

        # Totalizadores
        for titulo, col in [("Total P", 34), ("Total F", 35), ("Total J", 36), ("% Freq", 37)]:
            aplicar_header(ws.cell(row=1, column=col), titulo)
            ws.column_dimensions[get_column_letter(col)].width = 9

        ws.row_dimensions[1].height = 25
        ws.column_dimensions["A"].width = 5
        ws.column_dimensions["B"].width = 30

        for row in range(2, N_ALUNOS + 2):
            ws.cell(row=row, column=1, value=row - 1)
            ws.cell(row=row, column=2).value = f"=Turma!B{row}"

            # Fórmulas de contagem: dias ficam em C-AG (cols 3-33)
            ws.cell(row=row, column=34).value = f'=COUNTIF(C{row}:AG{row},"P")'
            ws.cell(row=row, column=35).value = f'=COUNTIF(C{row}:AG{row},"F")'
            ws.cell(row=row, column=36).value = f'=COUNTIF(C{row}:AG{row},"J")'
            ws.cell(row=row, column=37).value = (
                f"=IF((AH{row}+AI{row}+AJ{row})=0,"
                f'"-",AH{row}/(AH{row}+AI{row}+AJ{row}))'
            )
            ws.cell(row=row, column=37).number_format = "0.0%"

            for col in range(1, 38):
                estilizar_dado(ws.cell(row=row, column=col))

        # Alerta vermelho para % Freq < 75%
        ws.conditional_formatting.add(
            f"AK2:AK{N_ALUNOS + 1}",
            CellIsRule(operator="lessThan", formula=["0.75"], fill=red_fill)
        )

        ws.freeze_panes = "C2"
```

- [ ] **Step 3.2: Chamar `add_presenca_sheets(wb)` em `main()`, antes de `wb.save()`**

```python
def main():
    saida = os.path.join(os.path.dirname(__file__), "..", "diario-de-classe.xlsx")
    wb = openpyxl.Workbook()
    wb.remove(wb.active)

    add_turma(wb)
    add_presenca_sheets(wb)      # ← adicionar esta linha

    wb.save(saida)
    print(f"Gerado: {os.path.abspath(saida)}")
```

- [ ] **Step 3.3: Executar e verificar**

```bash
python scripts/criar_diario.py
```

Abrir o arquivo e confirmar: 10 abas de presença (Fev a Nov), colunas de dias 1-31, colunas Total P/F/J e % Freq com fórmulas, linha de aluno com frequência < 75% fica em vermelho ao preencher.

- [ ] **Step 3.4: Commit**

```bash
git add scripts/criar_diario.py diario-de-classe.xlsx
git commit -m "feat: abas de presença mensal com totalizadores e alerta 75%"
```

---

## Task 4: Abas de Notas por Bimestre

**Files:**
- Modify: `scripts/criar_diario.py`
- Output: `diario-de-classe.xlsx` (atualizado)

- [ ] **Step 4.1: Adicionar função `add_notas_sheets` antes de `main()`**

```python
def add_notas_sheets(wb):
    red_fill    = _fill("FFC7CE")
    yellow_fill = _fill("FFEB9C")
    green_fill  = _fill("C6EFCE")

    for b in BIMESTRES:
        ws = wb.create_sheet(f"Notas - {b}")
        ws.sheet_properties.tabColor = "375623"

        headers = ["Nº", "Nome"] + DISCIPLINAS + ["Média Geral"]
        larguras = [5, 30, 14, 14, 12, 12, 12, 10, 13]
        for col, (h, w) in enumerate(zip(headers, larguras), 1):
            aplicar_header(ws.cell(row=1, column=col), h)
            ws.column_dimensions[get_column_letter(col)].width = w

        ws.row_dimensions[1].height = 25

        for row in range(2, N_ALUNOS + 2):
            ws.cell(row=row, column=1, value=row - 1)
            ws.cell(row=row, column=2).value = f"=Turma!B{row}"
            # Média das 6 disciplinas (cols C=3 a H=8)
            ws.cell(row=row, column=9).value = f'=IFERROR(AVERAGE(C{row}:H{row}),"")'
            ws.cell(row=row, column=9).number_format = "0.0"
            for col in range(1, 10):
                estilizar_dado(ws.cell(row=row, column=col))

        # Formatação condicional: vermelho < 5, amarelo 5–6.9, verde >= 7
        grade_range = f"C2:I{N_ALUNOS + 1}"
        ws.conditional_formatting.add(
            grade_range, CellIsRule(operator="lessThan", formula=["5"], fill=red_fill)
        )
        ws.conditional_formatting.add(
            grade_range, CellIsRule(operator="between", formula=["5", "6.9"], fill=yellow_fill)
        )
        ws.conditional_formatting.add(
            grade_range, CellIsRule(operator="greaterThanOrEqual", formula=["7"], fill=green_fill)
        )

        ws.freeze_panes = "C2"
```

- [ ] **Step 4.2: Adicionar chamada em `main()`**

```python
    add_notas_sheets(wb)         # ← após add_presenca_sheets(wb)
```

- [ ] **Step 4.3: Executar e verificar**

```bash
python scripts/criar_diario.py
```

Abrir o arquivo, ir em "Notas - B1" e digitar nota 4 numa célula → fundo vermelho. Nota 6 → amarelo. Nota 8 → verde. Coluna Média Geral calcula automaticamente.

- [ ] **Step 4.4: Commit**

```bash
git add scripts/criar_diario.py diario-de-classe.xlsx
git commit -m "feat: abas de notas bimestrais com formatação condicional"
```

---

## Task 5: Abas Recuperação, Contatos e Ocorrências

**Files:**
- Modify: `scripts/criar_diario.py`
- Output: `diario-de-classe.xlsx` (atualizado)

- [ ] **Step 5.1: Adicionar as três funções antes de `main()`**

```python
def add_recuperacao(wb):
    ws = wb.create_sheet("Recuperação")
    ws.sheet_properties.tabColor = "FF0000"

    headers = [
        ("Bimestre",       12),
        ("Nº Aluno",        9),
        ("Nome do Aluno",  30),
        ("Disciplina",     14),
        ("Nota Original",  14),
        ("Nota Recup.",    14),
        ("Situação Final", 16),
    ]
    for col, (titulo, largura) in enumerate(headers, 1):
        aplicar_header(ws.cell(row=1, column=col), titulo)
        ws.column_dimensions[get_column_letter(col)].width = largura

    ws.row_dimensions[1].height = 25

    # Pré-popular 40 linhas em branco com fórmula de Situação Final
    for row in range(2, 42):
        ws.cell(row=row, column=7).value = (
            f'=IF(F{row}="","",IF(F{row}>=5,"Aprovado","Em recuperação"))'
        )
        green_fill = _fill("C6EFCE")
        red_fill   = _fill("FFC7CE")
        for col in range(1, 8):
            estilizar_dado(ws.cell(row=row, column=col))

    ws.conditional_formatting.add(
        "G2:G41",
        CellIsRule(operator="equal", formula=['"Aprovado"'], fill=_fill("C6EFCE"))
    )
    ws.conditional_formatting.add(
        "G2:G41",
        CellIsRule(operator="equal", formula=['"Em recuperação"'], fill=_fill("FFC7CE"))
    )

    ws.freeze_panes = "A2"


def add_contatos(wb):
    ws = wb.create_sheet("Contatos e Reuniões")
    ws.sheet_properties.tabColor = "FF9900"

    headers = [
        ("Data",          12),
        ("Nº Aluno",       9),
        ("Nome do Aluno", 30),
        ("Responsável",   25),
        ("Tipo",          14),
        ("Assunto",       40),
        ("Encaminhamento",40),
    ]
    for col, (titulo, largura) in enumerate(headers, 1):
        aplicar_header(ws.cell(row=1, column=col), titulo)
        ws.column_dimensions[get_column_letter(col)].width = largura

    ws.row_dimensions[1].height = 25

    for row in range(2, 102):  # 100 linhas
        for col in range(1, 8):
            estilizar_dado(ws.cell(row=row, column=col), centralizar=(col in [1, 2, 5]))

    ws.freeze_panes = "A2"


def add_ocorrencias(wb):
    ws = wb.create_sheet("Ocorrências")
    ws.sheet_properties.tabColor = "9900FF"

    headers = [
        ("Data",           12),
        ("Nº Aluno",        9),
        ("Nome do Aluno",  30),
        ("Descrição",      50),
        ("Providência",    50),
    ]
    for col, (titulo, largura) in enumerate(headers, 1):
        aplicar_header(ws.cell(row=1, column=col), titulo)
        ws.column_dimensions[get_column_letter(col)].width = largura

    ws.row_dimensions[1].height = 25

    for row in range(2, 102):
        for col in range(1, 6):
            c = ws.cell(row=row, column=col)
            c.border = _border()
            c.alignment = Alignment(
                horizontal="center" if col in [1, 2] else "left",
                vertical="top",
                wrap_text=True
            )
        ws.row_dimensions[row].height = 40

    ws.freeze_panes = "A2"
```

- [ ] **Step 5.2: Adicionar chamadas em `main()`**

```python
    add_recuperacao(wb)
    add_contatos(wb)
    add_ocorrencias(wb)
```

- [ ] **Step 5.3: Executar e verificar**

```bash
python scripts/criar_diario.py
```

Confirmar: aba "Recuperação" com fórmula de Situação Final (digitar "Aprovado" ou "Em recuperação" aparece colorido). Abas "Contatos e Reuniões" e "Ocorrências" com layout limpo.

- [ ] **Step 5.4: Commit**

```bash
git add scripts/criar_diario.py diario-de-classe.xlsx
git commit -m "feat: abas Recuperação, Contatos e Ocorrências"
```

---

## Task 6: Aba Resumo Anual

**Files:**
- Modify: `scripts/criar_diario.py`
- Output: `diario-de-classe.xlsx` (atualizado)

- [ ] **Step 6.1: Adicionar função `add_resumo` antes de `main()`**

```python
def add_resumo(wb):
    ws = wb.create_sheet("Resumo Anual")
    ws.sheet_properties.tabColor = "404040"

    # Linha 1: título
    ws.merge_cells("A1:L1")
    t = ws.cell(row=1, column=1, value="RESUMO ANUAL — 1º ANO")
    t.font = Font(bold=True, size=14, color=COR_BRANCO)
    t.fill = _fill("404040")
    t.alignment = _center()
    ws.row_dimensions[1].height = 35

    # Linha 2: cabeçalhos
    headers = ["Nº", "Nome", "Freq. %"] + [f"Média {b}" for b in BIMESTRES] + ["Média Final", "Situação"]
    larguras = [5, 30, 9] + [11] * 4 + [12, 16]
    for col, (h, w) in enumerate(zip(headers, larguras), 1):
        aplicar_header(ws.cell(row=2, column=col), h)
        ws.column_dimensions[get_column_letter(col)].width = w

    ws.row_dimensions[2].height = 25

    for row in range(3, N_ALUNOS + 3):
        aluno_row = row - 1  # linha correspondente nas abas de notas (começa em 2)
        ws.cell(row=row, column=1, value=row - 2)
        ws.cell(row=row, column=2).value = f"=Turma!B{aluno_row + 1}"

        # Frequência: média das % de todos os meses (cols AK nas abas de presença)
        freq_refs = "+".join([f"'Presença - {m}'!AK{aluno_row + 1}" for m in MESES])
        ws.cell(row=row, column=3).value = f"=IFERROR(({freq_refs})/{len(MESES)},\"\")"
        ws.cell(row=row, column=3).number_format = "0.0%"

        # Médias por bimestre
        for b_idx, b in enumerate(BIMESTRES):
            col = 4 + b_idx
            ws.cell(row=row, column=col).value = f"='Notas - {b}'!I{aluno_row + 1}"
            ws.cell(row=row, column=col).number_format = "0.0"

        # Média Final
        ws.cell(row=row, column=8).value = f"=IFERROR(AVERAGE(D{row}:G{row}),\"\")"
        ws.cell(row=row, column=8).number_format = "0.0"

        # Situação
        ws.cell(row=row, column=9).value = (
            f'=IF(H{row}="","",IF(AND(H{row}>=5,C{row}>=0.75),"Aprovado",'
            f'IF(H{row}>=5,"Rec. Freq.","Recuperação")))'
        )

        for col in range(1, 10):
            estilizar_dado(ws.cell(row=row, column=col))

    # Cores para Situação
    ws.conditional_formatting.add(
        f"I3:I{N_ALUNOS + 2}",
        CellIsRule(operator="equal", formula=['"Aprovado"'], fill=_fill("C6EFCE"))
    )
    ws.conditional_formatting.add(
        f"I3:I{N_ALUNOS + 2}",
        CellIsRule(operator="equal", formula=['"Recuperação"'], fill=_fill("FFC7CE"))
    )
    ws.conditional_formatting.add(
        f"I3:I{N_ALUNOS + 2}",
        CellIsRule(operator="equal", formula=['"Rec. Freq."'], fill=_fill("FFEB9C"))
    )

    ws.freeze_panes = "C3"
```

- [ ] **Step 6.2: Adicionar chamada em `main()` — Resumo Anual deve ser a última aba**

```python
    add_resumo(wb)
```

- [ ] **Step 6.3: Reorganizar ordem das abas em `main()` para ficar:**

```python
def main():
    saida = os.path.join(os.path.dirname(__file__), "..", "diario-de-classe.xlsx")
    wb = openpyxl.Workbook()
    wb.remove(wb.active)

    add_turma(wb)
    add_presenca_sheets(wb)
    add_notas_sheets(wb)
    add_recuperacao(wb)
    add_contatos(wb)
    add_ocorrencias(wb)
    add_resumo(wb)

    wb.save(saida)
    print(f"Gerado: {os.path.abspath(saida)}")

if __name__ == "__main__":
    main()
```

- [ ] **Step 6.4: Executar e verificar**

```bash
python scripts/criar_diario.py
```

Abrir o arquivo, confirmar aba "Resumo Anual" com fórmulas de média e situação. Preencher notas fictícias em "Notas - B1" e verificar que o Resumo Anual reflete automaticamente.

- [ ] **Step 6.5: Commit**

```bash
git add scripts/criar_diario.py diario-de-classe.xlsx
git commit -m "feat: aba Resumo Anual com médias, frequência e situação final"
```

---

## Task 7: planejamento-semanal.xlsx

**Files:**
- Create: `scripts/criar_planejamento.py`
- Output: `planejamento-semanal.xlsx`

- [ ] **Step 7.1: Criar scripts/criar_planejamento.py completo**

```python
# scripts/criar_planejamento.py
import os
import openpyxl
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
from openpyxl.formatting.rule import CellIsRule
from openpyxl.utils import get_column_letter

DISCIPLINAS = ["Português", "Matemática", "Ciências", "História", "Geografia", "Artes"]
BIMESTRES   = ["B1", "B2", "B3", "B4"]

COR_AZUL   = "1F4E79"
COR_BRANCO = "FFFFFF"

def _fill(cor):
    return PatternFill(fill_type="solid", fgColor=cor)

def _font(bold=False, cor=COR_BRANCO, size=10):
    return Font(bold=bold, color=cor, size=size)

def _center(wrap=False):
    return Alignment(horizontal="center", vertical="center", wrap_text=wrap)

def _border():
    s = Side(style="thin")
    return Border(left=s, right=s, top=s, bottom=s)

def aplicar_header(cell, valor, cor=COR_AZUL):
    cell.value = valor
    cell.font      = _font(bold=True)
    cell.fill      = _fill(cor)
    cell.alignment = _center(wrap=True)
    cell.border    = _border()

def add_calendario(wb):
    ws = wb.create_sheet("Calendário Letivo")
    ws.sheet_properties.tabColor = "1F4E79"

    # Seção: Bimestres
    ws.merge_cells("A1:D1")
    t = ws.cell(row=1, column=1, value="CALENDÁRIO LETIVO — 1º ANO")
    t.font = Font(bold=True, size=13, color=COR_BRANCO)
    t.fill = _fill(COR_AZUL)
    t.alignment = _center()
    ws.row_dimensions[1].height = 30

    bim_headers = ["Bimestre", "Início", "Término", "Obs."]
    bim_widths  = [16, 14, 14, 40]
    for col, (h, w) in enumerate(zip(bim_headers, bim_widths), 1):
        aplicar_header(ws.cell(row=2, column=col), h)
        ws.column_dimensions[get_column_letter(col)].width = w

    bimestres_datas = [
        ("1º Bimestre", "03/02/2025", "25/04/2025", ""),
        ("2º Bimestre", "28/04/2025", "04/07/2025", ""),
        ("3º Bimestre", "28/07/2025", "26/09/2025", ""),
        ("4º Bimestre", "29/09/2025", "05/12/2025", ""),
    ]
    for r, (b, ini, fim, obs) in enumerate(bimestres_datas, 3):
        for col, val in enumerate([b, ini, fim, obs], 1):
            c = ws.cell(row=r, column=col, value=val)
            c.border    = _border()
            c.alignment = _center() if col < 4 else Alignment(horizontal="left", vertical="center")

    # Seção: Datas importantes
    ws.cell(row=8, column=1, value="DATAS IMPORTANTES").font = Font(bold=True, size=11)
    ws.row_dimensions[8].height = 25

    di_headers = ["Data", "Evento", "Tipo"]
    di_widths  = [14, 50, 25]
    for col, (h, w) in enumerate(zip(di_headers, di_widths), 1):
        aplicar_header(ws.cell(row=9, column=col), h)
        ws.column_dimensions[get_column_letter(col)].width = w

    eventos = [
        ("25/01/2025", "Início do ano letivo", "Institucional"),
        ("01/05/2025", "Dia do Trabalho — Feriado", "Feriado"),
        ("12/06/2025", "Corpus Christi", "Feriado"),
        ("25/07/2025", "Retorno 2º semestre", "Institucional"),
        ("07/09/2025", "Independência — Feriado", "Feriado"),
        ("12/10/2025", "N. Sra. Aparecida — Feriado", "Feriado"),
        ("02/11/2025", "Finados — Feriado", "Feriado"),
        ("15/11/2025", "Proclamação da República — Feriado", "Feriado"),
    ]
    for r, (data, evento, tipo) in enumerate(eventos, 10):
        for col, val in enumerate([data, evento, tipo], 1):
            c = ws.cell(row=r, column=col, value=val)
            c.border    = _border()
            c.alignment = _center() if col != 2 else Alignment(horizontal="left", vertical="center")

    # Linhas em branco para a professora adicionar feriados municipais
    for r in range(18, 30):
        for col in range(1, 4):
            ws.cell(row=r, column=col).border = _border()

    ws.freeze_panes = "A3"


def add_planejamento_bimestre(wb, bimestre, n_semanas):
    ws = wb.create_sheet(f"{bimestre} - Planejamento")
    ws.sheet_properties.tabColor = "375623"

    # Linha 1: título
    n_cols = 2 + len(DISCIPLINAS) * 4
    ws.merge_cells(f"A1:{get_column_letter(n_cols)}1")
    t = ws.cell(row=1, column=1, value=f"PLANEJAMENTO SEMANAL — {bimestre}")
    t.font = Font(bold=True, size=12, color=COR_BRANCO)
    t.fill = _fill(COR_AZUL)
    t.alignment = _center()
    ws.row_dimensions[1].height = 30

    # Linha 2: cabeçalhos fixos
    aplicar_header(ws.cell(row=2, column=1), "Semana")
    aplicar_header(ws.cell(row=2, column=2), "Período")
    ws.column_dimensions["A"].width = 9
    ws.column_dimensions["B"].width = 22

    # Cabeçalhos de disciplinas (cada disciplina ocupa 4 colunas)
    sub_headers = ["Conteúdo / Tema", "Objetivo", "Atividade / Recurso", "Realizado?"]
    sub_widths  = [30, 25, 25, 11]
    COR_DISC = ["2E75B6", "375623", "7030A0", "C55A11", "C00000", "4472C4"]

    for d_idx, disc in enumerate(DISCIPLINAS):
        col_base = 3 + d_idx * 4
        # Merge para nome da disciplina
        ws.merge_cells(
            start_row=2, start_column=col_base,
            end_row=2,   end_column=col_base + 3
        )
        aplicar_header(ws.cell(row=2, column=col_base), disc, cor=COR_DISC[d_idx])

        # Sub-headers na linha 3
        for s_idx, (sh, sw) in enumerate(zip(sub_headers, sub_widths)):
            aplicar_header(ws.cell(row=3, column=col_base + s_idx), sh, cor=COR_DISC[d_idx])
            ws.column_dimensions[get_column_letter(col_base + s_idx)].width = sw

    ws.row_dimensions[2].height = 28
    ws.row_dimensions[3].height = 28

    # Linhas de semanas
    for sem in range(1, n_semanas + 1):
        row = 3 + sem
        ws.cell(row=row, column=1, value=f"Semana {sem}")
        ws.cell(row=row, column=2, value="")  # professora preenche o período

        for col in range(1, n_cols + 1):
            c = ws.cell(row=row, column=col)
            c.border = _border()
            is_realizado = (col - 3) % 4 == 3 and col >= 3
            c.alignment = Alignment(
                horizontal="center" if col == 1 or is_realizado else "left",
                vertical="top",
                wrap_text=True
            )
        ws.row_dimensions[row].height = 60

    # Semana de recuperação (última linha, destaque)
    row_recup = 3 + n_semanas + 1
    ws.merge_cells(f"A{row_recup}:{get_column_letter(n_cols)}{row_recup}")
    c = ws.cell(row=row_recup, column=1, value="SEMANA DE RECUPERAÇÃO / AVALIAÇÃO")
    c.font = Font(bold=True, color=COR_BRANCO)
    c.fill = _fill("FF0000")
    c.alignment = _center()
    c.border = _border()
    ws.row_dimensions[row_recup].height = 25

    ws.freeze_panes = "C4"


def main():
    saida = os.path.join(os.path.dirname(__file__), "..", "planejamento-semanal.xlsx")
    wb = openpyxl.Workbook()
    wb.remove(wb.active)

    add_calendario(wb)
    # Semanas por bimestre (ajuste conforme o calendário da escola)
    semanas_por_bimestre = {"B1": 10, "B2": 10, "B3": 10, "B4": 9}
    for b, n in semanas_por_bimestre.items():
        add_planejamento_bimestre(wb, b, n)

    wb.save(saida)
    print(f"Gerado: {os.path.abspath(saida)}")

if __name__ == "__main__":
    main()
```

- [ ] **Step 7.2: Executar e verificar**

```bash
python scripts/criar_planejamento.py
```

Abrir `planejamento-semanal.xlsx` e confirmar: aba "Calendário Letivo" com datas, abas "B1 - Planejamento" a "B4 - Planejamento" com colunas por disciplina, sub-colunas de Conteúdo / Objetivo / Atividade / Realizado?, e linha de recuperação ao final.

- [ ] **Step 7.3: Commit**

```bash
git add scripts/criar_planejamento.py planejamento-semanal.xlsx
git commit -m "feat: planejamento-semanal.xlsx com calendário e estrutura por bimestre/disciplina"
```

---

## Task 8: Atualizar ROADMAP.md

**Files:**
- Modify: `.claude/ROADMAP.md`

- [ ] **Step 8.1: Atualizar seção FEITO e EM ANDAMENTO no ROADMAP**

Mover todos os itens de "EM ANDAMENTO" e "TODO" para "FEITO", e atualizar "Próximas fases":

```markdown
## FEITO

- [x] Levantamento de requisitos com a professora
- [x] Design aprovado das duas planilhas
- [x] Documento de spec criado
- [x] Plano de implementação criado
- [x] Criação da pasta `.claude` com ROADMAP
- [x] Setup Python + openpyxl
- [x] diario-de-classe.xlsx — aba Turma
- [x] diario-de-classe.xlsx — abas Presença mensais (Fev-Nov)
- [x] diario-de-classe.xlsx — abas Notas B1-B4 com formatação condicional
- [x] diario-de-classe.xlsx — aba Recuperação
- [x] diario-de-classe.xlsx — aba Contatos e Reuniões
- [x] diario-de-classe.xlsx — aba Ocorrências
- [x] diario-de-classe.xlsx — aba Resumo Anual
- [x] planejamento-semanal.xlsx — Calendário Letivo
- [x] planejamento-semanal.xlsx — Planejamento B1-B4

## PRÓXIMAS FASES (aguardando)

- [ ] Professora adiciona o livro didático à pasta do projeto
- [ ] Preencher conteúdo real das aulas no planejamento semanal conforme livro e metodologia da escola
- [ ] Ajustes de datas no Calendário Letivo conforme calendário municipal real
```

- [ ] **Step 8.2: Commit final**

```bash
git add .claude/ROADMAP.md
git commit -m "docs: atualizar ROADMAP com implementação concluída"
```

---

## Self-Review Checklist

- [x] **Cobertura do spec:** Turma ✓, Presença ✓, Notas ✓, Recuperação ✓, Contatos ✓, Ocorrências ✓, Resumo Anual ✓, Calendário ✓, Planejamento B1-B4 ✓
- [x] **Placeholders:** nenhum TBD ou TODO no código
- [x] **Consistência de tipos:** funções utilitárias (`_fill`, `_font`, `aplicar_header`, `estilizar_dado`) usadas consistentemente em todos os tasks
- [x] **Aluno especial:** campo na aba Turma com "Al. Especial" e valor padrão "Não" — professora edita para o aluno específico
- [x] **Frequência 75%:** alerta condicional implementado na Task 3
- [x] **Bimestres:** 4 abas Notas + 4 abas Planejamento
- [x] **Fora do escopo:** conteúdo do livro aguarda entrega futura — estrutura pronta para receber
