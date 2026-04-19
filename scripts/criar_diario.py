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

def main():
    saida = os.path.join(os.path.dirname(__file__), "..", "diario-de-classe.xlsx")
    wb = openpyxl.Workbook()
    wb.remove(wb.active)

    add_turma(wb)
    add_presenca_sheets(wb)
    add_notas_sheets(wb)

    wb.save(saida)
    print(f"Gerado: {os.path.abspath(saida)}")

if __name__ == "__main__":
    main()
