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
