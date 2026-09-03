# -*- coding: utf-8 -*-
"""Media de comissoes dos ultimos 12 meses - base de calculo de ferias.
VALDICK (Trancoso), RENALDO (Centro) e MANOEL (Trancoso)."""
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

OUT = "/home/user/site-farmacia/relatorios/contabilidade/MEDIA_COMISSOES_12_MESES_FERIAS.xlsx"
F = "Arial"
MONEY = 'R$ #,##0.00;-R$ #,##0.00;"-"'
def font(sz=10, b=False, color="000000", it=False):
    return Font(name=F, size=sz, bold=b, color=color, italic=it)
FILL_TIT = PatternFill("solid", fgColor="1F3864")
FILL_SEC = PatternFill("solid", fgColor="D9E2F3")
FILL_HDR = PatternFill("solid", fgColor="2E5496")
FILL_IN = PatternFill("solid", fgColor="FFF2CC")
FILL_TOT = PatternFill("solid", fgColor="E2EFDA")
FILL_LIQ = PatternFill("solid", fgColor="C6E0B4")
FILL_ALERT = PatternFill("solid", fgColor="FFF0F0")
thin = Side(style="thin", color="BFBFBF")
BOX = Border(left=thin, right=thin, top=thin, bottom=thin)
BLUE = "0000FF"

MESES = ["set/2025", "out/2025", "nov/2025", "dez/2025", "jan/2026", "fev/2026",
         "mar/2026", "abr/2026", "mai/2026", "jun/2026", "jul/2026", "ago/2026"]
MESES_MANOEL = ["mai/2025", "jun/2025", "jul/2025", "ago/2025", "set/2025", "out/2025",
                "nov/2025", "dez/2025", "jan/2026", "fev/2026", "mar/2026", "abr/2026"]

CASOS = [
 dict(nome="VALDICK GONÇALVES RODRIGUES", cod="162", loja="TRANCOSO", meses=MESES,
      periodo="01/09/2025 a 31/08/2026", ferias="férias de 01 a 30/09/2026",
      dobro=True,
      conhecidos={"jul/2026": (908.70, None), "ago/2026": (886.30, 260.00)}),
 dict(nome="RENALDO RODRIGUES", cod="151", loja="CENTRO (MATRIZ)", meses=MESES,
      periodo="01/09/2025 a 31/08/2026", ferias="férias de 01 a 30/09/2026",
      dobro=False,
      conhecidos={"fev/2026": (855.00, 395.00), "mar/2026": (942.00, 253.00),
                  "abr/2026": (960.00, 426.00), "jul/2026": (1200.00, None),
                  "ago/2026": (847.94, 96.00)}),
 dict(nome="MANOEL SILVA", cod="92", loja="TRANCOSO", meses=MESES_MANOEL,
      periodo="01/05/2025 a 30/04/2026", ferias="férias gozadas em maio/2026",
      dobro=True,
      conhecidos={"fev/2026": (845.00, 202.00), "mar/2026": (1350.00, 240.00),
                  "abr/2026": (150.00, 263.00)}),
]

wb = openpyxl.Workbook()
primeiro = True
for caso in CASOS:
    ws = wb.active if primeiro else wb.create_sheet()
    ws.title = caso["nome"].split()[0][:20]
    primeiro = False
    ws.sheet_view.showGridLines = False
    for col, w in zip("ABCDEF", [14, 18, 18, 18, 18, 60]):
        ws.column_dimensions[col].width = w
    ws.merge_cells("A1:F1")
    c = ws.cell(1, 1, f"MÉDIA DE COMISSÕES DOS ÚLTIMOS 12 MESES — {caso['nome']}")
    c.font = font(13, True, "FFFFFF"); c.fill = FILL_TIT
    c.alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[1].height = 26
    ws.merge_cells("A2:F2")
    c = ws.cell(2, 1, f"Loja {caso['loja']} · código {caso['cod']} · período {caso['periodo']} · base para {caso['ferias']}")
    c.font = font(10, True, "1F3864"); c.alignment = Alignment(horizontal="center")
    hdr = ["MÊS", "COMISSÃO APURADA", "INCENTIVOS", "COMISSÃO PAGA", "TOTAL VARIÁVEL", "OBSERVAÇÃO"]
    for i, h in enumerate(hdr, 1):
        c = ws.cell(4, i, h); c.font = font(9, True, "FFFFFF"); c.fill = FILL_HDR
        c.alignment = Alignment(horizontal="center", wrap_text=True); c.border = BOX
    ws.row_dimensions[4].height = 28
    ws.freeze_panes = "A5"
    r = 5
    ini = r
    for m in caso["meses"]:
        ws.cell(r, 1, m).font = font(10, True)
        conhecido = caso["conhecidos"].get(m)
        for i in (2, 3):
            c = ws.cell(r, i, conhecido[i - 2] if conhecido else None)
            c.number_format = MONEY; c.font = font(10, False, BLUE); c.fill = FILL_IN
        mult = 2 if caso["dobro"] else 1
        c = ws.cell(r, 4, f"=ROUND(B{r}*{mult},2)")
        c.number_format = MONEY; c.font = font(10)
        c = ws.cell(r, 5, f"=D{r}+C{r}")
        c.number_format = MONEY; c.font = font(10, True); c.fill = FILL_TOT
        obs = ""
        if conhecido:
            obs = "Valor localizado nas planilhas que você já enviou — CONFERIR com o relatório do InovaFarma."
        c = ws.cell(r, 6, obs); c.font = font(9, it=True)
        c.alignment = Alignment(wrap_text=True, vertical="center")
        for i in range(1, 7):
            ws.cell(r, i).border = BOX
        r += 1
    fim = r - 1
    ws.cell(r, 1, "SOMA 12 MESES").font = font(10, True)
    for i in (2, 3, 4, 5):
        L = get_column_letter(i)
        c = ws.cell(r, i, f"=SUM({L}{ini}:{L}{fim})")
        c.number_format = MONEY; c.font = font(10, True)
    for i in range(1, 7):
        ws.cell(r, i).fill = FILL_TOT; ws.cell(r, i).border = BOX
    soma = r
    r += 1
    ws.cell(r, 1, "MÉDIA MENSAL (÷ 12)").font = font(11, True)
    for i in (2, 3, 4, 5):
        L = get_column_letter(i)
        c = ws.cell(r, i, f"=ROUND({L}{soma}/12,2)")
        c.number_format = MONEY; c.font = font(11, True)
    for i in range(1, 7):
        ws.cell(r, i).fill = FILL_LIQ; ws.cell(r, i).border = BOX
    ws.row_dimensions[r].height = 22
    c = ws.cell(r, 6, "É esta média (coluna TOTAL VARIÁVEL) que entra na base das férias, somada ao salário.")
    c.font = font(9, True, "1F3864"); c.alignment = Alignment(wrap_text=True, vertical="center")
    ws.auto_filter.ref = f"A4:F{fim}"
    r += 2
    passos = [
     "COMO PUXAR NO INOVAFARMA (tela Comissão de Vendedores):",
     f"1. EMPRESA: selecione a loja {caso['loja']}.",
     "2. TIPO: PRODUTOS VENDIDOS. VISUALIZAÇÃO: pode usar RESUMIDO POR VENDEDOR — para a média só interessam os totais do mês.",
     f"3. PERÍODO: rode MÊS A MÊS (do dia 1 ao último dia de cada mês), de {caso['periodo']}.",
     f"4. VENDEDOR: {caso['nome']} (código {caso['cod']}).",
     "5. COMISSÃO: TODOS OS PRODUTOS. Clique em VISUALIZA e anote TOTAL COMISSÃO VENDEDOR e TOTAL INCENTIVO VENDEDOR.",
     "6. Lance os dois valores nas colunas amarelas deste quadro; a soma e a média saem sozinhas.",
     "",
     "OBSERVAÇÕES:",
     "• Se o sistema deixar rodar o período inteiro de uma vez, ainda assim é preciso o valor mês a mês: a média de férias é a soma dos 12 meses dividida por 12.",
     "• O aviso da tela lembra que vendas da Farmácia Popular não geram comissão — elas já ficam de fora do relatório.",
     "• Meses em que a pessoa esteve de férias ou afastada entram com o valor que houve; se não houve nada, deixe zero.",
    ]
    if caso["dobro"]:
        passos.insert(7, "• TRANCOSO paga o dobro da comissão apurada: a coluna COMISSÃO PAGA já multiplica por 2. É esse valor que vale para a média, porque é o que foi efetivamente pago.")
    for t in passos:
        ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=6)
        c = ws.cell(r, 1, t)
        neg = t.endswith(":")
        c.font = font(9, neg, "C00000" if neg else "000000")
        c.alignment = Alignment(wrap_text=True, vertical="center")
        if neg:
            c.fill = FILL_ALERT
        r += 1

wb.save(OUT)
print("ok", OUT)
