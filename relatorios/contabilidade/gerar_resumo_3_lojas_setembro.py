# -*- coding: utf-8 -*-
"""Resumo consolidado das 3 lojas - competencia SETEMBRO/2026, pagamento 05/10/2026.
Gerado a partir das configuracoes de gerar_relatorio_setembro_2026.py."""
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

OUT = "/home/user/site-farmacia/relatorios/contabilidade/RESUMO_3_LOJAS_SETEMBRO_2026_pgto_05-10-2026.xlsx"
F = "Arial"
MONEY = 'R$ #,##0.00;-R$ #,##0.00;"-"'
def font(sz=10, b=False, color="000000", it=False):
    return Font(name=F, size=sz, bold=b, color=color, italic=it)
FILL_TIT = PatternFill("solid", fgColor="1F3864")
FILL_SEC = PatternFill("solid", fgColor="D9E2F3")
FILL_HDR = PatternFill("solid", fgColor="2E5496")
FILL_TOT = PatternFill("solid", fgColor="E2EFDA")
FILL_LIQ = PatternFill("solid", fgColor="C6E0B4")
FILL_IN = PatternFill("solid", fgColor="FFF2CC")
FILL_ALERT = PatternFill("solid", fgColor="FFF0F0")
thin = Side(style="thin", color="BFBFBF")
BOX = Border(left=thin, right=thin, top=thin, bottom=thin)

FATOR = 6 / 24          # setembro/2026: 4 domingos + 07/09 + 08/09 / 24 dias uteis
HORA = 1621 / 220
HE = round(HORA * 1.5, 4)     # 11,0523
NOT_ = round(HORA * 0.2, 4)   # 1,4736
DIA = round(1621 / 30, 2)     # 54,03

def E(nome, sal, com, he=0.0, notu=0.0, aux=0.0, meta=0.0, inc=0.0, vt6=0.0, obs="",
      bonif=0.0, dsrhe=0.0, salfam=0.0, insuf=0.0, vales=0.0, inss=0.0, irrf=0.0, liq_betel=None):
    dsr = round((com + he) * FATOR, 2)
    prov = round(sal + he + notu + com + dsr + dsrhe + aux + meta + bonif + inc + salfam + insuf, 2)
    desc = round(-inc + vt6 + vales + inss + irrf, 2)
    return dict(nome=nome, sal=sal, he=he, notu=notu, com=com, dsr=dsr, dsrhe=dsrhe, aux=aux,
                meta=meta, bonif=bonif, salfam=salfam, insuf=insuf, inc=inc, prov=prov,
                vt6=vt6, adiant_inc=-inc, adiant_sal=vales, inss=inss, irrf=irrf, ferias=None,
                desc=desc, liq=round(prov + desc, 2), liq_betel=liq_betel,
                dif=round(round(prov + desc, 2) - (liq_betel if liq_betel is not None else 0), 2), obs=obs)

import importlib.util, os
_spec = importlib.util.spec_from_file_location(
    "setembro", os.path.join(os.path.dirname(os.path.abspath(__file__)), "gerar_relatorio_setembro_2026.py"))
SET = importlib.util.module_from_spec(_spec); _spec.loader.exec_module(SET)

def _valor_ponto(tipo, qtd, sal):
    hora = sal / 220
    if tipo == "HE50":  return round(round(hora * 1.5, 4) * qtd, 2), 0.0
    if tipo == "HE100": return round(round(hora * 2.0, 4) * qtd, 2), 0.0
    if tipo == "FER":   return round(round(sal / 30, 2) * qtd, 2), 0.0
    if tipo == "NOT":   return 0.0, round(round(hora * 0.2, 4) * qtd, 2)
    if tipo == "VALOR": return 0.0, float(qtd)
    return 0.0, 0.0

OBS_SET = {
 "DEAN": "Não recebe comissão.",
 "AGNOR": "Comissão fixa de R$ 2.000,00 — já entra mesmo sem o apurado lançado.",
 "EDEY": "Adicional noturno fixo de R$ 350,00 — conferir com o ponto.",
 "JOEL": "Férias até 03/09 (pagas no recibo de agosto); voltou em 04/09 — salário de 27 dias.",
 "SARA": "Somar aqui o saldo dela na loja Centro (código 43).",
 "ARIANE": "Não recebe comissão.",
 "GENECIR": "Feriados de 07/09 e 08/09 trabalhados (R$ 108,06). Adiantamento de R$ 500,00 em 09/09. VT de R$ 270,00.",
 "RENALDO": "Férias de 01 a 30/09 — salário zerado; férias e 1/3 em recibo próprio.",
 "PEDRO": "Registrado na matriz. Nesta folha segue com R$ 1.621,00; o salário de R$ 4.000,00 começa na competência de outubro (pagamento 05/11). Não recebe comissão. O holerite de agosto dele não veio da Betel — pedir junto.",
 "UILLIAN": "Não recebe comissão.",
 "VALDICK": "Férias de 01 a 30/09 — salário zerado; férias e 1/3 em recibo próprio.",
 "TAMILES": "Atestado de 02 a 08/09, abonado — sem desconto.",
}

def _loja(cfg):
    out = []
    for n in cfg["EMP"]:
        sal = cfg["SALARIO"].get(n, 0.0)
        base_sal = 1621.0 if sal == 0 else sal
        com_inova = 0.0 if n in cfg["SEM_COMISSAO"] else round(cfg["INOVA"][n][4] * cfg["mult"], 2)
        fixa = cfg.get("COM_FIXA", {}).get(n)
        com = max(com_inova, fixa) if fixa else com_inova
        he = notu = 0.0
        for func, tipo, qtd, _o in cfg.get("PONTO", []):
            if func == n and qtd:
                a, b = _valor_ponto(tipo, qtd, base_sal)
                he += a; notu += b
        out.append(E(n, sal, com, he=round(he, 2), notu=round(notu, 2),
                     aux=cfg["AUXGER"].get(n, 0.0), meta=cfg["METACX"].get(n, 0.0),
                     vt6=cfg["VT6"].get(n, 0.0), vales=cfg.get("VALES", {}).get(n, 0.0),
                     obs=OBS_SET.get(n, "")))
    return out

ARRAIAL, CENTRO, TRANCOSO = _loja(SET.ARRAIAL), _loja(SET.CENTRO), _loja(SET.TRANCOSO)
LOJAS = [("ARRAIAL", ARRAIAL), ("CENTRO", CENTRO), ("TRANCOSO", TRANCOSO)]

COLS = [("FUNCIONÁRIO", "nome", 15), ("SALÁRIO", "sal", 12), ("HORAS EXTRAS", "he", 12),
        ("AD. NOTURNO", "notu", 12), ("COMISSÃO", "com", 12), ("DSR", "dsr", 11),
        ("AUX. GERÊNCIA", "aux", 12), ("PRÊMIO META CX", "meta", 12), ("INCENTIVOS", "inc", 12),
        ("PRÊMIOS / BONIFIC.", "bonif", 13), ("DSR S/ HORA EXTRA", "dsrhe", 12),
        ("SAL. FAMÍLIA", "salfam", 11), ("INSUF. SALDO", "insuf", 11),
        ("TOTAL PROVENTOS", "prov", 14), ("ADIANT. INCENT.", "adiant_inc", 13),
        ("ADIANTAMENTO SALARIAL / VALES", "adiant_sal", 15),
        ("VALE TRANSP.", "vt6", 12), ("INSS", "inss", 12), ("IRRF", "irrf", 12),
        ("TOTAL DESCONTOS", "desc", 14),
        ("LÍQUIDO PARCIAL", "liq", 14),
        ("FÉRIAS PAGAS À PARTE (contas a pagar)", "ferias", 16), ("OBSERVAÇÃO", "obs", 70)]

wb = openpyxl.Workbook()
ws = wb.active; ws.title = "RESUMO 3 LOJAS AGO.26"
ws.sheet_view.showGridLines = False
for i, (_, _, w) in enumerate(COLS, 1):
    ws.column_dimensions[get_column_letter(i)].width = w
N = len(COLS)
ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=N)
c = ws.cell(1, 1, "RESUMO PARA A CONTABILIDADE — TRÊS LOJAS")
c.font = font(14, True, "FFFFFF"); c.fill = FILL_TIT
c.alignment = Alignment(horizontal="center", vertical="center")
ws.row_dimensions[1].height = 26
ws.merge_cells(start_row=2, start_column=1, end_row=2, end_column=N)
c = ws.cell(2, 1, "Farmácia Tropical · competência SETEMBRO/2026 (01/09 a 30/09) · pagamento em 05/10/2026 · PARCIAL: sem comissões, INSS e IRRF")
c.font = font(10, True, "1F3864"); c.alignment = Alignment(horizontal="center")
r = 4
for i, (h, _, _) in enumerate(COLS, 1):
    c = ws.cell(r, i, h); c.font = font(9, True, "FFFFFF"); c.fill = FILL_HDR
    c.alignment = Alignment(horizontal="center", wrap_text=True); c.border = BOX
ws.row_dimensions[r].height = 30
r += 1
ws.freeze_panes = "A5"
ini_geral = r
linhas_loja = []
for loja, emps in LOJAS:
    for i in range(1, N + 1):
        ws.cell(r, i).fill = FILL_SEC
    c = ws.cell(r, 1, f"LOJA {loja}"); c.font = font(11, True, "1F3864")
    ws.row_dimensions[r].height = 19
    r += 1
    ini = r
    KEYS = [k for _, k, _ in COLS]
    L = {k: get_column_letter(i) for i, k in enumerate(KEYS, 1)}
    for e in emps:
        for i, (_, k, _) in enumerate(COLS, 1):
            if k == "prov":
                v = f"=SUM({L['sal']}{r}:{L['insuf']}{r})"
            elif k == "desc":
                v = f"=SUM({L['adiant_inc']}{r}:{L['irrf']}{r})"
            elif k == "liq":
                v = f"={L['prov']}{r}+{L['desc']}{r}"

            else:
                v = e[k]
            c = ws.cell(r, i, v)
            c.border = BOX
            if k == "nome":
                c.font = font(10, True)
            elif k == "obs":
                c.font = font(9, it=True)
                c.alignment = Alignment(wrap_text=True, vertical="center")
            else:
                c.number_format = MONEY
                c.font = font(10, True if k == "liq" else False)
                if k == "liq":
                    c.fill = FILL_LIQ
                elif k == "dif":
                    c.fill = FILL_TOT; c.font = font(10, True)
                elif k in ("prov", "desc"):
                    c.fill = FILL_TOT
                elif k in ("ferias",):
                    c.font = font(10, False, "0000FF"); c.fill = FILL_IN
        r += 1
    fim = r - 1
    linhas_loja.append((loja, ini, fim))
    ws.cell(r, 1, f"TOTAL {loja}").font = font(10, True)
    for i, (_, k, _) in enumerate(COLS, 1):
        if k in ("nome", "obs"):
            continue
        L = get_column_letter(i)
        c = ws.cell(r, i, f"=SUM({L}{ini}:{L}{fim})")
        c.number_format = MONEY; c.font = font(10, True)
    for i in range(1, N + 1):
        ws.cell(r, i).fill = FILL_TOT; ws.cell(r, i).border = BOX
    linhas_loja[-1] = (loja, ini, fim, r)
    r += 1
r += 1
ws.cell(r, 1, "TOTAL GERAL — 3 LOJAS").font = font(11, True)
for i, (_, k, _) in enumerate(COLS, 1):
    if k in ("nome", "obs"):
        continue
    L = get_column_letter(i)
    partes = "+".join(f"{L}{t}" for _, _, _, t in linhas_loja)
    c = ws.cell(r, i, f"={partes}")
    c.number_format = MONEY; c.font = font(11, True)
for i in range(1, N + 1):
    ws.cell(r, i).fill = FILL_LIQ; ws.cell(r, i).border = BOX
ws.row_dimensions[r].height = 22
r += 2
avisos = [
 "ESTE RESUMO AINDA É PARCIAL — falta, antes de mandar à Betel:",
 "• Comissões e incentivos de setembro: extrair o relatório do InovaFarma (01/09 a 30/09) das 3 lojas. Só o AGNOR já aparece com comissão, porque o fixo de R$ 2.000,00 não depende do apurado.",
 "• Prêmios do programa de metas de setembro (Metas 1 a 4): as metas de cada um já estão na aba METAS SET.26 de cada loja; os prêmios saem sozinhos quando a venda realizada, o pré-vencido vendido, a Meta 2 e a perfumaria da loja forem lançados.",
 "• Vales adiantados, convênio e faltas — por enquanto só o adiantamento do GENECIR está lançado.",
 "• Ponto de setembro: horas extras e quem mais trabalhou nos feriados de 07/09 e 08/09.",
 "• INSS e IRRF são calculados pela Betel.",
 "• Diárias dos folguistas SERGIO e ANA CELIA (contas a pagar, fora do holerite).",
 "",
 "CRITÉRIOS DE SETEMBRO:",
 "• DSR: 4 domingos + feriados de 07/09 (Independência) e 08/09 (municipal) = 6 ÷ 24 dias úteis = fator 0,25.",
 "• Feriado trabalhado sem folga = 1 salário-dia a mais (R$ 54,03 no salário de R$ 1.621,00).",
 "• TRANCOSO paga o dobro da comissão apurada; incentivos simples.",
 "• Férias pagas à parte em setembro: RENALDO (Centro) e VALDICK (Trancoso), de 01 a 30/09 — lançar o valor do recibo na coluna FÉRIAS PAGAS À PARTE.",
]
for t in avisos:
    ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=N)
    c = ws.cell(r, 1, t)
    c.font = font(9, t.endswith(":"), "C00000" if t.startswith("O QUE") else "000000")
    c.alignment = Alignment(wrap_text=True, vertical="center")
    if t.startswith("O QUE"):
        c.fill = FILL_ALERT
    r += 1
ws.auto_filter.ref = f"A4:{get_column_letter(N)}{linhas_loja[-1][3]}"
wb.save(OUT)
print("ok", OUT)
