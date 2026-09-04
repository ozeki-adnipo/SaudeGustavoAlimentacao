#!/usr/bin/env python3
"""Gera/regenera 'plano-alimentar-pos-infarto.xlsx' a partir deste script.

- Abas 1 e 2 (texto de orientação) ficam definidas diretamente aqui.
- Aba 3 ("Alimentos e Quantidades") é montada a partir de `dados_alimentos.json`,
  que é a fonte de verdade da lista de alimentos.

Para acrescentar um alimento novo: edite `dados_alimentos.json` (adicione um
objeto na lista, respeitando a categoria) e rode este script de novo:

    python3 gerar_planilha.py

Isso recria o arquivo .xlsx inteiro, sempre com a mesma formatação. É assim
que a skill `adicionar-alimento` funciona por baixo dos panos.

Sem fórmulas — é uma planilha de referência/consulta, não um modelo de cálculo.
"""
import json
import os

import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

HERE = os.path.dirname(os.path.abspath(__file__))
JSON_PATH = os.path.join(HERE, "dados_alimentos.json")
OUT_PATH = os.path.join(HERE, "plano-alimentar-pos-infarto.xlsx")

FONT_NAME = "Arial"

# ---- paleta ----
COLOR_TITLE_BG = "1F4E5F"      # azul petróleo escuro (títulos de aba/seção)
COLOR_TITLE_FG = "FFFFFF"
COLOR_SECTION_BG = "D9E6EC"    # azul bem claro (linhas de seção)
COLOR_WARN_BG = "FCE4D6"       # laranja claro (avisos/pendências)
COLOR_LIBERADO = "C6E0B4"      # verde
COLOR_MODERAR = "FFE699"       # amarelo
COLOR_EVITAR = "F4B6B6"        # vermelho claro
COLOR_CATEGORY_BG = "1F4E5F"
COLOR_CATEGORY_FG = "FFFFFF"

CLASS_COLORS = {
    "Liberado": COLOR_LIBERADO,
    "Moderar": COLOR_MODERAR,
    "Evitar": COLOR_EVITAR,
}

THIN = Side(style="thin", color="B7B7B7")
BORDER_ALL = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)

# Ordem fixa das categorias na aba 3. Uma categoria nova (que não esteja
# nesta lista) é acrescentada automaticamente ao final, na ordem em que
# aparecer pela primeira vez no JSON.
CATEGORY_ORDER = [
    "Frutas",
    "Grãos e Cereais",
    "Proteínas",
    "Laticínios",
    "Bebidas",
    "Doces e Sobremesas",
    "Gorduras e Temperos",
    "Vegetais e Verduras",
]


def style_title(ws, text, span, row=1, height=26):
    ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=span)
    cell = ws.cell(row=row, column=1, value=text)
    cell.font = Font(name=FONT_NAME, size=14, bold=True, color=COLOR_TITLE_FG)
    cell.fill = PatternFill("solid", fgColor=COLOR_TITLE_BG)
    cell.alignment = Alignment(horizontal="left", vertical="center", indent=1)
    ws.row_dimensions[row].height = height


def header_row(ws, row, headers, widths=None):
    for i, h in enumerate(headers, start=1):
        cell = ws.cell(row=row, column=i, value=h)
        cell.font = Font(name=FONT_NAME, size=11, bold=True, color=COLOR_TITLE_FG)
        cell.fill = PatternFill("solid", fgColor=COLOR_TITLE_BG)
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.border = BORDER_ALL
    if widths:
        for i, w in enumerate(widths, start=1):
            ws.column_dimensions[get_column_letter(i)].width = w


def build_aba1(wb):
    ws1 = wb.active
    ws1.title = "Cuidados na Alimentação"
    ws1.sheet_view.showGridLines = False

    style_title(ws1, "Cuidados na Alimentação — pós-infarto (25/08/2026)", span=2)
    header_row(ws1, 2, ["Seção", "Orientação"], widths=[26, 95])
    ws1.freeze_panes = "A3"

    rows_aba1 = [
        ("Perfil de risco atual",
         "Homem, 48 anos, IAM em 25/08/2026 (3 procedimentos, FE preservada 72% com "
         "hipocinesia segmentar leve). Dislipidemia mista: LDL 138 (meta pós-IAM <50), "
         "HDL 34 (baixo), triglicerídeos 343 (muito alto), colesterol total 222 (alto). "
         "Placa carotídea leve (1–15%) já mostra aterosclerose além do evento agudo."),
        ("Reduzir",
         "Gordura saturada e gordura trans (frituras, embutidos, carnes gordas, manteiga, "
         "produtos industrializados com 'gordura vegetal hidrogenada')."),
        ("Reduzir",
         "Açúcar e carboidrato refinado — prioridade alta por causa dos triglicerídeos "
         "muito altos (343 mg/dL): refrigerante comum, doces, pão/arroz/massa branca em excesso."),
        ("Reduzir",
         "Sódio — evitar sal em excesso, temperos prontos, embutidos, enlatados e "
         "fast-food; ler rótulo (mg de sódio por porção)."),
        ("Reduzir",
         "Álcool — evitar ou reduzir ao mínimo; álcool piora diretamente os triglicerídeos."),
        ("Priorizar",
         "Fibras: grãos integrais, leguminosas (feijão, lentilha, grão-de-bico), vegetais "
         "e frutas variadas."),
        ("Priorizar",
         "Ômega-3: peixes (preferir 2-3x/semana), azeite de oliva extra virgem, oleaginosas "
         "em pequena quantidade (castanhas, nozes)."),
        ("Priorizar",
         "Proteínas magras: frango sem pele, peixe, ovos, leguminosas; carnes vermelhas com "
         "moderação."),
        ("Atenção especial — triglicerídeos",
         "343 mg/dL é muito alto: o maior impacto individual costuma vir de cortar açúcar "
         "e álcool, não só gordura. Evitar refrigerante comum, suco industrializado, doces "
         "e bebida alcoólica quase por completo até reavaliação médica."),
        ("Atenção especial — LDL",
         "Meta pós-IAM é LDL <50 mg/dL (bem mais rígida que a meta geral da população). "
         "Reforça a importância de manter o cuidado mesmo quando o valor parecer "
         "'razoável' pelos padrões gerais."),
        ("Aviso pendente — função renal",
         "Há suspeita ainda não confirmada de proteinúria e duas elevações de creatinina/TFG "
         "durante a internação (TFG caiu para 57 na véspera da alta). Enquanto não houver "
         "confirmação médica, o plano NÃO aplica restrição renal (sódio/potássio/proteína) — "
         "apenas o cuidado cardiovascular abaixo. Revisar esta aba assim que o médico confirmar "
         "ou descartar o problema renal."),
        ("Aviso pendente — medicação",
         "Lista de medicações da alta ainda não preenchida. Existem interações conhecidas "
         "entre alimentos e medicações comuns pós-IAM (ex.: anticoagulantes como varfarina "
         "e vegetais verde-escuros ricos em vitamina K). Revisar esta planilha assim que a "
         "lista de medicações estiver disponível."),
        ("Leitura de rótulos",
         "Ao comprar um produto industrializado, olhar sempre: sódio por porção, açúcares "
         "totais/adicionados, e a lista de ingredientes procurando 'gordura hidrogenada' ou "
         "'gordura trans'. Prefira produtos com menos ingredientes e sem açúcar adicionado."),
    ]

    r = 3
    for secao, texto in rows_aba1:
        is_warning = secao.startswith("Aviso pendente")
        c1 = ws1.cell(row=r, column=1, value=secao)
        c2 = ws1.cell(row=r, column=2, value=texto)
        fill = PatternFill("solid", fgColor=COLOR_WARN_BG if is_warning else COLOR_SECTION_BG)
        c1.font = Font(name=FONT_NAME, size=10.5, bold=True)
        c1.alignment = Alignment(horizontal="left", vertical="top", wrap_text=True)
        c1.fill = fill
        c1.border = BORDER_ALL
        c2.font = Font(name=FONT_NAME, size=10.5)
        c2.alignment = Alignment(horizontal="left", vertical="top", wrap_text=True)
        c2.border = BORDER_ALL
        ws1.row_dimensions[r].height = 46
        r += 1


def build_aba2(wb):
    ws2 = wb.create_sheet("Plano Geral - Restaurante")
    ws2.sheet_view.showGridLines = False

    style_title(ws2, "Plano Geral — o que comer, inclusive fora de casa", span=2)
    header_row(ws2, 2, ["Seção", "Conteúdo"], widths=[26, 95])
    ws2.freeze_panes = "A3"

    rows_aba2 = [
        ("Modelo de prato",
         "1/2 do prato: vegetais e salada (à vontade, tempero simples com azeite/limão). "
         "1/4 do prato: proteína magra grelhada/assada (peixe, frango sem pele, ovo, "
         "leguminosas). 1/4 do prato: carboidrato integral (arroz integral, batata-doce, "
         "mandioca, pão integral) em porção moderada."),
        ("O que pedir",
         "Grelhado, assado ou cozido — sem manteiga extra ou molho cremoso. Peça o molho "
         "à parte para controlar a quantidade. Troque frituras por grelhados/assados. "
         "Troque refrigerante comum por água ou refrigerante zero. Sobremesa só ocasional "
         "e em porção pequena (dividir com alguém, se possível)."),
        ("O que evitar",
         "Frituras (empanados, milanesas), molhos à base de creme de leite/manteiga, "
         "embutidos (bacon, linguiça), pratos 'à parmegiana' ou gratinados com muito "
         "queijo, refrigerante comum, sobremesas grandes ou muito açucaradas."),
        ("Perguntas úteis ao garçom",
         "\"Esse prato pode vir grelhado, sem manteiga extra?\" · \"O molho leva creme de "
         "leite ou manteiga?\" · \"Dá para trocar a guarnição (fritas/farofa) por salada ou "
         "legumes?\" · \"Tem opção de refrigerante zero ou água com gás?\""),
        ("Exemplos de pratos que funcionam",
         "Peixe grelhado + legumes salteados + arroz integral. Frango grelhado sem pele + "
         "salada variada. Prato de grãos/leguminosas (feijoada de grão-de-bico, saladas com "
         "lentilha) + vegetais. Omelete de claras/ovo com legumes + salada."),
        ("Em festas/eventos",
         "Priorize proteína grelhada e salada primeiro, para 'preencher' o prato com "
         "opções seguras. Álcool: evitar ou limitar ao mínimo por causa dos "
         "triglicerídeos. Doces: só uma porção pequena, ocasionalmente."),
    ]

    r = 3
    for secao, texto in rows_aba2:
        c1 = ws2.cell(row=r, column=1, value=secao)
        c2 = ws2.cell(row=r, column=2, value=texto)
        c1.font = Font(name=FONT_NAME, size=10.5, bold=True)
        c1.alignment = Alignment(horizontal="left", vertical="top", wrap_text=True)
        c1.fill = PatternFill("solid", fgColor=COLOR_SECTION_BG)
        c1.border = BORDER_ALL
        c2.font = Font(name=FONT_NAME, size=10.5)
        c2.alignment = Alignment(horizontal="left", vertical="top", wrap_text=True)
        c2.border = BORDER_ALL
        ws2.row_dimensions[r].height = 60
        r += 1


def load_alimentos():
    with open(JSON_PATH, encoding="utf-8") as f:
        alimentos = json.load(f)

    # ordena por categoria (seguindo CATEGORY_ORDER, categorias novas vão ao
    # final na ordem de primeira aparição) e mantém a ordem original dentro
    # de cada categoria.
    seen_order = list(CATEGORY_ORDER)
    for item in alimentos:
        if item["categoria"] not in seen_order:
            seen_order.append(item["categoria"])

    alimentos_ordenados = sorted(
        alimentos, key=lambda item: seen_order.index(item["categoria"])
    )
    return alimentos_ordenados


def build_aba3(wb):
    ws3 = wb.create_sheet("Alimentos e Quantidades")
    ws3.sheet_view.showGridLines = False

    style_title(ws3, "Alimentos e Quantidades — lista de referência", span=6)

    ws3.row_dimensions[2].height = 20
    legend = [
        (2, "Liberado", COLOR_LIBERADO),
        (3, "Moderar", COLOR_MODERAR),
        (4, "Evitar", COLOR_EVITAR),
    ]
    for col, text, color in legend:
        cell = ws3.cell(row=2, column=col, value=text)
        cell.font = Font(name=FONT_NAME, size=9, bold=True)
        cell.fill = PatternFill("solid", fgColor=color)
        cell.alignment = Alignment(horizontal="center", vertical="center")
        cell.border = BORDER_ALL

    headers3 = ["Categoria", "Alimento", "Porção", "Frequência", "Classificação", "Observação"]
    widths3 = [16, 26, 16, 20, 15, 46]
    header_row(ws3, 3, headers3, widths=widths3)
    ws3.freeze_panes = "A4"

    alimentos = load_alimentos()

    r = 4
    current_cat = None
    for item in alimentos:
        categoria = item["categoria"]
        if categoria != current_cat:
            ws3.merge_cells(start_row=r, start_column=1, end_row=r, end_column=6)
            cat_cell = ws3.cell(row=r, column=1, value=categoria)
            cat_cell.font = Font(name=FONT_NAME, size=10.5, bold=True, color=COLOR_CATEGORY_FG)
            cat_cell.fill = PatternFill("solid", fgColor=COLOR_CATEGORY_BG)
            cat_cell.alignment = Alignment(horizontal="left", vertical="center", indent=1)
            for col in range(1, 7):
                ws3.cell(row=r, column=col).border = BORDER_ALL
            ws3.row_dimensions[r].height = 20
            current_cat = categoria
            r += 1

        classif = item["classificacao"]
        values = [None, item["alimento"], item["porcao"], item["frequencia"], classif, item["observacao"]]
        for col, val in enumerate(values, start=1):
            if col == 1:
                continue
            cell = ws3.cell(row=r, column=col, value=val)
            cell.font = Font(name=FONT_NAME, size=10.5, bold=(col == 5))
            cell.alignment = Alignment(
                horizontal="center" if col in (3, 4, 5) else "left",
                vertical="center", wrap_text=True, indent=(1 if col in (2, 6) else 0)
            )
            cell.border = BORDER_ALL
            if col == 5:
                cell.fill = PatternFill("solid", fgColor=CLASS_COLORS.get(classif, "FFFFFF"))
        ws3.row_dimensions[r].height = 30
        r += 1

    r += 1
    ws3.merge_cells(start_row=r, start_column=1, end_row=r, end_column=6)
    note = ws3.cell(row=r, column=1,
                    value="Lista com os alimentos acrescentados até agora — não é uma lista "
                          "completa. Use a skill 'adicionar-alimento' (ex.: \"adicione "
                          "sucrilhos\") para ir acrescentando novos itens aos poucos.")
    note.font = Font(name=FONT_NAME, size=9.5, italic=True)
    note.alignment = Alignment(horizontal="left", vertical="center", wrap_text=True)
    ws3.row_dimensions[r].height = 30


def main():
    wb = openpyxl.Workbook()
    build_aba1(wb)
    build_aba2(wb)
    build_aba3(wb)
    wb.save(OUT_PATH)
    print("Salvo em", OUT_PATH)


if __name__ == "__main__":
    main()
