import os
from datetime import datetime
from urllib.parse import urlparse

from openpyxl import Workbook, load_workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from playwright.sync_api import sync_playwright

from login import login, URL as LOGIN_URL

_origin = urlparse(LOGIN_URL)
LOANS_URL = f"{_origin.scheme}://{_origin.netloc}/loans"

DRY_RUN = True
VALOR_LIMITE = 10_000.00
MAX_PROPOSTAS = 5  # quantas propostas da lista processar nesta simulacao


def parse_valor_brl(texto):
    texto = texto.strip().replace("R$", "").strip()
    texto = texto.replace(".", "").replace(",", ".")
    return float(texto)


def _cartao_da_secao(page, titulo):
    secao = page.locator(
        f'xpath=//div[contains(@class,"text-xl") and contains(@class,"mb-2") '
        f'and normalize-space(.)="{titulo}"]'
    )
    return secao.locator('xpath=following-sibling::div[contains(@class,"ajin-card")][1]')


def _campo(cartao, rotulo):
    campo = cartao.locator(
        f'xpath=.//*[contains(@class,"ajin-label")]'
        f'[normalize-space(text())="{rotulo}"]'
        f'/following-sibling::*[contains(@class,"ajin-value")][1]'
    )
    return campo.inner_text().strip()


def ler_proposta(page):
    cartao_proposta = _cartao_da_secao(page, "Proposta")
    contrato = _campo(cartao_proposta, "Contrato")
    valor = parse_valor_brl(_campo(cartao_proposta, "Valor do Contrato"))
    liquido = parse_valor_brl(_campo(cartao_proposta, "Líquido"))

    cartao_pessoais = _cartao_da_secao(page, "Dados Pessoais")
    nome = _campo(cartao_pessoais, "Name")

    cartao_log = _cartao_da_secao(page, "Log do Registro")
    data_proposta = _campo(cartao_log, "Data de Cadastro:")

    return contrato, nome, valor, liquido, data_proposta


def aprovar_proposta_real(page):
    """NAO TESTADO. So deve ser chamado com DRY_RUN=False, apos validacao manual."""
    page.click('button:has-text("Ações")')
    page.get_by_text("Aprovação Supervisor", exact=True).click()
    page.fill("textarea[name='note']", "Aprovado via RPA")
    # TODO: mapear o botao final de confirmar/enviar - etapa nunca executada.


def processar_propostas(page):
    resultados = []

    page.goto(LOANS_URL)
    page.wait_for_load_state("networkidle")
    page.wait_for_timeout(1000)

    linhas = page.locator("table.app-table-search tbody tr.cursor-pointer")
    total = linhas.count()
    n = min(total, MAX_PROPOSTAS)
    print(f"Processando {n} de {total} propostas listadas (DRY_RUN={DRY_RUN})...")

    for i in range(n):
        page.goto(LOANS_URL)
        page.wait_for_load_state("networkidle")
        page.wait_for_timeout(800)

        linha = page.locator("table.app-table-search tbody tr.cursor-pointer").nth(i)
        linha.locator('button[aria-haspopup="menu"]').first.click()
        page.wait_for_timeout(400)
        page.get_by_role("menuitem", name="Visualizar", exact=True).click()
        page.wait_for_load_state("networkidle")
        page.wait_for_timeout(1000)

        try:
            contrato, nome, valor, liquido, data_proposta = ler_proposta(page)
        except Exception as e:
            print(f"[{i}] Nao foi possivel ler a proposta: {e}")
            resultados.append(
                {
                    "contrato": "?",
                    "nome": "",
                    "valor": None,
                    "liquido": None,
                    "data_proposta": "",
                    "data_aprovacao_supervisor": "",
                    "aprovado": False,
                }
            )
            continue

        aprovado = valor <= VALOR_LIMITE
        decisao = "APROVARIA" if aprovado else "PULA (valor > limite)"
        print(f"[{i}] Contrato {contrato} - Valor do Contrato: R$ {valor:,.2f} -> {decisao}")

        data_aprovacao_supervisor = ""
        if aprovado:

            data_aprovacao_supervisor = datetime.now().strftime("%d/%m/%Y %H:%M")
            if DRY_RUN:
                print("    [DRY-RUN] nao clicou em Acoes/Aprovacao Supervisor.")
            else:
                aprovar_proposta_real(page)

        resultados.append(
            {
                "contrato": contrato,
                "nome": nome,
                "valor": valor,
                "liquido": liquido,
                "data_proposta": data_proposta,
                "data_aprovacao_supervisor": data_aprovacao_supervisor,
                "aprovado": aprovado,
            }
        )

    return resultados


HEADER_FILL = PatternFill("solid", fgColor="1F4E78")
HEADER_FONT = Font(color="FFFFFF", bold=True, size=11)
APROVADO_FILL = PatternFill("solid", fgColor="C6EFCE")
APROVADO_FONT = Font(color="006100")
SEM_FILL = PatternFill(fill_type=None)
FONTE_PADRAO = Font(color="000000")
THIN_BORDER = Border(*(Side(style="thin", color="D9D9D9"),) * 4)
MOEDA_FORMATO = '"R$" #,##0.00'

COLUNAS_RELATORIO = [
    ("Código do Contrato", 22),
    ("Nome da Pessoa", 34),
    ("Valor Líquido", 18),
    ("Data e Horário da Proposta", 24),
    ("Data e Horário da Aprovação do Supervisor", 30),
    ("Decisão", 22),
]


def _ano_da_proposta(item):
    """Ano usado para escolher a aba: ano da 'Data de Cadastro' real da proposta."""
    try:
        return datetime.strptime(item.get("data_proposta", ""), "%d/%m/%Y %H:%M").year
    except (TypeError, ValueError):
        return datetime.now().year


def _obter_ou_criar_aba(wb, nome_aba):
    if nome_aba in wb.sheetnames:
        return wb[nome_aba]

    # Reaproveita a aba padrao "Sheet" vazia da primeira criacao do workbook.
    if wb.sheetnames == ["Sheet"] and wb["Sheet"].max_row == 1 and wb["Sheet"]["A1"].value is None:
        ws = wb["Sheet"]
        ws.title = nome_aba
    else:
        ws = wb.create_sheet(nome_aba)

    for col_idx, (titulo, largura) in enumerate(COLUNAS_RELATORIO, start=1):
        celula = ws.cell(row=1, column=col_idx, value=titulo)
        celula.fill = HEADER_FILL
        celula.font = HEADER_FONT
        celula.alignment = Alignment(horizontal="center", vertical="center")
        celula.border = THIN_BORDER
        ws.column_dimensions[get_column_letter(col_idx)].width = largura
    ws.freeze_panes = "A2"
    return ws


def _escrever_linha(ws, row_idx, item):
    aprovado = item.get("aprovado", False)
    decisao_texto = "Aprovado" if aprovado else "Não Aprovado"

    valores = [
        item.get("contrato") or "-",
        item.get("nome") or "-",
        item.get("liquido"),
        item.get("data_proposta") or "-",
        item.get("data_aprovacao_supervisor") or "",
        decisao_texto,
    ]
    for col_idx, valor in enumerate(valores, start=1):
        celula = ws.cell(row=row_idx, column=col_idx, value=valor)
        celula.border = THIN_BORDER
        if col_idx == 3 and isinstance(valor, (int, float)):
            celula.number_format = MOEDA_FORMATO
        if col_idx in (4, 5, 6):
            celula.alignment = Alignment(horizontal="center")
        celula.fill = APROVADO_FILL if aprovado else SEM_FILL
        celula.font = APROVADO_FONT if aprovado else FONTE_PADRAO


def gerar_relatorio(resultados, caminho):

    if os.path.exists(caminho):
        wb = load_workbook(caminho)
    else:
        wb = Workbook()

    for item in resultados:
        nome_aba = str(_ano_da_proposta(item))
        ws = _obter_ou_criar_aba(wb, nome_aba)

        contrato = item.get("contrato") or "-"
        row_idx = None
        if contrato != "-":
            for r in range(2, ws.max_row + 1):
                if ws.cell(row=r, column=1).value == contrato:
                    row_idx = r
                    break
        if row_idx is None:
            row_idx = ws.max_row + 1

        _escrever_linha(ws, row_idx, item)

    wb._sheets.sort(key=lambda ws: ws.title)
    wb.save(caminho)


if __name__ == "__main__":
    os.makedirs("screenshots", exist_ok=True)
    with sync_playwright() as p:
        browser, page = login(p)
        resultados = processar_propostas(page)

        caminho_relatorio = "resultado_simulacao.xlsx"
        gerar_relatorio(resultados, caminho_relatorio)

        print(f"\nRelatorio salvo em {caminho_relatorio} ({len(resultados)} propostas).")
        browser.close()
