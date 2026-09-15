"""
RPA - QualiBank: simulacao de liberacao de propostas.

Fluxo:
 1. Login (login.py) - ja fecha o popup de notificacao ao entrar.
 2. Abre a lista de emprestimos (/loans).
 3. Para cada proposta da lista: abre o menu de tres pontos, clica em
    "Visualizar", le o "Valor do Contrato" na secao "Proposta" e decide:
      - valor <= R$ 10.000,00  -> aprovaria
      - valor >  R$ 10.000,00  -> pula para a proxima

MODO DE TESTE (DRY_RUN=True, padrao): a decisao de aprovar e SO REGISTRADA
em log/planilha (resultado_simulacao.xlsx). O robo NAO clica em "Acoes"
nem em "Aprovacao Supervisor" e NAO aprova nenhuma proposta de verdade.

A coluna "Data e Horario da Aprovacao do Supervisor" do relatorio NAO vem
do sistema (nenhuma proposta foi aprovada de verdade) - e o horario em
que o robo simulou a decisao de aprovar, preenchido so quando DRY_RUN
decide "Aprovado". A "Data e Horario da Proposta" vem da secao real
"Log do Registro" -> "Data de Cadastro" de cada proposta.

A funcao aprovar_proposta_real() faz o clique de aprovacao real (Acoes ->
Aprovacao Supervisor -> observacao "Aprovado via RPA"), mas foi escrita
apenas a partir da especificacao recebida - NUNCA foi executada nem
testada contra o site, porque isso aprovaria uma proposta real em
producao. Falta inclusive mapear o botao final de confirmar/enviar.
Antes de rodar com DRY_RUN=False, valide esse fluxo manualmente (de
preferencia em ambiente de homologacao).
"""

import os
from datetime import datetime
from urllib.parse import urlparse

from openpyxl import Workbook
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
            # Nao existe aprovacao real do supervisor (DRY_RUN sempre ativo aqui).
            # Este horario e o momento em que O ROBO SIMULOU a aprovacao, nao uma
            # aprovacao de verdade registrada no sistema.
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
THIN_BORDER = Border(*(Side(style="thin", color="D9D9D9"),) * 4)
MOEDA_FORMATO = '"R$" #,##0.00'


def gerar_relatorio(resultados, caminho):
    wb = Workbook()
    ws = wb.active
    ws.title = "Simulacao de Propostas"

    colunas = [
        ("Código do Contrato", 22),
        ("Nome da Pessoa", 34),
        ("Valor Contrato", 18),
        ("Valor Líquido", 18),
        ("Data e Horário da Proposta", 24),
        ("Data e Horário da Aprovação do Supervisor", 30),
        ("Decisão", 22),
    ]

    for col_idx, (titulo, largura) in enumerate(colunas, start=1):
        celula = ws.cell(row=1, column=col_idx, value=titulo)
        celula.fill = HEADER_FILL
        celula.font = HEADER_FONT
        celula.alignment = Alignment(horizontal="center", vertical="center")
        celula.border = THIN_BORDER
        ws.column_dimensions[get_column_letter(col_idx)].width = largura
    ws.freeze_panes = "A2"

    for row_idx, item in enumerate(resultados, start=2):
        aprovado = item.get("aprovado", False)
        decisao_texto = "Aprovado" if aprovado else "Não Aprovado"

        valores = [
            item.get("contrato") or "-",
            item.get("nome") or "-",
            item.get("valor"),
            item.get("liquido"),
            item.get("data_proposta") or "-",
            item.get("data_aprovacao_supervisor") or "",
            decisao_texto,
        ]
        for col_idx, valor in enumerate(valores, start=1):
            celula = ws.cell(row=row_idx, column=col_idx, value=valor)
            celula.border = THIN_BORDER
            if col_idx in (3, 4) and isinstance(valor, (int, float)):
                celula.number_format = MOEDA_FORMATO
            if col_idx in (5, 6, 7):
                celula.alignment = Alignment(horizontal="center")
            if aprovado:
                celula.fill = APROVADO_FILL
                celula.font = APROVADO_FONT

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
