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
em log/CSV. O robo NAO clica em "Acoes" nem em "Aprovacao Supervisor" e
NAO aprova nenhuma proposta de verdade.

A funcao aprovar_proposta_real() faz o clique de aprovacao real (Acoes ->
Aprovacao Supervisor -> observacao "Aprovado via RPA"), mas foi escrita
apenas a partir da especificacao recebida - NUNCA foi executada nem
testada contra o site, porque isso aprovaria uma proposta real em
producao. Falta inclusive mapear o botao final de confirmar/enviar.
Antes de rodar com DRY_RUN=False, valide esse fluxo manualmente (de
preferencia em ambiente de homologacao).
"""

import csv
import os
from urllib.parse import urlparse

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
        f'xpath=.//div[contains(@class,"ajin-label")]'
        f'[normalize-space(text())="{rotulo}"]'
        f'/following-sibling::div[contains(@class,"ajin-value")][1]'
    )
    return campo.inner_text().strip()


def ler_proposta(page):
    cartao = _cartao_da_secao(page, "Proposta")
    contrato = _campo(cartao, "Contrato")
    valor = parse_valor_brl(_campo(cartao, "Valor do Contrato"))
    return contrato, valor


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
            contrato, valor = ler_proposta(page)
        except Exception as e:
            print(f"[{i}] Nao foi possivel ler a proposta: {e}")
            resultados.append({"contrato": "?", "valor": "", "decisao": f"ERRO: {e}"})
            continue

        decisao = "APROVARIA" if valor <= VALOR_LIMITE else "PULA (valor > limite)"
        print(f"[{i}] Contrato {contrato} - Valor do Contrato: R$ {valor:,.2f} -> {decisao}")

        if decisao == "APROVARIA":
            if DRY_RUN:
                print("    [DRY-RUN] nao clicou em Acoes/Aprovacao Supervisor.")
            else:
                aprovar_proposta_real(page)

        resultados.append({"contrato": contrato, "valor": valor, "decisao": decisao})

    return resultados


if __name__ == "__main__":
    os.makedirs("screenshots", exist_ok=True)
    with sync_playwright() as p:
        browser, page = login(p)
        resultados = processar_propostas(page)

        with open("resultado_simulacao.csv", "w", newline="", encoding="utf-8") as f:
            writer = csv.DictWriter(f, fieldnames=["contrato", "valor", "decisao"])
            writer.writeheader()
            writer.writerows(resultados)

        print(f"\nResultado salvo em resultado_simulacao.csv ({len(resultados)} propostas).")
        browser.close()
