import os
import re
from datetime import datetime
from urllib.parse import urlparse

from openpyxl import Workbook, load_workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from playwright.sync_api import sync_playwright

from login import login, URL as LOGIN_URL

_origin = urlparse(LOGIN_URL)
LOANS_URL = f"{_origin.scheme}://{_origin.netloc}/loans"

PASTA_PROJETO = os.path.dirname(os.path.abspath(__file__))
CAMINHO_RELATORIO_SIMULACAO = os.path.join(PASTA_PROJETO, "resultado_simulacao.xlsx")
CAMINHO_RELATORIO_REAL = os.path.join(PASTA_PROJETO, "resultado_aprovacoes.xlsx")

DRY_RUN = True
VALOR_LIMITE = 10_000.00
MAX_PROPOSTAS = None  # None = extrai todas as propostas da lista
ITENS_POR_PAGINA = 50


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


def ler_data_aprovacao_promotora(linha):
    """Le a data/hora do status 'Aguardando Aprovação Promotora' direto na linha da lista."""
    nota = linha.locator('ajin-status-label[name="operationStatus"] .ajin-note').first
    if nota.count() == 0:
        return ""
    return nota.inner_text().strip()


def ler_loja(linha):
    """Le o nome da Loja direto na linha da lista (coluna Registro).
    Nao existe campo equivalente na pagina de detalhe da proposta."""
    valor = linha.locator(
        'xpath=.//span[contains(@class,"ajin-label")][normalize-space(text())="Loja:"]'
        '/following-sibling::span[contains(@class,"ajin-value")][1]'
    )
    if valor.count() == 0:
        return ""
    return valor.inner_text().strip()


def aprovar_proposta_real(page, contrato):
    """So deve ser chamado com DRY_RUN=False, apos validacao manual.
    Se a aprovacao em si (Acoes -> Aprovacao Supervisor -> Confirmar) falhar,
    a excecao propaga para o chamador tratar. Uma falha so nos prints de
    evidencia (depois do Confirmar) NAO deve ser reportada como aprovacao
    falha, senao um contrato ja aprovado de verdade seria reprocessado."""
    page.click('button:has-text("Ações")')
    page.get_by_text("Aprovação Supervisor", exact=True).click()
    page.wait_for_timeout(500)
    page.fill("textarea[name='note']", "Aprovado via RPA")
    page.get_by_role("button", name="Confirmar", exact=True).click()
    page.wait_for_load_state("networkidle")
    page.wait_for_timeout(1000)

    try:
        page.screenshot(path=f"screenshots/aprovacao_{contrato}_proposta.png")
        page.get_by_text("Histórico", exact=True).click()
        page.wait_for_load_state("networkidle")
        page.wait_for_timeout(1000)
        page.screenshot(path=f"screenshots/aprovacao_{contrato}_historico.png")
    except Exception as e:
        print(f"    [AVISO] aprovacao do contrato {contrato} concluida, mas falha ao capturar prints de evidencia: {e}")


def _limpar_busca(page):
    """Garante que a lista nao fique presa a um filtro de busca residual
    (ex.: ficou preenchida com o nome do ultimo cliente visualizado)."""
    campo_busca = page.locator('input[placeholder="Pesquisar"]')
    if campo_busca.count() == 0:
        return
    if campo_busca.input_value():
        campo_busca.fill("")
        page.wait_for_load_state("networkidle")
        page.wait_for_timeout(500)


def _selecionar_todas_lojas(page):
    """Abre o filtro 'Loja', marca 'Selecionar Todos' e atualiza a lista,
    garantindo que a extracao enxergue propostas de todas as lojas."""
    botao_loja = page.locator(
        'xpath=//div[contains(@class,"text-base") and contains(@class,"text-secondary") '
        'and normalize-space(text())="Loja"]/following-sibling::button[@aria-haspopup="menu"][1]'
    )
    if botao_loja.count() == 0:
        return
    botao_loja.click()
    page.wait_for_timeout(500)

    page.get_by_role("menuitem", name="Selecionar Todos", exact=True).click()
    page.wait_for_timeout(500)

    botao_refresh = page.locator('button:has(mat-icon[data-mat-icon-name="refresh"])').first
    botao_refresh.click()
    page.wait_for_load_state("networkidle")
    page.wait_for_timeout(1000)


def _total_propostas(page):
    page.locator("table.app-table-search tbody tr.cursor-pointer").first.wait_for(
        state="visible", timeout=30000
    )
    texto = page.locator("ajin-search-count").inner_text()
    numeros = re.findall(r"\d+", texto)
    total = int(numeros[-1]) if numeros else 0

    # Salvaguarda: a contagem do cabecalho pode nao ter carregado ainda;
    # nunca reportar menos do que ja esta renderizado na pagina atual.
    linhas_na_pagina = page.locator("table.app-table-search tbody tr.cursor-pointer").count()
    return max(total, linhas_na_pagina)


def _ir_proxima_pagina(page):
    botao = page.locator(
        'ajin-search-pagination button:has(mat-icon[data-mat-icon-name="arrow_forward_ios"])'
    )
    botao.click()
    page.wait_for_load_state("networkidle")
    page.wait_for_timeout(800)


MAX_TENTATIVAS = 5
ESPERA_ENTRE_TENTATIVAS_SEGUNDOS = 120


def _abrir_e_ler_proposta(page, i):
    pagina = i // ITENS_POR_PAGINA
    posicao = i % ITENS_POR_PAGINA

    page.goto(LOANS_URL)
    page.wait_for_load_state("networkidle")
    page.wait_for_timeout(800)
    _limpar_busca(page)

    for _ in range(pagina):
        _ir_proxima_pagina(page)

    linha = page.locator("table.app-table-search tbody tr.cursor-pointer").nth(posicao)
    data_aprovacao_promotora = ler_data_aprovacao_promotora(linha)
    loja = ler_loja(linha)

    linha.locator('button[aria-haspopup="menu"]').first.click()
    page.wait_for_timeout(400)
    page.get_by_role("menuitem", name="Visualizar", exact=True).click()
    page.wait_for_load_state("networkidle")
    page.wait_for_timeout(1000)

    contrato, nome, valor, liquido, data_proposta = ler_proposta(page)
    return contrato, nome, valor, liquido, data_proposta, data_aprovacao_promotora, loja


CHECKPOINT_A_CADA = 20


def processar_propostas(page, caminho_relatorio=None):
    """Se caminho_relatorio for informado, salva o relatorio periodicamente
    durante a execucao (a cada CHECKPOINT_A_CADA propostas), para nao perder
    o progresso caso o script seja interrompido no meio de um lote grande."""
    resultados = []

    page.goto(LOANS_URL)
    page.wait_for_load_state("networkidle")
    page.wait_for_timeout(1000)
    _limpar_busca(page)
    _selecionar_todas_lojas(page)

    total = _total_propostas(page)
    n = total if MAX_PROPOSTAS is None else min(total, MAX_PROPOSTAS)
    print(f"Processando {n} de {total} propostas listadas (DRY_RUN={DRY_RUN})...")

    try:
        for i in range(n):
            _processar_uma_proposta(page, i, resultados, caminho_relatorio)
    finally:
        if caminho_relatorio and resultados:
            gerar_relatorio(resultados, caminho_relatorio)
            print(f"    [CHECKPOINT FINAL] relatorio salvo com {len(resultados)} propostas.")

    return resultados


def _processar_uma_proposta(page, i, resultados, caminho_relatorio):
    erro = None
    dados = None
    for tentativa in range(1, MAX_TENTATIVAS + 1):
        try:
            dados = _abrir_e_ler_proposta(page, i)
            erro = None
            break
        except Exception as e:
            erro = e
            print(f"[{i}] Tentativa {tentativa}/{MAX_TENTATIVAS} falhou: {e}")
            if tentativa < MAX_TENTATIVAS:
                print(
                    f"    Aguardando {ESPERA_ENTRE_TENTATIVAS_SEGUNDOS}s antes de "
                    "tentar novamente (mesma sessao, sem novo login)..."
                )
                page.wait_for_timeout(ESPERA_ENTRE_TENTATIVAS_SEGUNDOS * 1000)

    if erro is not None:
        print(f"[{i}] Desistindo apos {MAX_TENTATIVAS} tentativas: {erro}")
        resultados.append(
            {
                "contrato": "?",
                "nome": "",
                "loja": "",
                "valor": None,
                "liquido": None,
                "data_proposta": "",
                "data_aprovacao_promotora": "",
                "data_aprovacao_supervisor": "",
                "aprovado": False,
            }
        )
        return

    contrato, nome, valor, liquido, data_proposta, data_aprovacao_promotora, loja = dados

    elegivel = liquido <= VALOR_LIMITE
    decisao = "APROVARIA" if elegivel else "PULA (valor > limite)"
    print(f"[{i}] Contrato {contrato} - Valor Líquido: R$ {liquido:,.2f} -> {decisao}")

    data_aprovacao_supervisor = ""
    aprovado = False
    if elegivel:
        if DRY_RUN:
            data_aprovacao_supervisor = datetime.now().strftime("%d/%m/%Y %H:%M")
            aprovado = True
            print("    [DRY-RUN] nao clicou em Acoes/Aprovacao Supervisor.")
        else:
            try:
                aprovar_proposta_real(page, contrato)
                data_aprovacao_supervisor = datetime.now().strftime("%d/%m/%Y %H:%M")
                aprovado = True
            except Exception as e:
                data_aprovacao_supervisor = f"ERRO: {e}"
                print(f"    [ERRO] falha ao aprovar de verdade o contrato {contrato}: {e}")

    resultados.append(
        {
            "contrato": contrato,
            "nome": nome,
            "loja": loja,
            "valor": valor,
            "liquido": liquido,
            "data_proposta": data_proposta,
            "data_aprovacao_promotora": data_aprovacao_promotora,
            "data_aprovacao_supervisor": data_aprovacao_supervisor,
            "aprovado": aprovado,
        }
    )

    if caminho_relatorio and (i + 1) % CHECKPOINT_A_CADA == 0:
        gerar_relatorio(resultados, caminho_relatorio)
        print(f"    [CHECKPOINT] relatorio salvo com {len(resultados)} propostas processadas ate agora.")


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
    ("Data e Horário Aguardando Aprovação Promotora", 32),
    ("Data e Horário da Aprovação do Supervisor", 30),
    ("Decisão", 22),
    ("Loja", 26),
]


def _ano_da_proposta(item):
    """Ano usado para escolher a aba: ano da 'Data de Cadastro' real da proposta."""
    try:
        return datetime.strptime(item.get("data_proposta", ""), "%d/%m/%Y %H:%M").year
    except (TypeError, ValueError):
        return datetime.now().year


def _escrever_cabecalho(ws):
    """Escreve os cabecalhos de COLUNAS_RELATORIO que ainda nao existem na
    aba. Cobre tanto uma aba nova quanto uma aba criada por uma execucao
    anterior do script, que pode nao ter as colunas mais recentes."""
    for col_idx, (titulo, largura) in enumerate(COLUNAS_RELATORIO, start=1):
        if ws.cell(row=1, column=col_idx).value:
            continue
        celula = ws.cell(row=1, column=col_idx, value=titulo)
        celula.fill = HEADER_FILL
        celula.font = HEADER_FONT
        celula.alignment = Alignment(horizontal="center", vertical="center")
        celula.border = THIN_BORDER
        ws.column_dimensions[get_column_letter(col_idx)].width = largura


def _obter_ou_criar_aba(wb, nome_aba):
    if nome_aba in wb.sheetnames:
        ws = wb[nome_aba]
        _escrever_cabecalho(ws)
        return ws

    # Reaproveita a aba padrao "Sheet" vazia da primeira criacao do workbook.
    if wb.sheetnames == ["Sheet"] and wb["Sheet"].max_row == 1 and wb["Sheet"]["A1"].value is None:
        ws = wb["Sheet"]
        ws.title = nome_aba
    else:
        ws = wb.create_sheet(nome_aba)

    ws.freeze_panes = "A2"
    _escrever_cabecalho(ws)
    return ws


def _escrever_linha(ws, row_idx, item):
    aprovado = item.get("aprovado", False)
    decisao_texto = "Aprovado" if aprovado else "Não Aprovado"

    valores = [
        item.get("contrato") or "-",
        item.get("nome") or "-",
        item.get("liquido"),
        item.get("data_proposta") or "-",
        item.get("data_aprovacao_promotora") or "",
        item.get("data_aprovacao_supervisor") or "",
        decisao_texto,
        item.get("loja") or "-",
    ]
    for col_idx, valor in enumerate(valores, start=1):
        celula = ws.cell(row=row_idx, column=col_idx, value=valor)
        celula.border = THIN_BORDER
        if col_idx == 3 and isinstance(valor, (int, float)):
            celula.number_format = MOEDA_FORMATO
        if col_idx in (4, 5, 6, 7):
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
    os.makedirs(os.path.join(PASTA_PROJETO, "screenshots"), exist_ok=True)
    caminho_relatorio = CAMINHO_RELATORIO_SIMULACAO if DRY_RUN else CAMINHO_RELATORIO_REAL

    with sync_playwright() as p:
        browser, page = login(p)
        resultados = processar_propostas(page, caminho_relatorio=caminho_relatorio)

        gerar_relatorio(resultados, caminho_relatorio)

        print(f"\nRelatorio salvo em {caminho_relatorio} ({len(resultados)} propostas).")
        browser.close()
