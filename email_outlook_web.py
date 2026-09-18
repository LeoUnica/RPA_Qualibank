"""
RPA - QualiBank: envio do relatorio de execucao via Outlook na Web (OWA).

Metodo definitivo de envio do e-mail de encerramento do RPA, via automacao
de navegador (Playwright) sobre o Outlook Web (outlook.office.com), usado
porque o Outlook Desktop nao esta configurado com nenhuma conta nesta
maquina.

Loga com a conta de automacao dedicada (OUTLOOK_EMAIL/OUTLOOK_PASSWORD no
.env) e usa um PERFIL DE NAVEGADOR PERSISTENTE (pasta outlook_profile/,
fora do versionamento) para manter a sessao entre execucoes - assim o login
completo (usuario+senha) so roda de fato quando a sessao expirou ou na
primeira vez. Se a conta pedir verificacao adicional (MFA) que o robo nao
consiga concluir sozinho, a execucao aguarda um tempo para conclusao manual
antes de desistir.
"""

import os
import time
from typing import List, Optional

from dotenv import load_dotenv
from playwright.sync_api import Page, sync_playwright

load_dotenv()

OUTLOOK_EMAIL = os.environ["OUTLOOK_EMAIL"]
OUTLOOK_PASSWORD = os.environ["OUTLOOK_PASSWORD"]

OWA_URL = "https://outlook.office.com/mail/"
PASTA_PROJETO = os.path.dirname(os.path.abspath(__file__))
PASTA_PERFIL_OUTLOOK = os.path.join(PASTA_PROJETO, "outlook_profile")

TIMEOUT_LOGIN_SEGUNDOS = 900  # tempo maximo aguardando login concluir (cobre eventual MFA manual)


def _mailbox_carregada(page: Page) -> bool:
    """Heuristica para saber se a caixa de entrada do OWA carregou (ou seja,
    o login/MFA ja foi concluido)."""
    if "login.microsoftonline.com" in page.url or "login.live.com" in page.url:
        return False
    candidatos = [
        page.get_by_role("button", name="Novo email"),
        page.get_by_role("button", name="Novo e-mail"),
        page.get_by_role("button", name="New mail"),
        page.get_by_role("button", name="Nova mensagem"),
    ]
    return any(c.count() > 0 for c in candidatos)


def _preencher_credenciais_microsoft(page: Page, email: str, senha: str, logger=None) -> None:
    """Preenche o fluxo padrao de login da Microsoft (tela de e-mail, tela
    de senha e o prompt 'Continuar conectado?'). Cada etapa e opcional -
    se a tela nao aparecer (ex.: conta ja pre-selecionada), so segue para a
    proxima. Se travar em algo que o robo nao sabe resolver (ex.: MFA),
    quem chama esta funcao ainda vai aguardar via _aguardar_login."""
    campo_email = page.locator('input[type="email"]')
    if campo_email.count() > 0:
        campo_email.first.fill(email)
        page.locator("#idSIButton9").click()
        page.wait_for_timeout(1500)
        if logger:
            logger.info("E-mail da conta de automação preenchido na tela de login.")

    campo_senha = page.locator('input[type="password"]')
    try:
        campo_senha.first.wait_for(state="visible", timeout=15000)
    except Exception:
        return  # provavelmente pediu algo alem de usuario/senha (ex.: MFA) - deixa para o poll de espera
    campo_senha.first.fill(senha)
    page.locator("#idSIButton9").click()
    page.wait_for_timeout(2000)
    if logger:
        logger.info("Senha da conta de automação enviada.")

    # Prompt "Continuar conectado?" - clica em Sim/Yes, se aparecer.
    try:
        page.locator("#idSIButton9").click(timeout=8000)
    except Exception:
        pass


def _aguardar_login(page: Page, logger=None) -> None:
    """Aguarda ate a caixa de entrada do OWA carregar. Cobre tanto o tempo
    normal de redirecionamento pos-login quanto uma eventual necessidade de
    verificacao adicional (MFA) resolvida manualmente na janela do robo."""
    inicio = time.time()
    while time.time() - inicio < TIMEOUT_LOGIN_SEGUNDOS:
        if _mailbox_carregada(page):
            if logger:
                logger.info("Sessão do Outlook Web autenticada.")
            return
        page.wait_for_timeout(2000)

    raise TimeoutError(
        "Login no Outlook Web não foi concluído a tempo "
        f"({TIMEOUT_LOGIN_SEGUNDOS}s). Verifique a conta de automação "
        "(pode estar pedindo verificação adicional) na janela do navegador."
    )


def _clicar_novo_email(page: Page) -> None:
    for nome in ("Novo email", "Novo e-mail", "New mail", "Nova mensagem"):
        botao = page.get_by_role("button", name=nome)
        if botao.count() > 0:
            botao.first.click()
            return
    raise RuntimeError("Não encontrei o botão 'Novo email' no Outlook Web.")


def _preencher_destinatario(page: Page, destinatario: str) -> None:
    for rotulo in ("Para", "To"):
        campo = page.locator(f'div[aria-label="{rotulo}"]')
        if campo.count() > 0:
            campo.first.click()
            page.keyboard.type(destinatario)
            page.wait_for_timeout(1200)
            page.keyboard.press("Enter")
            page.wait_for_timeout(500)
            return
    raise RuntimeError("Não encontrei o campo 'Para' no Outlook Web.")


def _preencher_assunto(page: Page, assunto: str) -> None:
    for rotulo in ("Adicionar um assunto", "Add a subject", "Assunto", "Subject"):
        campo = page.get_by_placeholder(rotulo)
        if campo.count() > 0:
            campo.first.fill(assunto)
            return
    raise RuntimeError("Não encontrei o campo 'Assunto' no Outlook Web.")


def _preencher_corpo_html(page: Page, corpo_html: str) -> None:
    """Insere o HTML no corpo da mensagem simulando um "colar" de conteudo
    rico (document.execCommand('insertHTML', ...)) em vez de escrever
    diretamente em innerHTML. O OWA mantem um modelo interno do editor
    separado do DOM visivel: escrever via innerHTML muda o que aparece na
    tela (inclusive nos nossos screenshots durante a composicao), mas ao
    clicar em Enviar o OWA serializa a partir desse modelo interno - que
    nunca chega a registrar as mudancas de innerHTML, entao a maior parte
    do CSS (fundo, borda, cor) e descartada no envio. Passar pelo
    execCommand aciona o mesmo caminho que um "colar" real do usuario,
    que o OWA processa e registra corretamente."""
    for rotulo in ("Corpo da mensagem", "Message body"):
        corpo = page.locator(f'div[aria-label="{rotulo}"]')
        if corpo.count() > 0:
            corpo.first.click()
            page.keyboard.press("Control+A")
            corpo.first.evaluate(
                "(el, html) => { document.execCommand('insertHTML', false, html); }",
                corpo_html,
            )
            return
    raise RuntimeError("Não encontrei a área de corpo da mensagem no Outlook Web.")


def _input_anexo_generico(page: Page):
    """O OWA tem varios <input type=file> ocultos na tela de composicao (um
    deles so aceita imagem, para inserir imagem inline no corpo). Usa o
    primeiro que nao tem essa restricao, que e o de anexo de arquivo comum."""
    inputs = page.locator('input[type="file"]')
    for i in range(inputs.count()):
        candidato = inputs.nth(i)
        accept = candidato.get_attribute("accept") or ""
        if "image" not in accept:
            return candidato
    return None


def _anexar_arquivos(page: Page, anexos: List[str]) -> None:
    if not anexos:
        return
    input_arquivo = _input_anexo_generico(page)
    if input_arquivo is None:
        raise RuntimeError("Não encontrei o input de anexos genérico no Outlook Web.")
    caminhos_absolutos = [os.path.abspath(a) for a in anexos if os.path.isfile(a)]
    for caminho in caminhos_absolutos:
        input_arquivo.set_input_files(caminho)
        page.wait_for_timeout(1500)


def _clicar_enviar(page: Page) -> None:
    for nome in ("Enviar", "Send"):
        botao = page.locator(f'button[aria-label="{nome}"]')
        if botao.count() > 0:
            botao.first.click()
            return
    raise RuntimeError("Não encontrei o botão 'Enviar' no Outlook Web.")


def enviar_email_outlook_web(
    corpo_html: str,
    anexos: List[str],
    assunto: str,
    destinatario: str,
    headless: bool = False,
    logger=None,
) -> None:
    """Envia o relatorio de execucao usando o Outlook Web automatizado via
    Playwright, logando com a conta de automacao (OUTLOOK_EMAIL/
    OUTLOOK_PASSWORD) e reaproveitando a sessao salva em outlook_profile/."""
    os.makedirs(PASTA_PERFIL_OUTLOOK, exist_ok=True)

    with sync_playwright() as p:
        contexto = p.chromium.launch_persistent_context(
            PASTA_PERFIL_OUTLOOK,
            headless=headless,
        )
        try:
            page = contexto.pages[0] if contexto.pages else contexto.new_page()
            page.goto(OWA_URL)
            page.wait_for_timeout(2000)
            if not _mailbox_carregada(page):
                _preencher_credenciais_microsoft(page, OUTLOOK_EMAIL, OUTLOOK_PASSWORD, logger=logger)
            _aguardar_login(page, logger=logger)

            _clicar_novo_email(page)
            page.wait_for_timeout(1000)
            _preencher_destinatario(page, destinatario)
            _preencher_assunto(page, assunto)
            _preencher_corpo_html(page, corpo_html)
            _anexar_arquivos(page, anexos)
            page.wait_for_timeout(500)
            _clicar_enviar(page)
            page.wait_for_timeout(2000)
        finally:
            contexto.close()
