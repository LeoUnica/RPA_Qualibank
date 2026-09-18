"""
RPA - QualiBank: envio do relatorio de execucao com fallback.

Metodo principal: Outlook desktop (COM / win32com), com checagem de conta
configurada e tentativas automaticas. Se falhar em todas as tentativas,
usa o Outlook Web (Playwright) como plano B.
"""

import os
import time
from datetime import datetime
from typing import List, Optional

from relatorio_execucao import DESTINATARIO_PADRAO

TENTATIVAS_DESKTOP = 3
ESPERA_ENTRE_TENTATIVAS_SEGUNDOS = 10


def _log(logger, nivel: str, msg: str) -> None:
    if logger is not None:
        getattr(logger, nivel)(msg)
    else:
        print(msg)


def _enviar_outlook_desktop(
    corpo_html: str, anexos: List[str], assunto: str, destinatario: str
) -> None:
    """Envia pelo Outlook clássico instalado na maquina (COM). Falha com
    RuntimeError se o Outlook nao tiver nenhuma conta configurada."""
    import pythoncom
    import win32com.client as win32

    pythoncom.CoInitialize()
    try:
        outlook = win32.Dispatch("Outlook.Application")
        if outlook.Session.Accounts.Count == 0:
            raise RuntimeError("Outlook desktop sem nenhuma conta configurada.")

        mail = outlook.CreateItem(0)
        mail.To = destinatario
        mail.Subject = assunto
        mail.HTMLBody = corpo_html
        for caminho in anexos:
            if os.path.isfile(caminho):
                mail.Attachments.Add(os.path.abspath(caminho))
        mail.Send()
    finally:
        pythoncom.CoUninitialize()


def enviar_relatorio(
    corpo_html: str,
    anexos: List[str],
    assunto: Optional[str] = None,
    destinatario: str = DESTINATARIO_PADRAO,
    logger=None,
    headless_web: bool = False,
) -> str:
    """Envia o relatorio pelo Outlook desktop (com tentativas) e, se falhar,
    pelo Outlook Web. Retorna 'desktop' ou 'web' conforme o metodo usado;
    levanta excecao se ambos falharem."""
    if assunto is None:
        assunto = f"[RPA Qualibank] Execução Concluída - {datetime.now().strftime('%d/%m/%Y %H:%M')}"

    for tentativa in range(1, TENTATIVAS_DESKTOP + 1):
        try:
            _enviar_outlook_desktop(corpo_html, anexos, assunto, destinatario)
            _log(logger, "info", "E-mail enviado com sucesso via Outlook desktop.")
            return "desktop"
        except Exception as e:
            _log(
                logger, "warning",
                f"Falha no envio via Outlook desktop (tentativa {tentativa}/{TENTATIVAS_DESKTOP}): {e}",
            )
            if tentativa < TENTATIVAS_DESKTOP:
                time.sleep(ESPERA_ENTRE_TENTATIVAS_SEGUNDOS)

    _log(logger, "warning", "Usando o Outlook Web como plano B.")
    # Import tardio: email_outlook_web exige OUTLOOK_EMAIL/OUTLOOK_PASSWORD no .env.
    from email_outlook_web import enviar_email_outlook_web

    enviar_email_outlook_web(
        corpo_html, anexos, assunto=assunto, destinatario=destinatario,
        headless=headless_web, logger=logger,
    )
    _log(logger, "info", "E-mail enviado com sucesso via Outlook Web (plano B).")
    return "web"
