
import logging
import os
import sys

from dotenv import load_dotenv
from playwright.sync_api import sync_playwright

load_dotenv()

URL = os.environ["QUALI_URL"]
ACCESS_ID = os.environ["QUALI_ACCESS_ID"]
PASSWORD = os.environ["QUALI_PASSWORD"]

logger = logging.getLogger("rpa_qualibank.login")


def login(playwright):
    browser = playwright.chromium.launch(headless=False)
    page = browser.new_page()

    page.goto(URL)

    page.click("#accessId")
    page.fill("#accessId", ACCESS_ID)

    page.click("#password")
    page.fill("#password", PASSWORD)

    page.get_by_text("Login", exact=True).click()

    page.wait_for_load_state("networkidle")
    logger.info(f"Login realizado. URL apos login: {page.url}")

    close_notificacao(page)

    page.screenshot(path="screenshots/apos_login.png")

    return browser, page


def close_notificacao(page, timeout=8000):
    dialog = page.locator("notification-reader")
    try:
        dialog.wait_for(state="visible", timeout=timeout)
    except Exception:
        logger.info("Nenhuma notificação pendente ao entrar no portal.")
        return False

    corpo = dialog.locator("div.overflow-y-auto").first
    corpo.evaluate("el => el.scrollTop = el.scrollHeight")
    page.wait_for_timeout(300)

    botao_confirmar = dialog.locator('button[mat-flat-button][color="primary"]')
    botao_confirmar.click(timeout=timeout)
    dialog.wait_for(state="hidden", timeout=timeout)
    logger.info("Notificação de aviso fechada com sucesso.")
    return True


if __name__ == "__main__":
    os.makedirs("screenshots", exist_ok=True)
    with sync_playwright() as p:
        browser, page = login(p)
        input("Login realizado. Pressione Enter para fechar o navegador...")
        browser.close()
