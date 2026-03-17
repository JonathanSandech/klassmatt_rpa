"""Setup e helpers do Playwright para automação do Klassmatt."""

import asyncio
from playwright.async_api import async_playwright, Browser, BrowserContext, Page, Dialog

from config import (
    KLASSMATT_HOME, NAVIGATION_TIMEOUT, ACTION_TIMEOUT,
    PROFILE_DIR, SLOW_MO, HEADLESS, VIEWPORT_WIDTH, VIEWPORT_HEIGHT,
)
from logger import log


# ── Handler de dialogs JS (alerts do ASP.NET) ──
async def _handle_dialog(dialog: Dialog) -> None:
    """Aceita automaticamente alerts/confirms/prompts do ASP.NET."""
    log.debug(f"Dialog detectado ({dialog.type}): {dialog.message}")
    await dialog.accept()


async def launch_browser() -> tuple[object, BrowserContext, Page]:
    """Inicializa Playwright com Chrome em modo visível.

    Retorna (playwright, context, page).
    O usuário deve fazer login manual na primeira vez — o contexto
    persistente preserva cookies/sessão entre execuções.
    """
    pw = await async_playwright().start()

    context = await pw.chromium.launch_persistent_context(
        user_data_dir=PROFILE_DIR,
        headless=HEADLESS,
        slow_mo=SLOW_MO,
        viewport={"width": VIEWPORT_WIDTH, "height": VIEWPORT_HEIGHT},
        args=["--start-maximized"],
        accept_downloads=True,
        permissions=[],
    )
    context.set_default_timeout(ACTION_TIMEOUT)
    context.set_default_navigation_timeout(NAVIGATION_TIMEOUT)

    page = context.pages[0] if context.pages else await context.new_page()

    # Registrar handler para dialogs JS inesperados
    page.on("dialog", _handle_dialog)

    log.info("Browser iniciado com perfil persistente")

    return pw, context, page


async def navigate_home(page: Page) -> None:
    """Navega para a página inicial do Klassmatt."""
    await page.goto(KLASSMATT_HOME, wait_until="networkidle", timeout=NAVIGATION_TIMEOUT)
    log.debug("Navegou para home do Klassmatt")


async def verificar_sessao(page: Page, timeout: int = 10_000) -> bool:
    """Verifica se a sessão está ativa checando indicadores na página.

    Retorna True se logado, False se sessão expirou.
    Padrão do bot_omie: poll por indicadores de login.
    """
    try:
        indicadores = [
            "text=Acompanhamento das Solicitações (Worklist)",
            "text=Menu Principal",
            "text=Bem-vindo",
        ]
        start = asyncio.get_event_loop().time()
        while (asyncio.get_event_loop().time() - start) * 1000 < timeout:
            for selector in indicadores:
                try:
                    if await page.locator(selector).first.is_visible():
                        log.debug("Sessão ativa confirmada")
                        return True
                except Exception:
                    continue
            await page.wait_for_timeout(500)

        log.warning("Sessão possivelmente expirada — nenhum indicador encontrado")
        return False
    except Exception as e:
        log.warning(f"Erro ao verificar sessão: {e}")
        return False


async def fechar_popups(page: Page) -> None:
    """Fecha popups inesperados que possam bloquear a automação.

    Padrão do bot_omie: try/except silencioso para cada popup conhecido.
    """
    popups_conhecidos = [
        "text=OK",
        "text=Fechar",
        "input[value='OK']",
    ]
    for selector in popups_conhecidos:
        try:
            btn = page.locator(selector).first
            if await btn.is_visible():
                await btn.click(timeout=2000)
                log.debug(f"Popup fechado: {selector}")
                await page.wait_for_timeout(500)
        except Exception:
            pass


async def safe_click(page: Page, selector: str, timeout: int | None = None) -> None:
    """Clica em um elemento com wait automático."""
    timeout = timeout or ACTION_TIMEOUT
    await page.click(selector, timeout=timeout)


async def safe_fill(page: Page, selector: str, value: str, timeout: int | None = None) -> None:
    """Preenche um campo de texto com wait automático."""
    timeout = timeout or ACTION_TIMEOUT
    await page.fill(selector, str(value), timeout=timeout)


async def wait_for_text(page: Page, text: str, timeout: int = 5000) -> bool:
    """Aguarda um texto aparecer na página. Retorna True se encontrou."""
    try:
        await page.wait_for_selector(f"text={text}", timeout=timeout)
        return True
    except Exception:
        return False


async def page_contains_text(page: Page, text: str) -> bool:
    """Verifica se a página contém determinado texto."""
    content = await page.content()
    return text in content


async def retry_action(coro_factory, max_retries: int = 3, delay_ms: int = 2000):
    """Executa uma ação com retry e backoff linear.

    coro_factory: callable que retorna uma coroutine (para poder recriar a cada tentativa).
    """
    last_error = None
    for attempt in range(1, max_retries + 1):
        try:
            return await coro_factory()
        except Exception as e:
            last_error = e
            log.warning(f"Tentativa {attempt}/{max_retries} falhou: {e}")
            if attempt < max_retries:
                await asyncio.sleep(delay_ms / 1000 * attempt)
    raise last_error
