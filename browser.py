"""Setup e helpers do Playwright para automação do Klassmatt."""

import asyncio
from playwright.async_api import async_playwright, Browser, BrowserContext, Page, Dialog

from config import (
    KLASSMATT_HOME, NAVIGATION_TIMEOUT, ACTION_TIMEOUT,
    PROFILE_DIR, SLOW_MO, HEADLESS, VIEWPORT_WIDTH, VIEWPORT_HEIGHT,
)
from logger import log


# ── Handler de dialogs JS (alerts do ASP.NET) ──
last_dialog_message: str = ""


async def _handle_dialog(dialog: Dialog) -> None:
    """Aceita automaticamente alerts/confirms/prompts do ASP.NET."""
    global last_dialog_message
    last_dialog_message = dialog.message
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
        args=[f"--window-size={VIEWPORT_WIDTH},{VIEWPORT_HEIGHT}"],
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


async def hide_overlays(page: Page) -> None:
    """Esconde overlays do Klassmatt que interceptam clicks do Playwright.

    O div#div1 é um painel de descrição expandida que fica sobre os botões
    na parte inferior da página, bloqueando Playwright clicks.
    """
    try:
        await page.evaluate(
            """() => {
                const div1 = document.querySelector('#div1');
                if (div1) div1.style.pointerEvents = 'none';
                const pg2 = document.querySelector('#pg-2');
                if (pg2) pg2.style.pointerEvents = 'none';
            }"""
        )
    except Exception:
        pass


async def navigate_home(page: Page) -> None:
    """Navega para a página inicial do Klassmatt via link interno.

    Usa o link "Principal" do dlmenu no cabeçalho (via __doPostBack).
    Não usa page.goto() diretamente pois o Klassmatt bloqueia navegação
    por URL direta sem o parâmetro k= ('ACESSO NÃO AUTORIZADO À PAGINA').
    """
    # Tentar clicar no link "Principal" no header dlmenu (existe em todas as páginas)
    try:
        principal = page.locator("a[href*='dlmenu']", has_text="Principal").first
        await principal.click(timeout=5_000)
        await page.wait_for_load_state("networkidle")
        log.debug("Navegou para home via link Principal")
        return
    except Exception:
        pass

    # Fallback: tentar "Voltar" do dlmenu (volta um nível na hierarquia)
    try:
        voltar = page.locator("a[href*='dlmenu']", has_text="Voltar").first
        await voltar.click(timeout=5_000)
        await page.wait_for_load_state("networkidle")
        log.debug("Navegou via link Voltar (dlmenu)")
        return
    except Exception:
        pass

    # Fallback: tentar "Menu Principal" (existe na home/worklist)
    try:
        menu_link = page.locator("a:has-text('Menu Principal')").first
        await menu_link.click(timeout=5_000)
        await page.wait_for_load_state("networkidle")
        log.debug("Navegou para home via link Menu Principal")
        return
    except Exception:
        pass

    # Último recurso: goto direto (funciona no login inicial)
    log.warning("Navegando via goto direto — pode causar erro de acesso se sessão ativa")
    await page.goto(KLASSMATT_HOME, wait_until="networkidle", timeout=NAVIGATION_TIMEOUT)
    log.debug("Navegou para home via goto")


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
            "text=Principal",
            "text=Sair",
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
    """Clica em um elemento com wait automático.

    Se o click normal falhar por interceptação de outro elemento,
    tenta via JavaScript (bypassa overlays ASP.NET).
    Após JS click, aguarda navegação se o click disparou postback.
    """
    timeout = timeout or ACTION_TIMEOUT
    try:
        await page.click(selector, timeout=timeout)
    except Exception as e:
        if "intercepts pointer events" in str(e):
            log.debug(f"Click interceptado em '{selector}' — usando JS fallback")
            el = page.locator(selector).first
            await el.evaluate("el => el.click()")
            # JS click pode disparar __doPostBack — aguardar navegação
            await page.wait_for_timeout(500)
            try:
                await page.wait_for_load_state("networkidle", timeout=15_000)
            except Exception:
                pass
        else:
            raise


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
