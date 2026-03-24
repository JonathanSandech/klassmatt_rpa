"""Busca, criação e finalização de item no Klassmatt."""

from playwright.async_api import Page

from config import SELECTORS
from browser import safe_click, safe_fill, wait_for_text
from logger import log


async def search_and_select_sin(page: Page, sin: str) -> None:
    """Busca um SIN na worklist e seleciona o item encontrado."""
    log.info(f"Buscando SIN: {sin}")

    # Preencher campo de busca
    await safe_fill(page, SELECTORS["sin_search"], str(sin))

    # Clicar em Filtrar
    await safe_click(page, SELECTORS["sin_filter_btn"])
    await page.wait_for_load_state("networkidle")
    await page.wait_for_timeout(2000)  # Buffer pós-networkidle para grid popular

    # Verificar se há resultados na grid (estrutura de divs, não table)
    result_row = page.locator(SELECTORS["sin_result"]).first

    try:
        await result_row.wait_for(state="visible", timeout=25_000)
    except Exception:
        raise RuntimeError(
            f"SIN {sin} não encontrado na worklist — nenhum resultado em #DIVResultado após filtrar. "
            f"Verifique se o SIN existe e se o filtro 'Todas as Solicitações' está ativo."
        )

    # Navegar clicando no link abreSIN (estrutura real da worklist)
    sin_link = page.locator(f"a[href*='abreSIN({sin})']")
    try:
        await sin_link.wait_for(state="visible", timeout=5_000)
        await sin_link.click()
    except Exception:
        # Fallback: clicar no resultado diretamente
        await result_row.click()
    await page.wait_for_load_state("networkidle")

    log.info(f"SIN {sin} selecionado")


async def atuar_no_item(page: Page) -> str | None:
    """Clica em 'Atuar no Item' (ou 'Atuar na SIN') e aguarda navegação.

    Detecta o estado da página SIN_Item_Resultante ANTES de clicar:
    - Botão disabled (APROVACAO-TECNICA etc.) → retorna status para skip
    - 'Atuar na SIN' (CATALOGACAO-MODEC) → fluxo criar item
    - 'Atuar no Item' enabled (FINALIZACAO) → fluxo normal

    Retorna None se conseguiu entrar na edição, ou o status string se
    o item não pode ser editado (caller deve pular).
    """
    # Se já estamos em ITEM_Edita.aspx, não precisa clicar
    if "ITEM_Edita.aspx" in page.url:
        log.debug("Já em ITEM_Edita.aspx — pulando 'Atuar no Item'")
        return None

    # Override confirm para aceitar "outro usuário atuando" automaticamente
    await page.evaluate("() => { window.confirm = () => true; }")

    # Detectar estado dos botões e status ANTES de clicar
    page_state = await page.evaluate("""() => {
        const status = document.querySelector("input[id$='txtStatus']");
        const statusVal = status ? status.value : '';

        // Buscar botões de ação (butAcao3 = principal, butAcao2 = secundário)
        const btn3 = document.querySelector('#butAcao3');
        const btn2 = document.querySelector('#butAcao2');

        // Buscar botão genérico "Atuar no Item"
        const atuarItem = document.querySelector("input[value='Atuar no Item']");

        return {
            status: statusVal,
            btn3Value: btn3 ? btn3.value : null,
            btn3Disabled: btn3 ? btn3.disabled : true,
            btn3Visible: btn3 ? btn3.offsetParent !== null : false,
            btn2Value: btn2 ? btn2.value : null,
            atuarItemExists: !!atuarItem,
            atuarItemDisabled: atuarItem ? atuarItem.disabled : true,
        };
    }""")

    status = page_state.get("status", "")
    btn3_value = page_state.get("btn3Value")
    btn3_disabled = page_state.get("btn3Disabled", True)
    btn3_visible = page_state.get("btn3Visible", False)
    atuar_disabled = page_state.get("atuarItemDisabled", True)

    log.debug(f"Estado da página: status={status}, btn3={btn3_value}, disabled={btn3_disabled}")

    # Botão "Atuar no Item" existe mas está disabled → item em etapa não editável
    if page_state.get("atuarItemExists") and atuar_disabled:
        log.info(f"Botão 'Atuar no Item' disabled — item em '{status}'")
        return status or "disabled"

    # "Atuar na SIN" (CATALOGACAO-MODEC) — fluxo de criação de item
    if btn3_value == "Atuar na SIN" and btn3_visible and not btn3_disabled:
        log.info("SIN em CATALOGACAO-MODEC — executando fluxo Atuar na SIN → Criar item")
        await page.evaluate("() => { document.querySelector('#butAcao3').click(); }")
        await page.wait_for_load_state("networkidle")
        await page.wait_for_timeout(1500)

        criar_btn = page.locator("#butAcao2")
        if await criar_btn.count() > 0:
            await page.evaluate("() => { document.querySelector('#butAcao2').click(); }")
            await page.wait_for_load_state("networkidle")
            await page.wait_for_timeout(1000)

            finalizar = page.locator("#butFinaliza")
            if await finalizar.count() > 0:
                await page.evaluate("() => { document.querySelector('#butFinaliza').click(); }")
                await page.wait_for_load_state("networkidle")
                await page.wait_for_timeout(1000)

            salvar = page.locator("#butSalvar")
            if await salvar.count() > 0:
                await page.evaluate("() => { document.querySelector('#butSalvar').click(); }")
                await page.wait_for_load_state("networkidle")
                await page.wait_for_timeout(1000)

            sim = page.locator("#butSim")
            if await sim.count() > 0:
                await page.evaluate("() => { document.querySelector('#butSim').click(); }")
                await page.wait_for_load_state("networkidle")
                await page.wait_for_timeout(1000)

            log.info("Item criado via fluxo CATALOGACAO-MODEC")
        else:
            log.warning("Botão 'Criar item' não encontrado após 'Atuar na SIN'")
        return None

    # Fluxo normal: Atuar no Item (FINALIZACAO) — botão enabled
    if page_state.get("atuarItemExists") and not atuar_disabled:
        await safe_click(page, SELECTORS["atuar_no_item_btn"])
        await page.wait_for_load_state("networkidle")
        await page.wait_for_timeout(1000)
        log.debug("Clicou em 'Atuar no Item'")
        return None

    # Nenhum botão de ação encontrado
    log.warning(f"Nenhum botão de ação encontrado (status={status})")
    return status or "unknown"


async def check_item_already_processed(page: Page) -> str | None:
    """Verifica se o item já avançou além de FINALIZACAO no workflow.

    Itens em APROVACAO-TECNICA, APROVACAO-FINAL, etc. já foram remetidos
    e não têm botão 'Remeter Modec'.
    Retorna o status string se já processado, None se está em FINALIZACAO.
    """
    status_el = page.locator("input[id$='txtStatus']")
    if await status_el.count() > 0:
        status = await status_el.input_value()
        if status and status != "FINALIZACAO":
            log.info(f"Item já em status '{status}' — não está em FINALIZACAO")
            return status
    return None


async def criar_item(page: Page) -> None:
    """Cria o item se o botão 'Criar item' estiver presente.

    Se o item já foi criado (ex: etapa FINALIZACAO), pula esta etapa.
    Sequência quando necessário: Criar Item → Finalizar → Salvar → Sim
    """
    criar_btn = page.locator(SELECTORS["criar_item_btn"])
    if await criar_btn.count() > 0 and await criar_btn.is_visible():
        log.info("Criando item...")

        await safe_click(page, SELECTORS["criar_item_btn"])
        await page.wait_for_load_state("networkidle")

        await safe_click(page, SELECTORS["finalizar_btn"])
        await page.wait_for_load_state("networkidle")

        await safe_click(page, SELECTORS["salvar_btn"])
        await page.wait_for_load_state("networkidle")

        await safe_click(page, SELECTORS["sim_btn"])
        await page.wait_for_load_state("networkidle")

        log.info("Item criado com sucesso")
    else:
        log.info("Item já criado — pulando etapa de criação")


async def finalizar_e_remeter(page: Page) -> None:
    """Finaliza o item e remete para MODEC.

    Adapta o fluxo conforme os botões disponíveis na página:
    - Se 'Finalizar' visível: Finalizar → Atuar no Item → Remeter Modec → Sim
    - Se 'Remeter Modec' já visível (etapa FINALIZACAO): Remeter Modec → Sim
    """
    log.info("Finalizando e remetendo para MODEC...")

    remeter_btn = page.locator(SELECTORS["remeter_modec_btn"])
    finalizar_btn = page.locator(SELECTORS["finalizar_btn"])

    # Se "Remeter Modec" já está visível, clicar direto
    if await remeter_btn.count() > 0 and await remeter_btn.is_visible():
        log.info("Botão 'Remeter Modec' já disponível — remetendo direto")
        await safe_click(page, SELECTORS["remeter_modec_btn"])
        await page.wait_for_load_state("networkidle")

    elif await finalizar_btn.count() > 0 and await finalizar_btn.is_visible():
        # Fluxo completo: Finalizar → Atuar → Remeter
        await safe_click(page, SELECTORS["finalizar_btn"])
        await page.wait_for_load_state("networkidle")

        await page.wait_for_selector(SELECTORS["atuar_no_item_btn"], timeout=10_000)
        await safe_click(page, SELECTORS["atuar_no_item_btn"])
        await page.wait_for_load_state("networkidle")

        await page.wait_for_selector(SELECTORS["remeter_modec_btn"], timeout=10_000)
        await safe_click(page, SELECTORS["remeter_modec_btn"])
        await page.wait_for_load_state("networkidle")
    else:
        # Verificar se o item já foi remetido (status avançou)
        status_el = page.locator("input[id$='txtStatus']")
        status = await status_el.input_value() if await status_el.count() > 0 else "desconhecido"
        if status and status != "FINALIZACAO":
            log.info(f"Item já em status '{status}' — já foi remetido anteriormente")
            return

        # Tentar recarregar a página e verificar novamente
        await page.wait_for_timeout(2000)
        remeter_btn2 = page.locator(SELECTORS["remeter_modec_btn"])
        finalizar_btn2 = page.locator(SELECTORS["finalizar_btn"])
        if await remeter_btn2.count() > 0:
            await safe_click(page, SELECTORS["remeter_modec_btn"])
            await page.wait_for_load_state("networkidle")
        elif await finalizar_btn2.count() > 0:
            await safe_click(page, SELECTORS["finalizar_btn"])
            await page.wait_for_load_state("networkidle")
        else:
            log.warning(f"Nem 'Finalizar' nem 'Remeter Modec' encontrados (status={status}) — pulando remessa")
            return

    # Confirmar com "Sim" (se aparecer diálogo de confirmação)
    sim_btn = page.locator(SELECTORS["sim_btn"])
    try:
        await sim_btn.wait_for(state="visible", timeout=5_000)
        await safe_click(page, SELECTORS["sim_btn"])
        await page.wait_for_load_state("networkidle")
    except Exception:
        log.debug("Botão 'Sim' não apareceu — confirmação não necessária")

    log.info("Item remetido para MODEC")
