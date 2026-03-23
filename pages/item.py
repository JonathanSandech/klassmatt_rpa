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


async def atuar_no_item(page: Page) -> None:
    """Clica em 'Atuar no Item' e aguarda navegação para página de edição.

    Para SINs em CATALOGACAO-MODEC, o abreSIN() abre SIN_Resumo com botão
    'Atuar na SIN' (butAcao3) em vez de 'Atuar no Item'. Nesse caso:
    1. Atuar na SIN → aparece 'Criar item'
    2. Criar item → vai para DescricaoV3
    3. Finalizar → volta para ITEM_Edita
    4. Salvar → Sim → item criado em FINALIZACAO
    """
    # Se já estamos em ITEM_Edita.aspx, não precisa clicar
    if "ITEM_Edita.aspx" in page.url:
        log.debug("Já em ITEM_Edita.aspx — pulando 'Atuar no Item'")
        return

    # Override confirm para aceitar "outro usuário atuando" automaticamente
    await page.evaluate("() => { window.confirm = () => true; }")

    # Verificar se tem "Atuar na SIN" (CATALOGACAO-MODEC) em vez de "Atuar no Item"
    atuar_sin = await page.evaluate("""() => {
        const btn = document.querySelector('#butAcao3');
        return btn && btn.value === 'Atuar na SIN' && btn.offsetParent !== null;
    }""")

    if atuar_sin:
        log.info("SIN em CATALOGACAO-MODEC — executando fluxo Atuar na SIN → Criar item")
        # 1. Atuar na SIN
        await page.evaluate("() => { document.querySelector('#butAcao3').click(); }")
        await page.wait_for_load_state("networkidle")
        await page.wait_for_timeout(1500)

        # 2. Criar item (butAcao2)
        criar_btn = page.locator("#butAcao2")
        if await criar_btn.count() > 0:
            await page.evaluate("() => { document.querySelector('#butAcao2').click(); }")
            await page.wait_for_load_state("networkidle")
            await page.wait_for_timeout(1000)

            # 3. Finalizar na DescricaoV3
            finalizar = page.locator("#butFinaliza")
            if await finalizar.count() > 0:
                await page.evaluate("() => { document.querySelector('#butFinaliza').click(); }")
                await page.wait_for_load_state("networkidle")
                await page.wait_for_timeout(1000)

            # 4. Salvar na ITEM_Edita
            salvar = page.locator("#butSalvar")
            if await salvar.count() > 0:
                await page.evaluate("() => { document.querySelector('#butSalvar').click(); }")
                await page.wait_for_load_state("networkidle")
                await page.wait_for_timeout(1000)

            # 5. Sim (confirmar criação)
            sim = page.locator("#butSim")
            if await sim.count() > 0:
                await page.evaluate("() => { document.querySelector('#butSim').click(); }")
                await page.wait_for_load_state("networkidle")
                await page.wait_for_timeout(1000)

            log.info("Item criado via fluxo CATALOGACAO-MODEC")
        else:
            log.warning("Botão 'Criar item' não encontrado após 'Atuar na SIN'")
        return

    # Fluxo normal: Atuar no Item (FINALIZACAO)
    await safe_click(page, SELECTORS["atuar_no_item_btn"])
    await page.wait_for_load_state("networkidle")
    await page.wait_for_timeout(1000)
    log.debug("Clicou em 'Atuar no Item'")


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
