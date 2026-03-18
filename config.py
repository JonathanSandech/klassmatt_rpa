"""Configurações do RPA Klassmatt MODEC."""

import os
from pathlib import Path
from dotenv import load_dotenv

load_dotenv()

# ── Paths ──
BASE_DIR = Path(__file__).resolve().parent
EXCEL_PATH = Path(os.getenv("EXCEL_PATH", r"C:\Users\sandechin\Downloads\Robo\BASE_ANDRE_FINAL_REVISADA.xlsx"))
DOCUMENTS_DIR = Path(os.getenv("DOCUMENTS_DIR", r"C:\Users\sandechin\Downloads\Robo\Documentos - André"))
PROGRESS_FILE = BASE_DIR / "progress.json"
LOG_FILE = BASE_DIR / "klassmatt_rpa.log"
PROFILE_DIR = os.getenv("PROFILE_DIR", str(BASE_DIR / "playwright_profile"))

# ── URLs ──
KLASSMATT_HOME = os.getenv("KLASSMATT_HOME", "https://modec.klassmatt.com.br/MenuPrincipal.aspx")
# Token na URL é gerado automaticamente após login — navegar para HOME faz redirect pro login se necessário

# ── Valores fixos definidos pela MODEC ──
RELATIONSHIP_TYPE = "CÓDIGO ANTIGO"
RELATIONSHIP_STATUS = "ATIVO ERP"
RELATIONSHIP_COMMENT = "ZBRA"
PDM_CATEGORY = "PARTES E PECAS"

# ── Colunas do Excel (BASE_ANDRE_FINAL_REVISADA.xlsx) ──
# Mapeamento: nome usado no código → nome da coluna no Excel
EXCEL_COLUMNS = {
    "sin": "SIN",
    "ncm": "NCM",
    "empresa": "Empresa",
    "part_number": "Part Number",
    "unspsc": "UNSPSC",
    "codigo_60": "Código 60",
    "documento": "Documento",
    "pdm": "PDM",
    # Atributos técnicos: Atrib_1_Valor até Atrib_30_Valor
}
MAX_ATTRIBUTES = 30

# ── Retry ──
MAX_RETRIES = 3
RETRY_DELAY_MS = 2000

# ── Timeouts (ms) ──
NAVIGATION_TIMEOUT = 60_000
ACTION_TIMEOUT = 30_000
SHORT_WAIT = 2_000

# ── Browser ──
SLOW_MO = int(os.getenv("SLOW_MO", "100"))
HEADLESS = os.getenv("HEADLESS", "false").lower() == "true"
VIEWPORT_WIDTH = int(os.getenv("VIEWPORT_WIDTH", "1920"))
VIEWPORT_HEIGHT = int(os.getenv("VIEWPORT_HEIGHT", "1080"))

# ── ASP.NET Element IDs ──
# PAD usa ctl00$Body$... que no HTML vira ctl00_Body_...
SELECTORS = {
    # Worklist
    "worklist_link": "text=Acompanhamento das Solicitações (Worklist)",
    "worklist_filter_dropdown": "select:has(option[value='SOMENTE_REC_ACAO'])",

    # Busca de SIN
    "sin_search": "textarea[name$='txtValor']",
    "sin_filter_btn": "#butFiltrar",
    "sin_result": "#DIVResultado .result",
    "atuar_no_item_btn": "input[value='Atuar no Item']",

    # Criação de item
    "criar_item_btn": "input[value='Criar item']",
    "finalizar_btn": "input[value='Finalizar']",
    "salvar_btn": "input[value='Salvar']",
    "sim_btn": "input[value='Sim']",

    # Abas
    "tab_fiscal": "a:has-text('Fiscal')",
    "tab_referencias": "a:has-text('Referências')",
    "tab_classificacoes": "a:has-text('Classificações')",
    "tab_relacionamentos": "a:has-text('Relacionamentos')",
    "tab_midias": "a:has-text('Mídias')",
    "tab_descricoes": "a:has-text('Descrições')",

    # Fiscal
    "ncm_input": "#txtNCMTIPI",

    # Referências
    "ref_add_btn": "input[id$='Imagebutton22']",
    "ref_empresa_input": "#txtNome",
    "ref_partnumber_input": "#txtReferencia",
    "ref_exibe_d2_checkbox": "#ckExibeD2",
    "ref_salvar_btn": "#btnSalvar",
    "ref_duplicate_text": "Referência igual em fabricante",

    # UNSPSC
    "unspsc_btn": "#ibutUNSPSC",
    "unspsc_input": "#txtCodigoUnspsc",
    "unspsc_pesquisar_btn": "input[value='Pesquisar']",
    "unspsc_selecionar_btn": "input[value='Selecionar']",

    # Relacionamentos
    "rel_add_btn": "input[id$='Imagebutton7']",
    "rel_tipo_input": "input[name*='tabRelaciona'][id='txtTipo']",
    "rel_codigo_input": "#txtCodigoRel",
    "rel_status_input": "input[name*='tabRelaciona'][id='txtStatus']",
    "rel_comentario_input": "#txtComentario",
    "rel_save_btn": "#ibutUpdateRelac",

    # Mídias
    "media_add_link": "a:has-text('Adicionar Mídia')",
    "media_file_input": "input[id$='file']",
    "media_titulo_input": "input[id$='txtTitulo']",
    "media_salvar_btn": "input[value='Salvar']",
    "media_fechar_btn": "#cmdFechar",

    # Descrições / PDM
    "editar_descricao_link": "a:has-text('Editar Descrição')",
    "alterar_padrao_btn": "input[value='Alterar Padrão']",
    "definir_padrao_btn": "input[value='Definir Padrão']",
    "partes_pecas_link": "a:has-text('PARTES E PECAS')",

    # Atributos (IDs são genéricos; usar name para diferenciar rows)
    "attr_na_checkbox_tpl": "input[name$='dgDadosTecnicos$ctl{idx}$ckIsNA']",
    "attr_edit_btn_tpl": "input[name$='dgDadosTecnicos$ctl{idx}$btnAddEdit']",

    # Popup atributos
    "popup_letter_tpl": ".txt-letra:has-text('{letter}')",
    "popup_value_tpl": "a.nodeStyle:has-text('{value}')",
    "popup_select_btn": "#btnSelecionar",

    # Finalização
    "remeter_modec_btn": "input[value='Remeter Modec']",
}
