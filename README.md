# RPA Klassmatt MODEC

Bot de automação (RPA) para cadastro em massa de itens no sistema web [Klassmatt](https://modec.klassmatt.com.br), desenvolvido para a MODEC. Migrado do Power Automate Desktop para **Python + Playwright**.

## Funcionalidades

- Cadastro automatizado de itens a partir de uma planilha Excel
- Workflow de 12 etapas por item:
  1. Buscar SIN na worklist
  2. Atuar no item
  3. Criar item (Finalizar → Salvar → Sim)
  4. Classificação UNSPSC
  5. Dados fiscais (NCM)
  6. Referências (Empresa + Part Number)
  7. Relacionamento (Código Antigo)
  8. Upload de documentos
  9. Validação da descrição SAP (limite 40 chars / Exibe D2)
  10. Alteração de PDM
  11. Preenchimento de atributos técnicos (até 30)
  12. Finalizar e remeter para MODEC
- **Retomada automática** — progresso salvo por SIN em `progress.json`
- **Retry com backoff** — até 3 tentativas por item com recuperação de erro
- **Tracking visual no Excel** — linhas coloridas por status:
  - 🟢 Verde: sucesso
  - 🔴 Vermelho: erro
  - 🟠 Laranja: referência duplicada
- Detecção de sessão expirada com pausa para login manual

## Arquitetura Multi-VM

O processo completo de cadastro roda em **4 VMs paralelas**, cada uma responsável por uma etapa do pipeline:

```
VM .111 ─┐
VM .112 ─┤──→  \\MODEC SHARED\  ──→  VM .110 (este bot)
VM .11x ─┘     (diretório compartilhado)     │
                                              ▼
                                        Klassmatt Web
```

- **VMs upstream (.111, .112, ...)** — Executam etapas anteriores (download de documentos, processamento, etc.)
- **VM .110 (este bot)** — Etapa final: cadastro dos itens no sistema Klassmatt

### Diretório Compartilhado

Todas as VMs acessam o diretório de rede:

```
\\WKS-TESTUSER5\Users\sandechin\Desktop\MODEC SHARED\
├── downloads/                          # PDFs baixados pelas VMs upstream
├── documentos_baixados_<VM>.xlsx       # Manifesto de downloads por VM
│   Colunas: PART NUMBER, TAG, Nome do Documento, Data do Documento, Caminho
├── ptC_resultado_<VM>.xlsx             # Resultados de processamento por VM
│   Colunas: PART NUMBER, G09, Tempo (s)
```

As planilhas auxiliares (`documentos_baixados_*.xlsx`, `ptC_resultado_*.xlsx`) contêm informações adicionais do pipeline. Para produção, configure `DOCUMENTS_DIR` apontando para a pasta `downloads/` do shared.

## Pré-requisitos

- Python 3.10+
- Google Chrome instalado
- Acesso ao sistema Klassmatt MODEC

## Instalação

```bash
cd klassmatt_rpa
pip install -r requirements.txt
python -m playwright install chromium
```

### Dependências

| Pacote        | Uso                          |
|---------------|------------------------------|
| playwright    | Automação do navegador       |
| openpyxl      | Leitura/escrita do Excel     |
| python-dotenv | Variáveis de ambiente (.env) |

## Configuração

Crie um arquivo `.env` na pasta `klassmatt_rpa/` (opcional — há valores padrão em `config.py`):

```env
EXCEL_PATH=C:\caminho\para\planilha.xlsx
DOCUMENTS_DIR=C:\caminho\para\documentos
PROFILE_DIR=C:\caminho\para\perfil_playwright
KLASSMATT_HOME=https://modec.klassmatt.com.br/MenuPrincipal.aspx

# Produção (lendo do diretório compartilhado):
# DOCUMENTS_DIR=\\WKS-TESTUSER5\Users\sandechin\Desktop\MODEC SHARED\downloads
# SHARED_DIR=\\WKS-TESTUSER5\Users\sandechin\Desktop\MODEC SHARED
```

### Planilha Excel

A planilha deve conter as seguintes colunas (configuráveis em `config.py`):

| Coluna       | Descrição                        |
|--------------|----------------------------------|
| SIN          | Identificador do item            |
| NCM          | Código fiscal NCM                |
| Empresa      | Empresa para referência          |
| Part Number  | Número da peça                   |
| UNSPSC       | Classificação UNSPSC             |
| Código 60    | Código antigo (relacionamento)   |
| Documento    | Nome(s) do(s) arquivo(s) a anexar|
| PDM          | Categoria PDM                    |
| Atrib_1_Valor ... Atrib_30_Valor | Atributos técnicos |

## Execução

```bash
cd klassmatt_rpa
python main.py
```

### Primeiro uso

Na primeira execução, o navegador abrirá e será necessário **fazer login manualmente** no Klassmatt. O perfil do navegador é salvo em `./playwright_profile/`, preservando a sessão para as próximas execuções.

### Retomada

Se o processo for interrompido, basta executar novamente. O bot lê o `progress.json` e pula automaticamente os itens já processados com sucesso.

## Estrutura do Projeto

```
klassmatt_rpa/
├── main.py              # Orquestrador principal
├── config.py            # Configurações, caminhos, seletores
├── browser.py           # Setup do Playwright e helpers
├── excel_handler.py     # Leitura do Excel e coloração de linhas
├── state.py             # Controle de progresso (progress.json)
├── logger.py            # Logging (arquivo + console)
├── requirements.txt     # Dependências Python
├── pages/               # Page Objects (um por seção do formulário)
│   ├── worklist.py      #   Navegação e filtro da worklist
│   ├── item.py          #   Busca SIN, criar, finalizar, remeter
│   ├── classifications.py #  Popup de classificação UNSPSC
│   ├── fiscal.py        #   Dados fiscais (NCM)
│   ├── references.py    #   Referências (empresa + part number)
│   ├── relationships.py #   Relacionamentos (código antigo)
│   ├── media.py         #   Upload de documentos
│   ├── descriptions.py  #   Validação descrição SAP + PDM
│   └── attributes.py    #   Atributos técnicos (até 30)
├── progress.json        # Estado de progresso (gerado em runtime)
└── klassmatt_rpa.log    # Log detalhado (gerado em runtime)
```

## Logs

- **Console**: nível INFO — acompanhamento em tempo real
- **Arquivo** (`klassmatt_rpa.log`): nível DEBUG — diagnóstico completo

## Tratamento de Erros

- Cada item é tentado até **3 vezes** com backoff linear
- Em caso de falha, o bot tenta navegar de volta à worklist
- Se a navegação falhar, uma nova aba é criada como fallback
- Referências duplicadas são tratadas separadamente (pintadas de laranja e puladas)
- Itens com documentos faltantes são pulados antes do processamento

## Changelog

### 2026-03-17 — Debug e correções com MCP Playwright

Sessão intensiva de debug usando MCP Playwright para inspecionar a UI do Klassmatt em tempo real. O bot passou de **0/29 itens** para **5/6 OK** (~83% de sucesso, os erros restantes são de dados).

#### Bug crítico: Seletor case-sensitive no Relacionamento (`relationships.py`)

**Sintoma**: Todos os 29 itens falhavam com timeout no botão salvar do Relacionamento.

**Causa**: O seletor `input[type='image'][id$='updateRelac']` usava `u` minúsculo, mas o ID real do elemento é `ibutUpdateRelac` (com `U` maiúsculo). CSS `[id$=...]` é case-sensitive.

**Fix**: Substituiu seletores hardcoded por referências ao `SELECTORS["rel_save_btn"]` (`#ibutUpdateRelac`) e `SELECTORS["rel_add_btn"]` do `config.py`.

#### Bug: Tabela de atributos (`dgDadosTecnicos`) na página errada

**Sintoma**: Timeout em `input[id$='dgDadosTecnicos_ctl03_btnAddEdit']` — elemento não existe.

**Causa 3 bugs sobrepostos**:
1. **Página errada** — A tabela `dgDadosTecnicos` fica em `ITEM_Edita_DescricaoV3.aspx` (acessada via Descrições → Editar Descrição), não em `ITEM_Edita.aspx`
2. **Seletores `id` vs `name`** — No ASP.NET DataGrid, todos os botões `btnAddEdit` têm o mesmo `id="btnAddEdit"`. A diferenciação entre rows é pelo atributo `name` (e.g., `ctl00$Body$dgDadosTecnicos$ctl02$btnAddEdit`). Corrigido para `input[name$='dgDadosTecnicos$ctl{idx}$btnAddEdit']`
3. **Índice off-by-one** — Atributo 1 corresponde a `ctl02` (não `ctl03`). O `ctl01` é a row do header

**Fix**: Navegação para `ITEM_Edita_DescricaoV3.aspx`, seletores por `name`, índice corrigido para `loop_index + 1`.

#### Bug: Árvore de atributos com letras do alfabeto (`Dt_EditaArvore.aspx`)

**Sintoma**: Valor "PORCA BORBOLETA" não encontrado na árvore — a árvore mostra letras (A-Z) como nós.

**Causa**: A árvore de taxonomia (`Dt_EditaArvore.aspx`) é aberta via `window.open()` pelo JS `AbreJanTaxonomia()`. Tem estrutura hierárquica:
- Nível 0: nome do dado técnico (ex: "NOME VALIDO")
- Nível 1: letras do alfabeto ([0-9], A, B, C, ..., Z)
- Nível 2+: valores reais (ex: "PORCA BORBOLETA" sob "P")

Clicar numa letra dispara `__doPostBack` que recarrega a popup inteira. O código original usava `page.evaluate()` para clicar, que retornava antes do postback completar — a busca subsequente operava no DOM antigo.

**Fix**: Abordagem híbrida:
- **JS `evaluate`** para encontrar e clicar na letra (rápido para DOM grande) + `wait_for_load_state("networkidle")` para esperar o postback
- **JS `evaluate`** para buscar o valor entre 1900+ nós expandidos
- **Playwright `locator.click()`** para o botão "Selecionar" (que fecha a popup)

#### Outros fixes menores

- **Mídias `cmdFechar`** (`media.py`): O botão Fechar tem `onclick="window.close()"` que destrói a página antes do Playwright completar o click. Adicionado try/except
- **NCM readonly** (`fiscal.py`): Itens parcialmente processados têm o campo NCM como `readonly`. O bot agora detecta e pula
- **Volta da DescricaoV3** (`attributes.py`): Após preencher atributos, navega de volta via `butSIN_Voltar` → `Atuar no Item` para que `finalizar_e_remeter` funcione
