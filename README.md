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
