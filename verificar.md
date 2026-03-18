# NCM — RESOLVIDO

## Problema

NCMs da planilha eram rejeitados pelo Klassmatt: "o NCM informado é inválido ou está inativo!"

## Causa

**Formato errado.** O Excel tinha `73181500` (sem pontos), mas o Klassmatt espera `7318.15.00` (com pontos no padrão `XXXX.XX.XX`).

## Solução

Adicionada função `_format_ncm()` em `pages/fiscal.py` que converte automaticamente:
- `73181500` → `7318.15.00`
- `84841000` → `8484.10.00`
- `84799090` → `8479.90.90`

## Teste

Testado manualmente via MCP Playwright no SIN 474470:
- `7318.15.00` → aceito, classificação carregou: "PARAFUSOS, PINOS OU PERNOS..."
- IPI e II preenchidos automaticamente (6,50% e 16,00%)
