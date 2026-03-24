# Verificar — Status e Pontos de Atenção

## Resolvido (sessão 2026-03-17/18)

### NCM formato XXXX.XX.XX
- **Problema**: Excel tinha `73181500`, Klassmatt esperava `7318.15.00`
- **Fix**: `_format_ncm()` em `pages/fiscal.py` converte automaticamente
- **Retroativo**: `fix_ncm.py` corrigiu os 29 itens já processados (Retornar Etapa → NCM → Remeter)
- **Status**: 29/29 OK

### Overlay div1 bloqueando clicks
- **Fix**: `hide_overlays()` + `safe_click()` com JS fallback
- **Status**: Resolvido

### Referência "BAKER H" sem autocomplete
- **Fix**: Múltiplos fallbacks de autocomplete + Klassmatt cria fabricante novo via confirm dialog
- **Status**: Resolvido

### NCM readonly em itens parcialmente processados
- **Fix**: `is_editable()` em vez de `get_attribute("readonly")`
- **Status**: Resolvido

### Seletor updateRelac case-sensitive
- **Fix**: `#ibutUpdateRelac` (U maiúsculo) centralizado no config.py
- **Status**: Resolvido

### iButAddRef vs Imagebutton22
- **Fix**: `#iButAddRef` para ADD, `Imagebutton22` para EDIT
- **Status**: Resolvido

---

## Ficar de olho ao rodar 1300 itens

### Sessão expirando
- **Risco**: ~3min/item = ~65 horas para 1300 itens. A sessão Klassmatt vai expirar várias vezes.
- **Comportamento atual**: `verificar_sessao()` detecta expiração e pausa 60s para re-login manual.
- **O que fazer**: Monitorar se o bot está pausando muito. Se a sessão expirar durante a noite sem ninguém para re-logar, o bot vai ficar travado.
- **Melhoria possível**: Implementar re-login automático via SSO.

### Empresas novas no autocomplete
- **Risco**: Nomes de empresa diferentes dos 29 itens testados podem não ter match no autocomplete.
- **Comportamento atual**: Tenta 4 estratégias de fallback. Se nenhuma funciona, o confirm dialog "Deseja cadastrá-lo?" cria o fabricante novo.
- **O que fazer**: Verificar no log se muitas empresas estão sendo criadas como novas (`Autocomplete empresa '...' não encontrado`). Pode indicar erro na planilha.

### Atributos não encontrados na árvore
- **Risco**: Valores de atributos do Excel podem não existir na taxonomia do Klassmatt.
- **Comportamento atual**: Log warning com valores disponíveis, pula o atributo e continua.
- **O que fazer**: Verificar no log `Valor '...' não encontrado na árvore`. Se muitos, pode indicar PDM errado ou valores desatualizados.

### NCMs com formato diferente de 8 dígitos
- **Risco**: Algum NCM na planilha pode ter 6, 7 ou 10 dígitos.
- **Comportamento atual**: `_format_ncm()` só formata se tiver exatamente 8 dígitos. Outros passam como estão e podem ser rejeitados.
- **O que fazer**: Verificar no log `NCM ... rejeitado`. Se muitos, checar a planilha.

### Memória / estabilidade do Playwright
- **Risco**: 65+ horas de Playwright rodando pode acumular memory leaks.
- **Comportamento atual**: Se o browser crashar, o bot tenta recriar a página.
- **O que fazer**: Monitorar se o bot fica mais lento com o tempo. Se travar, reiniciar (`progress.json` garante retomada).

### Rate limiting do Klassmatt
- **Risco**: "Ocorreu uma exceção durante o processamento" por navegação rápida.
- **Comportamento atual**: 5s delay entre itens. Retry com backoff.
- **O que fazer**: Se o erro aparecer com frequência, aumentar o delay em `main.py` (linha do `asyncio.sleep(5)`).

### Referência igual em fabricante diferente
- **Risco**: Ao salvar referência, pode redirecionar para página de aviso com Voltar/Continuar.
- **Comportamento atual**: `descriptions.py` verifica e clica Continuar. Timeout curto (10s) no save.
- **O que fazer**: Se itens ficarem travados no passo de referência, verificar log por `Aviso detectado`.

### Excel aberto durante execução
- **Risco**: Se a planilha estiver aberta no Excel, o bot não consegue salvar cores.
- **Comportamento atual**: `PermissionError` e retry.
- **O que fazer**: Fechar o Excel antes de rodar.

---

## Resultados sessão 2026-03-20 (580 itens, 2 VMs)

### Números finais
- **412 OK** | 70 skipped | 43 error | 580 total
- Throughput: ~24 itens/hora (~2.4 min/item)
- Sessão estável por 10h sem expiração

### Problemas confirmados em itens OK (154 de 412)

#### NCM rejeitado pelo Klassmatt — 18 SINs (CRÍTICO)
- NCM preenchido pelo bot mas Klassmatt rejeitou → **campo ficou VAZIO**
- NCMs afetados: `84799090`, `73181500`, `84841000`
- Causa: NCM sem ponto era formatado mas Klassmatt ainda rejeitava (formato ou código inválido no cadastro fiscal)
- **Ação**: Corrigir manualmente ou via fix_ncm.py

#### Atributos/dados técnicos incompletos — 62 SINs (MÉDIO)
- Alert do Klassmatt: "É necessário preencher/verificar os dados técnicos destacados!"
- Bot aceitou o dialog e continuou → item marcado OK com atributos faltantes
- Causa: Valor do Excel não existe na árvore taxonomica (ex: "CHAVE FLUXO", "PORCA SEXTAVADA", "BOMBA AUXILIAR OLEO LUBRIFICANTE")
- **Ação**: Verificar se atributos são obrigatórios ou opcionais no fluxo

#### Sem Referência — 74 SINs (MÉDIO)
- Coluna "Empresa" vazia na planilha → step de Referência pulado silenciosamente
- `if item.get("empresa")` em main.py pula sem logar warning
- **Ação**: Preencher empresa na planilha ou confirmar que referência não é obrigatória

### Erros principais (43 total)
- **Adicionar Mídia timeout**: 34 erros (79%) — link não renderiza após uploads consecutivos. Itens com 5+ docs mais afetados. Retry 0% eficaz.
- **Execution context destroyed**: 5 erros permanentes (97% recuperado via retry)
- **Página de erro Klassmatt**: 3 erros — degradação noturna do servidor
- **Outros**: 1 timeout Relacionamento

### Gargalo de performance
- **UNSPSC**: ~45-50% do tempo total por item (popup de busca/seleção)
- Otimizar este step dobraria o throughput

### Lições para próxima wave
1. Desabilitar QuickEdit Mode no terminal Windows (pausa ao clicar na janela)
2. Klassmatt degrada após ~19h — preferir horário comercial
3. Retry de "Adicionar Mídia" é inútil — implementar fail-fast
4. Validar planilha antes: empresa vazia = sem referência

---

## Comando para verificar progresso durante execução

```powershell
# Resumo do progresso
python -c "import json; d=json.load(open('progress.json')); ok=sum(1 for v in d['items'].values() if v['status']=='ok'); err=sum(1 for v in d['items'].values() if v['status']=='error'); print(f'OK: {ok} | Erro: {err} | Total: {len(d[\"items\"])}')"

# Últimas linhas do log
powershell Get-Content klassmatt_rpa.log -Tail 20

# NCMs rejeitados
findstr "rejeitado" klassmatt_rpa.log

# Empresas sem autocomplete
findstr "Autocomplete empresa" klassmatt_rpa.log

# Atributos não encontrados
findstr "não encontrado na árvore" klassmatt_rpa.log
```
