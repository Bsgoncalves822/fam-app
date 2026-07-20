# FAM Metal — Reconciliação Bancária

**Cliente:** FAM Metal Indústria Metal Mecânica LTDA — CNPJ: 04.957.294/0001-03  
**Dev:** Bryan · ORPROCON / Contadores Digitais · Tubarão SC  
**Usuária final:** Yasmin (contadora)  
**Stack:** Excel `.xlsm` + VBA (Module1) + Python CLI (`extract_comprovantes_vN.py`)  
**Repositório:** github.com/Bsgoncalves822/fam-app (public, branch: main)

---

## O que é

Ferramenta mensal de ETL e conciliação bancária. Substitui o Flask/fam-app original. Todo o fluxo roda dentro do Excel via macros VBA + script Python standalone para extração de PDFs.

O arquivo principal é `FAM_Reconciliacao.xlsm`. O script Python fica na mesma pasta; o VBA o localiza por wildcard (`extract_comprovantes*.py`, pega o mais recente por timestamp).

---

## Estrutura do workbook

| Aba | Conteúdo |
|-----|----------|
| `RESUMO` | Dashboard + 4 botões de macro |
| `SICREDI` / `BB` / `CEF` / `DAYCOVAL` / `SICOOB` | Uma aba por banco; colunas detectadas por cabeçalho, não por posição |
| `COMPROVANTES` | Tabela flat de 15 colunas com dados extraídos dos PDFs |
| `_FORNECEDORES` | ~1.695 linhas CNPJ + Nome |
| `_FUNCIONARIOS` | ~2.372 nomes |
| `_CHEQUES` | ~199 linhas |
| `_MAPEAMENTO` | Sub-contas de roteamento → banco + conta (não é tabela de classificação) |

### Colunas de enriquecimento (cada aba de banco)

`CNPJ` / `PARTICIPANTE` / `FOLHA` / `CHEQUE Nº` / `COMPROVANTE` / `CODIGO` / `STATUS`

### Colunas COMPROVANTES (15)

`Data` / `Valor` / `Tipo` / `Label` / `Destinatário` / `CNPJ Dest` / `CPF Dest` / `Banco Dest` / `Solicitante` / `Nº Controle` / `ID Transação` / `Data Venc` / `Descrição` / `Arquivo` / `Obs`

---

## Macros (Module1 — `FAM_Extracao_v20.bas`)

| Botão | Sub | O que faz |
|-------|-----|-----------|
| Atualizar Participantes | `AtualizarParticipantes` | Recarrega arrays de fornecedores/funcionários na memória |
| Importar Extratos | `ImportarExtratos` | Detecta banco pelo nome do arquivo, detecta formato (OFX/XLS/XLSX/HTML), parseia e popula aba do banco |
| Importar Comprovantes | `ImportarComprovantes` | 3 fases: index ZIP → extração seletiva de PDFs → backfill de participante |
| Validar Comprovantes | `ValidarComprovantes` | Match data+valor em memória entre banco e COMPROVANTES; sem fórmulas ao vivo |

### Auxiliares (Module1+)

| Sub/Function | Descrição |
|---|---|
| `LimparDados` | Limpa linhas 3–502 de todas as abas de banco |
| `LimparComprovantes` | Limpa `COMPROVANTES!A3:O3000` |
| `MarcarOK/Revisar/Pendente` | Seta STATUS nas linhas selecionadas (funciona só em abas de banco) |
| `AtualizarResumo` | `CalculateFull` e ativa RESUMO |
| `IrParaResumo` | Navega para RESUMO |

---

## Lógica de resolução (`ResolverLinha`)

Ordem de prioridade:

| Passo | Regra |
|-------|-------|
| 0 | Auto-transferência — CNPJ ou nome da FAM |
| 0.5 | Padrão cheque ALW |
| 1 | CNPJ exato → `_FORNECEDORES` (injeta fórmula INDEX/MATCH ao vivo) |
| 2 | CPF → fuzzy → `_FUNCIONARIOS` (FOLHA=SIM) |
| 3 | Nome-após-doc fuzzy → `_FORNECEDORES` |
| 4 | CNPJ/CPF bruto mantido (REVISAR) |
| 6.5 | CNPJ encontrado no texto da transação mas não no cadastro → nome extraído do texto → STATUS=OK |
| 7 | CPF não encontrado → sinal do valor determina CATEGORIA (positivo → recebimento, negativo → fornecedor); ambos ficam REVISAR |

Todos os lookups usam arrays em memória (`mFornArr`/`mFuncArr`). Coluna CNPJ forçada para formato Texto (`@`) antes de gravar.

---

## Importação de extratos

### Detecção de banco/formato

- Banco detectado pelo nome do arquivo
- Formato detectado por conteúdo (sniff de headers OFX / tags HTML — CEF varia mês a mês)

### Parsers implementados

| Banco | Formatos | Observações |
|-------|----------|-------------|
| Sicredi | XLS, OFX | — |
| BB | XLSX, OFX | Valor como texto BR sem sinal; sinal na coluna D/C separada (`ValorSignado()`) |
| CEF | OFX, HTML | Formato inconsistente — detectar por conteúdo, não extensão; sem nome/CNPJ no texto |
| Daycoval | — | Parser ainda não implementado (aguarda arquivo de amostra) |
| Sicoob | — | Parser ainda não implementado (aguarda arquivo de amostra) |

**BB OFX memo:** padrão `"PIX - ENVIADO - dd/mm hh:mm NOME"` — usar último `" - "`, strip do prefixo datetime.

**OFX gap:** `ImportarOFX` captura só MEMO, descarta CHECKNUM. Para linhas "ESTORNO DE DÉBITO" sem participante no MEMO, CHECKNUM é o único campo identificador. Fix pendente: concatenar CHECKNUM na Descrição/Histórico.

---

## Importação de comprovantes (`ImportarComprovantes`)

### Fluxo em 3 fases

1. **Phase 1 — Index (~instantâneo):** `py index <zip> <csv>` — lista o ZIP via `parse_filename()`, sem abrir PDFs
2. **Phase 2 — Extração seletiva (só linhas REVISAR/PENDENTE com match de comprovante):** `py full <zip> <csv> <filter.txt>` — abre apenas os PDFs matchados com pdfplumber + ProcessPoolExecutor
3. **Phase 3 — `ResolverViaComprovantes`:** backfill de CNPJ/PARTICIPANTE nas linhas do banco com comprovante matchado mas sem participante

### `extract_comprovantes_vN.py`

- CLI standalone (sem Flask)
- Mesmo regex do fam-app/app.py (incluindo quirks `[a??]`)
- CSV gravado em **cp1252** (não UTF-8) para `Open`/`Line Input` do VBA ler acentos corretamente
- HEADERS: 15 colunas — `index_only()` deve bater exatamente
- Arquivo sentinela `.done` sempre gravado (OK ou `ERRO:` + traceback)
- Pre-flight `PythonAcessivel()` com timeout 10s antes de qualquer extração
- VBA grava `.bat` temporário por invocação (nunca string raw de cmd — bug de arg-shift descoberto)

**Caminho dos comprovantes:** `C:\Users\bryan\Downloads\FAM-RECIBOS`

---

## Exportação TXT (SCI Único)

Formato CELESP compact, 16 campos, comma-separated, sem aspas:

| Campo | Conteúdo |
|-------|----------|
| 1 | Sequência |
| 2 | Data (AAAAMMDD) |
| 3 | Conta débito |
| 4 | Conta crédito |
| 5 | Valor (sempre ABS — direção via débito/crédito, nunca negativo) |
| 6 | Histórico padrão (3026=Pagamento, 3708=Recebimento; blank quando não há código) |
| 7 | Complemento |
| 8 | Nº doc (prefixo DCTO / "DCTO0" fallback) |
| 9 | Lote |
| 10 | CNPJ participante lado débito |
| 11 | CNPJ participante lado crédito |
| 12–15 | Blank |
| 16 | "A" |

**Regra débito/crédito:** valor negativo (saída) → débito=código de classificação, crédito=conta bancária. Valor positivo (entrada) → crédito=código, débito=conta bancária.

### Códigos de conta confirmados

| Categoria | Conta |
|-----------|-------|
| Folha | 6775 |
| Fornecedor | 148 |
| Recebimento | 16 |
| Adiantamento | 4159 |
| Sicredi Consórcio | 6293 |
| Caixa Consórcio | 7772 |
| BB / BB Consórcio | 700 |
| Empréstimo (todos) | 8329 |
| FGTS | 6791 |
| DARF | 5 |
| IOF | 571 |
| Tarifa | 565 |

### Gatilhos de empréstimo

**Sicredi (prefixo C-):** C40730840 / C40721253 / C60731942 / C40733453 / C40733458 / C60732002 / C60731887 + `LIQUIDACAO DE PARCELA` / `LIQUIDACAO CONTRATO` / `LIBERACAO CREDITO` / `AMORTIZACAO CONTRATO`

**BB:** memos `BB GIRO` / `EMPRÉSTIMO` / `FINAME - AMORTIZAÇÃO` / `FINAME - ESTORNO AMORT` + série de ref 827.9xx.xxx.xxx

### Consórcio

- Crédito = conta banco de origem; débito = conta consórcio
- Cada banco tem seu próprio código de consórcio (não compartilhado)
- `GatilhoConsorcio` deve rodar **antes** de `IsSelfTransfer`
- BB Consórcio: sempre débito=DEBITO/crédito=700 independente de sinal

---

## Restrições arquiteturais (vbaProject.bin corrompido — erro 429)

| Proibido | Alternativa |
|----------|-------------|
| `ThisWorkbook.Path` | `ActiveWorkbook.Path` |
| `WScript.Shell` / `CreateObject` | `Shell()` nativo + polling de `.done` |
| `Scripting.Dictionary` | `Collection` nativa do VBA |
| `XLOOKUP` | `INDEX/MATCH` (compatibilidade pré-365) |
| String raw de cmd para Python | Arquivo `.bat` temporário por invocação |

**Outras regras:** `Calc=Manual` envolve todas as escritas em bulk. `NormalizarTexto` usa `Chr()` para acentos. CNPJ: `NumberFormat = "@"` antes de gravar. Datas: `ParseDataBR()` com `DateSerial` (nunca `CDate`/`IsDate`). Números BR: `ParseNumeroCSV()` com detecção explícita de ponto/vírgula. Helpers declarados após declarações de módulo.

---

## Pendências conhecidas

| Item | Status |
|------|--------|
| Parsers Daycoval / Sicoob | Aguarda arquivo de amostra |
| OFX: concatenar CHECKNUM na Descrição | Fix mapeado, não implementado |
| `_REGRAS_DC` (pares débito/crédito por fórmula) | Nova aba necessária; `_MAPEAMENTO` não serve para isso |
| Cross-bank Transferência (match Collection entre abas) | Não construído |
| BB Rende Fácil conta 3882 (aplicação financeira) | Sem gatilho ainda |
| `ClassificarLinha`: excluir PIX recebido do próprio CNPJ FAM | `IsSelfTransfer` existe em Extracao mas não no gerador TXT |
| Preferência de beneficiário final em comprovante (vs. processador intermediário) | Aguarda PDF de amostra |
| Adiantamento: gatilho via texto ("Citi bank") | Frágil — marcado REVISAR |
| ReceitaWS API | Existia no fam-app, não portado para VBA |
| Fórmulas RESUMO com refs de coluna inteira | Atual usa ranges limitados (ex: `-2` manual) |
| Botões RESUMO | Podem sumir por sobrescrita do Drive Sync — recriar manualmente se necessário |
| Feedback loop: correções → DB | Não construído |
| Log de histórico de execução | Não construído |

### Verificações pendentes antes de ir para produção

- **Campo 006 (Histórico) realmente blank:** teste com um TXT com campo 6 vazio rodando no Único (layout oficial marca como obrigatório — em tensão com deixar blank)
- **Posição do CNPJ participante:** layout oficial PDF = campos 118/119; CELESP working file = campos 10/11. Confirmar qual instalação do SCI da FAM usa antes de construir escrita de CNPJ
- **Dropdown CODIGO:** v5 mostra "148,6575,5" — 6575 conflita com código folha confirmado 6775; reconciliar
- **Código participante FAM:** possivelmente 254.365.396; "FAM STTEL E LOCAÇÕES" (CNPJ 21430094000109) é entidade separada

---

## Padrões de texto em extratos

```
PAGAMENTO PIX {CNPJ/CPF} {NOME} - PIX_DEB     → fornecedor ou funcionário
LIQUIDACAO BOLETO {CNPJ} {NOME} -              → fornecedor
Pagamento de Boleto {NOME} - {NUMERO}          → nome do fornecedor no Histórico
CHEQUE COMPE SICREDI - ALW{NUMERO}             → cheque → consulta _CHEQUES
DEBITO TED/IB {CNPJ} {NOME} - I00{ref}        → TED para fornecedor
PIX - ENVIADO - dd/mm hh:mm NOME              → BB OFX (usar último " - ")
ESTORNO DE DÉBITO - BB ADMIN CONSÓRCIO SA      → BB Consórcio (funciona por coincidência)
C40xxxxxxx / C60xxxxxxx                        → número de contrato Sicredi
```

---

## Precisão (execução 2025 completa — referência histórica fam-app)

| Status | Linhas | % |
|--------|--------|---|
| OK |~92% |
| REVISAR | ~8% |
| NAO ENCONTRADO |~0% |

Fontes de identificação: CADASTRO (maioria), PDF (~2.000), COMPLEMENTO (~600), RECEITA_WS (~370), FOLHA (~178), CHEQUE (~66).
