# ETL Consolidador

Junta as planilhas soltas do cliente em **uma base única e limpa**, pronta para o
Power BI.

O problema que ele resolve: o cliente tem `Vendas_janeiro.xlsx`,
`vendas fev (2).csv` e `VENDAS-MARCO_final_v3.xlsx` — cada um com nome de coluna
diferente, linha de título no topo, número em formato brasileiro e data em três
formatos. Alguém gasta 4 horas por mês juntando isso na mão.

## Como usar

```bash
pip install pandas openpyxl
```

1. Jogue as planilhas do cliente na pasta `entrada/`
2. Rode:

```bash
python etl_consolidador.py
```

3. Pegue o resultado em `saida/`:
   - `base_consolidada.csv` — a base limpa (`;` como separador, UTF-8 com BOM,
     decimal com vírgula → abre direto no Excel e no Power BI)
   - `relatorio_qualidade.txt` — o que foi lido, descartado e por quê

Para apontar outras pastas:

```bash
python etl_consolidador.py --entrada "C:\dados do cliente" --saida "C:\bi"
```

## O que ele trata sozinho

| Bagunça na origem | O que o script faz |
|---|---|
| `Qtde`, `QT`, `Quant`, `Unidades` | traduz tudo para `quantidade` |
| Título e logo nas primeiras linhas | acha sozinho a linha do cabeçalho |
| `R$ 1.234,56` e `1,234.56` no mesmo lote | converte os dois para número |
| `05/03/2026`, `2026-03-05`, `05.03.2026`, serial do Excel | vira data ISO |
| Linha `TOTAL GERAL` no rodapé | descarta (senão dobra o faturamento) |
| Registros repetidos entre arquivos | remove duplicatas |
| Encoding latin-1 / cp1252 / UTF-8 | detecta automaticamente |
| Separador `;` `,` tab `\|` | detecta automaticamente |
| Aba que não é de vendas | ignora e registra no relatório |
| `valor_total` faltando | recalcula por `quantidade × valor_unitario` |
| `valor_total` que não bate com a conta | marca na coluna `revisar` |

## Adaptando para cada cliente

Praticamente tudo que muda de cliente para cliente está no topo do arquivo,
em `MAPA_COLUNAS`. Achou uma coluna nova (`"VL. LIQUIDO"`)? Acrescente o apelido
na lista do campo padrão correspondente:

```python
"valor_total": ["valor total", "total", "valor", "vl total", "vl liquido"],
```

Não precisa mexer em mais nada. As outras constantes ajustáveis:

- `COLUNAS_OBRIGATORIAS` — linha sem nenhuma delas é considerada lixo
- `MIN_COLUNAS_RECONHECIDAS` — quantas colunas do mapa uma aba precisa ter para
  ser aceita como planilha de vendas (padrão: 3)

## Testando sem dados reais

```bash
python gerar_exemplos.py
```

Cria em `entrada/` quatro arquivos propositalmente bagunçados (um em latin-1 com
decimal brasileiro, um em UTF-8 padrão americano, um cheio de `R$` e duplicatas,
e um que nem é planilha de vendas). Rodando o ETL em cima deles o resultado
esperado é **105 linhas** e **R$ 22.632,70** — os mesmos números que o
`gerar_exemplos.py` imprime ao criar os arquivos.

## Colocando no automático

Para rodar sozinho toda segunda-feira às 7h, agende no Windows:

```bash
schtasks /create /tn "ETL Cliente" /tr "python \"E:\mais projetos\negocio-bi\02-etl-consolidador\etl_consolidador.py\"" /sc weekly /d MON /st 07:00
```

O Power BI então só precisa apontar para `saida/base_consolidada.csv` e atualizar.
