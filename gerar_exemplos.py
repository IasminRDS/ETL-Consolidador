"""
Cria em  entrada/  planilhas propositalmente baguncadas, do jeito que chegam
do cliente de verdade: cabecalho diferente em cada arquivo, linha de titulo em
cima, numero em formato BR e US, data em tres formatos, duplicatas e lixo.

Serve para testar o etl_consolidador.py sem precisar de dados reais.
"""

import random
from pathlib import Path

random.seed(7)

ENTRADA = Path(__file__).parent / "entrada"
ENTRADA.mkdir(exist_ok=True)

PRODUTOS = [
    ("7891000100103", "Dipirona 500mg 20cp", "Genericos", 12.90),
    ("7891000100202", "Losartana 50mg 30cp", "Genericos", 18.50),
    ("7891000100305", "Protetor solar FPS 50", "Dermocosmeticos", 74.90),
    ("7891000100408", "Omeprazol 20mg 28cp", "Genericos", 22.40),
    ("7891000100500", "Amoxicilina 500mg 21cp", "Medicamentos", 41.80),
    ("7891000100604", "Vitamina D 2000UI 60cp", "Medicamentos", 48.60),
]
VENDEDORES = ["Ana", "Bruno", "Carla", "Diego"]


def brl(v):          # 1234.5 -> "1.234,50"
    return f"{v:,.2f}".replace(",", "@").replace(".", ",").replace("@", ".")


def usd(v):          # 1234.5 -> "1,234.50"
    return f"{v:,.2f}"


linhas_geradas = []


def vendas(n, mes, ano=2026):
    out = []
    for _ in range(n):
        cod, nome, cat, preco = random.choice(PRODUTOS)
        qtd = random.randint(1, 12)
        dia = random.randint(1, 28)
        out.append({
            "dia": dia, "mes": mes, "ano": ano, "cod": cod, "nome": nome,
            "cat": cat, "qtd": qtd, "preco": preco,
            "total": round(preco * qtd, 2),
            "vend": random.choice(VENDEDORES),
            "doc": f"CF{random.randint(10000, 99999)}",
        })
    linhas_geradas.extend(out)
    return out


# --------------------------------------------------------- arquivo 1 (BR, ;)
# Cabecalho em portugues abreviado, duas linhas de titulo no topo, decimal BR.
a1 = vendas(40, 5)
txt = ["DROGARIA SAO BENTO - RELATORIO DE VENDAS", "Periodo: MAIO/2026", "",
       "Data;Cupom;Cod Produto;Descricao;Grupo;Qtde;Preco Unitario;Valor Total;Vendedor"]
for r in a1:
    txt.append(f"{r['dia']:02d}/{r['mes']:02d}/{r['ano']};{r['doc']};{r['cod']};"
               f"{r['nome']};{r['cat']};{r['qtd']};{brl(r['preco'])};"
               f"{brl(r['total'])};{r['vend']}")
txt += ["", "TOTAL GERAL;;;;;;;" + brl(sum(r["total"] for r in a1)) + ";"]
(ENTRADA / "Vendas_MAIO_2026.csv").write_text(
    "\n".join(txt), encoding="latin-1", errors="ignore")

# ------------------------------------------------ arquivo 2 (US, virgula, ISO)
# Exportacao de outro sistema: nomes em ingles-ish, decimal com ponto, data ISO.
a2 = vendas(35, 6)
txt = ["DATA,NOTA,SKU,ITEM,DEPARTAMENTO,QUANT,UNITARIO,FATURAMENTO,OPERADOR"]
for r in a2:
    txt.append(f"{r['ano']}-{r['mes']:02d}-{r['dia']:02d},{r['doc']},{r['cod']},"
               f"\"{r['nome']}\",{r['cat']},{r['qtd']},{usd(r['preco'])},"
               f"{usd(r['total'])},{r['vend']}")
(ENTRADA / "vendas junho (2).csv").write_text("\n".join(txt), encoding="utf-8")

# ------------------------------------- arquivo 3 (sujo: R$, vazias, duplicatas)
a3 = vendas(30, 7)
repetidas = random.sample(a3, 6)          # duplicatas de verdade
txt = ["", "Controle interno - Julho", "",
       "DT EMISSAO;PEDIDO;REFERENCIA;MERCADORIA;FAMILIA;QT;PRECO VENDA;VLR TOTAL;ATENDENTE;OBS"]
for r in a3 + repetidas:
    txt.append(f"{r['dia']:02d}.{r['mes']:02d}.{r['ano']};{r['doc']};{r['cod']};"
               f"{r['nome']};{r['cat']};{r['qtd']};R$ {brl(r['preco'])};"
               f"R$ {brl(r['total'])};{r['vend']};")
txt.insert(8, ";;;;;;;;;")                 # linha em branco no meio
txt.append(";;;;;;;;;")                    # rodape vazio
(ENTRADA / "VENDAS-JULHO_final_v3.csv").write_text("\n".join(txt), encoding="utf-8")

# ------------------------------------- arquivo 4 (fora do padrao: sera pulado)
(ENTRADA / "anotacoes do gerente.csv").write_text(
    "Observacao;Responsavel\nPedir mais protetor solar;Ana\n", encoding="utf-8")

print(f"Criados 4 arquivos em {ENTRADA}")
print(f"Linhas de venda unicas geradas: {len(linhas_geradas)} "
      f"(+6 duplicatas propositais)")
print(f"Faturamento real esperado: R$ {brl(sum(r['total'] for r in linhas_geradas))}")
