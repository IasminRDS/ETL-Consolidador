"""
ETL Consolidador — junta as planilhas soltas do cliente em uma base unica e limpa.

O problema que ele resolve: o cliente tem "Vendas_janeiro.xlsx", "vendas fev (2).csv",
"VENDAS-MARCO_final_v3.xlsx" — cada um com nome de coluna diferente, linhas de titulo
no topo, numero em formato brasileiro e data em tres formatos. Alguem gasta 4 horas
por mes juntando isso na mao.

O que ele faz:
  1. Le tudo que estiver em  entrada/  (.csv, .xlsx, .xls)
  2. Acha sozinho a linha do cabecalho (ignora titulo/logo/linhas em branco)
  3. Traduz os nomes de coluna para um padrao unico (mapa de sinonimos abaixo)
  4. Converte "1.234,56" -> 1234.56 e "05/03/2026" -> 2026-03-05
  5. Remove duplicatas e linhas vazias
  6. Grava a base limpa + um relatorio de qualidade dizendo o que foi descartado

Uso:
    python etl_consolidador.py                  # entrada/ -> saida/
    python etl_consolidador.py --entrada C:\\dados --saida C:\\bi

Requisitos:  pip install pandas openpyxl
             (openpyxl so e necessario se houver arquivos .xlsx)
"""

from __future__ import annotations

import argparse
import re
import sys
import unicodedata
from datetime import datetime
from pathlib import Path

try:
    import pandas as pd
except ImportError:
    sys.exit("Falta a biblioteca pandas. Rode:  pip install pandas openpyxl")


# ---------------------------------------------------------------------------
# CONFIGURACAO — e aqui que voce adapta o script para cada cliente.
# Chave = nome padrao de saida.  Valor = todos os apelidos que ja apareceram.
# ---------------------------------------------------------------------------
MAPA_COLUNAS: dict[str, list[str]] = {
    "data":        ["data", "dt", "data venda", "data da venda", "emissao",
                    "data emissao", "dt emissao", "competencia"],
    "documento":   ["documento", "nf", "nota", "nota fiscal", "cupom",
                    "num cupom", "pedido", "num pedido"],
    "cliente":     ["cliente", "nome cliente", "razao social", "comprador"],
    "produto":     ["produto", "descricao", "descricao produto", "item",
                    "mercadoria", "nome produto"],
    "codigo":      ["codigo", "cod", "cod produto", "sku", "ean", "referencia"],
    "categoria":   ["categoria", "grupo", "familia", "linha", "departamento",
                    "secao"],
    "quantidade":  ["quantidade", "qtd", "qtde", "qt", "quant", "unidades"],
    "valor_unitario": ["valor unitario", "preco", "preco unitario", "vl unit",
                       "unitario", "preco venda"],
    "valor_total": ["valor total", "total", "valor", "vl total", "vlr total",
                    "faturamento", "receita", "valor liquido"],
    "custo":       ["custo", "custo total", "cmv", "valor custo", "custo unitario"],
    "vendedor":    ["vendedor", "operador", "atendente", "responsavel"],
    "forma_pagto": ["forma pagamento", "pagamento", "forma de pagamento",
                    "meio pagamento", "condicao"],
}

COLUNAS_NUMERICAS = ["quantidade", "valor_unitario", "valor_total", "custo"]
COLUNAS_DATA = ["data"]
# Uma linha sem NENHUMA destas colunas preenchidas e considerada lixo:
COLUNAS_OBRIGATORIAS = ["data", "valor_total"]
# Quantas colunas do mapa uma aba precisa ter para ser aceita como planilha
# de vendas (evita puxar a aba "anotacoes do gerente" junto):
MIN_COLUNAS_RECONHECIDAS = 3

SEPARADORES_CSV = [";", ",", "\t", "|"]
CODIFICACOES = ["utf-8-sig", "utf-8", "latin-1", "cp1252"]


# ---------------------------------------------------------------------------
# Normalizacao de texto
# ---------------------------------------------------------------------------
def normalizar(texto: object) -> str:
    """'Qtde. Vendida ' -> 'qtde vendida' (sem acento, sem pontuacao, minusculo)."""
    s = str(texto)
    s = unicodedata.normalize("NFKD", s).encode("ascii", "ignore").decode()
    s = s.lower().replace("_", " ")
    s = re.sub(r"[^a-z0-9 ]", " ", s)
    return re.sub(r"\s+", " ", s).strip()


_LOOKUP = {normalizar(apelido): padrao
           for padrao, apelidos in MAPA_COLUNAS.items()
           for apelido in apelidos}


def traduzir_colunas(colunas) -> tuple[dict[str, str], list[str]]:
    """Devolve ({original: padronizado}, [colunas que nao reconheci])."""
    renomear, desconhecidas = {}, []
    for c in colunas:
        chave = normalizar(c)
        if chave in _LOOKUP:
            renomear[c] = _LOOKUP[chave]
        else:
            desconhecidas.append(str(c))
    return renomear, desconhecidas


# ---------------------------------------------------------------------------
# Conversao de valores
# ---------------------------------------------------------------------------
def para_numero(valor: object) -> float | None:
    """Aceita 'R$ 1.234,56', '1,234.56', '(120,00)' e devolve float."""
    if valor is None or (isinstance(valor, float) and pd.isna(valor)):
        return None
    if isinstance(valor, (int, float)):
        return float(valor)

    s = str(valor).strip()
    if not s or s in {"-", "--", "n/a", "N/A"}:
        return None

    negativo = s.startswith("(") and s.endswith(")")
    s = re.sub(r"[^\d,.\-]", "", s)  # tira R$, espacos, %, etc.
    if not s or s in {"-", ".", ","}:
        return None

    # Decide quem e o separador decimal pelo ULTIMO simbolo que aparece.
    if "," in s and "." in s:
        if s.rfind(",") > s.rfind("."):      # 1.234,56  (padrao BR)
            s = s.replace(".", "").replace(",", ".")
        else:                                 # 1,234.56  (padrao US)
            s = s.replace(",", "")
    elif "," in s:
        # "1,5" e decimal; "1,500" com 3 casas costuma ser milhar
        inteiro, _, frac = s.rpartition(",")
        s = s.replace(",", "" if len(frac) == 3 and inteiro else ".")

    try:
        n = float(s)
    except ValueError:
        return None
    return -n if negativo else n


FORMATOS_DATA = ["%d/%m/%Y", "%d/%m/%y", "%Y-%m-%d", "%d-%m-%Y",
                 "%d.%m.%Y", "%Y/%m/%d", "%d/%m/%Y %H:%M:%S", "%Y-%m-%d %H:%M:%S"]


def para_data(valor: object):
    if valor is None or (isinstance(valor, float) and pd.isna(valor)):
        return None
    if isinstance(valor, (pd.Timestamp, datetime)):
        return pd.Timestamp(valor).normalize()

    s = str(valor).strip()
    if not s:
        return None
    for fmt in FORMATOS_DATA:
        try:
            return pd.Timestamp(datetime.strptime(s, fmt)).normalize()
        except ValueError:
            continue
    # Numero serial do Excel (dias desde 30/12/1899)
    try:
        n = float(s)
        if 20000 < n < 60000:
            return pd.Timestamp("1899-12-30") + pd.Timedelta(days=int(n))
    except ValueError:
        pass
    return None


# ---------------------------------------------------------------------------
# Leitura dos arquivos
# ---------------------------------------------------------------------------
def achar_cabecalho(bruto: pd.DataFrame, limite: int = 12) -> int:
    """Retorna o indice da linha que mais parece um cabecalho."""
    melhor_linha, melhor_nota = 0, -1
    for i in range(min(limite, len(bruto))):
        celulas = [c for c in bruto.iloc[i].tolist() if str(c).strip() not in ("", "nan")]
        if len(celulas) < 2:
            continue
        reconhecidas = sum(1 for c in celulas if normalizar(c) in _LOOKUP)
        # Cabecalho tem muitas celulas de texto e poucas numericas
        textuais = sum(1 for c in celulas if para_numero(c) is None)
        nota = reconhecidas * 3 + textuais
        if nota > melhor_nota:
            melhor_linha, melhor_nota = i, nota
    return melhor_linha


def ler_csv(caminho: Path) -> list[tuple[str, pd.DataFrame]]:
    """
    Le o CSV com o modulo csv, nao com pd.read_csv.

    Motivo: o pandas fixa a largura da tabela pela primeira linha do arquivo.
    Planilha de cliente quase sempre comeca com um titulo de uma celula so
    ("RELATORIO DE VENDAS"), e ai todas as linhas de dados viram "bad lines".
    Lendo na mao, cada linha e independente e a largura e o maximo encontrado.
    """
    import csv as _csv

    texto = None
    for enc in CODIFICACOES:
        try:
            texto = caminho.read_text(encoding=enc)
            break
        except (UnicodeDecodeError, LookupError):
            continue
    if texto is None:
        raise ValueError("nao consegui decodificar o arquivo em "
                         f"{'/'.join(CODIFICACOES)}")

    # Separador = o que mais aparece nas primeiras linhas com conteudo.
    amostra = [l for l in texto.splitlines()[:40] if l.strip()][:20]
    sep = max(SEPARADORES_CSV, key=lambda s: sum(l.count(s) for l in amostra))
    if not any(l.count(sep) for l in amostra):
        raise ValueError("nao identifiquei o separador de colunas")

    linhas = [l for l in _csv.reader(texto.splitlines(), delimiter=sep)]
    linhas = [l for l in linhas if any(str(c).strip() for c in l)]
    if len(linhas) < 2:
        raise ValueError("arquivo sem linhas de dados")

    largura = max(len(l) for l in linhas)
    linhas = [l + [""] * (largura - len(l)) for l in linhas]

    bruto = pd.DataFrame(linhas, dtype=str)
    pos = achar_cabecalho(bruto)
    df = bruto.iloc[pos + 1:].copy()
    df.columns = [str(c) for c in bruto.iloc[pos]]
    return [(caminho.name, df)]


def ler_excel(caminho: Path) -> list[tuple[str, pd.DataFrame]]:
    try:
        planilhas = pd.read_excel(caminho, sheet_name=None, header=None, dtype=str)
    except ImportError:
        raise ValueError("arquivo Excel exige 'pip install openpyxl'")
    resultado = []
    for aba, bruto in planilhas.items():
        if bruto.empty:
            continue
        linha = achar_cabecalho(bruto)
        df = bruto.iloc[linha + 1:].copy()
        df.columns = [str(c) for c in bruto.iloc[linha]]
        resultado.append((f"{caminho.name} [{aba}]", df))
    return resultado


def ler_arquivo(caminho: Path) -> list[tuple[str, pd.DataFrame]]:
    if caminho.suffix.lower() == ".csv":
        return ler_csv(caminho)
    if caminho.suffix.lower() in (".xlsx", ".xlsm", ".xls"):
        return ler_excel(caminho)
    raise ValueError(f"extensao nao suportada: {caminho.suffix}")


# ---------------------------------------------------------------------------
# Pipeline
# ---------------------------------------------------------------------------
def consolidar(entrada: Path, saida: Path) -> None:
    arquivos = sorted(p for p in entrada.iterdir()
                      if p.is_file() and not p.name.startswith("~$")
                      and p.suffix.lower() in (".csv", ".xlsx", ".xlsm", ".xls"))
    if not arquivos:
        sys.exit(f"Nenhuma planilha encontrada em {entrada}")

    partes: list[pd.DataFrame] = []
    log: list[str] = []
    total_lido = 0

    print(f"Lendo {len(arquivos)} arquivo(s) de {entrada}\n")
    for caminho in arquivos:
        try:
            blocos = ler_arquivo(caminho)
        except Exception as e:
            print(f"  [ERRO ] {caminho.name}: {e}")
            log.append(f"ERRO ao ler {caminho.name}: {e}")
            continue

        for origem, df in blocos:
            renomear, desconhecidas = traduzir_colunas(df.columns)
            # Uma aba so entra na base se parecer mesmo uma planilha de vendas:
            # pelo menos 3 colunas reconhecidas E uma das obrigatorias.
            reconhecidas = set(renomear.values())
            if len(reconhecidas) < MIN_COLUNAS_RECONHECIDAS or not (
                    reconhecidas & set(COLUNAS_OBRIGATORIAS)):
                print(f"  [PULA ] {origem}: nao parece planilha de vendas "
                      f"({len(reconhecidas)} coluna(s) reconhecida(s))")
                log.append(f"{origem}: ignorado — colunas vistas: "
                           f"{list(df.columns)[:8]}")
                continue

            df = df.rename(columns=renomear)
            df = df[[c for c in df.columns if c in MAPA_COLUNAS]]
            df = df.loc[:, ~df.columns.duplicated()]

            for col in COLUNAS_NUMERICAS:
                if col in df:
                    df[col] = df[col].map(para_numero)
            for col in COLUNAS_DATA:
                if col in df:
                    df[col] = df[col].map(para_data)
            for col in df.columns:
                if col not in COLUNAS_NUMERICAS + COLUNAS_DATA:
                    df[col] = df[col].astype(str).str.strip().replace(
                        {"nan": None, "None": "", "": None})

            df.insert(0, "origem", origem)
            total_lido += len(df)
            partes.append(df)

            if desconhecidas:
                log.append(f"{origem}: colunas ignoradas (sem mapeamento) -> "
                           f"{desconhecidas}")
            print(f"  [ OK  ] {origem}: {len(df)} linhas, "
                  f"{len(renomear)} colunas reconhecidas"
                  + (f", {len(desconhecidas)} ignoradas" if desconhecidas else ""))

    if not partes:
        sys.exit("\nNada foi consolidado. Confira o MAPA_COLUNAS no topo do script.")

    base = pd.concat(partes, ignore_index=True)

    # ---- limpeza final ----
    presentes = [c for c in COLUNAS_OBRIGATORIAS if c in base.columns]
    antes = len(base)
    if presentes:
        base = base.dropna(subset=presentes, how="all")
    vazias = antes - len(base)

    # Rodape de planilha ("TOTAL GERAL", "SUBTOTAL", "SOMA") vem com valor
    # preenchido e o resto vazio — se entrar na base, dobra o faturamento.
    antes = len(base)
    texto_cols = [c for c in ["produto", "documento", "cliente", "categoria",
                              "codigo", "vendedor"] if c in base.columns]
    if texto_cols:
        junto = base[texto_cols].fillna("").agg(" ".join, axis=1).map(normalizar)
        e_agregado = junto.str.contains(
            r"\b(?:total|subtotal|soma|totais|acumulado)\b", regex=True)
        sem_identificacao = junto.str.strip() == ""
        if "data" in base.columns:
            sem_identificacao &= base["data"].isna()
        base = base[~(e_agregado | sem_identificacao)]
    totais = antes - len(base)

    chave = [c for c in ["data", "documento", "produto", "quantidade", "valor_total"]
             if c in base.columns]
    antes = len(base)
    if chave:
        base = base.drop_duplicates(subset=chave, keep="first")
    duplicadas = antes - len(base)

    # Recalcula valor_total quando der, e marca as linhas suspeitas
    if {"quantidade", "valor_unitario"} <= set(base.columns):
        calculado = base["quantidade"] * base["valor_unitario"]
        if "valor_total" in base.columns:
            faltando = base["valor_total"].isna()
            base.loc[faltando, "valor_total"] = calculado[faltando]
            divergente = (
                base["valor_total"].notna() & calculado.notna()
                & ((base["valor_total"] - calculado).abs() > 0.02)
            )
            base["revisar"] = divergente.map({True: "valor divergente", False: ""})
        else:
            base["valor_total"] = calculado

    if "data" in base.columns:
        base = base.sort_values("data", na_position="last")
        base["ano"] = base["data"].dt.year.astype("Int64")   # 2026, nao 2026,00
        base["mes"] = base["data"].dt.strftime("%Y-%m")
        base["data"] = base["data"].dt.strftime("%Y-%m-%d")

    ordem = (["origem"] + [c for c in MAPA_COLUNAS if c in base.columns]
             + [c for c in ["ano", "mes", "revisar"] if c in base.columns])
    base = base[ordem]

    # ---- gravacao ----
    saida.mkdir(parents=True, exist_ok=True)
    destino = saida / "base_consolidada.csv"
    base.to_csv(destino, index=False, sep=";", encoding="utf-8-sig",
                decimal=",", float_format="%.2f")

    a_revisar = int((base["revisar"] != "").sum()) if "revisar" in base else 0
    relatorio = [
        "RELATORIO DE QUALIDADE — ETL Consolidador",
        f"Gerado em {datetime.now():%d/%m/%Y %H:%M}",
        "",
        f"Arquivos lidos .............. {len(arquivos)}",
        f"Linhas brutas ............... {total_lido}",
        f"Linhas vazias descartadas ... {vazias}",
        f"Linhas de total/rodape ...... {totais}",
        f"Linhas duplicadas removidas . {duplicadas}",
        f"Linhas na base final ........ {len(base)}",
        f"Linhas marcadas p/ revisao .. {a_revisar}",
        "",
        "Colunas da base final:",
        "  " + ", ".join(base.columns),
        "",
    ]
    if "valor_total" in base.columns:
        soma = f"{base['valor_total'].sum():,.2f}"
        soma = soma.replace(",", "@").replace(".", ",").replace("@", ".")
        relatorio += [f"Soma de valor_total ......... R$ {soma}", ""]
    if log:
        relatorio += ["Ocorrencias:", *(f"  - {l}" for l in log)]
    else:
        relatorio += ["Nenhuma ocorrencia."]

    (saida / "relatorio_qualidade.txt").write_text(
        "\n".join(relatorio), encoding="utf-8")

    print(f"\nBase final: {len(base)} linhas  ->  {destino}")
    print(f"Relatorio:  {saida / 'relatorio_qualidade.txt'}")
    if a_revisar:
        print(f"Atencao: {a_revisar} linha(s) com valor_total divergente "
              f"de quantidade x preco (coluna 'revisar').")


def main() -> None:
    aqui = Path(__file__).parent
    p = argparse.ArgumentParser(description="Consolida planilhas em uma base unica.")
    p.add_argument("--entrada", type=Path, default=aqui / "entrada")
    p.add_argument("--saida", type=Path, default=aqui / "saida")
    args = p.parse_args()

    if not args.entrada.exists():
        sys.exit(f"Pasta de entrada nao existe: {args.entrada}")
    consolidar(args.entrada, args.saida)


if __name__ == "__main__":
    main()
