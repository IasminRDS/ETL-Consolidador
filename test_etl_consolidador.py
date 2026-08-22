"""
Testes do ETL Consolidador.

Cada caso corresponde a uma linha da tabela "O que ele trata sozinho" do
README — a bagunça que aparece nas planilhas do cliente e o que o script
deve fazer com ela. O teste de ponta a ponta roda o pipeline inteiro sobre
os arquivos que `gerar_exemplos.py` cria, em pastas temporárias.

    python -m unittest -v
"""

import shutil
import subprocess
import sys
import tempfile
import unittest
from datetime import datetime
from pathlib import Path

import pandas as pd

import etl_consolidador as etl


class TesteNumero(unittest.TestCase):
    """`R$ 1.234,56` e `1,234.56` no mesmo lote."""

    def test_padrao_brasileiro(self):
        self.assertAlmostEqual(etl.para_numero("R$ 1.234,56"), 1234.56)

    def test_padrao_americano(self):
        self.assertAlmostEqual(etl.para_numero("1,234.56"), 1234.56)

    def test_decimal_simples_com_virgula(self):
        self.assertAlmostEqual(etl.para_numero("1,5"), 1.5)

    def test_virgula_de_milhar_sem_decimal(self):
        # "1,500" com três casas depois da vírgula é milhar, não 1.5
        self.assertAlmostEqual(etl.para_numero("1,500"), 1500.0)

    def test_negativo_entre_parenteses(self):
        self.assertAlmostEqual(etl.para_numero("(120,00)"), -120.0)

    def test_simbolos_e_espacos_sao_ignorados(self):
        self.assertAlmostEqual(etl.para_numero("  R$  89,90 "), 89.90)

    def test_numero_ja_numerico_passa_direto(self):
        self.assertAlmostEqual(etl.para_numero(42), 42.0)

    def test_vazios_e_marcadores_viram_none(self):
        for v in ("", "   ", "-", "n/a", "N/A", None):
            self.assertIsNone(etl.para_numero(v), f"esperava None para {v!r}")

    def test_texto_puro_vira_none(self):
        self.assertIsNone(etl.para_numero("TOTAL GERAL"))


class TesteData(unittest.TestCase):
    """`05/03/2026`, `2026-03-05`, `05.03.2026` e o serial do Excel."""

    esperado = pd.Timestamp("2026-03-05")

    def test_formatos_equivalentes(self):
        for texto in ("05/03/2026", "2026-03-05", "05-03-2026",
                      "05.03.2026", "2026/03/05"):
            self.assertEqual(etl.para_data(texto), self.esperado, texto)

    def test_data_com_hora_e_normalizada(self):
        self.assertEqual(etl.para_data("05/03/2026 14:30:00"), self.esperado)

    def test_serial_do_excel(self):
        serial = (self.esperado - pd.Timestamp("1899-12-30")).days
        self.assertEqual(etl.para_data(str(serial)), self.esperado)

    def test_datetime_nativo(self):
        self.assertEqual(etl.para_data(datetime(2026, 3, 5, 9, 0)), self.esperado)

    def test_texto_invalido_vira_none(self):
        self.assertIsNone(etl.para_data("sem data"))
        self.assertIsNone(etl.para_data(""))


class TesteColunas(unittest.TestCase):
    """`Qtde`, `QT`, `Quant`, `Unidades` viram todos `quantidade`."""

    def test_apelidos_caem_no_campo_padrao(self):
        mapa, _ = etl.traduzir_colunas(["Qtde", "VALOR TOTAL", "Data"])
        self.assertEqual(mapa.get("Qtde"), "quantidade")
        self.assertEqual(mapa.get("VALOR TOTAL"), "valor_total")
        self.assertEqual(mapa.get("Data"), "data")

    def test_acento_e_caixa_nao_importam(self):
        mapa, _ = etl.traduzir_colunas(["QUANTIDADE", "quantidade", "Quantidade"])
        self.assertTrue(all(v == "quantidade" for v in mapa.values()))

    def test_coluna_desconhecida_fica_de_fora(self):
        mapa, desconhecidas = etl.traduzir_colunas(["Data", "Coluna Estranha"])
        self.assertNotIn("Coluna Estranha", mapa)
        self.assertIn("Coluna Estranha", desconhecidas)


class TesteCabecalho(unittest.TestCase):
    """Título e logo nas primeiras linhas: achar sozinho a linha do cabeçalho."""

    def test_pula_titulo_no_topo(self):
        bruto = pd.DataFrame([
            ["RELATÓRIO DE VENDAS", None, None],
            [None, None, None],
            ["Data", "Produto", "Valor Total"],
            ["05/03/2026", "Mouse", "R$ 89,90"],
        ])
        self.assertEqual(etl.achar_cabecalho(bruto), 2)

    def test_cabecalho_ja_na_primeira_linha(self):
        bruto = pd.DataFrame([
            ["Data", "Produto", "Valor Total"],
            ["05/03/2026", "Mouse", "R$ 89,90"],
        ])
        self.assertEqual(etl.achar_cabecalho(bruto), 0)


class TestePipeline(unittest.TestCase):
    """Ponta a ponta sobre os arquivos bagunçados de `gerar_exemplos.py`."""

    @classmethod
    def setUpClass(cls):
        # gerar_exemplos.py é um script: roda na importação e escreve sempre
        # em entrada/, ao lado dele. Chamamos por subprocess para não
        # importar efeito colateral, e mandamos a saída para uma pasta
        # temporária. A semente é fixa (random.seed(7)), então os números
        # abaixo são estáveis.
        raiz = Path(__file__).resolve().parent
        subprocess.run([sys.executable, str(raiz / "gerar_exemplos.py")],
                       check=True, capture_output=True)
        cls.entrada = raiz / "entrada"
        cls.tmp = tempfile.mkdtemp()
        cls.saida = Path(cls.tmp) / "saida"
        etl.consolidar(cls.entrada, cls.saida)
        cls.base = pd.read_csv(
            cls.saida / "base_consolidada.csv", sep=";", decimal=",", encoding="utf-8-sig"
        )

    @classmethod
    def tearDownClass(cls):
        shutil.rmtree(cls.tmp, ignore_errors=True)

    def test_gera_os_dois_arquivos_de_saida(self):
        self.assertTrue((self.saida / "base_consolidada.csv").exists())
        self.assertTrue((self.saida / "relatorio_qualidade.txt").exists())

    def test_numeros_do_readme(self):
        # O README promete 105 linhas e R$ 22.632,70 para estes exemplos.
        self.assertEqual(len(self.base), 105)
        self.assertAlmostEqual(self.base["valor_total"].sum(), 22632.70, places=2)

    def test_linha_de_total_geral_e_descartada(self):
        # Se o rodapé "TOTAL GERAL" entrasse, o faturamento dobraria.
        texto = " ".join(str(v) for v in self.base.astype(str).values.ravel())
        self.assertNotIn("TOTAL GERAL", texto.upper())

    def test_sem_duplicatas_entre_arquivos(self):
        self.assertFalse(self.base.duplicated().any())

    def test_datas_saem_em_iso(self):
        datas = self.base["data"].dropna().astype(str)
        self.assertTrue((datas.str.match(r"^\d{4}-\d{2}-\d{2}")).all())

    def test_colunas_obrigatorias_presentes_e_preenchidas(self):
        for coluna in etl.COLUNAS_OBRIGATORIAS:
            self.assertIn(coluna, self.base.columns)
            self.assertFalse(self.base[coluna].isna().any(), coluna)

    def test_valores_sao_numericos(self):
        self.assertTrue(pd.api.types.is_numeric_dtype(self.base["valor_total"]))

    def test_aba_que_nao_e_de_vendas_fica_no_relatorio(self):
        relatorio = (self.saida / "relatorio_qualidade.txt").read_text(encoding="utf-8")
        self.assertTrue(len(relatorio.strip()) > 0)


if __name__ == "__main__":
    unittest.main(verbosity=2)
