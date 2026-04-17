import csv


class DataImporter:
    """Responsável por carregar tabelas e extrair colunas/valores."""

    @staticmethod
    def formatar_excecao(exc: Exception, contexto: str) -> str:
        return f"{contexto}: {exc}"

    def carregar_tabela(self, caminho):
        if caminho.lower().endswith(".csv"):
            try:
                import pandas as pd

                return pd.read_csv(caminho)
            except ImportError:
                try:
                    with open(caminho, newline="", encoding="utf-8-sig") as arquivo:
                        return list(csv.DictReader(arquivo))
                except (OSError, UnicodeDecodeError, ValueError) as exc:
                    raise RuntimeError(self.formatar_excecao(exc, "Falha ao carregar CSV")) from exc
            except (OSError, UnicodeDecodeError, ValueError) as exc:
                raise RuntimeError(self.formatar_excecao(exc, "Falha ao carregar CSV")) from exc

        try:
            import pandas as pd

            return pd.read_excel(caminho)
        except ImportError as exc:
            raise RuntimeError(
                "Falha ao carregar Excel. Instale/repare 'pandas', 'numpy' e 'openpyxl' no ambiente."
            ) from exc
        except (OSError, ValueError) as exc:
            raise RuntimeError(self.formatar_excecao(exc, "Falha ao carregar Excel")) from exc

    @staticmethod
    def obter_colunas(tabela):
        if tabela is None:
            return []
        if hasattr(tabela, "columns"):
            return list(tabela.columns)
        if isinstance(tabela, list) and tabela:
            return list(tabela[0].keys())
        return []

    @staticmethod
    def obter_valores_coluna(tabela, coluna):
        if hasattr(tabela, "__getitem__") and hasattr(tabela, "columns"):
            return [str(v) for v in tabela[coluna].dropna().tolist()]
        if isinstance(tabela, list):
            vals = []
            for linha in tabela:
                valor = linha.get(coluna)
                if valor is not None and str(valor).strip() != "":
                    vals.append(str(valor))
            return vals
        return []
