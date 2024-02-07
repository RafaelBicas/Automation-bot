import pandas as pd
import openpyxl

class Arquivo_Excel:
    def __init__(self, arquivo, label_col, value_col1, value_col2):
        """
        Inicializa a classe com um arquivo Excel e organiza seu conteúdo em um dicionário.

        Args:
            arquivo (str): O nome do arquivo Excel.
            label_col (str): O caractere que representa a coluna dos labels.
            value_col1 (str): O caractere que representa a primeira coluna dos valores.
            value_col2 (str): O caractere que representa a segunda coluna dos valores.
        """
        self.arquivo = arquivo
        self.label_col = label_col
        self.value_col1 = value_col1
        self.value_col2 = value_col2
        self.dicionario = self._ler_excel()

    def _ler_excel(self):
        try:
            # Lê o arquivo Excel
            df = pd.read_excel(self.arquivo)

            # Inicializa o dicionário
            dicionario = {}

            # Itera pelas linhas do DataFrame
            for index, row in df.iterrows():
                label = row[self.label_col]
                value1 = row[self.value_col1]
                value2 = row[self.value_col2]

                # Se o label ainda não estiver no dicionário, cria uma nova lista
                if label not in dicionario:
                    dicionario[label] = []

                # Adiciona os valores à lista correspondente
                dicionario[label].append([value1, value2])

            return dicionario
        except Exception as e:
            print(f"Erro ao ler o arquivo Excel: {str(e)}")
            return {}

    def get_dicionario(self):
        """
        Retorna o dicionário organizado a partir do arquivo Excel.

        Returns:
            dict: O dicionário onde as chaves são os labels e os valores são listas de pares de valores correspondentes.
        """
        return self.dicionario
