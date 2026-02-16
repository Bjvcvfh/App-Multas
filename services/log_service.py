import os
import pandas as pd

class LogService:
    def __init__(self, log_csv_path: str):
        self.path = log_csv_path
        os.makedirs(os.path.dirname(self.path), exist_ok=True)

        # Cabeçalho fixo que você pediu
        self.columns = [
            "id_registro", "data_registro",
            "motorista_id", "nome_motorista", "telefone",
            "placa", "uf", "cidade",
            "data_multa", "hora_multa",
            "codigo_multa", "descricao_multa",
            "valor_base", "pontos",
            "valor_com_indicacao", "valor_sem_indicacao",
            "decisao_indicar", "gravidade_multa",
        ]

    def registrar(self, row: dict):
        # garante as colunas, mesmo se faltar algo (evita quebrar Power BI)
        out = {c: row.get(c, "") for c in self.columns}
        df = pd.DataFrame([out])
        file_exists = os.path.exists(self.path)
        df.to_csv(self.path, mode="a", header=not file_exists, index=False, encoding="utf-8")