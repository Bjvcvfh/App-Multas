import os
import pandas as pd
from datetime import datetime

from utils.helpers import parse_money_to_float, format_brl

MESSAGE_TEMPLATE = (
    "Bom dia {nome_motorista}, tudo bem?\n\n"
    "O senhor levou uma multa no dia {data_multa} em {cidade}/{uf} às {hora_multa} com a placa {placa}.\n\n"
    "Multa por {descricao_multa}, no valor de {valor_base} e {pontos} pontos na carteira.\n\n"
    "Preciso saber se posso indicar os pontos na sua carteira.\n"
    "Se indicar o valor da multa cai pra {valor_com_indicacao} e sem indicar o valor da multa sobe pra {valor_sem_indicacao}.\n\n"
    "Sobre o pagamento o senhor pode discutir com o RH sobre parcelamentos pra acertarem da melhor forma."
)

class MultaService:
    def __init__(self, motoristas_csv: str, tipos_multa_csv: str):
        if not os.path.exists(motoristas_csv):
            raise FileNotFoundError(f"motoristas.csv não encontrado: {motoristas_csv}")
        if not os.path.exists(tipos_multa_csv):
            raise FileNotFoundError(f"tipos_multa.csv não encontrado: {tipos_multa_csv}")

        self.motoristas_df = self._load_csv(motoristas_csv)
        self.tipos_df = self._load_csv(tipos_multa_csv)

        for col in ["Nome Curto", "TELEFONE"]:
            if col not in self.motoristas_df.columns:
                raise RuntimeError(f"motoristas.csv precisa ter coluna: {col}")

        for col in ["COD_MULTA", "DESCRICAO", "VALOR", "PONTOS", "GRAVIDADE"]:
            if col not in self.tipos_df.columns:
                raise RuntimeError(f"tipos_multa.csv precisa ter coluna: {col}")

        self.tipos_df["COD_MULTA"] = self.tipos_df["COD_MULTA"].astype(str).str.strip()
        self.tipos_df["_valor_float"] = self.tipos_df["VALOR"].apply(parse_money_to_float)

    def _load_csv(self, path: str) -> pd.DataFrame:
        try:
            df = pd.read_csv(path, sep=";", encoding="utf-8-sig")
        except UnicodeDecodeError:
            df = pd.read_csv(path, sep=";", encoding="latin-1")
        df.columns = df.columns.astype(str).str.replace("\ufeff", "", regex=False).str.strip()
        return df

    def listar_motoristas(self) -> list[str]:
        return self.motoristas_df["Nome Curto"].astype(str).tolist()

    def buscar_motorista(self, nome: str) -> dict:
        df = self.motoristas_df
        row = df.loc[df["Nome Curto"].astype(str) == str(nome)]
        if row.empty:
            raise RuntimeError("Motorista não encontrado no motoristas.csv")
        r = row.iloc[0]
        motorista_id = str(r["Cód. Motorista"]) if "Cód. Motorista" in df.columns else ""
        return {
            "motorista_id": motorista_id,
            "nome_motorista": str(r["Nome Curto"]),
            "telefone": str(r["TELEFONE"]),
        }

    def buscar_multa_por_cod(self, cod_multa: str) -> dict:
        row = self.tipos_df.loc[self.tipos_df["COD_MULTA"] == str(cod_multa).strip()]
        if row.empty:
            raise RuntimeError(f"COD_MULTA {cod_multa} não encontrado no tipos_multa.csv")
        m = row.iloc[0]
        valor_base = float(m["_valor_float"])
        pontos = int(m["PONTOS"])
        gravidade = str(m["GRAVIDADE"]).strip()
        return {
            "codigo_multa": str(m["COD_MULTA"]).strip(),
            "descricao_multa": str(m["DESCRICAO"]).strip(),
            "valor_base_num": valor_base,
            "valor_base": format_brl(valor_base),
            "pontos": pontos,
            "gravidade_multa": gravidade
        }

    def calcular_valores(self, valor_base: float) -> tuple[float, float]:
        # regras definidas por você:
        # com indicação = valor * 0,6
        # sem indicação = (valor * 0,8)*3
        v_com = round(valor_base * 0.6, 2)
        v_sem = round((valor_base * 0.8) * 3, 2)
        return v_com, v_sem

    def gerar_mensagem(self, motorista_nome: str, extracao: dict, multa_atual: dict) -> str:
        motor = self.buscar_motorista(motorista_nome)

        valor_base = float(multa_atual["valor_base_num"])
        v_com, v_sem = self.calcular_valores(valor_base)

        msg = MESSAGE_TEMPLATE.format(
            nome_motorista=motor["nome_motorista"],
            data_multa=extracao["data_multa"],
            cidade=extracao.get("cidade", ""),
            uf=extracao.get("uf", ""),
            hora_multa=extracao["hora_multa"],
            placa=extracao["placa"],
            descricao_multa=multa_atual["descricao_multa"],
            valor_base=format_brl(valor_base),
            pontos=int(multa_atual["pontos"]),
            valor_com_indicacao=format_brl(v_com),
            valor_sem_indicacao=format_brl(v_sem),
        )
        return msg