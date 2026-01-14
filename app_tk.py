import os
import re
import sys
import uuid
import shutil
import tempfile
import subprocess
from datetime import datetime
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from pathlib import Path

import pandas as pd


# =========================
# CONFIG
# =========================
BASE_DIR = getattr(sys, "_MEIPASS", os.path.abspath(os.path.dirname(__file__)))

DATA_DIR = os.path.join(BASE_DIR, "data")
TEMPLATES_DIR = os.path.join(BASE_DIR, "templates")
OUTPUT_DIR = os.path.join(BASE_DIR, "output")

MOTORISTAS_CSV = os.path.join(DATA_DIR, "motoristas.csv")
TIPOS_MULTA_CSV = os.path.join(DATA_DIR, "tipos_multa.csv")
TERMO_TEMPLATE_DOCX = os.path.join(TEMPLATES_DIR, "termo_multa_modelo.docx")

DOWNLOADS_DIR = Path.home() / "Downloads"
LOG_CSV_PATH = os.path.join(OUTPUT_DIR, "logs_multas.csv")

MESSAGE_TEMPLATE = (
    "Bom dia {nome_motorista}, tudo bem?\n\n"
    "O senhor levou uma multa no dia {data_multa} em {cidade}/{uf} às {hora_multa} com a placa {placa}.\n\n"
    "Multa por {descricao_multa}, no valor de {valor_base} e {pontos} pontos na carteira.\n\n"
    "Preciso saber se posso indicar os pontos na sua carteira.\n"
    "Se indicar o valor da multa cai pra {valor_com_indicacao} e sem indicar o valor da multa sobe pra {valor_sem_indicacao}.\n\n"
    "Sobre o pagamento o senhor pode discutir com o RH sobre parcelamentos pra acertarem da melhor forma."
)


# =========================
# UTILITÁRIOS
# =========================
def append_log_csv(path: str, row: dict):
    df = pd.DataFrame([row])
    file_exists = os.path.exists(path)
    df.to_csv(
        path,
        mode="a",
        header=not file_exists,
        index=False,
        encoding="utf-8"
    )


def resource_path(relative_path: str) -> str:
    base = getattr(sys, "_MEIPASS", os.path.abspath(os.path.dirname(__file__)))
    return os.path.join(base, relative_path)


def ensure_dirs():
    os.makedirs(DATA_DIR, exist_ok=True)
    os.makedirs(TEMPLATES_DIR, exist_ok=True)
    os.makedirs(OUTPUT_DIR, exist_ok=True)


def parse_money_to_float(x) -> float:
    if pd.isna(x):
        return 0.0
    s = str(x).strip().replace("R$", "").replace(" ", "")
    s = re.sub(r"[^0-9,\.\-]", "", s)
    if s.count(",") > 0 and s.count(".") > 0:
        s = s.replace(".", "").replace(",", ".")
    elif s.count(",") > 0 and s.count(".") == 0:
        s = s.replace(",", ".")
    try:
        return float(s)
    except:
        return 0.0


def format_brl(v: float) -> str:
    s = f"{v:,.2f}"
    s = s.replace(",", "X").replace(".", ",").replace("X", ".")
    return f"R$ {s}"


def sanitize_filename(name: str) -> str:
    name = re.sub(r"[^\w\-. ]", "", name, flags=re.UNICODE).strip()
    name = name.replace(" ", "_")
    return name[:120] if len(name) > 120 else name


def data_por_extenso_ptbr(dt: datetime) -> str:
    meses = [
        "janeiro", "fevereiro", "março", "abril", "maio", "junho",
        "julho", "agosto", "setembro", "outubro", "novembro", "dezembro"
    ]
    return f"{dt.day} de {meses[dt.month - 1].capitalize()} de {dt.year}"


def load_csv(path: str, sep=";") -> pd.DataFrame:
    if not os.path.exists(path):
        raise FileNotFoundError(path)
    try:
        df = pd.read_csv(path, sep=sep, encoding="utf-8-sig")
    except UnicodeDecodeError:
        df = pd.read_csv(path, sep=sep, encoding="latin-1")
    df.columns = df.columns.astype(str).str.replace("\ufeff", "", regex=False).str.strip()
    return df


def extrair_texto_pdf(pdf_path: str) -> str:
    import pdfplumber
    parts = []
    with pdfplumber.open(pdf_path) as pdf:
        for page in pdf.pages:
            parts.append(page.extract_text() or "")
    return "\n".join(parts)


def extrair_cidade_uf_por_linhas(text: str):
    lines = [re.sub(r"\s+", " ", ln).strip() for ln in text.splitlines()]
    lines = [ln for ln in lines if ln]

    header_idx = None
    for i, ln in enumerate(lines):
        if "NOME DO MUNICIPIO UF" in ln.upper() or "NOME DO MUNICÍPIO UF" in ln.upper():
            header_idx = i
            break

    if header_idx is None:
        return None, None

    for j in range(header_idx + 1, min(header_idx + 8, len(lines))):
        cand = lines[j]
        if ")" in cand:
            cand = cand.split(")", 1)[1]
            cand = cand.lstrip()

        m = re.match(r"^(.+?)\s+([A-Z]{2})$", cand.upper())
        if m:
            cidade = m.group(1).title().strip()
            uf = m.group(2).strip().upper()
            return cidade, uf

        m2 = re.search(r"(.+?)\s+([A-Z]{2})\b$", cand.upper())
        if m2:
            cidade = m2.group(1).title().strip()
            uf = m2.group(2).strip().upper()
            return cidade, uf

    return None, None


def extrair_campos_notificacao(pdf_path: str) -> dict:
    text = extrair_texto_pdf(pdf_path)

    placa = None
    m = re.search(
        r"\bPLACA\b.*?\n\s*([A-Z]{3}[0-9A-Z][0-9]{2}[0-9A-Z]|[A-Z]{3}\s*\-?\s*\d{4}|[A-Z0-9]{7})\b",
        text,
        re.IGNORECASE
    )
    if m:
        placa = re.sub(r"\s|\-", "", m.group(1)).upper()

    data_multa = None
    hora_multa = None
    m = re.search(
        r"DATA\s+HORA.*?\b(\d{2}/\d{2}/\d{4})\s+(\d{2}:\d{2})\b",
        text,
        re.IGNORECASE | re.DOTALL
    )
    if m:
        data_multa, hora_multa = m.group(1), m.group(2)

    codigo_4d = None
    desdobramento = None
    valor_pdf = None
    m = re.search(
        r"C[ÓO]DIGO\s+DA\s+INFRA[CÇ][AÃ]O\s+DESDOBRAMENTO\s+VALOR\s+DA\s+MULTA\s*\n\s*(\d{4})\s+(\d)\s+(R\$\s*[0-9\.\,]+)",
        text,
        re.IGNORECASE
    )
    if m:
        codigo_4d, desdobramento, valor_pdf = m.group(1), m.group(2), m.group(3)

    cidade, uf = extrair_cidade_uf_por_linhas(text)

    descricao_pdf = ""
    m = re.search(
        r"DESCRI[CÇ][AÃ]O\s+DA\s+INFRA[CÇ][AÃ]O\s*\n\s*(.+?)(?:\n[A-Z ]{5,}|$)",
        text,
        re.IGNORECASE | re.DOTALL
    )
    if m:
        descricao_pdf = " ".join(m.group(1).split())

    missing = [k for k, v in {
        "placa": placa,
        "data_multa": data_multa,
        "hora_multa": hora_multa,
        "codigo_infracao": codigo_4d,
        "desdobramento": desdobramento,
        "valor_pdf": valor_pdf
    }.items() if not v]

    if missing:
        raise RuntimeError(
            "Não consegui extrair do PDF os campos:\n"
            f"{missing}\n\n"
            "📄 TEXTO EXTRAÍDO DO PDF (para diagnóstico):\n\n"
            + text[:6000]
        )

    return {
        "placa": placa,
        "data_multa": data_multa,
        "hora_multa": hora_multa,
        "cidade": cidade or "",
        "uf": uf or "",
        "codigo_4d": codigo_4d,
        "desdobramento": desdobramento,
        "valor_pdf": valor_pdf,
        "descricao_pdf": descricao_pdf
    }


def codigo_pdf_para_cod_multa(codigo_4d: str, desdobramento: str) -> str:
    base = codigo_4d[:3]
    sufixo = codigo_4d[3:]
    return f"{base}-{sufixo}{desdobramento}"


def gerar_termo_docx(template_docx: str, context: dict, out_docx_path: str):
    from docxtpl import DocxTemplate
    doc = DocxTemplate(template_docx)
    doc.render(context)
    doc.save(out_docx_path)
    return out_docx_path


def docx_to_pdf(docx_path: str, pdf_path: str):
    os.makedirs(os.path.dirname(pdf_path), exist_ok=True)

    # 1) Windows + Word via docx2pdf
    try:
        if sys.platform.startswith("win"):
            from docx2pdf import convert
            import pythoncom
            pythoncom.CoInitialize()
            try:
                convert(docx_path, pdf_path)
            finally:
                pythoncom.CoUninitialize()

            if os.path.exists(pdf_path):
                return pdf_path
    except Exception:
        pass

    # 2) LibreOffice (soffice)
    out_dir = os.path.dirname(pdf_path)
    cmd = ["soffice", "--headless", "--convert-to", "pdf", "--outdir", out_dir, docx_path]
    try:
        subprocess.run(cmd, check=True, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
    except Exception as e:
        raise RuntimeError(
            "Falha ao converter DOCX para PDF.\n"
            "Instale Word+docx2pdf (Windows) OU LibreOffice (soffice no PATH).\n"
            f"Erro: {e}"
        )

    base = os.path.splitext(os.path.basename(docx_path))[0]
    generated = os.path.join(out_dir, base + ".pdf")
    if not os.path.exists(generated):
        raise RuntimeError("LibreOffice rodou, mas o PDF não foi encontrado.")
    os.replace(generated, pdf_path)
    return pdf_path


def merge_pdfs(pdf_paths: list, out_path: str):
    from pypdf import PdfWriter, PdfReader
    writer = PdfWriter()
    for p in pdf_paths:
        reader = PdfReader(p)
        for page in reader.pages:
            writer.add_page(page)
    with open(out_path, "wb") as f:
        writer.write(f)
    return out_path


def open_downloads_folder():
    try:
        if sys.platform.startswith("win"):
            os.startfile(str(DOWNLOADS_DIR))
    except Exception:
        pass


# =========================
# APP TKINTER
# =========================
class MultasApp(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("App Multas — PDF + Termo")
        self.geometry("820x600")
        self.resizable(False, False)

        try:
            self.iconbitmap(resource_path("assets/app.ico"))
        except Exception:
            pass

        ensure_dirs()

        self.pdf_path = None
        self.extracao = None
        self.motoristas_df = None
        self.tipos_df = None

        self._load_data()
        self._build_ui()

    def _load_data(self):
        if not os.path.exists(TERMO_TEMPLATE_DOCX):
            messagebox.showerror("Template ausente", f"Template não encontrado:\n{TERMO_TEMPLATE_DOCX}")
            self.destroy()
            return

        try:
            self.motoristas_df = load_csv(MOTORISTAS_CSV)
            self.tipos_df = load_csv(TIPOS_MULTA_CSV)
        except Exception as e:
            messagebox.showerror("Erro", f"Erro ao ler CSV:\n{e}")
            self.destroy()
            return

        for col in ["Nome Curto", "TELEFONE"]:
            if col not in self.motoristas_df.columns:
                messagebox.showerror("Erro", f"motoristas.csv precisa ter coluna: {col}")
                self.destroy()
                return

        for col in ["COD_MULTA", "DESCRICAO", "VALOR", "PONTOS", "GRAVIDADE"]:
            if col not in self.tipos_df.columns:
                messagebox.showerror("Erro", f"tipos_multa.csv precisa ter coluna: {col}")
                self.destroy()
                return

        self.tipos_df["COD_MULTA"] = self.tipos_df["COD_MULTA"].astype(str).str.strip()
        self.tipos_df["_valor_float"] = self.tipos_df["VALOR"].apply(parse_money_to_float)

    def _build_ui(self):
        pad = 10

        top = ttk.Frame(self, padding=pad)
        top.pack(fill="x")
        ttk.Label(top, text="1) Anexar Notificação (PDF)", font=("Segoe UI", 11, "bold")).pack(anchor="w")

        row = ttk.Frame(top)
        row.pack(fill="x", pady=(6, 0))

        self.pdf_label = ttk.Label(row, text="Nenhum PDF selecionado", width=70)
        self.pdf_label.pack(side="left")
        ttk.Button(row, text="Selecionar PDF", command=self.on_select_pdf).pack(side="right")

        mid = ttk.Frame(self, padding=pad)
        mid.pack(fill="x")

        ttk.Label(mid, text="2) Motorista", font=("Segoe UI", 11, "bold")).pack(anchor="w")

        self.motorista_var = tk.StringVar()
        motoristas_list = self.motoristas_df["Nome Curto"].astype(str).tolist()
        self.motorista_combo = ttk.Combobox(mid, textvariable=self.motorista_var, values=motoristas_list, state="readonly", width=60)
        self.motorista_combo.pack(anchor="w", pady=(6, 0))
        if motoristas_list:
            self.motorista_combo.current(0)

        ttk.Label(mid, text="3) Indicar pontos?", font=("Segoe UI", 11, "bold")).pack(anchor="w", pady=(12, 0))
        self.indicar_var = tk.StringVar(value="SIM")

        rrow = ttk.Frame(mid)
        rrow.pack(anchor="w", pady=(6, 0))
        ttk.Radiobutton(rrow, text="SIM", variable=self.indicar_var, value="SIM").pack(side="left", padx=(0, 12))
        ttk.Radiobutton(rrow, text="NÃO", variable=self.indicar_var, value="NÃO").pack(side="left")

        bottom = ttk.Frame(self, padding=pad)
        bottom.pack(fill="both", expand=True)

        ttk.Label(bottom, text="Prévia dos dados extraídos do PDF", font=("Segoe UI", 11, "bold")).pack(anchor="w")

        self.preview = tk.Text(bottom, height=14, wrap="word")
        self.preview.pack(fill="both", expand=True, pady=(6, 0))
        self.preview.insert("1.0", "Anexe um PDF para extrair automaticamente os campos.\n")

        actions = ttk.Frame(self, padding=pad)
        actions.pack(fill="x")

        self.btn_msg = ttk.Button(actions, text="📝 Gerar somente Mensagem (TXT)", command=self.on_generate_message)
        self.btn_msg.pack(fill="x", pady=(0, 8))

        self.btn_pdf = ttk.Button(actions, text="📄 Gerar PDF Final (Termo + Notificação)", command=self.on_generate_pdf_final)
        self.btn_pdf.pack(fill="x")

    def on_select_pdf(self):
        path = filedialog.askopenfilename(
            title="Selecione a notificação (PDF)",
            filetypes=[("PDF", "*.pdf")]
        )
        if not path:
            return

        self.pdf_path = path
        self.pdf_label.config(text=os.path.basename(path))

        try:
            self.extracao = extrair_campos_notificacao(path)
            self._render_preview(self.extracao)
        except Exception as e:
            self.extracao = None
            messagebox.showerror("Erro ao ler PDF", str(e))
            self.preview.delete("1.0", "end")
            self.preview.insert("1.0", "Não foi possível extrair dados do PDF.\n")

    def _render_preview(self, info: dict):
        cod_multa = codigo_pdf_para_cod_multa(info["codigo_4d"], info["desdobramento"])

        self.preview.delete("1.0", "end")
        self.preview.insert("1.0", f"Placa: {info['placa']}\n")
        self.preview.insert("end", f"Data: {info['data_multa']}\n")
        self.preview.insert("end", f"Hora: {info['hora_multa']}\n")
        self.preview.insert("end", f"Cidade/UF: {info['cidade']}/{info['uf']}\n")
        self.preview.insert("end", f"Código da infração (PDF): {info['codigo_4d']}\n")
        self.preview.insert("end", f"Desdobramento: {info['desdobramento']}\n")
        self.preview.insert("end", f"COD_MULTA (interno): {cod_multa}\n")
        self.preview.insert("end", f"Valor (PDF): {info['valor_pdf']}\n")
        self.preview.insert("end", f"Descrição (PDF): {info['descricao_pdf']}\n")

    def _get_motorista(self):
        motorista_nome = self.motorista_var.get().strip()
        if not motorista_nome:
            raise RuntimeError("Selecione um motorista.")

        mrow = self.motoristas_df.loc[self.motoristas_df["Nome Curto"].astype(str) == motorista_nome]
        if mrow.empty:
            raise RuntimeError("Motorista não encontrado no CSV.")
        mrow = mrow.iloc[0]

        telefone = str(mrow["TELEFONE"])
        motorista_id = str(mrow["Cód. Motorista"]) if "Cód. Motorista" in self.motoristas_df.columns else ""
        return motorista_nome, telefone, motorista_id

    def _get_multa_por_pdf(self, info: dict):
        cod_multa = codigo_pdf_para_cod_multa(info["codigo_4d"], info["desdobramento"])
        linha = self.tipos_df.loc[self.tipos_df["COD_MULTA"] == cod_multa]
        if linha.empty:
            raise RuntimeError(f"COD_MULTA {cod_multa} não encontrado no tipos_multa.csv.")
        multa = linha.iloc[0]

        descricao_multa = str(multa["DESCRICAO"]).strip()
        valor_base = float(multa["_valor_float"])
        pontos = int(multa["PONTOS"])
        gravidade = str(multa["GRAVIDADE"]).strip()

        return cod_multa, descricao_multa, valor_base, pontos, gravidade

    def on_generate_message(self):
        try:
            if not self.pdf_path or not self.extracao:
                messagebox.showwarning("Atenção", "Selecione um PDF válido primeiro.")
                return

            motorista_nome, telefone, motorista_id = self._get_motorista()
            info = self.extracao

            cod_multa, descricao_multa, valor_base, pontos, gravidade = self._get_multa_por_pdf(info)

            valor_com = round(valor_base * 0.6, 2)
            valor_sem = round((valor_base * 0.8) * 3, 2)

            msg = MESSAGE_TEMPLATE.format(
                nome_motorista=motorista_nome,
                data_multa=info["data_multa"],
                cidade=info["cidade"],
                uf=info["uf"],
                hora_multa=info["hora_multa"],
                placa=info["placa"],
                descricao_multa=descricao_multa,
                valor_base=format_brl(valor_base),
                pontos=pontos,
                valor_com_indicacao=format_brl(valor_com),
                valor_sem_indicacao=format_brl(valor_sem),
            )

            # salva TXT no Downloads usando data_multa (do PDF)
            data_nome = info["data_multa"].replace("/", "-")
            out_txt = DOWNLOADS_DIR / f"Mensagem {sanitize_filename(motorista_nome)} {data_nome}.txt"
            out_txt.write_text(msg, encoding="utf-8")

            # ✅ NÃO registra no log quando gerar somente mensagem
            messagebox.showinfo("OK", f"Mensagem gerada em:\n{out_txt}")
            open_downloads_folder()

        except Exception as e:
            messagebox.showerror("Erro", str(e))

    def on_generate_pdf_final(self):
        if not self.pdf_path or not self.extracao:
            messagebox.showwarning("Atenção", "Selecione um PDF válido primeiro.")
            return

        try:
            motorista_nome, telefone, motorista_id = self._get_motorista()
            indicar = self.indicar_var.get().strip()
            info = self.extracao

            cod_multa, descricao_multa, valor_base, pontos, gravidade = self._get_multa_por_pdf(info)

            valor_com = round(valor_base * 0.6, 2)
            valor_sem = round((valor_base * 0.8) * 3, 2)

            reg_id = uuid.uuid4().hex[:10]
            now = datetime.now()

            tmp_dir = tempfile.mkdtemp(prefix="multas_")
            termo_docx = os.path.join(tmp_dir, f"termo_{reg_id}.docx")
            termo_pdf = os.path.join(tmp_dir, f"termo_{reg_id}.pdf")

            try:
                marca_com = "X" if indicar == "SIM" else ""
                marca_sem = "X" if indicar == "NÃO" else ""

                context = {
                    "id_registro": reg_id,
                    "data_hoje": data_por_extenso_ptbr(now),
                    "data_registro": now.strftime("%d/%m/%Y %H:%M"),

                    "motorista_id": motorista_id,
                    "nome_motorista": motorista_nome,
                    "telefone": telefone,

                    "placa": info["placa"],
                    "cidade": info["cidade"],
                    "uf": info["uf"],
                    "data_multa": info["data_multa"],
                    "hora_multa": info["hora_multa"],

                    "codigo_multa": cod_multa,
                    "descricao_multa": descricao_multa,

                    "valor_base": format_brl(valor_base),
                    "pontos": pontos,
                    "valor_com_indicacao": format_brl(valor_com),
                    "valor_sem_indicacao": format_brl(valor_sem),

                    "decisao_indicar": indicar,
                    "gravidade_multa": gravidade,

                    "marca_com_indicacao": marca_com,
                    "marca_sem_indicacao": marca_sem,
                }

                gerar_termo_docx(TERMO_TEMPLATE_DOCX, context, termo_docx)
                docx_to_pdf(termo_docx, termo_pdf)

                # final PDF no Downloads com data_multa (do PDF)
                data_nome = info["data_multa"].replace("/", "-")
                final_pdf_name = f"Autorização Desconto {sanitize_filename(motorista_nome)} {data_nome}.pdf"
                out_pdf = DOWNLOADS_DIR / final_pdf_name

                # merge termo + notificação direto no Downloads
                merge_pdfs([termo_pdf, self.pdf_path], str(out_pdf))

                # log no CSV (PDF gerado)
                log_row = {
                    "id_registro": reg_id,
                    "data_registro": now.strftime("%Y-%m-%d %H:%M:%S"),
                    "motorista_id": motorista_id,
                    "nome_motorista": motorista_nome,
                    "telefone": telefone,
                    "placa": info["placa"],
                    "uf": info["uf"],
                    "cidade": info["cidade"],
                    "data_multa": datetime.strptime(info["data_multa"], "%d/%m/%Y").strftime("%Y-%m-%d"),
                    "hora_multa": info["hora_multa"],
                    "codigo_multa": cod_multa,
                    "descricao_multa": descricao_multa,
                    "gravidade_multa": gravidade,
                    "valor_base": valor_base,
                    "pontos": pontos,
                    "valor_com_indicacao": valor_com,
                    "valor_sem_indicacao": valor_sem,
                    "decisao_indicar": indicar,
                }
                append_log_csv(LOG_CSV_PATH, log_row)

                messagebox.showinfo("OK", f"PDF Final gerado em:\n{out_pdf}")
                open_downloads_folder()

            finally:
                shutil.rmtree(tmp_dir, ignore_errors=True)

        except Exception as e:
            messagebox.showerror("Erro", str(e))


if __name__ == "__main__":
    app = MultasApp()
    app.mainloop()