import os
import sys
import shutil
import tempfile
import subprocess
from datetime import datetime
from pathlib import Path

from docxtpl import DocxTemplate

from utils.helpers import sanitize_filename, data_por_extenso_ptbr, format_brl
from services.multa_service import MultaService

def gerar_termo_docx(template_docx: str, context: dict, out_docx_path: str):
    doc = DocxTemplate(template_docx)
    doc.render(context)
    doc.save(out_docx_path)
    return out_docx_path

def docx_to_pdf(docx_path: str, pdf_path: str):
    os.makedirs(os.path.dirname(pdf_path), exist_ok=True)

    if not os.path.exists(docx_path):
        raise RuntimeError(f"DOCX não existe: {docx_path}")

    if sys.platform.startswith("win"):
        try:
            import pythoncom
            import win32com.client

            pythoncom.CoInitialize()
            try:
                word = win32com.client.DispatchEx("Word.Application")
                word.Visible = False
                word.DisplayAlerts = 0

                doc = word.Documents.Open(docx_path, ReadOnly=1)
                try:
                    # 17 = wdExportFormatPDF
                    doc.ExportAsFixedFormat(pdf_path, 17)
                finally:
                    doc.Close(False)
                    word.Quit()
            finally:
                pythoncom.CoUninitialize()

            if not os.path.exists(pdf_path):
                raise RuntimeError("Word executou, mas o PDF não foi gerado.")

            return pdf_path

        except Exception as e:
            raise RuntimeError(
                "Falha ao converter DOCX→PDF via Word (COM).\n\n"
                f"DOCX:\n{docx_path}\n\nPDF:\n{pdf_path}\n\n"
                f"Erro original: {e}"
            )

    raise RuntimeError("Conversão DOCX→PDF só configurada para Windows/Word.")

def merge_pdfs(pdf_paths: list[str], out_path: str):
    from pypdf import PdfWriter, PdfReader
    writer = PdfWriter()
    for p in pdf_paths:
        reader = PdfReader(p)
        for page in reader.pages:
            writer.add_page(page)
    with open(out_path, "wb") as f:
        writer.write(f)
    return out_path

def gerar_pdf_final(
    motoristas_csv: str,
    template_docx: str,
    pdf_notificacao: str,
    extracao: dict,
    multa_atual: dict,
    motorista_nome: str,
    indicar: str,
    output_dir: str
) -> dict:
    """
    - Gera termo preenchido (docx → pdf)
    - Mescla termo_pdf + pdf_notificacao => pdf_final
    - Retorna {pdf_final_path, log_row}
    """
    if not os.path.exists(template_docx):
        raise FileNotFoundError(template_docx)

    # precisa do motorista_id/telefone
    # (reutiliza MultaService só para buscar motorista)
    # - tipos_multa não precisa aqui; mas mantemos o padrão
    tipos_dummy = os.path.join(os.path.dirname(motoristas_csv), "tipos_multa.csv")
    service = MultaService(motoristas_csv, tipos_dummy)

    motor = service.buscar_motorista(motorista_nome)

    valor_base_num = float(multa_atual["valor_base_num"])
    v_com, v_sem = service.calcular_valores(valor_base_num)

    now = datetime.now()
    reg_id = now.strftime("%Y%m%d%H%M%S")
    data_nome = extracao["data_multa"].replace("/", "-")

    # X no template
    marca_com = "X" if indicar == "SIM" else ""
    marca_sem = "X" if indicar == "NÃO" else ""

    workdir = Path(tempfile.mkdtemp(prefix="multas_qt_"))
    termo_docx_path = str(workdir / f"termo_{reg_id}.docx")
    termo_pdf_path = str(workdir / f"termo_{reg_id}.pdf")

    try:
        context = {
            "id_registro": reg_id,
            "data_hoje": data_por_extenso_ptbr(now),  # EX: "13 de Janeiro de 2026"
            "data_registro": now.strftime("%d/%m/%Y %H:%M"),

            "motorista_id": motor["motorista_id"],
            "nome_motorista": motor["nome_motorista"],
            "telefone": motor["telefone"],

            "placa": extracao["placa"],
            "cidade": extracao.get("cidade", ""),
            "uf": extracao.get("uf", ""),
            "data_multa": extracao["data_multa"],
            "hora_multa": extracao["hora_multa"],

            "codigo_multa": multa_atual["codigo_multa"],
            "descricao_multa": multa_atual["descricao_multa"],
            "gravidade_multa": multa_atual["gravidade_multa"],

            "valor_base": format_brl(valor_base_num),
            "pontos": int(multa_atual["pontos"]),
            "valor_com_indicacao": format_brl(v_com),
            "valor_sem_indicacao": format_brl(v_sem),

            "decisao_indicar": indicar,
            "marca_com_indicacao": marca_com,
            "marca_sem_indicacao": marca_sem,
        }

        gerar_termo_docx(template_docx, context, termo_docx_path)
        docx_to_pdf(termo_docx_path, termo_pdf_path)

        # PDF final vai para output (um único arquivo)
        os.makedirs(output_dir, exist_ok=True)
        final_name = f"Autorização Desconto {sanitize_filename(motorista_nome)} {data_nome}.pdf"
        final_path = os.path.join(output_dir, final_name)

        merge_pdfs([termo_pdf_path, pdf_notificacao], final_path)

        # log row (somente campos que você definiu)
        # data_multa: converter dd/mm/yyyy -> yyyy-mm-dd
        dt_iso = datetime.strptime(extracao["data_multa"], "%d/%m/%Y").strftime("%Y-%m-%d")

        log_row = {
            "id_registro": reg_id,
            "data_registro": now.strftime("%Y-%m-%d %H:%M:%S"),
            "motorista_id": motor["motorista_id"],
            "nome_motorista": motor["nome_motorista"],
            "telefone": motor["telefone"],
            "placa": extracao["placa"],
            "uf": extracao.get("uf",""),
            "cidade": extracao.get("cidade",""),
            "data_multa": dt_iso,
            "hora_multa": extracao["hora_multa"],
            "codigo_multa": multa_atual["codigo_multa"],
            "descricao_multa": multa_atual["descricao_multa"],
            "valor_base": valor_base_num,
            "pontos": int(multa_atual["pontos"]),
            "valor_com_indicacao": v_com,
            "valor_sem_indicacao": v_sem,
            "decisao_indicar": indicar,
            "gravidade_multa": multa_atual["gravidade_multa"],
        }

        return {"pdf_final_path": final_path, "log_row": log_row}

    finally:
        shutil.rmtree(workdir, ignore_errors=True)