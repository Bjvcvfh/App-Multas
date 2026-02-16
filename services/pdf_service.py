import re
import pdfplumber

def extrair_texto_pdf(pdf_path: str) -> str:
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
        up = ln.upper()
        if "NOME DO MUNICIPIO UF" in up or "NOME DO MUNICÍPIO UF" in up:
            header_idx = i
            break

    if header_idx is None:
        return None, None

    for j in range(header_idx + 1, min(header_idx + 8, len(lines))):
        cand = lines[j]

        # pega só o que vem depois de ')', remove 1+ espaços à esquerda
        if ")" in cand:
            cand = cand.split(")", 1)[1]
            cand = cand.lstrip()

        # esperamos: "JACUPIRANGA SP"
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

    # PLACA
    placa = None
    m = re.search(r"\bPLACA\b.*?\n\s*([A-Z0-9]{7}|[A-Z]{3}\s*\-?\s*\d{4})\b", text, re.IGNORECASE)
    if m:
        placa = re.sub(r"\s|\-", "", m.group(1)).upper()

    # DATA + HORA
    data_multa = None
    hora_multa = None
    m = re.search(r"DATA\s+HORA.*?\b(\d{2}/\d{2}/\d{4})\s+(\d{2}:\d{2})\b", text, re.IGNORECASE | re.DOTALL)
    if m:
        data_multa, hora_multa = m.group(1), m.group(2)

    # CÓDIGO INFRAÇÃO + DESDOBRAMENTO + VALOR
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

    # CIDADE/UF
    cidade, uf = extrair_cidade_uf_por_linhas(text)

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
            "📄 TEXTO EXTRAÍDO DO PDF (diagnóstico, primeiros 6000 chars):\n\n"
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
    }

def codigo_pdf_para_cod_multa(codigo_4d: str, desdobramento: str) -> str:
    # 7455 + 0 -> 745-50
    base = codigo_4d[:3]
    sufixo = codigo_4d[3:]
    return f"{base}-{sufixo}{desdobramento}"