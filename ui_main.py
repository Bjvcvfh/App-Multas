import os
from pathlib import Path

from PySide6.QtCore import Qt
from PySide6.QtGui import QIcon, QPixmap
from PySide6.QtWidgets import (
    QMainWindow, QWidget, QLabel, QPushButton, QTextEdit,
    QVBoxLayout, QHBoxLayout, QFileDialog, QMessageBox,
    QComboBox, QGroupBox, QRadioButton, QButtonGroup
)

from utils.helpers import resource_path, ensure_dirs, sanitize_filename, get_persistent_app_dir
from services.pdf_service import extrair_campos_notificacao, codigo_pdf_para_cod_multa
from services.multa_service import MultaService
from services.doc_service import gerar_pdf_final
from services.log_service import LogService


class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()

        # ====== Paths base do app (compatível com .py e .exe)
        self.BASE_DIR = Path(resource_path("."))
        self.DATA_DIR = self.BASE_DIR / "data"
        self.TEMPLATES_DIR = self.BASE_DIR / "templates"
        self.OUTPUT_DIR = self.BASE_DIR / "output"
        self.ASSETS_DIR = self.BASE_DIR / "assets"
        self.APP_DIR = get_persistent_app_dir(
            app_name="AppMultas",
        )

        self.MOTORISTAS_CSV = str(self.DATA_DIR / "motoristas.csv")
        self.TIPOS_MULTA_CSV = str(self.DATA_DIR / "tipos_multa.csv")
        self.TERMO_TEMPLATE_DOCX = str(self.TEMPLATES_DIR / "termo_multa_modelo.docx")
        self.LOG_CSV_PATH = str(self.APP_DIR / "logs_multas.csv")

        ensure_dirs([self.DATA_DIR, self.TEMPLATES_DIR, self.OUTPUT_DIR, self.ASSETS_DIR])

        # ====== Estado
        self.pdf_path: str | None = None
        self.extracao: dict | None = None
        self.multa_atual: dict | None = None

        # ====== Services
        try:
            self.multa_service = MultaService(self.MOTORISTAS_CSV, self.TIPOS_MULTA_CSV)
            self.log_service = LogService(self.LOG_CSV_PATH)
        except Exception as e:
            QMessageBox.critical(self, "Erro ao iniciar", str(e))
            raise

        # ====== Janela
        self.setWindowTitle("App Multas — PySide6 (PDF + Termo)")
        self.setMinimumSize(980, 600)

        ico_path = resource_path("assets/app.ico")
        if os.path.exists(ico_path):
            self.setWindowIcon(QIcon(ico_path))

        # ====== UI
        self._build_ui()
        self._load_motoristas()
        
        self.setStyleSheet("""
        QMainWindow {
            background-color: #F5F7FA;
        }

        QGroupBox {
            font-weight: bold;
            border: 1px solid #DADADA;
            border-radius: 6px;
            margin-top: 10px;
        }

        QGroupBox::title {
            subcontrol-origin: margin;
            left: 10px;
            padding: 0 4px 0 4px;
        }

        QPushButton {
            background-color: #2563EB;
            color: white;
            border-radius: 6px;
            padding: 8px;
            font-size: 13px;
        }

        QPushButton:hover {
            background-color: #1D4ED8;
        }

        QPushButton:disabled {
            background-color: #9CA3AF;
        }

        QComboBox, QTextEdit {
            border: 1px solid #D1D5DB;
            border-radius: 5px;
            padding: 6px;
            background: white;
        }
        """)
        

    def _build_ui(self):
        central = QWidget()
        self.setCentralWidget(central)
        root = QVBoxLayout(central)
        root.setContentsMargins(14, 14, 14, 14)
        root.setSpacing(10)

        # ===== Seção PDF
        gb_pdf = QGroupBox("1) Anexar Notificação (PDF)")
        lay_pdf = QHBoxLayout(gb_pdf)

        self.lbl_pdf = QLabel("Nenhum PDF selecionado")
        self.lbl_pdf.setWordWrap(True)

        btn_select_pdf = QPushButton("Selecionar PDF")
        btn_select_pdf.clicked.connect(self.on_select_pdf)

        lay_pdf.addWidget(self.lbl_pdf, 1)
        lay_pdf.addWidget(btn_select_pdf)
        root.addWidget(gb_pdf)

        # ===== Seção motorista
        gb_motor = QGroupBox("2) Motorista")
        lay_motor = QVBoxLayout(gb_motor)

        self.cb_motorista = QComboBox()
        lay_motor.addWidget(self.cb_motorista)

        root.addWidget(gb_motor)

        # ===== Seção indicar pontos
        gb_ind = QGroupBox("3) Indicar pontos?")
        lay_ind = QHBoxLayout(gb_ind)

        self.rb_sim = QRadioButton("SIM")
        self.rb_nao = QRadioButton("NÃO")
        self.rb_sim.setChecked(True)

        self.ind_group = QButtonGroup()
        self.ind_group.addButton(self.rb_sim)
        self.ind_group.addButton(self.rb_nao)

        lay_ind.addWidget(self.rb_sim)
        lay_ind.addWidget(self.rb_nao)
        lay_ind.addStretch(1)
        root.addWidget(gb_ind)

        # ===== Preview
        gb_prev = QGroupBox("Prévia dos dados extraídos + multa encontrada")
        lay_prev = QVBoxLayout(gb_prev)
        self.txt_preview = QTextEdit()
        self.txt_preview.setReadOnly(True)
        self.txt_preview.setPlaceholderText("Selecione um PDF para extrair automaticamente os campos.")
        lay_prev.addWidget(self.txt_preview)
        root.addWidget(gb_prev, 1)

        # ===== Botões
        row_btn = QHBoxLayout()

        row_btn.addStretch(1)
        
        self.btn_msg = QPushButton("Gerar Mensagem")
        self.btn_pdf = QPushButton("Gerar PDF Final (e registrar log)")

        # padroniza tamanhos
        self.btn_msg.setMinimumHeight(40)
        self.btn_pdf.setMinimumHeight(40)
        self.btn_msg.setMinimumWidth(240)
        self.btn_pdf.setMinimumWidth(240)

        self.btn_msg.clicked.connect(self.on_gerar_mensagem)
        self.btn_pdf.clicked.connect(self.on_gerar_pdf_final)

        row_btn.addWidget(self.btn_msg)
        row_btn.addSpacing(20)  # espaço entre os botões
        row_btn.addWidget(self.btn_pdf)

        row_btn.addStretch(1)

        root.addLayout(row_btn)

        # Status
        self.lbl_status = QLabel("")
        self.lbl_status.setStyleSheet("color: #555;")
        root.addWidget(self.lbl_status)

    def _load_motoristas(self):
        self.cb_motorista.clear()
        for nome in self.multa_service.listar_motoristas():
            self.cb_motorista.addItem(nome)

    def _indicar_valor(self) -> str:
        return "SIM" if self.rb_sim.isChecked() else "NÃO"

    # =========================
    # Eventos
    # =========================
    def on_select_pdf(self):
        path, _ = QFileDialog.getOpenFileName(self, "Selecione a notificação (PDF)", "", "PDF (*.pdf)")
        if not path:
            return

        self.pdf_path = path
        self.lbl_pdf.setText(path)

        try:
            self.extracao = extrair_campos_notificacao(path)

            cod_multa = codigo_pdf_para_cod_multa(self.extracao["codigo_4d"], self.extracao["desdobramento"])
            self.multa_atual = self.multa_service.buscar_multa_por_cod(cod_multa)

            self._render_preview()
            self.lbl_status.setText("PDF lido e multa encontrada no tipos_multa.csv.")
        except Exception as e:
            self.extracao = None
            self.multa_atual = None
            QMessageBox.critical(self, "Erro ao ler PDF", str(e))
            self.txt_preview.setText("Falha ao extrair dados do PDF.")

    def _render_preview(self):
        if not self.extracao:
            self.txt_preview.setText("")
            return

        info = self.extracao
        cod_multa = codigo_pdf_para_cod_multa(info["codigo_4d"], info["desdobramento"])

        s = []
        s.append(f"Placa: {info['placa']}")
        s.append(f"Data: {info['data_multa']}")
        s.append(f"Hora: {info['hora_multa']}")
        s.append(f"Cidade/UF: {info.get('cidade','')}/{info.get('uf','')}")
        s.append(f"Código infração (PDF): {info['codigo_4d']}")
        s.append(f"Desdobramento: {info['desdobramento']}")
        s.append(f"COD_MULTA (interno): {cod_multa}")
        s.append(f"Valor (PDF): {info['valor_pdf']}")

        if self.multa_atual:
            s.append("")
            s.append("=== Dados puxados do tipos_multa.csv ===")
            s.append(f"Descrição: {self.multa_atual['descricao_multa']}")
            s.append(f"Valor base: {self.multa_atual['valor_base']}")
            s.append(f"Pontos: {self.multa_atual['pontos']}")
            s.append(f"Gravidade: {self.multa_atual['gravidade_multa']}")
        else:
            s.append("")
            s.append("⚠️ Não encontrei essa multa no tipos_multa.csv (COD_MULTA).")

        self.txt_preview.setText("\n".join(s))

    def on_gerar_mensagem(self):
        """
        Gera a mensagem e salva em TXT na pasta Downloads (NÃO registra log).
        """
        if not self.extracao or not self.multa_atual:
            QMessageBox.warning(
                self, "Atenção",
                "Selecione um PDF válido primeiro (e garanta que o COD_MULTA exista no tipos_multa.csv)."
            )
            return

        motorista_nome = self.cb_motorista.currentText().strip()

        try:
            msg = self.multa_service.gerar_mensagem(motorista_nome, self.extracao, self.multa_atual)

            downloads_dir = Path.home() / "Downloads"
            downloads_dir.mkdir(parents=True, exist_ok=True)

            data_nome = self.extracao["data_multa"].replace("/", "-")  # dd-mm-aaaa
            txt_name = f"Mensagem Multa {sanitize_filename(motorista_nome)} {data_nome}.txt"
            txt_path = downloads_dir / txt_name

            with open(txt_path, "w", encoding="utf-8") as f:
                f.write(msg)

            QMessageBox.information(self, "OK", f"Mensagem salva em:\n{txt_path}")
            self.lbl_status.setText("Mensagem salva em TXT (sem log).")

        except Exception as e:
            QMessageBox.critical(self, "Erro", str(e))

    def on_gerar_pdf_final(self):
        """
        Gera o PDF final (Termo + Notificação) em Downloads e registra no log.
        """
        if not self.pdf_path or not self.extracao or not self.multa_atual:
            QMessageBox.warning(
                self, "Atenção",
                "Selecione um PDF válido primeiro (e garanta COD_MULTA no tipos_multa.csv)."
            )
            return

        if not os.path.exists(self.TERMO_TEMPLATE_DOCX):
            QMessageBox.critical(self, "Template ausente", f"Template não encontrado:\n{self.TERMO_TEMPLATE_DOCX}")
            return

        motorista_nome = self.cb_motorista.currentText().strip()
        indicar = self._indicar_valor()

        try:
            downloads_dir = Path.home() / "Downloads"
            downloads_dir.mkdir(parents=True, exist_ok=True)

            result = gerar_pdf_final(
                motoristas_csv=self.MOTORISTAS_CSV,
                template_docx=self.TERMO_TEMPLATE_DOCX,
                pdf_notificacao=self.pdf_path,
                extracao=self.extracao,
                multa_atual=self.multa_atual,
                motorista_nome=motorista_nome,
                indicar=indicar,
                output_dir=str(downloads_dir),  # ✅ Downloads
            )

            # registra log SOMENTE aqui
            self.log_service.registrar(result["log_row"])

            QMessageBox.information(self, "OK", f"PDF final gerado em:\n{result['pdf_final_path']}")
            self.lbl_status.setText("PDF final gerado em Downloads e log registrado.")

            # abre Downloads no Windows
            try:
                if os.name == "nt":
                    os.startfile(str(downloads_dir))
            except Exception:
                pass

        except Exception as e:
            QMessageBox.critical(self, "Erro ao gerar PDF", str(e))