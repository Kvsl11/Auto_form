import os
import sys
import time
import ssl
import certifi
import threading
import requests
import subprocess
import urllib3
from pathlib import Path

import pandas as pd

# Selenium / Undetected ChromeDriver
import undetected_chromedriver as uc
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

from selenium.common.exceptions import (
    TimeoutException,
    NoSuchElementException,
    ElementClickInterceptedException,
    WebDriverException,
)

from selenium.webdriver.chrome.options import Options

# PySide6 / Qt
from PySide6.QtCore import QThread, Signal, Qt, QSize
from PySide6.QtGui import QFont, QAction, QTextOption, QPalette, QColor, QIcon
from PySide6.QtWidgets import (
    QApplication,
    QMainWindow,
    QWidget,
    QLabel,
    QLineEdit,
    QPushButton,
    QCheckBox,
    QFileDialog,
    QProgressBar,
    QPlainTextEdit,
    QGridLayout,
    QHBoxLayout,
    QVBoxLayout,
    QMessageBox,
    QSpacerItem,
    QSizePolicy,
    QStatusBar,
)


# ==========================
# CONFIGURAÇÕES INICIAIS
# ==========================

# Corrige SSL e suprime avisos
def apply_ssl_fix():
    """Garante que o ambiente Python reconheça certificados SSL e suprime avisos de requisição insegura."""
    try:
        os.environ["SSL_CERT_FILE"] = certifi.where()
        ssl._create_default_https_context = ssl._create_unverified_context
        # Suprime o aviso de InsecureRequestWarning
        urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)
    except Exception:
        pass

VERSAO = "1.0.0"

# URL padrão
DEFAULT_FORM_URL = "https://docs.google.com/forms/d/e/1FAIpQLSfLyptmo3NFUx8dxC7k0obmQxAXPuimBLC_L30xgZOsygvqpg/viewform"

# Caminho padrão do Excel
DEFAULT_EXCEL_PATH = str(Path.home() / "Downloads" / "Auto_teste.xlsx")


# Mapeamento de campos
FIELD_MAPPING_DEFAULT = {
    '#mG61Hd > div.RH5hzf.RLS9Fe > div > div.o3Dpx > div:nth-child(1) > div > div > div.vQES8d > div > div:nth-child(1) > div.ry3kXd > div.MocG8c.HZ3kWc.mhLiyf.LMgvRb.DEh1R.KKjvXb': "TECNOLOGIA",
    '#mG61Hd > div.RH5hzf.RLS9Fe > div > div.o3Dpx > div:nth-child(2) > div > div > div.vQES8d > div > div:nth-child(1) > div.ry3kXd > div.MocG8c.HZ3kWc.mhLiyf.LMgvRb.KKjvXb.DEh1R': "UNIDADE",
    '#mG61Hd > div.RH5hzf.RLS9Fe > div > div.o3Dpx > div:nth-child(3) > div > div > div.vQES8d > div > div:nth-child(1) > div.ry3kXd > div.MocG8c.HZ3kWc.mhLiyf.LMgvRb.KKjvXb.DEh1R': "SETOR",
    '#mG61Hd > div.RH5hzf.RLS9Fe > div > div.o3Dpx > div:nth-child(4) > div > div > div.vQES8d > div > div:nth-child(1) > div.ry3kXd > div.MocG8c.HZ3kWc.mhLiyf.LMgvRb.KKjvXb.DEh1R': "FRENTE",
    '#mG61Hd > div.RH5hzf.RLS9Fe > div > div.o3Dpx > div:nth-child(5) > div > div > div.vQES8d > div > div:nth-child(1) > div.ry3kXd > div.MocG8c.HZ3kWc.mhLiyf.LMgvRb.KKjvXb.DEh1R': "MODELO",
    '#mG61Hd > div.RH5hzf.RLS9Fe > div > div.o3Dpx > div:nth-child(6) > div > div > div.AgroKb > div > div.aCsJod.oJeWuf > div > div.Xb9hP > input': "FROTA",
    '#mG61Hd > div.RH5hzf.RLS9Fe > div > div.o3Dpx > div:nth-child(7) > div > div > div.AgroKb > div > div.RpC4Ne.oJeWuf > div.Pc9Gce.Wic03c > textarea': "QRM",
    '#mG61Hd > div.RH5hzf.RLS9Fe > div > div.o3Dpx > div:nth-child(8) > div > div > div.AgroKb > div > div.RpC4Ne.oJeWuf > div.Pc9Gce.Wic03c > textarea': "LOCAL (QTH)",
    '#mG61Hd > div.RH5hzf.RLS9Fe > div > div.o3Dpx > div:nth-child(9) > div > div > div.AgroKb > div > div.RpC4Ne.oJeWuf > div.Pc9Gce.Wic03c > textarea': "RESPONSÁVEL PELA O.S",
}


# ==========================
# WORKER (THREAD) DA AUTOMAÇÃO
# ==========================

class FormsWorker(QThread):
    log = Signal(str)                 # mensagens de log
    progress = Signal(int, int)       # atual, total
    status = Signal(str)              # texto de status
    finished = Signal(int, int, str)  # sucessos, falhas, motivo

    def __init__(self, form_url: str, excel_path: str, field_mapping: dict, headless: bool, keep_open: bool):
        super().__init__()
        self.form_url = form_url.strip()
        self.excel_path = excel_path.strip()
        self.field_mapping = field_mapping or {}
        self.headless = headless
        self.keep_open = keep_open
        self._stop_event = threading.Event()
        self.driver = None

    def request_stop(self):
        # Sinaliza a parada e tenta forçar o fechamento do driver imediatamente
        self._stop_event.set()
        self.log.emit("🛑 Parada solicitada. Tentando fechar o navegador...")
        try:
            if self.driver:
                # O quit() é crucial para interromper chamadas bloqueantes do Selenium (wait.until, driver.get)
                self.driver.quit()
        except Exception:
            pass
        finally:
            self.driver = None

    def stopped(self) -> bool:
        return self._stop_event.is_set()

    def _try_fill_field(self, wait: WebDriverWait, entry_selector: str, column_name: str, valor: str) -> bool:
        if self.stopped():
            return False

        valor = (valor or "").strip()
        if not valor:
            self.log.emit(f"  -> ℹ️ Aviso: Valor vazio para '{column_name}'. Pulando.")
            return True

        try:
            # 1) Tenta localizar pelo CSS
            field = wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, entry_selector)))

            # 2) Campo texto/input
            tag = field.tag_name.lower()
            if tag in ["input", "textarea"]:
                field.clear()
                field.send_keys(valor)
                self.log.emit(f"  -> ✅ Preenchido '{column_name}' com sucesso.")
                return True

            # 3) Dropdown/Opção
            self.log.emit(f"  -> 🖱️ Tentando selecionar opção '{valor}' para '{column_name}'...")
            try:
                field.click()
                option_in_menu_xpath = f'//div[@role="option"]//span[normalize-space(text())="{valor}"]'
                option_element = wait.until(EC.element_to_be_clickable((By.XPATH, option_in_menu_xpath)))
                option_element.click()
                time.sleep(0.3)
                self.log.emit(f"  -> ✅ Selecionado via Dropdown.")
                return True
            except Exception as e:
                # Loga o aviso intermediário (falha no dropdown, tenta opção visível)
                if self.stopped() and isinstance(e, WebDriverException):
                    raise # Re-raise para ser pego pelo bloco exterior e tratado como parada
                self.log.emit(f"    ⚠️ Aviso: Falha como dropdown. Tentando como opção visível. Detalhe: {e.__class__.__name__}")


            # 3b) Rádio/Checkbox/Opção visível
            option_xpath_visible = (
                f'//div[contains(@role, "radio")]/div/div/div[3]/div | '
                f'//label/div/div[2]/div[1]/span[normalize-space(text())="{valor}"] | '
                f'//div[@role="listitem"]//span[normalize-space(text())="{valor}"]'
            )
            visible_option = wait.until(EC.element_to_be_clickable((By.XPATH, option_xpath_visible)))
            visible_option.click()
            time.sleep(0.3)
            self.log.emit(f"  -> ✅ Selecionado como OPÇÃO visível.")
            return True

        except WebDriverException:
            # Captura a exceção de "conexão recusada" quando o driver é fechado.
            if self.stopped():
                # Se a flag de parada estiver ativa, trata como parada suave
                return False
            # Caso contrário, é um erro de driver inesperado e crítico
            self.log.emit(f"  -> ❌ ERRO DE COMUNICAÇÃO (Preenchimento): Não foi possível preencher '{column_name}' com '{valor}'. Driver falhou.")
            return False
        except Exception as e:
            # Captura outras exceções (como TimeoutException, etc.)
            self.log.emit(f"  -> ❌ ERRO FINAL (Preenchimento): Não foi possível preencher '{column_name}' com '{valor}'. Detalhe: {e.__class__.__name__}")
            return False


    def run(self):
        apply_ssl_fix()

        if not self.form_url.startswith("http"):
            self.finished.emit(0, 0, "URL inválida")
            return

        # Leitura da planilha
        try:
            self.status.emit("📚 Lendo arquivo Excel...")
            df = pd.read_excel(self.excel_path)
            df.columns = df.columns.str.replace('\n', ' ', regex=False).str.strip()
            cols_mapped = set(self.field_mapping.values())
            missing = [c for c in cols_mapped if c not in df.columns]
            if missing:
                self.log.emit(f"🚨 ERRO CRÍTICO: Colunas mapeadas não encontradas: {', '.join(missing)}")
                self.finished.emit(0, 0, "colunas ausentes")
                return
            total = len(df)
            self.log.emit(f"📊 Planilha lida com sucesso. Total de {total} registros.")
        except FileNotFoundError:
            self.finished.emit(0, 0, "arquivo não encontrado")
            return
        except Exception as e:
            self.log.emit(f"🚨 ERRO ao ler a planilha: {e}")
            self.finished.emit(0, 0, "erro leitura planilha")
            return

        # Configura o Chrome
        self.status.emit("🌐 Inicializando navegador...")
        try:
            chrome_options = Options()
            chrome_options.add_argument('--ignore-certificate-errors')
            if self.headless:
                chrome_options.add_argument('--headless=new')
                chrome_options.add_argument('--window-size=1920,1080')
                chrome_options.add_argument('--no-sandbox')
                chrome_options.add_argument('--disable-dev-shm-usage')

            self.driver = uc.Chrome(options=chrome_options)
            try:
                self.driver.maximize_window()
            except Exception:
                pass

            wait = WebDriverWait(self.driver, 15)
        except Exception as e:
            self.log.emit(f"🚨 ERRO ao iniciar o Chrome: {e}")
            self.finished.emit(0, 0, "erro chrome")
            return

        successes = 0
        failures = 0
        current_page_is_form = False

        try:
            for index, row in df.iterrows():
                if self.stopped():
                    self.finished.emit(successes, failures, "parado pelo usuário")
                    return

                self.log.emit(f"\n📝 Processando registro {index + 1}/{len(df)}...")
                self.progress.emit(index, len(df))
                self.status.emit(f"▶️ Processando registro {index + 1} de {len(df)}")

                # Garante estar no formulário limpo
                try:
                    if not current_page_is_form:
                        self.log.emit("  -> Recarregando formulário...")
                        self.driver.get(self.form_url)
                        wait.until(EC.presence_of_element_located((By.TAG_NAME, 'form')))
                        current_page_is_form = True
                except WebDriverException as e:
                    if self.stopped():
                        self.finished.emit(successes, failures, "parado pelo usuário")
                        return
                    self.log.emit(f"⚠️ Falha ao carregar o formulário. Tentativa de recuperação: {e}")
                    failures += 1
                    current_page_is_form = False
                    continue

                # Preenchimento de campos
                current_submission_failed = False
                for entry_selector, column_name in self.field_mapping.items():
                    if self.stopped():
                        self.finished.emit(successes, failures, "parado pelo usuário")
                        return
                    
                    valor = str(row[column_name]) if pd.notna(row[column_name]) else ""
                    
                    if not self._try_fill_field(wait, entry_selector, column_name, valor):
                        # Se _try_fill_field retorna False, é falha ou parada (tratada dentro da função)
                        if self.stopped():
                            self.finished.emit(successes, failures, "parado pelo usuário")
                            return
                        current_submission_failed = True
                        break

                if current_submission_failed:
                    failures += 1
                    current_page_is_form = False
                    self.log.emit(f"Registro {index + 1}: ❌ FALHA no preenchimento. Pulando.")
                    continue

                # Submissão
                try:
                    self.log.emit("  -> 📤 Tentando submeter...")
                    submit_button_xpath = '//div[@role="button"]//*[normalize-space(text())="Enviar"]'
                    submit_label = wait.until(EC.element_to_be_clickable((By.XPATH, submit_button_xpath)))
                    submit_label.find_element(By.XPATH, '..').click()

                    # Espera pela mensagem de sucesso
                    success_message_xpath = (
                        '//div[contains(text(), "Sua resposta foi registrada")] | '
                        '//div[contains(text(), "Sua resposta foi enviada")]'
                    )
                    wait.until(EC.presence_of_element_located((By.XPATH, success_message_xpath)))

                    self.log.emit(f"Registro {index + 1}: ✅ SUCESSO! Submetido.")
                    successes += 1
                    time.sleep(0.6)

                    # Prepara a próxima resposta
                    if index < len(df) - 1:
                        self.log.emit("  -> 🔄 Preparando próxima resposta...")
                        next_response_xpath = (
                            '//a[contains(text(), "Enviar outra resposta")] | '
                            '//div[@role="button"]//*[normalize-space(text())="Enviar outra resposta"]'
                        )
                        next_btn = WebDriverWait(self.driver, 8).until(
                            EC.element_to_be_clickable((By.XPATH, next_response_xpath))
                        )
                        next_btn.click()
                        current_page_is_form = True
                    else:
                        self.log.emit("  -> Fim da lista de registros.")

                except WebDriverException:
                    if self.stopped():
                        # Parada suave durante a submissão/espera de sucesso
                        self.finished.emit(successes, failures, "parado pelo usuário")
                        return
                    raise # É uma falha inesperada se não tiver sido parado

                except Exception as e:
                    if self.stopped():
                        # Caso a thread tenha sido parada e a exceção não tenha sido WebDriverException
                        self.finished.emit(successes, failures, "parado pelo usuário")
                        return
                    self.log.emit(f"Registro {index + 1}: ❌ FALHA na submissão. Erro: {e.__class__.__name__}")
                    failures += 1
                    current_page_is_form = False

            # Fim do loop
            reason = "concluído"
            self.finished.emit(successes, failures, reason)

        except Exception as e:
            if self.stopped():
                # Se parou aqui, é porque o driver foi fechado por request_stop()
                self.finished.emit(successes, failures, "parado pelo usuário (driver fechado)")
            else:
                # Erro não esperado
                self.log.emit(f"🚨 ERRO CRÍTICO no loop principal: {e}")
                self.finished.emit(successes, failures, f"erro inesperado: {e.__class__.__name__}")
        
        finally:
            # Garante que o driver feche. Usamos 'self.driver' como flag de que o quit() não foi chamado.
            try:
                if self.driver and (not self.keep_open or self.stopped()):
                    self.driver.quit()
            except Exception:
                pass
            self.driver = None


# ==========================
# WORKER (THREAD) DA ATUALIZAÇÃO
# ==========================
class UpdateWorker(QThread):
    """Verifica a atualização em uma thread separada para não bloquear a UI."""
    result = Signal(str, str)  # Sinal emitido com (versao_online, erro_msg)

    def run(self):
        try:
            repo_url = "https://raw.githubusercontent.com/Kvsl11/Auto_form/main/version.txt"
            resposta = requests.get(repo_url, timeout=8, verify=False)
            resposta.raise_for_status()
            versao_online = resposta.text.strip()
            self.result.emit(versao_online, "")
        except requests.exceptions.RequestException as e:
            self.result.emit("", f"Falha de rede: {e}")
        except Exception as e:
            self.result.emit("", f"Erro inesperado: {e}")


# ==========================
# JANELA PRINCIPAL (UI)
# ==========================

class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle(f"Auto - Form v{VERSAO}")
        self.setMinimumSize(950, 700)
        self.worker: FormsWorker | None = None
        self.update_worker: UpdateWorker | None = None
        self.update_btn: QPushButton | None = None

        # Widgets
        self.url_edit = QLineEdit(DEFAULT_FORM_URL)
        self.path_edit = QLineEdit(DEFAULT_EXCEL_PATH)
        
        self.browse_btn = QPushButton("📂 Procurar Arquivo")
        self.headless_cb = QCheckBox("Executar Invisível (Headless)")
        self.keep_open_cb = QCheckBox("Manter navegador aberto após o fim")
        
        self.start_btn = QPushButton("▶ Iniciar Automação")
        self.stop_btn = QPushButton("■ Parar")
        
        self.start_btn.setObjectName("start_btn")
        self.stop_btn.setObjectName("stop_btn")

        self.progress_bar = QProgressBar()
        self.log_view = QPlainTextEdit()
        self.log_view.setReadOnly(True)
        self.log_view.setWordWrapMode(QTextOption.NoWrap)
        
        mono = QFont("Consolas" if sys.platform.startswith("win") else "Menlo", 10)
        self.log_view.setFont(mono)

        # Layouts
        title_label = QLabel("Auto - Form")
        title_font = QFont()
        title_font.setPointSize(18)
        title_font.setBold(True)
        title_label.setFont(title_font)
        title_label.setAlignment(Qt.AlignmentFlag.AlignCenter)

        top_grid = QGridLayout()
        top_grid.setSpacing(12)
        
        top_grid.addWidget(QLabel("🔗 URL do Formulário:"), 1, 0)
        top_grid.addWidget(self.url_edit, 1, 1, 1, 3)

        top_grid.addWidget(QLabel("📄 Arquivo Excel (.xlsx):"), 2, 0)
        top_grid.addWidget(self.path_edit, 2, 1, 1, 2)
        top_grid.addWidget(self.browse_btn, 2, 3)

        check_layout = QHBoxLayout()
        check_layout.addWidget(self.headless_cb)
        check_layout.addWidget(self.keep_open_cb)
        check_layout.addStretch(1)
        top_grid.addLayout(check_layout, 3, 1, 1, 3)
        top_grid.setRowStretch(3, 1)

        buttons_row = QHBoxLayout()
        buttons_row.addWidget(self.start_btn)
        buttons_row.addWidget(self.stop_btn)
        buttons_row.addSpacerItem(QSpacerItem(40, 20, QSizePolicy.Expanding, QSizePolicy.Minimum))

        center = QVBoxLayout()
        center.addWidget(title_label)
        center.addLayout(top_grid)
        center.addSpacing(15)
        center.addLayout(buttons_row)
        center.addWidget(QLabel("📈 Progresso:"))
        center.addWidget(self.progress_bar)
        center.addWidget(QLabel("📜 Logs:"))
        center.addWidget(self.log_view)

        container = QWidget()
        container.setLayout(center)
        self.setCentralWidget(container)

        # Status bar
        self.status = QStatusBar()
        self.setStatusBar(self.status)
        self.status.showMessage("Pronto para iniciar.")

        # Conexões
        self.browse_btn.clicked.connect(self.on_browse)
        self.start_btn.clicked.connect(self.on_start)
        self.stop_btn.clicked.connect(self.on_stop)

        # Estado inicial
        self.stop_btn.setEnabled(False)

        self.apply_dark_style()
        self._build_menu()
        self._check_for_updates()
        
    def _check_for_updates(self):
        self.status.showMessage("🔄 Verificando atualizações...")
        self.update_worker = UpdateWorker()
        self.update_worker.result.connect(self._handle_update_check_result)
        self.update_worker.start()

    def _handle_update_check_result(self, versao_online, error_msg):
        if error_msg:
            self.status.showMessage(f"⚠️ Falha ao verificar atualização: {error_msg}")
            return

        if versao_online != VERSAO:
            # Alterado para usar cor no texto em vez de emoji
            status_label = QLabel(f"Nova versão disponível: v{versao_online}")
            status_label.setStyleSheet("color: #ee8715; font-weight: bold;")

            self.update_btn = QPushButton("⬇ Atualizar agora")
            # --- ESTILO DO BOTÃO ALTERADO PARA A NOVA COR ---
            self.update_btn.setStyleSheet("""
                QPushButton {
                    background-color: transparent;
                    color: #ee8715;
                    border: 1px solid #ee8715;
                    border-radius: 5px;
                    padding: 3px 10px;
                    font-weight: bold;
                }
                QPushButton:hover {
                    background-color: #3B3C53;
                    color: #ffffff;
                }
            """)
            self.update_btn.clicked.connect(lambda: self._download_and_apply_update(versao_online))
            
            self.status.addPermanentWidget(status_label)
            self.status.addPermanentWidget(self.update_btn)
        else:
            self.status.showMessage(f"🟢 Atualizado — v{VERSAO}")

    def _download_and_apply_update(self, versao_online):
        script_url = "https://raw.githubusercontent.com/Kvsl11/Auto_form/main/main.py"
        
        if self.update_btn:
            self.update_btn.setText("⬇ Baixando...")
            self.update_btn.setEnabled(False)

        try:
            r = requests.get(script_url, timeout=15, verify=False)
            r.raise_for_status()

            local_path = os.path.abspath(sys.argv[0])
            with open(local_path, "wb") as f:
                f.write(r.content)

            QMessageBox.information(self, "Atualização Concluída", f"✅ Atualizado para v{versao_online}.\nO app será reiniciado.")
            
            # Inicia um novo processo e fecha o atual
            subprocess.Popen([sys.executable, local_path])
            self.close()

        except Exception as e:
            QMessageBox.critical(self, "Erro na Atualização", f"⚠️ Falha ao baixar ou salvar a atualização: {e}")
            if self.update_btn:
                self.update_btn.setText("⬇ Atualizar agora")
                self.update_btn.setEnabled(True)

    def _build_menu(self):
        menubar = self.menuBar()
        menubar.setStyleSheet("QMenuBar { background-color: #1A1B2C; color: #E7E6E6; } QMenuBar::item:selected { background-color: #43948C; color: #1A1B2C; }")
        
        arquivo_menu = menubar.addMenu("Arquivo")
        
        menu_style = """
        QMenu { 
            background-color: #2A2B3D; 
            color: #E7E6E6; 
            border: 1px solid #43948C; 
        } 
        QMenu::item:selected { 
            background-color: #43948C; 
            color: #1A1B2C; 
        }
        """
        arquivo_menu.setStyleSheet(menu_style)
        
        sair_action = QAction("Sair", self)
        sair_action.triggered.connect(self.close)
        arquivo_menu.addAction(sair_action)

    def apply_dark_style(self):
        BACKGROUND_COLOR = "#1A1B2C"
        BASE_COLOR = "#2A2B3D"
        TEXT_COLOR = "#E7E6E6"
        ACCENT_COLOR = "#43948C"
        DANGER_COLOR = "#B90000"
        BORDER_COLOR = "#43948C"
        STATUS_BACKGROUND = "#2A2B3D"

        QApplication.setStyle("Fusion")
        palette = self.palette()
        palette.setColor(QPalette.ColorRole.Window, QColor(BACKGROUND_COLOR))
        palette.setColor(QPalette.ColorRole.WindowText, QColor(TEXT_COLOR))
        palette.setColor(QPalette.ColorRole.Base, QColor(BASE_COLOR))
        palette.setColor(QPalette.ColorRole.Text, QColor(TEXT_COLOR))
        palette.setColor(QPalette.ColorRole.Button, QColor(BASE_COLOR))
        palette.setColor(QPalette.ColorRole.ButtonText, QColor(TEXT_COLOR))
        palette.setColor(QPalette.ColorRole.Highlight, QColor(ACCENT_COLOR))
        palette.setColor(QPalette.ColorRole.HighlightedText, QColor(BACKGROUND_COLOR))
        self.setPalette(palette)

        self.setStyleSheet(f"""
            QMainWindow, QWidget {{
                background-color: {BACKGROUND_COLOR};
            }}
            QLabel {{ color: {TEXT_COLOR}; font-size: 10pt; }}
            QCheckBox {{ color: {TEXT_COLOR}; spacing: 10px; }}
            QCheckBox::indicator {{ border: 1px solid {ACCENT_COLOR}; border-radius: 3px; width: 15px; height: 15px; }}
            QCheckBox::indicator:checked {{ background-color: {ACCENT_COLOR}; }}
            QLineEdit, QPlainTextEdit {{
                background-color: {BASE_COLOR};
                color: {TEXT_COLOR};
                border: 2px solid {BASE_COLOR};
                border-radius: 8px;
                padding: 10px;
                selection-background-color: {ACCENT_COLOR};
                selection-color: {BACKGROUND_COLOR};
            }}
            QLineEdit:focus, QPlainTextEdit:focus {{ border: 2px solid {BORDER_COLOR}; }}
            QPlainTextEdit {{ padding: 15px; min-height: 200px; }}
            QPushButton#start_btn {{
                background-color: {ACCENT_COLOR};
                color: {BACKGROUND_COLOR};
                border: none;
                border-radius: 10px;
                padding: 12px 25px;
                font-weight: bold;
                font-size: 11pt;
            }}
            QPushButton#start_btn:hover {{ background-color: #63ADA7; }}
            QPushButton#stop_btn {{
                background-color: {DANGER_COLOR};
                color: {TEXT_COLOR};
                border: none;
                border-radius: 10px;
                padding: 12px 25px;
                font-weight: bold;
                font-size: 11pt;
            }}
            QPushButton#stop_btn:hover {{ background-color: #E00000; }}
            QPushButton {{
                background-color: {BASE_COLOR};
                color: {ACCENT_COLOR};
                border: 2px solid {ACCENT_COLOR};
                border-radius: 10px;
                padding: 10px 15px;
                font-weight: bold;
            }}
            QPushButton:hover {{ background-color: #3B3C53; }}
            QPushButton:disabled {{
                background-color: {BASE_COLOR};
                color: #555555;
                border: 2px solid #555555;
            }}
            QProgressBar {{
                border: 2px solid {BORDER_COLOR};
                border-radius: 10px;
                text-align: center;
                color: {TEXT_COLOR};
                background-color: {BASE_COLOR};
                height: 30px;
                font-weight: bold;
                font-size: 10pt;
            }}
            QProgressBar::chunk {{
                background-color: {ACCENT_COLOR};
                border-radius: 7px;
                margin: 2px;
            }}
            QStatusBar {{
                color: {TEXT_COLOR};
                font-size: 10pt; 
                padding: 5px;
                background-color: {STATUS_BACKGROUND};
                border-top: 1px solid #000000;
            }}
        """)

    def append_log(self, text: str):
        self.log_view.appendPlainText(text)
        self.log_view.verticalScrollBar().setValue(self.log_view.verticalScrollBar().maximum())

    def on_browse(self):
        path, _ = QFileDialog.getOpenFileName(self, "Selecionar planilha Excel", str(Path.home()), "Arquivos Excel (*.xlsx)")
        if path:
            self.path_edit.setText(path)

    def on_start(self):
        form_url = self.url_edit.text().strip()
        excel_path = self.path_edit.text().strip()
        headless = self.headless_cb.isChecked()
        keep_open = self.keep_open_cb.isChecked()

        if not form_url.startswith("http"):
            QMessageBox.warning(self, "URL inválida", "Informe uma URL válida do Google Forms.")
            return

        if not excel_path or not excel_path.lower().endswith(".xlsx"):
            QMessageBox.warning(self, "Planilha inválida", "Selecione um arquivo .xlsx válido.")
            return

        self.log_view.clear()
        self.progress_bar.setValue(0)
        self.progress_bar.setFormat("%p%")
        self.status.showMessage("🚀 Iniciando automação...")
        self.toggle_controls(running=True)

        self.append_log("==================================================")
        self.append_log("   INICIANDO AUTOMAÇÃO DE GOOGLE FORMS")
        self.append_log("==================================================")

        self.worker = FormsWorker(
            form_url=form_url,
            excel_path=excel_path,
            field_mapping=FIELD_MAPPING_DEFAULT,
            headless=headless,
            keep_open=keep_open
        )
        self.worker.log.connect(self.append_log)
        self.worker.progress.connect(self.on_progress)
        self.worker.status.connect(self.status.showMessage)
        self.worker.finished.connect(self.on_finished)
        self.worker.start()

    def on_stop(self):
        if self.worker and self.worker.isRunning():
            self.stop_btn.setEnabled(False) 
            self.worker.request_stop()
            self.status.showMessage("🛑 Solicitando parada...")

    def on_progress(self, current: int, total: int):
        if total > 0:
            value = int(((current + 1) / total) * 100)
            self.progress_bar.setValue(value)
            self.progress_bar.setFormat(f"Processando {current + 1}/{total} (%p%)")

    def on_finished(self, successes: int, failures: int, reason: str):
        if self.worker:
            self.worker.wait()
            self.worker = None

        self.toggle_controls(running=False)
        
        if reason == "concluído":
            self.progress_bar.setValue(100)
            self.progress_bar.setFormat("Concluído (%p%)")
        else:
            self.progress_bar.setFormat(f"Interrompido ({reason})")

        summary = f"Finalizado ({reason}). Sucesso: {successes} | Falhas: {failures}"
        
        self.append_log("\n==============================")
        self.append_log("  ✨ RESUMO DA AUTOMAÇÃO ✨")
        self.append_log("==============================")
        self.append_log(f"✅ Enviados com sucesso: {successes}")
        self.append_log(f"❌ Falhas: {failures}")
        self.append_log("==============================")
        
        self.status.showMessage(f"✅ {summary}")

        BACKGROUND_COLOR = "#2A2B3D"
        TEXT_COLOR = "#E7E6E6"
        ACCENT_COLOR = "#43948C"
        
        msg = QMessageBox(self)
        msg.setWindowTitle("Concluído")
        msg.setText(summary)
        msg.setIcon(QMessageBox.Information)
        
        msg.setStyleSheet(f"""
            QMessageBox {{ background-color: {BACKGROUND_COLOR}; color: {TEXT_COLOR}; }}
            QMessageBox QLabel {{ color: {TEXT_COLOR}; font-size: 10pt; }}
            QPushButton {{
                background-color: {ACCENT_COLOR};
                color: #1A1B2C;
                border: none;
                border-radius: 8px;
                padding: 10px 15px;
                font-weight: bold;
            }}
            QPushButton:hover {{ background-color: #63ADA7; }}
        """)
        msg.exec()

    def toggle_controls(self, running: bool):
        self.start_btn.setEnabled(not running)
        self.stop_btn.setEnabled(running)
        self.url_edit.setEnabled(not running)
        self.path_edit.setEnabled(not running)
        self.browse_btn.setEnabled(not running)
        self.headless_cb.setEnabled(not running)
        self.keep_open_cb.setEnabled(not running)
        if self.update_btn:
            self.update_btn.setEnabled(not running)


def main():
    if not QApplication.instance():
        app = QApplication(sys.argv)
    else:
        app = QApplication.instance()
        
    apply_ssl_fix()
    window = MainWindow()
    window.show()
    sys.exit(app.exec())


if __name__ == "__main__":
    main()