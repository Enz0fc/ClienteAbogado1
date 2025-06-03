import sys
import traceback
import time
from PyQt6.QtWidgets import (
    QApplication, QWidget, QLabel, QPushButton, QVBoxLayout, QHBoxLayout,
    QLineEdit, QTextBrowser, QSizePolicy
)
from PyQt6.QtGui import QPixmap, QFont, QDesktopServices
from PyQt6.QtCore import Qt, QUrl, QThread, pyqtSignal

from funciones_de_API import get_authenticated_service, get_client_by_cedula, process_client_data, subir_archivos_a_drive
from crear_documentos import limpiar_carpeta, FORM_DATOS_NUEVOS_PARA_TRABAJADOR, Carta_Poder, Carta_Compromiso, Desistimiento_de_renuncia, Nota_de_Renuncia, documento_demanda, resource_path


class DocumentWorker(QThread):
    progress = pyqtSignal(str)
    finished = pyqtSignal(str)

    def __init__(self, cliente, drive_service, parent_folder_id):
        super().__init__()
        self.cliente = cliente
        self.drive_service = drive_service
        self.parent_folder_id = parent_folder_id

    def run(self):
        try:
            limpiar_carpeta('Generado')
            self.progress.emit("🗂️ Creando carpeta personal...")
            Carpeta_Personal = process_client_data(self.drive_service, self.parent_folder_id, self.cliente)

            documentos = [
                ("📝 Generando Formulario...", FORM_DATOS_NUEVOS_PARA_TRABAJADOR),
                ("✒️ Generando Carta de Poder...", Carta_Poder),
                ("🤝 Generando Carta de Compromiso...", Carta_Compromiso),
                ("📄 Generando Desistimiento de Renuncia...", Desistimiento_de_renuncia),
                ("📃 Generando Nota de Renuncia...", Nota_de_Renuncia),
                ("⚖️ Redactando documento de demanda...", documento_demanda)
            ]

            for mensaje, funcion in documentos:
                self.progress.emit(mensaje)
                funcion(self.cliente)
                time.sleep(0.1)

            self.progress.emit("☁️ Subiendo archivos a Drive...")
            subir_archivos_a_drive(self.drive_service, 'Generado', Carpeta_Personal)

            self.finished.emit("✅ Documentos generados y subidos exitosamente.")
        except Exception as e:
            self.finished.emit(f"❌ Error al generar documentos:\n{str(e)}")


class DemandaLaboralApp(QWidget):
    def __init__(self):
        super().__init__()
        self.verification_state = False
        self.cliente = {}

        self.SPREADSHEET_ID = '1hkNDtfxtE2qfUZY9CPwH0uvAFO09qFLqk9ihPZ3pY8A'
        self.SHEET_NAME = 'RespuesaClientes'
        self.PARENT_FOLDER_ID = '15wU5vReziVY9STabIHSEKxsaUpQtQJ1b'

        self.sheets_service = get_authenticated_service('sheets', 'v4')
        self.drive_service = get_authenticated_service('drive', 'v3')

        self.init_ui()

    def init_ui(self):
        self.setWindowTitle("Clientes Demanda Laboral")
        self.setFixedSize(650, 750)

        #logo_path = resource_path("Logo/logobernis.png")
        self.logo_label = QLabel()
        #pixmap = QPixmap(logo_path).scaled(250, 250, Qt.AspectRatioMode.KeepAspectRatio, Qt.TransformationMode.SmoothTransformation)
        #self.logo_label.setPixmap(pixmap)
        self.logo_label.setAlignment(Qt.AlignmentFlag.AlignCenter)

        self.title_label = QLabel("CASO NUEVO")
        self.title_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.title_label.setFont(QFont("Arial", 20, QFont.Weight.Bold))
        self.title_label.setStyleSheet("color: #0D47A1;")

        self.cedula_input = QLineEdit()
        self.cedula_input.setPlaceholderText("Ingrese número de cédula del cliente")
        self.cedula_input.setFixedHeight(40)
        self.cedula_input.setStyleSheet("padding: 10px; font-size: 16px;")

        self.verify_button = QPushButton("Verificar Cédula")
        self.verify_button.clicked.connect(self.toggle_verification)

        self.reset_button = QPushButton("Nueva Consulta")
        self.reset_button.clicked.connect(self.toggle_verification)
        self.reset_button.hide()

        self.generate_button = QPushButton("Generar Documentos")
        self.generate_button.clicked.connect(self.generate_documents)
        self.generate_button.hide()

        self.status_text = QTextBrowser()
        self.status_text.setOpenExternalLinks(True)
        self.status_text.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Expanding)
        self.status_text.setStyleSheet("font-size: 14px;")

        button_layout = QHBoxLayout()
        button_layout.addWidget(self.verify_button)
        button_layout.addWidget(self.reset_button)
        button_layout.addWidget(self.generate_button)

        layout = QVBoxLayout()
        layout.addWidget(self.logo_label)
        layout.addWidget(self.title_label)
        layout.addWidget(self.cedula_input)
        layout.addLayout(button_layout)
        layout.addWidget(self.status_text)
        self.setLayout(layout)

    def toggle_verification(self):
        if not self.verification_state:
            cedula = self.cedula_input.text()
            if not cedula:
                self.update_status("✅ Ingrese el N° de cédula del cliente")
                return

            self.set_loading(True, "🔍 Buscando cliente...")

            try:
                datos_cliente = get_client_by_cedula(
                    self.sheets_service,
                    self.SPREADSHEET_ID,
                    self.SHEET_NAME,
                    cedula
                )
                if datos_cliente:
                    self.cliente = datos_cliente
                    nombre = datos_cliente.get("Nombres y Apellidos completos como esta en tu Cedula.", "").upper()
                    self.update_status(f"✅ Cliente Encontrado: {nombre}")
                    self.verification_state = True
                else:
                    self.update_status(
                        f"❌ No se encontró cliente con cédula {cedula}. "
                        f'<a href="https://docs.google.com/forms/d/1Zkys_Yhf9_Q00TG6Et5zE_RD0SUaLTkNHjGPN80n_Yw">Registrar aquí</a>'
                    )
                    return
            except Exception as e:
                self.update_status(f"❌ Error: {str(e)}")

        else:
            self.verification_state = False
            self.cedula_input.clear()
            self.update_status("")

        self.update_buttons()
        self.set_loading(False)

    def update_buttons(self):
        self.verify_button.setVisible(not self.verification_state)
        self.reset_button.setVisible(self.verification_state)
        self.generate_button.setVisible(self.verification_state)

    def update_status(self, message):
        self.status_text.setHtml(message)

    def set_loading(self, loading, message=""):
        self.verify_button.setEnabled(not loading)
        self.reset_button.setEnabled(not loading)
        self.generate_button.setEnabled(not loading)
        if message:
            self.update_status(f"⏳ {message}")

    def generate_documents(self):
        self.set_loading(True)
        self.worker = DocumentWorker(self.cliente, self.drive_service, self.PARENT_FOLDER_ID)
        self.worker.progress.connect(self.update_status)
        self.worker.finished.connect(self.on_generation_finished)
        self.worker.start()

    def on_generation_finished(self, message):
        self.update_status(message)
        self.verification_state = False
        self.cedula_input.clear()
        self.update_buttons()
        self.set_loading(False)


def main():
    app = QApplication(sys.argv)
    window = DemandaLaboralApp()
    window.show()
    sys.exit(app.exec())


if __name__ == "__main__":
    main()
