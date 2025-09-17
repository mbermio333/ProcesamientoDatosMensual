# -*- coding: utf-8 -*-
import sys
import os
import json
import subprocess
import platform
from PyQt5.QtWidgets import (QApplication, QMainWindow, QWidget, QVBoxLayout, 
                             QHBoxLayout, QPushButton, QLabel, QTextEdit, 
                             QFileDialog, QProgressBar, QMessageBox, QGroupBox,
                             QTabWidget, QFrame, QComboBox, QLineEdit, QGridLayout)
from PyQt5.QtCore import Qt, QThread, pyqtSignal
from PyQt5.QtGui import QFont, QDoubleValidator

# Constantes para archivos de configuración separados
CONFIG_PROCESAMIENTO_FILE = "config_procesamiento.json"
CONFIG_OCUPACION_FILE = "config_ocupacion.json"

# Configuraciones por defecto para cada pestaña
DEFAULT_PATHS_PROCESAMIENTO = {
    "fm_path": "MedicionesFmCSV",
    "tv_path": "MedicionesTvCSV", 
    "output_path": "ReportesUnificados"
}

DEFAULT_PATHS_OCUPACION = {
    "fm_path": "MedicionesFmCSV",
    "tv_path": "MedicionesTvCSV", 
    "ocupacion_output_path": "ReportesOcupacion"
}

# Estilos globales para mantener consistencia (sin cambios)
GROUP_BOX_STYLE = """
    QGroupBox { 
        font-weight: bold; 
        border: 1px solid #cccccc;
        border-radius: 4px;
        margin-top: 10px;
        padding-top: 15px;
    }
    QGroupBox::title {
        subcontrol-origin: margin;
        left: 10px;
        padding: 0 5px 0 5px;
    }
"""

BUTTON_STYLE = """
    QPushButton { 
        background-color: #4CAF50; 
        color: white; 
        font-weight: bold; 
        padding: 8px;
        border: none;
        border-radius: 4px;
        min-width: 120px;
    }
    QPushButton:hover {
        background-color: #45a049;
    }
    QPushButton:disabled {
        background-color: #cccccc;
    }
"""

STOP_BUTTON_STYLE = """
    QPushButton { 
        background-color: #f44336; 
        color: white; 
        font-weight: bold; 
        padding: 8px;
        border: none;
        border-radius: 4px;
        min-width: 120px;
    }
    QPushButton:hover {
        background-color: #d32f2f;
    }
    QPushButton:disabled {
        background-color: #cccccc;
    }
"""

OPEN_BUTTON_STYLE = """
    QPushButton { 
        background-color: #2196F3; 
        color: white; 
        font-weight: bold; 
        padding: 8px;
        border: none;
        border-radius: 4px;
        min-width: 120px;
    }
    QPushButton:hover {
        background-color: #0b7dda;
    }
"""

CHANGE_BUTTON_STYLE = """
    QPushButton { 
        font-weight: bold; 
        background-color: #e0e0e0;
        border: 1px solid #ccc;
        border-radius: 3px;
    }
    QPushButton:hover {
        background-color: #d0d0d0;
    }
"""

PROGRESS_BAR_STYLE = """
    QProgressBar {
        border: 1px solid #ccc;
        border-radius: 4px;
        text-align: center;
        height: 20px;
    }
    QProgressBar::chunk {
        background-color: #4CAF50;
        width: 10px;
    }
"""

LOG_TEXT_STYLE = """
    font-family: 'Consolas', 'Monospace', monospace; 
    font-size: 10pt;
    background-color: #f8f8f8;
    border: 1px solid #ddd;
    border-radius: 3px;
"""

STATUS_LABEL_STYLE = """
    background-color: #f0f0f0; 
    padding: 8px; 
    border: 1px solid #ddd;
    border-radius: 3px;
    font-weight: bold;
"""

PATH_LABEL_STYLE = """
    background-color: #f8f8f8; 
    padding: 5px; 
    border: 1px solid #ddd;
    border-radius: 3px;
"""

class WorkerThread(QThread):
    """Hilo para ejecutar el procesamiento en segundo plano"""
    progress_signal = pyqtSignal(int)
    log_signal = pyqtSignal(str)
    finished_signal = pyqtSignal(bool)
    
    def __init__(self, fm_path, tv_path, output_path, mode="procesamiento"):
        super().__init__()
        self.running = True
        self.fm_path = fm_path
        self.tv_path = tv_path
        self.output_path = output_path
        self.mode = mode  # "procesamiento" o "ocupacion"
        
    def run(self):
        try:
            if self.mode == "procesamiento":
                # Importar y configurar el módulo principal de procesamiento
                import main
                
                # Ejecutar el procesamiento principal
                self.log_signal.emit("Iniciando procesamiento...")
                
                # Configurar rutas en el módulo main
                main.ruta_fm = self.fm_path
                main.ruta_tv = self.tv_path
                main.ruta_salida = self.output_path
                
                # Llamar a la función principal con nuestros callbacks
                resultado = main.procesar_datos(
                    callback_progreso=self.progress_signal.emit,
                    callback_log=self.log_signal.emit
                )
            else:
                # Importar y configurar el módulo principal de ocupación
                import main2
                
                # Ejecutar el procesamiento de ocupación
                self.log_signal.emit("Iniciando análisis de ocupación...")
                
                # Configurar rutas en el módulo main2
                main2.ruta_fm = self.fm_path
                main2.ruta_tv = self.tv_path
                main2.ruta_salida = self.output_path
                
                # Llamar a la función de ocupación con nuestros callbacks
                resultado = main2.procesar_ocupacion(
                    callback_progreso=self.progress_signal.emit,
                    callback_log=self.log_signal.emit
                )
                
            self.finished_signal.emit(resultado)
            
        except Exception as e:
            self.log_signal.emit(f"Error: {str(e)}")
            import traceback
            self.log_signal.emit(f"Traceback: {traceback.format_exc()}")
            self.finished_signal.emit(False)
    
    def stop(self):
        self.running = False
        self.log_signal.emit("Procesamiento detenido por el usuario")

class ProcesamientoTab(QWidget):
    """Pestaña de procesamiento"""
    def __init__(self, parent=None):
        super().__init__(parent)
        self.parent = parent
        self.config = self.parent.load_config("procesamiento")  # Cargar configuración específica
        self.initUI()
        
    def initUI(self):
        layout = QVBoxLayout(self)
        layout.setSpacing(10)
        
        # Grupo de configuración
        config_group = QGroupBox("Configuración de Directorios - Procesamiento")
        config_group.setStyleSheet(GROUP_BOX_STYLE)
        config_layout = QVBoxLayout()
        config_layout.setSpacing(8)
        
        # Rutas de entrada con mejor formato - usar valores de configuración específica
        paths = [
            ("Ruta FM:", self.config.get("fm_path", "MedicionesFmCSV"), "fm_path_label"),
            ("Ruta TV:", self.config.get("tv_path", "MedicionesTvCSV"), "tv_path_label"), 
            ("Ruta Salida:", self.config.get("output_path", "ReportesUnificados"), "output_path_label")
        ]
        
        for label_text, default_path, attr_name in paths:
            path_layout = QHBoxLayout()
            path_layout.setSpacing(5)
            
            # Crear etiqueta para el texto descriptivo
            label_desc = QLabel(label_text)
            label_desc.setMinimumWidth(80)
            label_desc.setStyleSheet("font-weight: bold;")
            
            # Crear etiqueta para la ruta (con texto truncado)
            label_ruta = QLabel(self.parent.truncar_texto(default_path))
            label_ruta.setStyleSheet(PATH_LABEL_STYLE)
            label_ruta.setMinimumWidth(300)
            label_ruta.setToolTip(default_path)
            setattr(self, attr_name, label_ruta)
            
            # Crear botón "..." para cambiar ruta
            btn_change = QPushButton("...")
            btn_change.setFixedSize(30, 30)
            btn_change.setStyleSheet(CHANGE_BUTTON_STYLE)
            
            # Conectar señal según el tipo de ruta
            if "fm" in attr_name:
                btn_change.clicked.connect(lambda: self.parent.change_path("fm", self, "procesamiento"))
            elif "tv" in attr_name:
                btn_change.clicked.connect(lambda: self.parent.change_path("tv", self, "procesamiento"))
            elif "output" in attr_name:
                btn_change.clicked.connect(lambda: self.parent.change_path("output", self, "procesamiento"))
            
            path_layout.addWidget(label_desc)
            path_layout.addWidget(label_ruta)
            path_layout.addWidget(btn_change)
            path_layout.addStretch(1)
            config_layout.addLayout(path_layout)
        
        config_group.setLayout(config_layout)
        layout.addWidget(config_group)
        # Botones de acción
        action_layout = QHBoxLayout()
        action_layout.setSpacing(10)
        
        self.start_btn = QPushButton("Iniciar Procesamiento")
        self.start_btn.setStyleSheet(BUTTON_STYLE)
        
        self.stop_btn = QPushButton("Detener")
        self.stop_btn.setStyleSheet(STOP_BUTTON_STYLE)
        self.stop_btn.setEnabled(False)
        
        # Nuevo botón para abrir carpeta de salida
        self.open_output_btn = QPushButton("Abrir Carpeta de Salida")
        self.open_output_btn.setStyleSheet(OPEN_BUTTON_STYLE)
        self.open_output_btn.clicked.connect(lambda: self.parent.open_output_folder("procesamiento"))
        
        self.start_btn.clicked.connect(lambda: self.parent.start_processing("procesamiento"))
        self.stop_btn.clicked.connect(self.parent.stop_processing)
        
        action_layout.addWidget(self.start_btn)
        action_layout.addWidget(self.stop_btn)
        action_layout.addWidget(self.open_output_btn)
        layout.addLayout(action_layout)
        
        # Barra de progreso
        self.progress_bar = QProgressBar()
        self.progress_bar.setStyleSheet(PROGRESS_BAR_STYLE)
        layout.addWidget(self.progress_bar)
        
        # Área de log
        log_group = QGroupBox("Log de Actividad - Procesamiento")
        log_group.setStyleSheet(GROUP_BOX_STYLE)
        log_layout = QVBoxLayout()
        self.log_text = QTextEdit()
        self.log_text.setReadOnly(True)
        self.log_text.setStyleSheet(LOG_TEXT_STYLE)
        log_layout.addWidget(self.log_text)
        log_group.setLayout(log_layout)
        layout.addWidget(log_group)
        
        # Estado
        self.status_label = QLabel("Listo para iniciar")
        self.status_label.setStyleSheet(STATUS_LABEL_STYLE)
        layout.addWidget(self.status_label)
        
        # Mensaje inicial
        self.log_text.append("Aplicación iniciada correctamente")
        self.log_text.append("1. Verifique las rutas de los directorios")
        self.log_text.append("2. Presione 'Iniciar Procesamiento' para comenzar")

class OcupacionTab(QWidget):
    """Pestaña de ocupación"""
    def __init__(self, parent=None):
        super().__init__(parent)
        self.parent = parent
        self.config = self.parent.load_config("ocupacion")  # Cargar configuración específica
        self.tv_umbral_general = None
        self.tv_umbrales_bandas = {}
        self.initUI()
        
    def initUI(self):
        layout = QVBoxLayout(self)
        layout.setSpacing(10)
        
        # Contenedor principal con two columnas
        main_container = QHBoxLayout()
        main_container.setSpacing(15)
        
        # Columna izquierda - Configuración de directorios (60% del ancho)
        left_column = QVBoxLayout()
        left_column.setSpacing(10)
        
        # Grupo de configuración
        config_group = QGroupBox("Configuración de Directorios - Ocupación")
        config_group.setStyleSheet(GROUP_BOX_STYLE)
        config_layout = QVBoxLayout()
        config_layout.setSpacing(8)
        
        # Rutas de entrada con mejor formato - usar valores de configuración específica
        paths = [
            ("Ruta FM:", self.config.get("fm_path", "MedicionesFmCSV"), "fm_path_label_ocup"),
            ("Ruta TV:", self.config.get("tv_path", "MedicionesTvCSV"), "tv_path_label_ocup"), 
            ("Ruta Salida:", self.config.get("ocupacion_output_path", "ReportesOcupacion"), "output_path_label_ocup")
        ]
        
        for label_text, default_path, attr_name in paths:
            path_layout = QHBoxLayout()
            path_layout.setSpacing(5)
            
            # Crear etiqueta para el texto descriptivo
            label_desc = QLabel(label_text)
            label_desc.setMinimumWidth(80)
            label_desc.setStyleSheet("font-weight: bold;")
            
            # Crear etiqueta para la ruta (con texto truncado)
            label_ruta = QLabel(self.parent.truncar_texto(default_path))
            label_ruta.setStyleSheet(PATH_LABEL_STYLE)
            label_ruta.setMinimumWidth(200)
            label_ruta.setToolTip(default_path)
            setattr(self, attr_name, label_ruta)
            
            # Crear botón "..." para cambiar ruta
            btn_change = QPushButton("...")
            btn_change.setFixedSize(30, 30)
            btn_change.setStyleSheet(CHANGE_BUTTON_STYLE)
            
            # Conectar señal según el tipo de ruta
            if "fm" in attr_name:
                btn_change.clicked.connect(lambda: self.parent.change_path("fm", self, "ocupacion"))
            elif "tv" in attr_name:
                btn_change.clicked.connect(lambda: self.parent.change_path("tv", self, "ocupacion"))
            elif "output" in attr_name:
                btn_change.clicked.connect(lambda: self.parent.change_path("ocupacion_output", self, "ocupacion"))
            
            path_layout.addWidget(label_desc)
            path_layout.addWidget(label_ruta)
            path_layout.addWidget(btn_change)
            path_layout.addStretch(1)
            config_layout.addLayout(path_layout)
        
        config_group.setLayout(config_layout)
        left_column.addWidget(config_group)
        left_column.addStretch(1)
        
        # Columna derecha - Umbrales (40% del ancho)
        right_column = QVBoxLayout()
        right_column.setSpacing(10)
        
        # Grupo de umbrales
        umbrales_group = QGroupBox("Umbrales")
        umbrales_group.setStyleSheet(GROUP_BOX_STYLE)
        umbrales_layout = QVBoxLayout()
        umbrales_layout.setSpacing(8)
        
        # Campo para FM
        fm_layout = QHBoxLayout()
        fm_layout.setSpacing(5)
        fm_label = QLabel("FM:")
        fm_label.setMinimumWidth(40)
        fm_label.setStyleSheet("font-weight: bold;")
        self.fm_umbral = QLineEdit()
        self.fm_umbral.setValidator(QDoubleValidator(0, 1000, 2))
        self.fm_umbral.setText("60")
        self.fm_umbral.setMaximumWidth(60)
        self.fm_umbral.setStyleSheet("padding: 3px;")
        fm_layout.addWidget(fm_label)
        fm_layout.addWidget(self.fm_umbral)
        fm_layout.addWidget(QLabel("dBµV/m"))
        fm_layout.addStretch(1)
        umbrales_layout.addLayout(fm_layout)
        
        # Separador
        separator = QFrame()
        separator.setFrameShape(QFrame.HLine)
        separator.setFrameShadow(QFrame.Sunken)
        separator.setStyleSheet("color: #e0e0e0;")
        umbrales_layout.addWidget(separator)
        
        # Campo para TV
        tv_layout = QHBoxLayout()
        tv_layout.setSpacing(5)
        tv_label = QLabel("TV:")
        tv_label.setMinimumWidth(40)
        tv_label.setStyleSheet("font-weight: bold;")
        self.tv_tipo_umbral = QComboBox()
        self.tv_tipo_umbral.addItems(["Umbral General", "Umbral por bandas"])
        self.tv_tipo_umbral.currentIndexChanged.connect(self.actualizar_campos_tv)
        self.tv_tipo_umbral.setMaximumWidth(150)
        self.tv_tipo_umbral.setStyleSheet("padding: 3px;")
        tv_layout.addWidget(tv_label)
        tv_layout.addWidget(self.tv_tipo_umbral)
        tv_layout.addStretch(1)
        umbrales_layout.addLayout(tv_layout)
        
        # Contenedor para campos dinámicos de TV
        self.tv_campos_widget = QWidget()
        self.tv_campos_layout = QVBoxLayout(self.tv_campos_widget)
        self.tv_campos_layout.setSpacing(5)
        self.tv_campos_layout.setContentsMargins(0, 5, 0, 0)
        umbrales_layout.addWidget(self.tv_campos_widget)
        
        umbrales_group.setLayout(umbrales_layout)
        right_column.addWidget(umbrales_group)
        right_column.addStretch(1)
        
        # Agregar columnas al contenedor principal
        main_container.addLayout(left_column, 3)
        main_container.addLayout(right_column, 2)
        
        layout.addLayout(main_container)
        
        # Inicializar campos de TV
        self.actualizar_campos_tv()
        
        # Botones de acción
        action_layout = QHBoxLayout()
        action_layout.setSpacing(10)
        
        self.start_btn = QPushButton("Iniciar Análisis de Ocupación")
        self.start_btn.setStyleSheet(BUTTON_STYLE)
        
        self.stop_btn = QPushButton("Detener")
        self.stop_btn.setStyleSheet(STOP_BUTTON_STYLE)
        self.stop_btn.setEnabled(False)
        
        # Nuevo botón para abrir carpeta de salida
        self.open_output_btn = QPushButton("Abrir Carpeta de Salida")
        self.open_output_btn.setStyleSheet(OPEN_BUTTON_STYLE)
        self.open_output_btn.clicked.connect(lambda: self.parent.open_output_folder("ocupacion"))
        
        self.start_btn.clicked.connect(lambda: self.parent.start_processing("ocupacion"))
        self.stop_btn.clicked.connect(self.parent.stop_processing)
        
        action_layout.addWidget(self.start_btn)
        action_layout.addWidget(self.stop_btn)
        action_layout.addWidget(self.open_output_btn)
        layout.addLayout(action_layout)
        
        # Barra de progreso
        self.progress_bar = QProgressBar()
        self.progress_bar.setStyleSheet(PROGRESS_BAR_STYLE)
        layout.addWidget(self.progress_bar)
        
        # Área de log
        log_group = QGroupBox("Log de Actividad - Ocupación")
        log_group.setStyleSheet(GROUP_BOX_STYLE)
        log_layout = QVBoxLayout()
        self.log_text = QTextEdit()
        self.log_text.setReadOnly(True)
        self.log_text.setStyleSheet(LOG_TEXT_STYLE)
        log_layout.addWidget(self.log_text)
        log_group.setLayout(log_layout)
        layout.addWidget(log_group)
        
        # Estado
        self.status_label = QLabel("Listo para iniciar")
        self.status_label.setStyleSheet(STATUS_LABEL_STYLE)
        layout.addWidget(self.status_label)
        
        # Mensaje inicial
        self.log_text.append("Aplicación iniciada correctamente")
        self.log_text.append("1. Verifique las rutas de los directorios")
        self.log_text.append("2. Configure los umbrales si es necesario")
        self.log_text.append("3. Presione 'Iniciar Análisis de Ocupación' para comenzar")
    
    def actualizar_campos_tv(self):
        """Actualiza los campos de TV según la selección del tipo de umbral"""
        # Limpiar layout actual de forma segura
        for i in reversed(range(self.tv_campos_layout.count())):
            item = self.tv_campos_layout.itemAt(i)
            if item.widget():
                item.widget().deleteLater()
            elif item.layout():
                # Limpiar layout hijo
                for j in reversed(range(item.layout().count())):
                    child_item = item.layout().itemAt(j)
                    if child_item.widget():
                        child_item.widget().deleteLater()
                # Eliminar el layout
                self.tv_campos_layout.removeItem(item)
        
        # Limpiar referencias
        self.tv_umbral_general = None
        self.tv_umbrales_bandas = {}
        
        tipo_seleccionado = self.tv_tipo_umbral.currentText()
        
        if tipo_seleccionado == "Umbral General":
            # Campo único para umbral general
            general_layout = QHBoxLayout()
            general_layout.setSpacing(5)
            general_label = QLabel("Umbral General:")
            general_label.setMinimumWidth(100)
            general_label.setStyleSheet("font-weight: bold;")
            self.tv_umbral_general = QLineEdit()
            self.tv_umbral_general.setValidator(QDoubleValidator(0, 1000, 2))
            self.tv_umbral_general.setText("45")
            self.tv_umbral_general.setMaximumWidth(60)
            self.tv_umbral_general.setStyleSheet("padding: 3px;")
            general_layout.addWidget(general_label)
            general_layout.addWidget(self.tv_umbral_general)
            general_layout.addWidget(QLabel("dBµV/m"))
            general_layout.addStretch(1)
            self.tv_campos_layout.addLayout(general_layout)
        else:
            # Umbral por bandas
            bandas = ["Banda I-III", "Banda III", "Banda IV-V"]
            
            for banda in bandas:
                banda_layout = QHBoxLayout()
                banda_layout.setSpacing(5)
                banda_label = QLabel(f"{banda}:")
                banda_label.setMinimumWidth(80)
                banda_label.setStyleSheet("font-weight: bold;")
                umbral_edit = QLineEdit()
                umbral_edit.setValidator(QDoubleValidator(0, 1000, 2))
                umbral_edit.setMaximumWidth(60)
                umbral_edit.setStyleSheet("padding: 3px;")
                
                # Valores por defecto según banda
                if banda == "Banda I-III":
                    umbral_edit.setText("47")
                elif banda == "Banda III":
                    umbral_edit.setText("56")
                else:  # Banda IV-V
                    umbral_edit.setText("64")
                
                self.tv_umbrales_bandas[banda] = umbral_edit
                
                banda_layout.addWidget(banda_label)
                banda_layout.addWidget(umbral_edit)
                banda_layout.addWidget(QLabel("dBµV/m"))
                banda_layout.addStretch(1)
                self.tv_campos_layout.addLayout(banda_layout)

# ... (el resto del código MainWindow y main() permanece igual)
    
    def obtener_umbrales(self):
        """Obtiene los valores de umbrales configurados"""
        umbrales = {
            "FM": float(self.fm_umbral.text()) if self.fm_umbral.text() else 60.0,
            "TV": {}
        }
        
        if self.tv_tipo_umbral.currentText() == "Umbral General":
            umbrales["TV"]["tipo"] = "general"
            if self.tv_umbral_general:
                umbrales["TV"]["valor"] = float(self.tv_umbral_general.text()) if self.tv_umbral_general.text() else 45.0
            else:
                umbrales["TV"]["valor"] = 45.0
        else:
            umbrales["TV"]["tipo"] = "bandas"
            umbrales["TV"]["valores"] = {}
            for banda, edit in self.tv_umbrales_bandas.items():
                if edit:
                    umbrales["TV"]["valores"][banda] = float(edit.text()) if edit.text() else 0.0
                else:
                    # Valores por defecto si no hay edit
                    if banda == "Banda I-III":
                        umbrales["TV"]["valores"][banda] = 47.0
                    elif banda == "Banda III":
                        umbrales["TV"]["valores"][banda] = 56.0
                    else:
                        umbrales["TV"]["valores"][banda] = 64.0
        
        return umbrales
    
class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.worker = None
        self.current_tab = None
        self.initUI()
        
    def load_config(self, mode):
        """Cargar configuración desde archivo JSON específico"""
        config_file = CONFIG_PROCESAMIENTO_FILE if mode == "procesamiento" else CONFIG_OCUPACION_FILE
        default_paths = DEFAULT_PATHS_PROCESAMIENTO if mode == "procesamiento" else DEFAULT_PATHS_OCUPACION
        
        if os.path.exists(config_file):
            try:
                with open(config_file, 'r') as f:
                    return json.load(f)
            except:
                return default_paths.copy()
        return default_paths.copy()
    
    def save_config(self, mode):
        """Guardar configuración en archivo JSON específico"""
        config_file = CONFIG_PROCESAMIENTO_FILE if mode == "procesamiento" else CONFIG_OCUPACION_FILE
        
        # Obtener la configuración actual de la pestaña correspondiente
        if mode == "procesamiento":
            config_data = {
                "fm_path": self.procesamiento_tab.fm_path_label.toolTip() or self.procesamiento_tab.fm_path_label.text().replace("...", ""),
                "tv_path": self.procesamiento_tab.tv_path_label.toolTip() or self.procesamiento_tab.tv_path_label.text().replace("...", ""),
                "output_path": self.procesamiento_tab.output_path_label.toolTip() or self.procesamiento_tab.output_path_label.text().replace("...", "")
            }
        else:
            config_data = {
                "fm_path": self.ocupacion_tab.fm_path_label_ocup.toolTip() or self.ocupacion_tab.fm_path_label_ocup.text().replace("...", ""),
                "tv_path": self.ocupacion_tab.tv_path_label_ocup.toolTip() or self.ocupacion_tab.tv_path_label_ocup.text().replace("...", ""),
                "ocupacion_output_path": self.ocupacion_tab.output_path_label_ocup.toolTip() or self.ocupacion_tab.output_path_label_ocup.text().replace("...", "")
            }
        
        try:
            with open(config_file, 'w') as f:
                json.dump(config_data, f)
        except Exception as e:
            self.log_message(f"Error guardando configuración: {str(e)}", mode)
    
    def truncar_texto(self, texto, max_caracteres=30):
        """Truncar texto largo para mostrar con puntos suspensivos"""
        if len(texto) > max_caracteres:
            return "..." + texto[-max_caracteres:]
        return texto
    
    def initUI(self):
        self.setWindowTitle("Sistema de Reportes Unificados - ARCOTEL")
        self.setGeometry(100, 100, 900, 700)
        
        # Widget central
        central_widget = QWidget()
        self.setCentralWidget(central_widget)
        
        # Layout principal
        layout = QVBoxLayout(central_widget)
        
        # Título
        title_label = QLabel("Sistema de Generación de Reportes Unificados")
        title_label.setFont(QFont("Arial", 16, QFont.Bold))
        title_label.setAlignment(Qt.AlignCenter)
        title_label.setStyleSheet("padding: 10px; background-color: #f0f0f0;")
        layout.addWidget(title_label)
        
        # Crear pestañas
        self.tabs = QTabWidget()
        
        # Pestaña de procesamiento
        self.procesamiento_tab = ProcesamientoTab(self)
        self.tabs.addTab(self.procesamiento_tab, "Procesamiento")
        
        # Pestaña de ocupación
        self.ocupacion_tab = OcupacionTab(self)
        self.tabs.addTab(self.ocupacion_tab, "Ocupación")
        
        # Conectar señal de cambio de pestaña
        self.tabs.currentChanged.connect(self.tab_changed)
        
        layout.addWidget(self.tabs)
        
        # Mensaje inicial
        self.log_message("Aplicación iniciada correctamente", "procesamiento")
        self.log_message("1. Verifique las rutas de los directorios", "procesamiento")
        self.log_message("2. Presione 'Iniciar Procesamiento' para comenzar", "procesamiento")
        
        self.log_message("Aplicación iniciada correctamente", "ocupacion")
        self.log_message("1. Verifique las rutas de los directorios", "ocupacion")
        self.log_message("2. Presione 'Iniciar Análisis de Ocupación' para comenzar", "ocupacion")
    
    def tab_changed(self, index):
        """Manejar cambio de pestaña"""
        self.current_tab = self.tabs.widget(index)
    
    def showEvent(self, event):
        """Se ejecuta cuando la ventana se muestra"""
        super().showEvent(event)
        self.log_message("Ventana principal visible", "procesamiento")
        self.log_message("Ventana principal visible", "ocupacion")
    
    def change_path(self, path_type, tab, mode):
        """Cambiar las rutas de los directorios"""
        dialog = QFileDialog()
        dialog.setFileMode(QFileDialog.Directory)
        
        # Obtener la ruta actual según el tipo y modo
        if mode == "procesamiento":
            if path_type == "fm":
                current_path = tab.fm_path_label.text().replace("...", "")
            elif path_type == "tv":
                current_path = tab.tv_path_label.text().replace("...", "")
            elif path_type == "output":
                current_path = tab.output_path_label.text().replace("...", "")
        else:  # ocupacion
            if path_type == "fm":
                current_path = tab.fm_path_label_ocup.text().replace("...", "")
            elif path_type == "tv":
                current_path = tab.tv_path_label_ocup.text().replace("...", "")
            elif path_type == "ocupacion_output":
                current_path = tab.output_path_label_ocup.text().replace("...", "")
        
        new_path = dialog.getExistingDirectory(self, f"Seleccionar directorio {path_type.upper()}", current_path)
        if not new_path:
            return
            
        truncated_text = self.truncar_texto(new_path)
        
        if mode == "procesamiento":
            if path_type == "fm":
                tab.fm_path_label.setText(truncated_text)
                tab.fm_path_label.setToolTip(new_path)
            elif path_type == "tv":
                tab.tv_path_label.setText(truncated_text)
                tab.tv_path_label.setToolTip(new_path)
            elif path_type == "output":
                tab.output_path_label.setText(truncated_text)
                tab.output_path_label.setToolTip(new_path)
        else:  # ocupacion
            if path_type == "fm":
                tab.fm_path_label_ocup.setText(truncated_text)
                tab.fm_path_label_ocup.setToolTip(new_path)
            elif path_type == "tv":
                tab.tv_path_label_ocup.setText(truncated_text)
                tab.tv_path_label_ocup.setToolTip(new_path)
            elif path_type == "ocupacion_output":
                tab.output_path_label_ocup.setText(truncated_text)
                tab.output_path_label_ocup.setToolTip(new_path)
        
        # Guardar la configuración específica para esta pestaña
        self.save_config(mode)
        self.log_message(f"Ruta {path_type} cambiada a: {new_path}", mode)
    
    
    def open_output_folder(self, mode):
        """Abrir la carpeta de salida en el explorador de archivos"""
        if mode == "procesamiento":
            output_path = self.procesamiento_tab.output_path_label.toolTip() or self.procesamiento_tab.output_path_label.text().replace("...", "")
        else:  # ocupacion
            output_path = self.ocupacion_tab.output_path_label_ocup.toolTip() or self.ocupacion_tab.output_path_label_ocup.text().replace("...", "")
        
        if not os.path.exists(output_path):
            self.log_message(f"La carpeta de salida no existe: {output_path}", mode)
            QMessageBox.warning(self, "Carpeta no encontrada", f"La carpeta de salida no existe:\n{output_path}")
            return
        
        try:
            # Abrir la carpeta según el sistema operativo
            if platform.system() == "Windows":
                os.startfile(output_path)
            elif platform.system() == "Darwin":  # macOS
                subprocess.Popen(["open", output_path])
            else:  # Linux
                subprocess.Popen(["xdg-open", output_path])
                
            self.log_message(f"Carpeta de salida abierta: {output_path}", mode)
        except Exception as e:
            error_msg = f"No se pudo abrir la carpeta: {str(e)}"
            self.log_message(error_msg, mode)
            QMessageBox.critical(self, "Error", error_msg)
    
    def start_processing(self, mode):
        """Iniciar el procesamiento"""
        if mode == "procesamiento":
            # Obtener rutas desde la pestaña de procesamiento
            fm_path = self.procesamiento_tab.fm_path_label.toolTip() or self.procesamiento_tab.fm_path_label.text().replace("...", "")
            tv_path = self.procesamiento_tab.tv_path_label.toolTip() or self.procesamiento_tab.tv_path_label.text().replace("...", "")
            output_path = self.procesamiento_tab.output_path_label.toolTip() or self.procesamiento_tab.output_path_label.text().replace("...", "")
        else:  # ocupacion
            # Obtener rutas desde la pestaña de ocupación
            fm_path = self.ocupacion_tab.fm_path_label_ocup.toolTip() or self.ocupacion_tab.fm_path_label_ocup.text().replace("...", "")
            tv_path = self.ocupacion_tab.tv_path_label_ocup.toolTip() or self.ocupacion_tab.tv_path_label_ocup.text().replace("...", "")
            output_path = self.ocupacion_tab.output_path_label_ocup.toolTip() or self.ocupacion_tab.output_path_label_ocup.text().replace("...", "")
        
        # Verificar que las rutas existan
        if not os.path.exists(fm_path):
            self.log_message(f"ERROR: La ruta FM no existe: {fm_path}", mode)
            return
            
        if not os.path.exists(tv_path):
            self.log_message(f"ERROR: La ruta TV no existe: {tv_path}", mode)
            return
        
        # Crear el hilo de trabajo con las rutas configuradas
        self.worker = WorkerThread(fm_path, tv_path, output_path, mode)
        
        # Conectar señales según la pestaña activa
        if mode == "procesamiento":
            self.worker.progress_signal.connect(self.procesamiento_tab.progress_bar.setValue)
            self.worker.progress_signal.connect(lambda value: self.procesamiento_tab.status_label.setText(f"Procesando... {value}%"))
            self.worker.log_signal.connect(lambda msg: self.log_message(msg, "procesamiento"))
            self.worker.finished_signal.connect(lambda success: self.processing_finished(success, "procesamiento"))
            
            self.procesamiento_tab.start_btn.setEnabled(False)
            self.procesamiento_tab.stop_btn.setEnabled(True)
        else:
            self.worker.progress_signal.connect(self.ocupacion_tab.progress_bar.setValue)
            self.worker.progress_signal.connect(lambda value: self.ocupacion_tab.status_label.setText(f"Procesando... {value}%"))
            self.worker.log_signal.connect(lambda msg: self.log_message(msg, "ocupacion"))
            self.worker.finished_signal.connect(lambda success: self.processing_finished(success, "ocupacion"))
            
            self.ocupacion_tab.start_btn.setEnabled(False)
            self.ocupacion_tab.stop_btn.setEnabled(True)
        
        self.worker.start()
        
        self.log_message(f"Procesamiento de {mode} iniciado...", mode)
    
    def stop_processing(self):
        """Detener el procesamiento"""
        if self.worker and self.worker.isRunning():
            self.worker.stop()
            self.worker.wait()
            
            # Determinar qué pestaña está activa
            current_tab_index = self.tabs.currentIndex()
            mode = "procesamiento" if current_tab_index == 0 else "ocupacion"
            
            self.log_message("Procesamiento detenido por el usuario", mode)
        
        # Reactivar botones en ambas pestañas
        self.procesamiento_tab.start_btn.setEnabled(True)
        self.procesamiento_tab.stop_btn.setEnabled(False)
        self.ocupacion_tab.start_btn.setEnabled(True)
        self.ocupacion_tab.stop_btn.setEnabled(False)
    
    def log_message(self, message, mode="procesamiento"):
        """Añadir mensaje al log de la pestaña especificada"""
        from datetime import datetime
        timestamp = datetime.now().strftime("%H:%M:%S")
        
        if mode == "procesamiento":
            self.procesamiento_tab.log_text.append(f"[{timestamp}] {message}")
            # Auto-scroll to bottom
            scrollbar = self.procesamiento_tab.log_text.verticalScrollBar()
            scrollbar.setValue(scrollbar.maximum())
        else:
            self.ocupacion_tab.log_text.append(f"[{timestamp}] {message}")
            # Auto-scroll to bottom
            scrollbar = self.ocupacion_tab.log_text.verticalScrollBar()
            scrollbar.setValue(scrollbar.maximum())
    
    def processing_finished(self, success, mode):
        """Procesamiento completado"""
        if mode == "procesamiento":
            self.procesamiento_tab.start_btn.setEnabled(True)
            self.procesamiento_tab.stop_btn.setEnabled(False)
            
            if success:
                self.log_message("Procesamiento completado con éxito!", mode)
                self.procesamiento_tab.status_label.setText("Completado con éxito")
                QMessageBox.information(self, "Éxito", "El procesamiento se completó correctamente.")
            else:
                self.log_message("Error en el procesamiento.", mode)
                self.procesamiento_tab.status_label.setText("Error en el procesamiento")
                QMessageBox.warning(self, "Error", "Ocurrió un error durante el procesamiento.")
        else:
            self.ocupacion_tab.start_btn.setEnabled(True)
            self.ocupacion_tab.stop_btn.setEnabled(False)
            
            if success:
                self.log_message("Análisis de ocupación completado con éxito!", mode)
                self.ocupacion_tab.status_label.setText("Completado con éxito")
                QMessageBox.information(self, "Éxito", "El análisis de ocupación se completó correctamente.")
            else:
                self.log_message("Error en el análisis de ocupación.", mode)
                self.ocupacion_tab.status_label.setText("Error en el procesamiento")
                QMessageBox.warning(self, "Error", "Ocurrió un error durante el análisis de ocupación.")
    
    def closeEvent(self, event):
        """Se ejecuta cuando la ventana se cierra"""
        self.save_config()
        super().closeEvent(event)

def main():
    # Configurar la aplicación
    app = QApplication(sys.argv)
    
    # Establecer estilo visual
    app.setStyle('Fusion')
    
    # Crear y mostrar la ventana
    window = MainWindow()
    window.show()
    
    # Forzar el enfoque en la ventana
    window.activateWindow()
    window.raise_()
    
    print("Interfaz gráfica iniciada - Verifica tu pantalla")
    
    # Ejecutar la aplicación
    sys.exit(app.exec_())

if __name__ == "__main__":
    print("Iniciando Sistema de Reportes Unificados...")
    main()