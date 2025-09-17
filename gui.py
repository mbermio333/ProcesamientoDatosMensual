# -*- coding: utf-8 -*-
import sys
import io

# Forzar UTF-8 en stdout (evita errores de consola con ñ, tildes, etc.)
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')
import os
import json
import subprocess
import platform
from PyQt5.QtWidgets import (QApplication, QMainWindow, QWidget, QVBoxLayout, 
                             QHBoxLayout, QPushButton, QLabel, QTextEdit, 
                             QFileDialog, QProgressBar, QMessageBox, QGroupBox,
                             QTabWidget, QFrame, QComboBox, QLineEdit, QGridLayout,
                             QScrollArea, QSizePolicy)
from PyQt5.QtCore import Qt, QThread, pyqtSignal
from PyQt5.QtGui import QFont, QDoubleValidator

# Constantes para archivos de configuración separados
CONFIG_PROCESAMIENTO_FILE = "config_procesamiento.json"
CONFIG_OCUPACION_FILE = "config_ocupacion.json"
CIUDADES_UMBRALES_FILE = "config_umbrales_ciudades.json"

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

# Umbrales por defecto para ciudades
DEFAULT_UMBRALES_CIUDADES = {
    "global": {
        "FM": 60.0,
        "TV": {
            "tipo": "general",
            "valor": 45.0,
            "valores": {
                "Banda I-III": 47.0,
                "Banda III": 56.0,
                "Banda IV-V": 64.0
            }
        }
    }
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
    ciudades_signal = pyqtSignal(list)  # Nueva señal para enviar lista de ciudades
    
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
                resultado, ciudades = main.procesar_datos(
                    callback_progreso=self.progress_signal.emit,
                    callback_log=self.log_signal.emit,
                    obtener_ciudades=True  # Nueva bandera para obtener ciudades
                )
                
                # Emitir la lista de ciudades encontradas
                if ciudades:
                    self.ciudades_signal.emit(ciudades)
                
            else:
                # Importar y configurar el módulo principal de ocupación
                import main2
                
                # Ejecutar el procesamiento de ocupación
                self.log_signal.emit("Iniciando análisis de ocupación...")
                
                # Configurar rutas en el módulo main2
                main2.ruta_fm = self.fm_path
                main2.ruta_tv = self.tv_path
                main2.ruta_salida = self.output_path
                
                # Obtener umbrales de la interfaz
                umbrales = self.parent().ocupacion_tab.obtener_umbrales_todos()
                
                # Llamar a la función de ocupación con nuestros callbacks
                resultado = main2.procesar_ocupacion(
                    callback_progreso=self.progress_signal.emit,
                    callback_log=self.log_signal.emit,
                    umbrales=umbrales  # Pasar umbrales por ciudad
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
        self.umbrales_ciudades = self.parent.load_umbrales_ciudades()
        self.ciudades = []
        self.ciudad_actual = "global"
        self.tv_umbral_general = None
        self.tv_umbrales_bandas = {}
        self.initUI()
        
    def initUI(self):
        layout = QVBoxLayout(self)
        layout.setSpacing(10)
        
        # Contenedor principal con dos columnas
        main_container = QHBoxLayout()
        main_container.setSpacing(15)
        
        # Columna izquierda - Configuración de directorios (50% del ancho)
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
        
        # Columna derecha - Selección de ciudad y umbrales (50% del ancho)
        right_column = QVBoxLayout()
        right_column.setSpacing(10)
        
        # Grupo de selección de ciudad
        ciudad_group = QGroupBox("Selección de Ciudad")
        ciudad_group.setStyleSheet(GROUP_BOX_STYLE)
        ciudad_layout = QVBoxLayout()
        ciudad_layout.setSpacing(8)
        
        # Dropdown para seleccionar ciudad
        ciudad_selector_layout = QHBoxLayout()
        ciudad_label = QLabel("Ciudad:")
        ciudad_label.setMinimumWidth(40)
        ciudad_label.setStyleSheet("font-weight: bold;")
        
        self.ciudad_combo = QComboBox()
        self.ciudad_combo.setMaximumWidth(200)
        self.ciudad_combo.setStyleSheet("padding: 3px;")
        self.ciudad_combo.currentTextChanged.connect(self.cambiar_ciudad)
        
        # Botón para actualizar lista de ciudades
        self.actualizar_ciudades_btn = QPushButton("Actualizar Ciudades")
        self.actualizar_ciudades_btn.setStyleSheet(CHANGE_BUTTON_STYLE)
        self.actualizar_ciudades_btn.clicked.connect(self.actualizar_lista_ciudades)
        
        ciudad_selector_layout.addWidget(ciudad_label)
        ciudad_selector_layout.addWidget(self.ciudad_combo)
        ciudad_selector_layout.addWidget(self.actualizar_ciudades_btn)
        ciudad_selector_layout.addStretch(1)
        ciudad_layout.addLayout(ciudad_selector_layout)
        
        ciudad_group.setLayout(ciudad_layout)
        right_column.addWidget(ciudad_group)
        
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
        self.fm_umbral.textChanged.connect(self.guardar_umbral_actual)
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
        self.tv_tipo_umbral.currentTextChanged.connect(self.guardar_umbral_actual)
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
        main_container.addLayout(left_column, 1)
        main_container.addLayout(right_column, 1)
        
        layout.addLayout(main_container)
        
        # Inicializar campos de TV y cargar ciudades
        self.actualizar_campos_tv()
        self.cargar_ciudades()
        
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
        self.log_text.append("2. Configure los umbrales para cada ciudad")
        self.log_text.append("3. Presione 'Iniciar Análisis de Ocupación' para comenzar")
    
    def cargar_ciudades(self):
        """Cargar la lista de ciudades desde el procesamiento anterior"""
        # Primero cargar "global" como opción por defecto
        self.ciudad_combo.clear()
        self.ciudad_combo.addItem("global")
        
        # Agregar ciudades desde la configuración
        for ciudad in self.umbrales_ciudades.keys():
            if ciudad != "global":
                self.ciudad_combo.addItem(ciudad)
        
        # Cargar configuración de la ciudad actual
        self.cargar_configuracion_ciudad()
    
    def actualizar_lista_ciudades(self):
        """Actualizar la lista de ciudades desde los archivos de salida del procesamiento"""
        try:
            output_path = self.output_path_label_ocup.toolTip() or self.output_path_label_ocup.text().replace("...", "")
            if not os.path.exists(output_path):
                self.parent.log_message("La carpeta de salida no existe para buscar ciudades", "ocupacion")
                return
            
            # Buscar archivos CSV en la carpeta de salida
            archivos = [f for f in os.listdir(output_path) if f.endswith('.csv')]
            ciudades_encontradas = set()
            
            for archivo in archivos:
                # Extraer nombre de ciudad del archivo (ej: "Quito_FM.csv" -> "Quito")
                nombre_base = os.path.splitext(archivo)[0]
                if '_' in nombre_base:
                    ciudad = nombre_base.split('_')[0]
                    ciudades_encontradas.add(ciudad)
            
            # Actualizar combo box
            self.ciudad_combo.clear()
            self.ciudad_combo.addItem("global")
            
            for ciudad in sorted(ciudades_encontradas):
                self.ciudad_combo.addItem(ciudad)
                if ciudad not in self.umbrales_ciudades:
                    # Crear entrada por defecto para nueva ciudad
                    self.umbrales_ciudades[ciudad] = {
                        "FM": 60.0,
                        "TV": {
                            "tipo": "general",
                            "valor": 45.0,
                            "valores": {
                                "Banda I-III": 47.0,
                                "Banda III": 56.0,
                                "Banda IV-V": 64.0
                            }
                        }
                    }
            
            self.parent.guardar_umbrales_ciudades(self.umbrales_ciudades)
            self.parent.log_message(f"Lista de ciudades actualizada: {len(ciudades_encontradas)} ciudades encontradas", "ocupacion")
            
        except Exception as e:
            self.parent.log_message(f"Error al actualizar lista de ciudades: {str(e)}", "ocupacion")
    
    def cambiar_ciudad(self, ciudad):
        """Cambiar la ciudad actual y cargar su configuración"""
        if ciudad != self.ciudad_actual:
            # Guardar configuración actual antes de cambiar
            self.guardar_umbral_actual()
            
            # Cambiar a nueva ciudad
            self.ciudad_actual = ciudad
            self.cargar_configuracion_ciudad()
            
            self.parent.log_message(f"Ciudad cambiada a: {ciudad}", "ocupacion")
    
    def cargar_configuracion_ciudad(self):
        """Cargar la configuración de umbrales para la ciudad actual"""
        if self.ciudad_actual not in self.umbrales_ciudades:
            # Crear configuración por defecto si no existe
            self.umbrales_ciudades[self.ciudad_actual] = {
                "FM": 60.0,
                "TV": {
                    "tipo": "general",
                    "valor": 45.0,
                    "valores": {
                        "Banda I-III": 47.0,
                        "Banda III": 56.0,
                        "Banda IV-V": 64.0
                    }
                }
            }
        
        config = self.umbrales_ciudades[self.ciudad_actual]
        
        # Cargar FM
        self.fm_umbral.setText(str(config["FM"]))
        
        # Cargar TV
        tv_config = config["TV"]
        if tv_config["tipo"] == "general":
            self.tv_tipo_umbral.setCurrentText("Umbral General")
            if self.tv_umbral_general:
                self.tv_umbral_general.setText(str(tv_config["valor"]))
        else:
            self.tv_tipo_umbral.setCurrentText("Umbral por bandas")
            for banda, edit in self.tv_umbrales_bandas.items():
                if edit and banda in tv_config["valores"]:
                    edit.setText(str(tv_config["valores"][banda]))
    
    def guardar_umbral_actual(self):
        """Guardar la configuración actual de umbrales para la ciudad actual"""
        if not self.ciudad_actual:
            return
        
        # Obtener valores actuales
        fm_valor = float(self.fm_umbral.text()) if self.fm_umbral.text() else 60.0
        
        tv_config = {
            "tipo": "general" if self.tv_tipo_umbral.currentText() == "Umbral General" else "bandas",
            "valor": 45.0,
            "valores": {
                "Banda I-III": 47.0,
                "Banda III": 56.0,
                "Banda IV-V": 64.0
            }
        }
        
        if tv_config["tipo"] == "general" and self.tv_umbral_general:
            tv_config["valor"] = float(self.tv_umbral_general.text()) if self.tv_umbral_general.text() else 45.0
        elif tv_config["tipo"] == "bandas":
            for banda, edit in self.tv_umbrales_bandas.items():
                if edit:
                    tv_config["valores"][banda] = float(edit.text()) if edit.text() else 0.0
        
        # Guardar en diccionario
        self.umbrales_ciudades[self.ciudad_actual] = {
            "FM": fm_valor,
            "TV": tv_config
        }
        
        # Guardar en archivo
        self.parent.guardar_umbrales_ciudades(self.umbrales_ciudades)
    
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
            self.tv_umbral_general.textChanged.connect(self.guardar_umbral_actual)
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
                umbral_edit.textChanged.connect(self.guardar_umbral_actual)
                
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
        
        # Cargar valores actuales después de crear los campos
        self.cargar_configuracion_ciudad()
    
    def obtener_umbrales_todos(self):
        """Obtener todos los umbrales configurados por ciudad"""
        return self.umbrales_ciudades
    
    def obtener_umbrales_ciudad(self, ciudad):
        """Obtener umbrales para una ciudad específica"""
        if ciudad in self.umbrales_ciudades:
            return self.umbrales_ciudades[ciudad]
        else:
            # Devolver umbrales globales por defecto
            return self.umbrales_ciudades.get("global", {
                "FM": 60.0,
                "TV": {
                    "tipo": "general",
                    "valor": 45.0,
                    "valores": {
                        "Banda I-III": 47.0,
                        "Banda III": 56.0,
                        "Banda IV-V": 64.0
                    }
                }
            })

class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.worker = None
        self.current_tab = None
        self.umbrales_ciudades = {}
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
    
    def load_umbrales_ciudades(self):
        """Cargar umbrales por ciudad desde archivo"""
        if os.path.exists(CIUDADES_UMBRALES_FILE):
            try:
                with open(CIUDADES_UMBRALES_FILE, 'r') as f:
                    return json.load(f)
            except:
                return DEFAULT_UMBRALES_CIUDADES.copy()
        return DEFAULT_UMBRALES_CIUDADES.copy()
    
    def guardar_umbrales_ciudades(self, umbrales):
        """Guardar umbrales por ciudad en archivo"""
        try:
            with open(CIUDADES_UMBRALES_FILE, 'w') as f:
                json.dump(umbrales, f, indent=4)
        except Exception as e:
            self.log_message(f"Error al guardar umbrales: {str(e)}", "ocupacion")
    
    def save_config(self, config, mode):
        """Guardar configuración en archivo JSON específico"""
        config_file = CONFIG_PROCESAMIENTO_FILE if mode == "procesamiento" else CONFIG_OCUPACION_FILE
        try:
            with open(config_file, 'w') as f:
                json.dump(config, f, indent=4)
        except Exception as e:
            self.log_message(f"Error al guardar configuración: {str(e)}", mode)
    
    def initUI(self):
        self.setWindowTitle("Sistema de Procesamiento de Mediciones")
        self.setGeometry(100, 100, 1200, 800)
        
        # Widget central y layout principal
        central_widget = QWidget()
        self.setCentralWidget(central_widget)
        main_layout = QVBoxLayout(central_widget)
        main_layout.setSpacing(10)
        main_layout.setContentsMargins(15, 15, 15, 15)
        
        # Crear pestañas
        self.tabs = QTabWidget()
        self.tabs.setStyleSheet("""
            QTabWidget::pane { border: 1px solid #cccccc; }
            QTabBar::tab { 
                background: #f0f0f0; 
                padding: 8px 12px; 
                border: 1px solid #cccccc; 
                border-bottom: none; 
                border-top-left-radius: 4px; 
                border-top-right-radius: 4px;
            }
            QTabBar::tab:selected { 
                background: #ffffff; 
                font-weight: bold;
            }
        """)
        
        # Crear las pestañas
        self.procesamiento_tab = ProcesamientoTab(self)
        self.ocupacion_tab = OcupacionTab(self)
        
        # Agregar pestañas
        self.tabs.addTab(self.procesamiento_tab, "Procesamiento")
        self.tabs.addTab(self.ocupacion_tab, "Ocupación")
        
        # Conectar señal de cambio de pestaña
        self.tabs.currentChanged.connect(self.cambiar_pestana)
        
        main_layout.addWidget(self.tabs)
        
        # Estado inicial
        self.current_tab = "procesamiento"
        
    def cambiar_pestana(self, index):
        """Manejar el cambio de pestaña"""
        if index == 0:
            self.current_tab = "procesamiento"
        else:
            self.current_tab = "ocupacion"
    
    def truncar_texto(self, texto, max_length=50):
        """Truncar texto largo para mostrar en la interfaz"""
        if len(texto) > max_length:
            return "..." + texto[-max_length:]
        return texto
    
    def change_path(self, path_type, tab_widget, mode):
        """Cambiar ruta de directorio"""
        current_path = ""
        if path_type == "fm":
            current_path = getattr(tab_widget, "fm_path_label" if mode == "procesamiento" else "fm_path_label_ocup").toolTip()
        elif path_type == "tv":
            current_path = getattr(tab_widget, "tv_path_label" if mode == "procesamiento" else "tv_path_label_ocup").toolTip()
        elif path_type == "output":
            current_path = getattr(tab_widget, "output_path_label" if mode == "procesamiento" else "output_path_label_ocup").toolTip()
        elif path_type == "ocupacion_output":
            current_path = getattr(tab_widget, "output_path_label_ocup").toolTip()
        
        # Diálogo para seleccionar directorio
        new_path = QFileDialog.getExistingDirectory(self, f"Seleccionar directorio {path_type.upper()}", current_path)
        
        if new_path:
            # Actualizar etiqueta y tooltip
            if path_type == "fm":
                label = getattr(tab_widget, "fm_path_label" if mode == "procesamiento" else "fm_path_label_ocup")
            elif path_type == "tv":
                label = getattr(tab_widget, "tv_path_label" if mode == "procesamiento" else "tv_path_label_ocup")
            elif path_type == "output":
                label = getattr(tab_widget, "output_path_label" if mode == "procesamiento" else "output_path_label_ocup")
            elif path_type == "ocupacion_output":
                label = getattr(tab_widget, "output_path_label_ocup")
            
            label.setText(self.truncar_texto(new_path))
            label.setToolTip(new_path)
            
            # Guardar configuración
            config = self.load_config(mode)
            if path_type == "fm":
                config["fm_path"] = new_path
            elif path_type == "tv":
                config["tv_path"] = new_path
            elif path_type == "output":
                config["output_path"] = new_path
            elif path_type == "ocupacion_output":
                config["ocupacion_output_path"] = new_path
            
            self.save_config(config, mode)
            
            self.log_message(f"Ruta {path_type.upper()} actualizada: {new_path}", mode)
    
    def open_output_folder(self, mode):
        """Abrir la carpeta de salida en el explorador de archivos"""
        if mode == "procesamiento":
            output_path = self.procesamiento_tab.output_path_label.toolTip()
        else:
            output_path = self.ocupacion_tab.output_path_label_ocup.toolTip()
        
        if os.path.exists(output_path):
            try:
                if platform.system() == "Windows":
                    os.startfile(output_path)
                elif platform.system() == "Darwin":  # macOS
                    subprocess.run(["open", output_path])
                else:  # Linux
                    subprocess.run(["xdg-open", output_path])
                self.log_message(f"Carpeta de salida abierta: {output_path}", mode)
            except Exception as e:
                self.log_message(f"Error al abrir carpeta: {str(e)}", mode)
        else:
            self.log_message("La carpeta de salida no existe", mode)
    
    def start_processing(self, mode):
        """Iniciar el procesamiento"""
        if self.worker and self.worker.isRunning():
            self.log_message("Ya hay un proceso en ejecución", mode)
            return
        
        # Obtener rutas según el modo
        if mode == "procesamiento":
            fm_path = self.procesamiento_tab.fm_path_label.toolTip()
            tv_path = self.procesamiento_tab.tv_path_label.toolTip()
            output_path = self.procesamiento_tab.output_path_label.toolTip()
            tab = self.procesamiento_tab
        else:
            fm_path = self.ocupacion_tab.fm_path_label_ocup.toolTip()
            tv_path = self.ocupacion_tab.tv_path_label_ocup.toolTip()
            output_path = self.ocupacion_tab.output_path_label_ocup.toolTip()
            tab = self.ocupacion_tab
        
        # Verificar que las rutas existan
        if not os.path.exists(fm_path):
            self.log_message(f"ERROR: La ruta FM '{fm_path}' no existe", mode)
            return
        if not os.path.exists(tv_path):
            self.log_message(f"ERROR: La ruta TV '{tv_path}' no existe", mode)
            return
        
        # Crear directorio de salida si no existe
        if not os.path.exists(output_path):
            try:
                os.makedirs(output_path)
                self.log_message(f"Directorio de salida creado: {output_path}", mode)
            except Exception as e:
                self.log_message(f"ERROR: No se pudo crear el directorio de salida: {str(e)}", mode)
                return
        
        # Configurar interfaz
        tab.start_btn.setEnabled(False)
        tab.stop_btn.setEnabled(True)
        tab.progress_bar.setValue(0)
        tab.status_label.setText("Procesando...")
        
        # Crear y configurar worker
        self.worker = WorkerThread(fm_path, tv_path, output_path, mode)
        self.worker.progress_signal.connect(lambda value: tab.progress_bar.setValue(value))
        self.worker.log_signal.connect(lambda msg: self.log_message(msg, mode))
        self.worker.finished_signal.connect(lambda success: self.processing_finished(success, mode))
        
        # Conectar señal de ciudades para la pestaña de ocupación
        if mode == "procesamiento":
            self.worker.ciudades_signal.connect(self.actualizar_ciudades_desde_procesamiento)
        
        # Iniciar worker
        self.worker.start()
        self.log_message("Procesamiento iniciado", mode)
    
    def actualizar_ciudades_desde_procesamiento(self, ciudades):
        """Actualizar la lista de ciudades en la pestaña de ocupación con las ciudades encontradas"""
        if ciudades:
            # Actualizar el diccionario de umbrales con las nuevas ciudades
            for ciudad in ciudades:
                if ciudad not in self.ocupacion_tab.umbrales_ciudades:
                    self.ocupacion_tab.umbrales_ciudades[ciudad] = {
                        "FM": 60.0,
                        "TV": {
                            "tipo": "general",
                            "valor": 45.0,
                            "valores": {
                                "Banda I-III": 47.0,
                                "Banda III": 56.0,
                                "Banda IV-V": 64.0
                            }
                        }
                    }
            
            # Guardar umbrales actualizados
            self.guardar_umbrales_ciudades(self.ocupacion_tab.umbrales_ciudades)
            
            # Actualizar combo box en la pestaña de ocupación
            self.ocupacion_tab.ciudad_combo.clear()
            self.ocupacion_tab.ciudad_combo.addItem("global")
            for ciudad in sorted(self.ocupacion_tab.umbrales_ciudades.keys()):
                if ciudad != "global":
                    self.ocupacion_tab.ciudad_combo.addItem(ciudad)
            
            self.log_message(f"Lista de ciudades actualizada desde procesamiento: {len(ciudades)} ciudades encontradas", "ocupacion")
    
    def stop_processing(self):
        """Detener el procesamiento"""
        if self.worker and self.worker.isRunning():
            self.worker.stop()
            self.worker.terminate()
            self.worker.wait()
            self.log_message("Procesamiento detenido por el usuario", self.current_tab)
            
            # Restablecer interfaz
            if self.current_tab == "procesamiento":
                self.procesamiento_tab.start_btn.setEnabled(True)
                self.procesamiento_tab.stop_btn.setEnabled(False)
                self.procesamiento_tab.status_label.setText("Procesamiento detenido")
            else:
                self.ocupacion_tab.start_btn.setEnabled(True)
                self.ocupacion_tab.stop_btn.setEnabled(False)
                self.ocupacion_tab.status_label.setText("Procesamiento detenido")
    
    def processing_finished(self, success, mode):
        """Manejar la finalización del procesamiento"""
        if mode == "procesamiento":
            tab = self.procesamiento_tab
        else:
            tab = self.ocupacion_tab
        
        tab.start_btn.setEnabled(True)
        tab.stop_btn.setEnabled(False)
        
        if success:
            tab.status_label.setText("Procesamiento completado con éxito")
            self.log_message("Procesamiento completado exitosamente", mode)
        else:
            tab.status_label.setText("Procesamiento falló")
            self.log_message("Procesamiento falló", mode)
    
    def log_message(self, message, mode):
        """Agregar mensaje al log de la pestaña correspondiente"""
        if mode == "procesamiento":
            self.procesamiento_tab.log_text.append(message)
            self.procesamiento_tab.log_text.ensureCursorVisible()
        else:
            self.ocupacion_tab.log_text.append(message)
            self.ocupacion_tab.log_text.ensureCursorVisible()

def main():
    app = QApplication(sys.argv)
    
    # Configurar estilo de la aplicación
    app.setStyle('Fusion')
    
    window = MainWindow()
    window.show()
    
    sys.exit(app.exec_())

if __name__ == '__main__':
    main()