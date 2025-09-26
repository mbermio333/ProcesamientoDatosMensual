# -*- coding: utf-8 -*-
import sys
import io
import pandas as pd
from datetime import datetime

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
                             QScrollArea, QSizePolicy, QDialog, QTableWidget, 
                             QTableWidgetItem, QHeaderView, QDialogButtonBox,
                             QCheckBox, QButtonGroup, QRadioButton)  # Agregar QCheckBox y QRadioButton
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
# Umbrales por defecto para ciudades - ELIMINAR "global"
DEFAULT_UMBRALES_CIUDADES = {
    # Eliminar completamente la entrada "global"
}

def normalizar_nombre_ciudad(nombre):
    """Normaliza el nombre de la ciudad para consistencia"""
    if not nombre or not isinstance(nombre, str):
        return ""
    
    nombre = nombre.lower().strip()
    
    # Manejar todas las variantes de "cañar" y "tambo"
    #if nombre in ["cañar", "cañar", "canar", "caã±ar", "tambo"]:
    #    return "TAMBO"  # Mostrar "TAMBO" en la interfaz, pero guardar en "cañar"
    
    # Mapeo de otras ciudades si es necesario
    mapeo_ciudades = {
        "zamora": "ZAMORA",
        "loja": "LOJA", 
        "macas": "MACAS",
        "tambo":"TAMBO",
        "machala": "MACHALA",
        "cuenca": "CUENCA"
    }

    return mapeo_ciudades.get(nombre, nombre.upper())

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

class AdvertenciaOcupacionCeroDialog(QDialog):
    def __init__(self, frecuencias_cero, parent=None):
        super().__init__(parent)
        self.frecuencias_cero = frecuencias_cero
        self.setWindowTitle("Advertencia - Frecuencias con Ocupación 0%")
        self.setModal(True)
        self.setMinimumSize(800, 500)
        self.initUI()
        
    def initUI(self):
        layout = QVBoxLayout(self)
        
        # Mensaje de advertencia
        mensaje_label = QLabel(
            "Se han detectado frecuencias autorizadas o no autorizadas con ocupación = 0%.\nLos porcentajes obtenidos podrian ser inconsistentes.\nSe recomienda reconsiderar el umbral.\n"
            #"Los archivos Excel han sido generados. ¿Qué desea hacer?"
        )
        mensaje_label.setStyleSheet("font-weight: bold; color: #d32f2f; font-size: 12pt;")
        mensaje_label.setAlignment(Qt.AlignCenter)
        layout.addWidget(mensaje_label)
        
        # Información adicional
        info_label = QLabel(
            "• FINALIZAR: Mantener los archivos generados y terminar el proceso\n"
            "• CANCELAR: Eliminar todos los archivos generados y cancelar el proceso"
        )
        info_label.setStyleSheet("color: #666; font-size: 10pt; margin: 10px;")
        layout.addWidget(info_label)
        
        # Tabla de frecuencias con ocupación 0%
        if self.frecuencias_cero:
            tabla_label = QLabel("Frecuencias detectadas con ocupación 0%:")
            tabla_label.setStyleSheet("font-weight: bold; margin-top: 10px;")
            layout.addWidget(tabla_label)
            
            self.tabla = QTableWidget()
            self.tabla.setColumnCount(6)
            self.tabla.setHorizontalHeaderLabels([
                "CIUDAD/PROVINCIA", "ESTACIÓN", "TIPO", "Frecuencia (MHz)", 
                "Ocupación (%)", "Level (dBµV/m)"
            ])
            
            # Configurar tabla
            self.tabla.setRowCount(len(self.frecuencias_cero))
            header = self.tabla.horizontalHeader()
            header.setSectionResizeMode(QHeaderView.ResizeToContents)
            header.setStretchLastSection(True)
            
            # Llenar tabla con datos
            for row, frecuencia in enumerate(self.frecuencias_cero):
                self.tabla.setItem(row, 0, QTableWidgetItem(frecuencia.get('ciudad', '')))
                self.tabla.setItem(row, 1, QTableWidgetItem(frecuencia.get('estacion', '')))
                self.tabla.setItem(row, 2, QTableWidgetItem(frecuencia.get('tipo', '')))
                self.tabla.setItem(row, 3, QTableWidgetItem(str(frecuencia.get('frecuencia', ''))))
                self.tabla.setItem(row, 4, QTableWidgetItem(str(frecuencia.get('ocupacion', ''))))
                self.tabla.setItem(row, 5, QTableWidgetItem(str(frecuencia.get('level', ''))))
            
            layout.addWidget(self.tabla)
        
        # Botones
        botones_layout = QHBoxLayout()
        
        self.btn_continuar = QPushButton("Finalizar Procesamiento")
        self.btn_continuar.setStyleSheet(BUTTON_STYLE)
        self.btn_continuar.clicked.connect(self.accept)
        self.btn_continuar.setToolTip("Mantener los archivos generados y terminar el proceso")
        
        self.btn_cancelar = QPushButton("Cancelar y Eliminar Archivos")
        self.btn_cancelar.setStyleSheet(STOP_BUTTON_STYLE)
        self.btn_cancelar.clicked.connect(self.reject)
        self.btn_cancelar.setToolTip("Eliminar todos los archivos generados y cancelar el proceso")
        
        botones_layout.addWidget(self.btn_continuar)
        botones_layout.addWidget(self.btn_cancelar)
        
        layout.addLayout(botones_layout)


class ObservacionTab(QWidget):
    """Pestaña de frecuencias en observación"""
    def __init__(self, parent=None):
        super().__init__(parent)
        self.parent = parent
        self.umbrales_ciudades = self.parent.load_umbrales_ciudades()
        self.datos_actuales = {"FM": [], "TV": []}
        self.config_data = self.cargar_configuracion()
        self.initUI()
    
    def cargar_configuracion(self):
        """Cargar la configuración desde config.json"""
        try:
            with open('config.json', 'r', encoding='utf-8') as f:
                return json.load(f)
        except Exception as e:
            self.log_text.append(f"❌ Error al cargar config.json: {str(e)}")
            return {"emisoras_por_ciudad": {}}
    
    def guardar_configuracion(self):
        """Guardar la configuración en config.json"""
        try:
            with open('config.json', 'w', encoding='utf-8') as f:
                json.dump(self.config_data, f, indent=4, ensure_ascii=False)
            return True
        except Exception as e:
            self.log_text.append(f"❌ Error al guardar config.json: {str(e)}")
            return False

    def initUI(self):
        layout = QVBoxLayout(self)
        layout.setSpacing(10)
        
        # Título de la pestaña
        titulo_label = QLabel("Frecuencias en Observación")
        titulo_label.setStyleSheet("font-size: 16pt; font-weight: bold; color: #2c3e50; margin: 10px;")
        titulo_label.setAlignment(Qt.AlignCenter)
        layout.addWidget(titulo_label)
        
        # Grupo de selección de ciudad
        ciudad_group = QGroupBox("Selección de Ciudad")
        ciudad_group.setStyleSheet(GROUP_BOX_STYLE)
        ciudad_layout = QVBoxLayout()
        
        # Layout horizontal para la selección de ciudad
        seleccion_layout = QHBoxLayout()
        
        ciudad_label = QLabel("Ciudad:")
        ciudad_label.setMinimumWidth(60)
        ciudad_label.setStyleSheet("font-weight: bold;")
        
        self.ciudad_combo = QComboBox()
        self.ciudad_combo.setMaximumWidth(250)
        self.ciudad_combo.setStyleSheet("padding: 5px; font-size: 10pt;")
        self.ciudad_combo.currentTextChanged.connect(self.cargar_frecuencias_ciudad)
        
        self.btn_actualizar = QPushButton("Actualizar")
        self.btn_actualizar.setStyleSheet(CHANGE_BUTTON_STYLE)
        self.btn_actualizar.clicked.connect(self.actualizar_datos)
        
        seleccion_layout.addWidget(ciudad_label)
        seleccion_layout.addWidget(self.ciudad_combo)
        seleccion_layout.addWidget(self.btn_actualizar)
        seleccion_layout.addStretch(1)
        
        ciudad_layout.addLayout(seleccion_layout)
        ciudad_group.setLayout(ciudad_layout)
        layout.addWidget(ciudad_group)
        
        # Información de resultados
        self.info_label = QLabel("Seleccione una ciudad para ver las frecuencias en observación")
        self.info_label.setAlignment(Qt.AlignCenter)
        self.info_label.setStyleSheet("color: #666; font-size: 11pt; padding: 10px; background-color: #f8f8f8; border-radius: 4px;")
        layout.addWidget(self.info_label)
        
        # Botones de acción para estados
        estados_group = QGroupBox("Gestión de Estados")
        estados_group.setStyleSheet(GROUP_BOX_STYLE)
        estados_layout = QHBoxLayout()
        
        #self.btn_guardar_estados = QPushButton("Guardar Estados")
        #self.btn_guardar_estados.setStyleSheet(BUTTON_STYLE)
        #self.btn_guardar_estados.clicked.connect(self.guardar_estados)
        #self.btn_guardar_estados.setEnabled(False)
        
        self.btn_guardar_config = QPushButton("Guardar en Configuración")
        self.btn_guardar_config.setStyleSheet(OPEN_BUTTON_STYLE)
        self.btn_guardar_config.clicked.connect(self.guardar_en_configuracion)
        self.btn_guardar_config.setEnabled(False)
        self.btn_guardar_config.setToolTip("Guardar los cambios en el archivo config.json")
        
        self.btn_exportar_excel = QPushButton("Exportar a Excel")
        self.btn_exportar_excel.setStyleSheet(OPEN_BUTTON_STYLE)
        self.btn_exportar_excel.clicked.connect(self.exportar_a_excel)
        self.btn_exportar_excel.setEnabled(False)
        
        #estados_layout.addWidget(self.btn_guardar_estados)
        estados_layout.addWidget(self.btn_guardar_config)
        estados_layout.addWidget(self.btn_exportar_excel)
        estados_layout.addStretch(1)
        
        estados_group.setLayout(estados_layout)
        layout.addWidget(estados_group)
        
        # Crear pestañas para FM y TV
        self.tabs_datos = QTabWidget()
        self.tabs_datos.setStyleSheet("""
            QTabWidget::pane { border: 1px solid #cccccc; }
            QTabBar::tab { 
                background: #f0f0f0; 
                padding: 6px 10px; 
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
        
        # Pestaña FM
        self.tab_fm = QWidget()
        layout_fm = QVBoxLayout(self.tab_fm)
        self.tabla_fm = QTableWidget()
        layout_fm.addWidget(self.tabla_fm)
        self.tabs_datos.addTab(self.tab_fm, "Frecuencias FM")
        
        # Pestaña TV
        self.tab_tv = QWidget()
        layout_tv = QVBoxLayout(self.tab_tv)
        self.tabla_tv = QTableWidget()
        layout_tv.addWidget(self.tabla_tv)
        self.tabs_datos.addTab(self.tab_tv, "Frecuencias TV")
        
        layout.addWidget(self.tabs_datos)
        
        # Área de log
        log_group = QGroupBox("Log de Actividad")
        log_group.setStyleSheet(GROUP_BOX_STYLE)
        log_layout = QVBoxLayout()
        self.log_text = QTextEdit()
        self.log_text.setReadOnly(True)
        self.log_text.setStyleSheet(LOG_TEXT_STYLE)
        self.log_text.setMaximumHeight(150)
        log_layout.addWidget(self.log_text)
        log_group.setLayout(log_layout)
        layout.addWidget(log_group)
        
        # Cargar las ciudades al inicializar
        self.cargar_ciudades()
    
    def cargar_ciudades(self):
        """Cargar la lista de ciudades desde la configuración"""
        self.ciudad_combo.clear()
        
        # Agregar ciudades desde la configuración, filtrando nombres vacíos y "global"
        ciudades_unicas = set()
        
        for ciudad in self.umbrales_ciudades.keys():
            if ciudad != "global":  # Excluir "global"
                ciudad_normalizada = normalizar_nombre_ciudad(ciudad)
                if ciudad_normalizada and ciudad_normalizada.strip():
                    # Evitar duplicados (como "cañar" y "tambo" que se mapean ambos a "TAMBO")
                    if ciudad_normalizada not in ciudades_unicas:
                        ciudades_unicas.add(ciudad_normalizada)
                        self.ciudad_combo.addItem(ciudad_normalizada)
        
        # Si no hay ciudades, agregar un mensaje
        if self.ciudad_combo.count() == 0:
            self.ciudad_combo.addItem("No hay ciudades configuradas")
    
    def actualizar_datos(self):
        """Actualizar los datos para la ciudad seleccionada - FORZAR RECARGA COMPLETA"""
        ciudad_actual = self.ciudad_combo.currentText()
        if ciudad_actual and ciudad_actual != "No hay ciudades configuradas":
            self.log_text.append("🔄 Actualizando datos desde archivos recientes...")
            
            # Limpiar completamente los datos previos
            self.datos_actuales = {"FM": [], "TV": []}
            self.info_label.setText("Actualizando datos...")
            
            # Limpiar las tablas visualmente
            self.tabla_fm.setRowCount(0)
            self.tabla_tv.setRowCount(0)
            
            # Forzar recarga completa
            self.cargar_frecuencias_ciudad(ciudad_actual)
            
            self.log_text.append("✅ Actualización completada")
    
    def cargar_frecuencias_ciudad(self, ciudad):
        """Cargar las frecuencias en observación para la ciudad seleccionada"""
        if not ciudad or ciudad == "No hay ciudades configuradas":
            return
        
        self.log_text.append(f"🔍 Buscando frecuencias en observación para: {ciudad}")
        
        try:
            # Importar el módulo main3
            import main3
            
            # Obtener la ruta de salida de ocupación
            ruta_salida_ocupacion = self.parent.ocupacion_tab.output_path_label_ocup.toolTip()
            
            if not os.path.exists(ruta_salida_ocupacion):
                self.log_text.append(f"❌ La ruta de salida no existe: {ruta_salida_ocupacion}")
                return
            
            # Llamar a la función de main3
            resultado = main3.procesar_frecuencias_observacion(
                ruta_salida_ocupacion, 
                ciudad,
                callback_log=self.log_text.append
            )
            
            if resultado.get('error'):
                self.log_text.append(f"❌ Error: {resultado['error']}")
                self.info_label.setText(f"Error: {resultado['error']}")
                return
            
            # Guardar datos actuales
            self.datos_actuales = resultado
            
            # Actualizar la información
            total_fm = resultado.get('total_fm', 0)
            total_tv = resultado.get('total_tv', 0)
            archivo = resultado.get('archivo_utilizado', 'N/A')
            
            self.info_label.setText(
                f"📊 Encontradas {total_fm} frecuencias FM y {total_tv} frecuencias TV en observación\n"
                f"📁 Archivo: {archivo}"
            )
            
            # Actualizar tablas
            self.actualizar_tabla_fm(resultado.get('FM', []))
            self.actualizar_tabla_tv(resultado.get('TV', []))
            
            # Habilitar botones
            tiene_datos = total_fm + total_tv > 0
            #self.btn_guardar_estados.setEnabled(tiene_datos)
            self.btn_guardar_config.setEnabled(tiene_datos)
            self.btn_exportar_excel.setEnabled(tiene_datos)
            
            self.log_text.append("✅ Datos cargados correctamente")
            
        except Exception as e:
            error_msg = f"❌ Error al cargar datos: {str(e)}"
            self.log_text.append(error_msg)
            self.info_label.setText(error_msg)
    
    def crear_checkbox_estado(self, estado_actual="Observación"):
        """Crear grupo de radio buttons para seleccionar estado"""
        widget = QWidget()
        layout = QHBoxLayout(widget)
        layout.setContentsMargins(0, 0, 0, 0)
        layout.setSpacing(5)
        
        grupo = QButtonGroup(widget)
        
        rb_observacion = QRadioButton("Obs")
        rb_observacion.setToolTip("Observación")
        rb_no_autorizada = QRadioButton("No Aut")
        rb_no_autorizada.setToolTip("No Autorizada")
        
        grupo.addButton(rb_observacion, 1)
        grupo.addButton(rb_no_autorizada, 2)
        
        # Establecer estado inicial
        if estado_actual == "No Autorizada":
            rb_no_autorizada.setChecked(True)
        else:
            rb_observacion.setChecked(True)
        
        layout.addWidget(rb_observacion)
        layout.addWidget(rb_no_autorizada)
        layout.addStretch(1)
        
        return widget, grupo
 
    def actualizar_tabla_fm(self, datos_fm):
        """Actualizar la tabla de frecuencias FM - Con campos editables para estación"""
        if not datos_fm:
            self.tabla_fm.setRowCount(0)
            self.tabla_fm.setColumnCount(1)
            self.tabla_fm.setHorizontalHeaderLabels(["No hay frecuencias FM en observación"])
            return
        
        # Definir columnas para FM (incluyendo Estado)
        columnas = [
            'Frecuencia (MHz)', 'Estación', 'Ocupación (%)', 'Level (dBµV/m)', 'Estado'
        ]
        
        self.tabla_fm.setRowCount(len(datos_fm))
        self.tabla_fm.setColumnCount(len(columnas))
        self.tabla_fm.setHorizontalHeaderLabels(columnas)
        
        # Guardar referencias a los grupos de botones
        self.grupos_fm = []
        
        for fila, dato in enumerate(datos_fm):
            for col, columna in enumerate(columnas):
                if columna == 'Estado':
                    # Crear radio buttons para estado
                    estado_actual = dato.get('Estado', 'Observación')
                    widget_estado, grupo = self.crear_checkbox_estado(estado_actual)
                    self.grupos_fm.append((fila, grupo))
                    self.tabla_fm.setCellWidget(fila, col, widget_estado)
                elif columna == 'Estación':
                    # Crear campo editable para el nombre de la estación
                    nombre_estacion = dato.get('Estación', '')
                    edit_estacion = QLineEdit(nombre_estacion)
                    edit_estacion.setStyleSheet("padding: 2px;")
                    edit_estacion.textChanged.connect(lambda text, f=fila: self.actualizar_nombre_estacion(f, text, 'FM'))
                    self.tabla_fm.setCellWidget(fila, col, edit_estacion)
                else:
                    valor = dato.get(columna, '')
                    # Formatear valores numéricos
                    if columna == 'Frecuencia (MHz)' and valor != '':
                        try:
                            valor = f"{float(valor):.2f}"
                        except:
                            pass
                    elif columna == 'Ocupación (%)' and valor != '':
                        try:
                            valor = f"{float(valor):.1f}%"
                        except:
                            pass
                    elif columna == 'Level (dBµV/m)' and valor != '':
                        try:
                            valor = f"{float(valor):.1f}"
                        except:
                            pass
                    
                    item = QTableWidgetItem(str(valor))
                    # Alinear números a la derecha
                    if columna in ['Frecuencia (MHz)', 'Ocupación (%)', 'Level (dBµV/m)']:
                        item.setTextAlignment(Qt.AlignRight | Qt.AlignVCenter)
                    self.tabla_fm.setItem(fila, col, item)
        
        # Ajustar el tamaño de las columnas
        header = self.tabla_fm.horizontalHeader()
        header.setSectionResizeMode(QHeaderView.ResizeToContents)
        header.setStretchLastSection(True)
    
    def actualizar_tabla_tv(self, datos_tv):
        """Actualizar la tabla de frecuencias TV - Con campos editables para estación y estado"""
        if not datos_tv:
            self.tabla_tv.setRowCount(0)
            self.tabla_tv.setColumnCount(1)
            self.tabla_tv.setHorizontalHeaderLabels(["No hay frecuencias TV en observación"])
            return
        
        # Definir columnas para TV (incluyendo Estado)
        columnas = [
            'Frecuencia (MHz)', 'Estación', 'Banda', 'Canal', 'Ocupación (%)', 'Level (dBµV/m)', 'Estado'
        ]
        
        self.tabla_tv.setRowCount(len(datos_tv))
        self.tabla_tv.setColumnCount(len(columnas))
        self.tabla_tv.setHorizontalHeaderLabels(columnas)
        
        # Guardar referencias a los grupos de botones
        self.grupos_tv = []
        
        for fila, dato in enumerate(datos_tv):
            for col, columna in enumerate(columnas):
                if columna == 'Estado':
                    # Crear radio buttons para estado
                    estado_actual = dato.get('Estado', 'Observación')
                    widget_estado, grupo = self.crear_checkbox_estado(estado_actual)
                    self.grupos_tv.append((fila, grupo))
                    self.tabla_tv.setCellWidget(fila, col, widget_estado)
                elif columna == 'Estación':
                    # Crear campo editable para el nombre de la estación
                    nombre_estacion = dato.get('Estación', '')
                    edit_estacion = QLineEdit(nombre_estacion)
                    edit_estacion.setStyleSheet("padding: 2px;")
                    edit_estacion.textChanged.connect(lambda text, f=fila: self.actualizar_nombre_estacion(f, text, 'TV'))
                    self.tabla_tv.setCellWidget(fila, col, edit_estacion)
                else:
                    valor = dato.get(columna, '')
                    # Formatear valores numéricos
                    if columna == 'Frecuencia (MHz)' and valor != '':
                        try:
                            valor = f"{float(valor):.2f}"
                        except:
                            pass
                    elif columna == 'Ocupación (%)' and valor != '':
                        try:
                            valor = f"{float(valor):.1f}%"
                        except:
                            pass
                    elif columna == 'Level (dBµV/m)' and valor != '':
                        try:
                            valor = f"{float(valor):.1f}"
                        except:
                            pass
                    
                    item = QTableWidgetItem(str(valor))
                    # Alinear números a la derecha
                    if columna in ['Frecuencia (MHz)', 'Ocupación (%)', 'Level (dBµV/m)']:
                        item.setTextAlignment(Qt.AlignRight | Qt.AlignVCenter)
                    self.tabla_tv.setItem(fila, col, item)
        
        # Ajustar el tamaño de las columnas
        header = self.tabla_tv.horizontalHeader()
        header.setSectionResizeMode(QHeaderView.ResizeToContents)
        header.setStretchLastSection(True)
        
        # Resaltar filas con ocupación > 0
        for fila in range(len(datos_tv)):
            for col in range(len(columnas)):
                if columnas[col] == 'Ocupación (%)':
                    item = self.tabla_tv.item(fila, col)
                    if item and '%' in item.text():
                        try:
                            ocupacion = float(item.text().replace('%', ''))
                            if ocupacion > 0:
                                item.setBackground(Qt.yellow)
                        except:
                            pass
    
    def actualizar_nombre_estacion(self, fila, texto, tipo):
        """Actualizar el nombre de la estación en los datos actuales"""
        if tipo == 'FM' and fila < len(self.datos_actuales.get('FM', [])):
            self.datos_actuales['FM'][fila]['Estación'] = texto
        elif tipo == 'TV' and fila < len(self.datos_actuales.get('TV', [])):
            self.datos_actuales['TV'][fila]['Estación'] = texto
    
    def obtener_estados_actuales(self):
        """Obtener los estados actuales seleccionados en las tablas"""
        estados_fm = []
        estados_tv = []
        
        # Obtener estados de FM
        if hasattr(self, 'grupos_fm'):
            for fila, grupo in self.grupos_fm:
                if grupo.checkedButton():
                    estado = "No Autorizada" if grupo.checkedButton().text() == "No Aut" else "Observación"
                    frecuencia = self.tabla_fm.item(fila, 0).text() if self.tabla_fm.item(fila, 0) else ""
                    # Obtener el nombre de la estación del campo editable
                    estacion_widget = self.tabla_fm.cellWidget(fila, 1)
                    nombre_estacion = estacion_widget.text() if estacion_widget else ""
                    
                    estados_fm.append({
                        'fila': fila,
                        'frecuencia': frecuencia,
                        'estacion': nombre_estacion,
                        'estado': estado
                    })
        
        # Obtener estados de TV (AGREGAR ESTA SECCIÓN)
        if hasattr(self, 'grupos_tv'):
            for fila, grupo in self.grupos_tv:
                if grupo.checkedButton():
                    estado = "No Autorizada" if grupo.checkedButton().text() == "No Aut" else "Observación"
                    frecuencia = self.tabla_tv.item(fila, 0).text() if self.tabla_tv.item(fila, 0) else ""
                    # Obtener el nombre de la estación del campo editable
                    estacion_widget = self.tabla_tv.cellWidget(fila, 1)
                    nombre_estacion = estacion_widget.text() if estacion_widget else ""
                    
                    estados_tv.append({
                        'fila': fila,
                        'frecuencia': frecuencia,
                        'estacion': nombre_estacion,
                        'estado': estado
                    })
        
        return estados_fm, estados_tv
    
    """def guardar_estados(self):
        #Guardar los estados seleccionados en un archivo JSON
        try:
            ciudad_actual = self.ciudad_combo.currentText()
            if not ciudad_actual or ciudad_actual == "No hay ciudades configuradas":
                QMessageBox.warning(self, "Advertencia", "Seleccione una ciudad primero.")
                return
            
            estados_fm, estados_tv = self.obtener_estados_actuales()
            
            datos_guardar = {
                'ciudad': ciudad_actual,
                'fecha_actualizacion': datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                'FM': estados_fm,
                'TV': estados_tv,
                'total_fm': len(estados_fm),
                'total_tv': len(estados_tv)
            }
            
            # Crear nombre de archivo
            nombre_archivo = f"estados_frecuencias_{ciudad_actual.lower().replace(' ', '_')}.json"
            
            with open(nombre_archivo, 'w', encoding='utf-8') as f:
                json.dump(datos_guardar, f, indent=4, ensure_ascii=False)
            
            self.log_text.append(f"✅ Estados guardados en: {nombre_archivo}")
            QMessageBox.information(self, "Éxito", f"Estados guardados correctamente en {nombre_archivo}")
            
        except Exception as e:
            error_msg = f"❌ Error al guardar estados: {str(e)}"
            self.log_text.append(error_msg)
            QMessageBox.critical(self, "Error", error_msg)"""
    
    def guardar_en_configuracion(self):
        """Guardar las frecuencias en observación en el archivo config.json"""
        try:
            ciudad_actual = self.ciudad_combo.currentText()
            if not ciudad_actual or ciudad_actual == "No hay ciudades configuradas":
                QMessageBox.warning(self, "Advertencia", "Seleccione una ciudad primero.")
                return
            
            estados_fm, estados_tv = self.obtener_estados_actuales()
            
            if not estados_fm and not estados_tv:
                QMessageBox.warning(self, "Advertencia", "No hay frecuencias para guardar.")
                return
            
            # Normalizar nombre de ciudad para buscar en config.json
            ciudad_normalizada = ciudad_actual.lower()
            
            # MAPEO ESPECIAL: Si la ciudad es "tambo", guardar en "cañar"
            ciudad_guardar = ciudad_normalizada
            """if ciudad_normalizada == "tambo":
                ciudad_guardar = "cañar"
                self.log_text.append(f"🔀 Ciudad 'TAMBO' mapeada a 'CAÑAR' para guardado en configuración")
            """
            # Asegurarse de que la ciudad existe en la configuración
            if ciudad_guardar not in self.config_data.get('emisoras_por_ciudad', {}):
                self.config_data['emisoras_por_ciudad'][ciudad_guardar] = {'FM': [], 'TV': []}
            
            # Procesar frecuencias FM
            for estado_fm in estados_fm:
                frecuencia = float(estado_fm['frecuencia'])
                nombre_estacion = estado_fm['estacion']
                
                # Agregar sufijo según el estado
                if estado_fm['estado'] == "No Autorizada":
                    nombre_completo = f"{nombre_estacion} (NO AUTORIZADA)"
                else:
                    nombre_completo = f"{nombre_estacion}_OBSERVACION"
                
                # Buscar si ya existe esta frecuencia
                frecuencia_existente = False
                for emisora in self.config_data['emisoras_por_ciudad'][ciudad_guardar]['FM']:
                    if abs(emisora['frecuencia'] - frecuencia) < 0.01:  # Tolerancia para comparación de floats
                        emisora['nombre'] = nombre_completo
                        frecuencia_existente = True
                        break
                
                # Si no existe, agregar nueva
                if not frecuencia_existente:
                    nueva_emisora = {
                        "nombre": nombre_completo,
                        "frecuencia": frecuencia,
                        "tipo": "FM"
                    }
                    self.config_data['emisoras_por_ciudad'][ciudad_guardar]['FM'].append(nueva_emisora)
            
            # Procesar frecuencias TV
            for estado_tv in estados_tv:
                frecuencia = float(estado_tv['frecuencia'])
                nombre_estacion = estado_tv['estacion']
                
                # Agregar sufijo según el estado
                if estado_tv['estado'] == "No Autorizada":
                    nombre_completo = f"{nombre_estacion} (NO AUTORIZADA)"
                else:
                    nombre_completo = f"{nombre_estacion}_OBSERVACION"
                
                # Buscar si ya existe esta frecuencia
                frecuencia_existente = False
                for emisora in self.config_data['emisoras_por_ciudad'][ciudad_guardar]['TV']:
                    if abs(emisora['frecuencia'] - frecuencia) < 0.01:  # Tolerancia para comparación de floats
                        emisora['nombre'] = nombre_completo
                        frecuencia_existente = True
                        break
                
                # Si no existe, agregar nueva
                if not frecuencia_existente:
                    nueva_emisora = {
                        "nombre": nombre_completo,
                        "frecuencia": frecuencia,
                        "tipo": "TV"
                    }
                    self.config_data['emisoras_por_ciudad'][ciudad_guardar]['TV'].append(nueva_emisora)
            
            # Guardar la configuración
            if self.guardar_configuracion():
                if ciudad_normalizada == "tambo":
                    mensaje_ciudad = "TAMBO (guardado en CAÑAR)"
                else:
                    mensaje_ciudad = ciudad_actual
                    
                self.log_text.append(f"✅ Configuración guardada en config.json para {mensaje_ciudad}")
                QMessageBox.information(self, "Éxito", f"Configuración guardada correctamente en config.json para {mensaje_ciudad}")
            else:
                raise Exception("No se pudo guardar el archivo config.json")
            
        except Exception as e:
            error_msg = f"❌ Error al guardar en configuración: {str(e)}"
            self.log_text.append(error_msg)
            QMessageBox.critical(self, "Error", error_msg)
    
    def exportar_a_excel(self):
        """Exportar las frecuencias con sus estados a Excel"""
        try:
            ciudad_actual = self.ciudad_combo.currentText()
            if not ciudad_actual or ciudad_actual == "No hay ciudades configuradas":
                QMessageBox.warning(self, "Advertencia", "Seleccione una ciudad primero.")
                return
            
            # Obtener datos actuales con estados
            datos_fm = self.datos_actuales.get('FM', [])
            datos_tv = self.datos_actuales.get('TV', [])
            
            # Actualizar datos con estados seleccionados
            estados_fm, estados_tv = self.obtener_estados_actuales()
            
            # Actualizar estados en datos FM
            for estado in estados_fm:
                if estado['fila'] < len(datos_fm):
                    datos_fm[estado['fila']]['Estado'] = estado['estado']
            
            # Actualizar estados en datos TV
            for estado in estados_tv:
                if estado['fila'] < len(datos_tv):
                    datos_tv[estado['fila']]['Estado'] = estado['estado']
            
            # Crear DataFrames
            df_fm = pd.DataFrame(datos_fm)
            df_tv = pd.DataFrame(datos_tv)
            
            # Crear nombre de archivo
            nombre_archivo = f"frecuencias_observacion_{ciudad_actual.lower().replace(' ', '_')}.xlsx"
            
            # Exportar a Excel
            with pd.ExcelWriter(nombre_archivo, engine='openpyxl') as writer:
                if not df_fm.empty:
                    df_fm.to_excel(writer, sheet_name='FM', index=False)
                if not df_tv.empty:
                    df_tv.to_excel(writer, sheet_name='TV', index=False)
            
            self.log_text.append(f"✅ Datos exportados a: {nombre_archivo}")
            QMessageBox.information(self, "Éxito", f"Datos exportados correctamente a {nombre_archivo}")
            
        except Exception as e:
            error_msg = f"❌ Error al exportar a Excel: {str(e)}"
            self.log_text.append(error_msg)
            QMessageBox.critical(self, "Error", error_msg)
    
    def exportar_a_excel(self):
        """Exportar las frecuencias con sus estados a Excel"""
        try:
            ciudad_actual = self.ciudad_combo.currentText()
            if not ciudad_actual or ciudad_actual == "No hay ciudades configuradas":
                QMessageBox.warning(self, "Advertencia", "Seleccione una ciudad primero.")
                return
            
            # Obtener datos actuales con estados
            datos_fm = self.datos_actuales.get('FM', [])
            datos_tv = self.datos_actuales.get('TV', [])
            
            # Actualizar datos con estados seleccionados
            estados_fm, estados_tv = self.obtener_estados_actuales()
            
            # Actualizar estados en datos FM
            for estado in estados_fm:
                if estado['fila'] < len(datos_fm):
                    datos_fm[estado['fila']]['Estado'] = estado['estado']
            
            # Actualizar estados en datos TV
            for estado in estados_tv:
                if estado['fila'] < len(datos_tv):
                    datos_tv[estado['fila']]['Estado'] = estado['estado']
            
            # Crear DataFrames
            df_fm = pd.DataFrame(datos_fm)
            df_tv = pd.DataFrame(datos_tv)
            
            # Crear nombre de archivo
            nombre_archivo = f"frecuencias_observacion_{ciudad_actual.lower().replace(' ', '_')}.xlsx"
            
            # Exportar a Excel
            with pd.ExcelWriter(nombre_archivo, engine='openpyxl') as writer:
                if not df_fm.empty:
                    df_fm.to_excel(writer, sheet_name='FM', index=False)
                if not df_tv.empty:
                    df_tv.to_excel(writer, sheet_name='TV', index=False)
            
            self.log_text.append(f"✅ Datos exportados a: {nombre_archivo}")
            QMessageBox.information(self, "Éxito", f"Datos exportados correctamente a {nombre_archivo}")
            
        except Exception as e:
            error_msg = f"❌ Error al exportar a Excel: {str(e)}"
            self.log_text.append(error_msg)
            QMessageBox.critical(self, "Error", error_msg)
    
    


# Modificar la clase WorkerThread para que reciba la referencia de la ventana principal
class WorkerThread(QThread):
    """Hilo para ejecutar el procesamiento en segundo plano"""
    progress_signal = pyqtSignal(int)
    log_signal = pyqtSignal(str)
    finished_signal = pyqtSignal(object)  # Cambiar a object para recibir diferentes tipos de datos
    ciudades_signal = pyqtSignal(list)
    
    def __init__(self, fm_path, tv_path, output_path, mode="procesamiento", parent_window=None):
        super().__init__()
        self.running = True
        self.fm_path = fm_path
        self.tv_path = tv_path
        self.output_path = output_path
        self.mode = mode
        self.parent_window = parent_window
        
    def run(self):
        try:
            if self.mode == "procesamiento":
                # Importar y configurar el módulo principal de procesamiento
                import main
                
                self.log_signal.emit("Iniciando procesamiento...")
                
                # Configurar rutas
                main.ruta_fm = self.fm_path
                main.ruta_tv = self.tv_path
                main.ruta_salida = self.output_path
                
                # Llamar a la función principal
                resultado, ciudades = main.procesar_datos(
                    callback_progreso=self.progress_signal.emit,
                    callback_log=self.log_signal.emit,
                    obtener_ciudades=True
                )
                
                # Emitir ciudades encontradas
                if ciudades:
                    self.ciudades_signal.emit(ciudades)
                
                # Para procesamiento, emitir booleano
                self.finished_signal.emit(resultado)
                
            else:  # modo == "ocupacion"
                # Importar y configurar el módulo principal de ocupación
                import main2
                
                self.log_signal.emit("Iniciando análisis de ocupación...")
                
                # Configurar rutas
                main2.ruta_fm = self.fm_path
                main2.ruta_tv = self.tv_path
                main2.ruta_salida = self.output_path
                
                # Obtener umbrales de la interfaz
                if self.parent_window and hasattr(self.parent_window, 'ocupacion_tab'):
                    umbrales = self.parent_window.ocupacion_tab.obtener_umbrales_todos()
                else:
                    umbrales = {}
                    self.log_signal.emit("⚠️  Usando umbrales por defecto")
                
                # Llamar a la función de ocupación
                resultado = main2.procesar_ocupacion(
                    callback_progreso=self.progress_signal.emit,
                    callback_log=self.log_signal.emit,
                    umbrales=umbrales
                )
                
                # Para ocupación, emitir el diccionario completo
                self.finished_signal.emit(resultado)
            
        except Exception as e:
            self.log_signal.emit(f"Error: {str(e)}")
            import traceback
            self.log_signal.emit(f"Traceback: {traceback.format_exc()}")
            
            # Emitir resultado de error según el modo
            if self.mode == "procesamiento":
                self.finished_signal.emit(False)
            else:
                self.finished_signal.emit({"existen_problemas": False, "error": str(e)})
    
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

# ... (código anterior sin cambios)

class OcupacionTab(QWidget):
    """Pestaña de ocupación"""
    def __init__(self, parent=None):
        super().__init__(parent)
        self.parent = parent
        self.config = self.parent.load_config("ocupacion")
        self.umbrales_ciudades = self.parent.load_umbrales_ciudades()
        self.ciudades = []
        # Eliminar "global" como ciudad por defecto, usar la primera ciudad disponible
        self.ciudad_actual = self.obtener_primera_ciudad() or ""
        self.tv_umbral_general = None
        self.tv_umbrales_bandas = {}
        self.initUI()
    
    def obtener_primera_ciudad(self):
        """Obtener la primera ciudad disponible (excluyendo 'global')"""
        ciudades = [ciudad for ciudad in self.umbrales_ciudades.keys() 
                   if ciudad != "global" and ciudad.strip()]
        return ciudades[0] if ciudades else ""
    
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
        
        # Columna derecha - Umbrales (50% del ancho)
        right_column = QVBoxLayout()
        right_column.setSpacing(10)
        
        # Grupo de umbrales (ahora incluye selección de ciudad)
        umbrales_group = QGroupBox("Umbrales")
        umbrales_group.setStyleSheet(GROUP_BOX_STYLE)
        umbrales_layout = QVBoxLayout()
        umbrales_layout.setSpacing(8)
        
        # Selección de ciudad dentro del grupo de umbrales
        ciudad_layout = QHBoxLayout()
        ciudad_label = QLabel("Ciudad:")
        ciudad_label.setMinimumWidth(40)
        ciudad_label.setStyleSheet("font-weight: bold;")
        
        self.ciudad_combo = QComboBox()
        self.ciudad_combo.setMaximumWidth(200)
        self.ciudad_combo.setStyleSheet("padding: 3px;")
        self.ciudad_combo.currentTextChanged.connect(self.cambiar_ciudad)
        
        # ELIMINAR BOTÓN ACTUALIZAR CIUDADES
        # self.actualizar_ciudades_btn = QPushButton("Actualizar Ciudades")
        # self.actualizar_ciudades_btn.setStyleSheet(CHANGE_BUTTON_STYLE)
        # self.actualizar_ciudades_btn.clicked.connect(self.actualizar_lista_ciudades)
        
        ciudad_layout.addWidget(ciudad_label)
        ciudad_layout.addWidget(self.ciudad_combo)
        # ciudad_layout.addWidget(self.actualizar_ciudades_btn)  # ELIMINAR ESTA LÍNEA
        ciudad_layout.addStretch(1)
        umbrales_layout.addLayout(ciudad_layout)
        
        # Separador
        separator = QFrame()
        separator.setFrameShape(QFrame.HLine)
        separator.setFrameShadow(QFrame.Sunken)
        separator.setStyleSheet("color: #e0e0e0;")
        umbrales_layout.addWidget(separator)
        
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
        # Eliminé la conexión a currentTextChanged para evitar conflictos
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
        self.ciudad_combo.clear()
        
        # Agregar ciudades desde la configuración, filtrando nombres vacíos y "global"
        for ciudad in self.umbrales_ciudades.keys():
            if ciudad != "global":  # Excluir "global"
                ciudad_normalizada = normalizar_nombre_ciudad(ciudad)
                if ciudad_normalizada and ciudad_normalizada.strip():
                    self.ciudad_combo.addItem(ciudad_normalizada)
        
        # Si no hay ciudades, agregar un mensaje
        if self.ciudad_combo.count() == 0:
            self.ciudad_combo.addItem("No hay ciudades configuradas")
            self.ciudad_actual = ""
        else:
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
                    ciudad = nombre_base.split('_')[0].strip()
                    ciudad_normalizada = normalizar_nombre_ciudad(ciudad)
                    if ciudad_normalizada and ciudad_normalizada != "GLOBAL":  # Excluir "GLOBAL"
                        ciudades_encontradas.add(ciudad_normalizada)
            
            # Actualizar combo box - ELIMINAR "global"
            self.ciudad_combo.clear()
            
            for ciudad in sorted(ciudades_encontradas):
                if ciudad and ciudad.strip():  # Filtrar nombres vacíos
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
            
            # Filtrar "global" antes de guardar
            if "global" in self.umbrales_ciudades:
                del self.umbrales_ciudades["global"]
                
            self.parent.guardar_umbrales_ciudades(self.umbrales_ciudades)
            self.parent.log_message(f"Lista de ciudades actualizada: {len(ciudades_encontradas)} ciudades encontradas", "ocupacion")
            
        except Exception as e:
            self.parent.log_message(f"Error al actualizar lista de ciudades: {str(e)}", "ocupacion")

    def cambiar_ciudad(self, ciudad):
        """Cambiar la ciudad actual y cargar su configuración"""
        if not ciudad or not ciudad.strip() or ciudad == "No hay ciudades configuradas":
            return  # No hacer nada si no hay ciudad válida
        
        if ciudad != self.ciudad_actual:
            # Guardar configuración actual antes de cambiar
            self.guardar_umbral_actual()
            
            # Cambiar a nueva ciudad
            self.ciudad_actual = ciudad
            self.cargar_configuracion_ciudad()
            
            self.parent.log_message(f"Ciudad cambiada a: {ciudad}", "ocupacion")
# ... (código anterior sin cambios)

    def cargar_configuracion_ciudad(self):
        """Cargar la configuración de umbrales para la ciudad actual"""
        if not self.ciudad_actual or not self.ciudad_actual.strip():
            return
        
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

    def actualizar_campos_tv(self):
        """Actualiza los campos de TV según la selección del tipo de umbral"""
        # Desconectar temporalmente la señal para evitar recursión
        try:
            self.tv_tipo_umbral.currentIndexChanged.disconnect(self.actualizar_campos_tv)
        except:
            pass
            
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
            
            # Obtener valor actual de la configuración
            valor_actual = "45"  # Valor por defecto
            if (self.ciudad_actual in self.umbrales_ciudades and 
                "TV" in self.umbrales_ciudades[self.ciudad_actual] and
                "valor" in self.umbrales_ciudades[self.ciudad_actual]["TV"]):
                valor_actual = str(self.umbrales_ciudades[self.ciudad_actual]["TV"]["valor"])
            
            self.tv_umbral_general.setText(valor_actual)
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
                
                # Obtener valor actual de la configuración
                valor_actual = "47"  # Valor por defecto
                if banda == "Banda III":
                    valor_actual = "56"
                elif banda == "Banda IV-V":
                    valor_actual = "64"
                
                # Intentar cargar valor guardado
                if (self.ciudad_actual in self.umbrales_ciudades and 
                    "TV" in self.umbrales_ciudades[self.ciudad_actual] and
                    "valores" in self.umbrales_ciudades[self.ciudad_actual]["TV"] and
                    banda in self.umbrales_ciudades[self.ciudad_actual]["TV"]["valores"]):
                    valor_actual = str(self.umbrales_ciudades[self.ciudad_actual]["TV"]["valores"][banda])
                
                umbral_edit.setText(valor_actual)
                
                self.tv_umbrales_bandas[banda] = umbral_edit
                
                banda_layout.addWidget(banda_label)
                banda_layout.addWidget(umbral_edit)
                banda_layout.addWidget(QLabel("dBµV/m"))
                banda_layout.addStretch(1)
                self.tv_campos_layout.addLayout(banda_layout)
        
        # Reconectar la señal después de actualizar los campos
        self.tv_tipo_umbral.currentIndexChanged.connect(self.actualizar_campos_tv)
        
        # Asegurarse de que los cambios se muestren inmediatamente
        self.tv_campos_widget.update()

    def guardar_umbral_actual(self):
        """Guardar la configuración actual de umbrales para la ciudad actual"""
        if not self.ciudad_actual or not self.ciudad_actual.strip():
            return
        
        # Obtener valores actuales
        try:
            fm_valor = float(self.fm_umbral.text()) if self.fm_umbral.text() else 60.0
        except:
            fm_valor = 60.0
        
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
            try:
                tv_config["valor"] = float(self.tv_umbral_general.text()) if self.tv_umbral_general.text() else 45.0
            except:
                tv_config["valor"] = 45.0
        elif tv_config["tipo"] == "bandas":
            for banda, edit in self.tv_umbrales_bandas.items():
                if edit:
                    try:
                        tv_config["valores"][banda] = float(edit.text()) if edit.text() else 0.0
                    except:
                        tv_config["valores"][banda] = 0.0
        
        # Guardar en diccionario
        self.umbrales_ciudades[self.ciudad_actual] = {
            "FM": fm_valor,
            "TV": tv_config
        }
        
        # Guardar en archivo
        self.parent.guardar_umbrales_ciudades(self.umbrales_ciudades)

    
    
    def obtener_umbrales_todos(self):
        """Obtener todos los umbrales configurados por ciudad"""
        # Filtrar ciudades con nombres vacíos y excluir "global"
        umbrales_filtrados = {}
        for ciudad, config in self.umbrales_ciudades.items():
            if ciudad and ciudad.strip() and ciudad != "global":  # Excluir "global"
                umbrales_filtrados[ciudad] = config
        return umbrales_filtrados
    
    def obtener_umbrales_ciudad(self, ciudad):
        """Obtener umbrales para una ciudad específica"""
        if ciudad in self.umbrales_ciudades:
            return self.umbrales_ciudades[ciudad]
        else:
            # Devolver umbrales por defecto (sin "global")
            return {
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

# ... (el resto del código MainWindow y main() permanece igual)

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
                    umbrales = json.load(f)
                    # Filtrar "global" al cargar
                    if "global" in umbrales:
                        del umbrales["global"]
                    return umbrales
            except:
                return DEFAULT_UMBRALES_CIUDADES.copy()
        return DEFAULT_UMBRALES_CIUDADES.copy()
    
    
    
    def guardar_umbrales_ciudades(self, umbrales):
        """Guardar umbrales por ciudad en archivo"""
        try:
            # Filtrar "global" antes de guardar
            umbrales_filtrados = {k: v for k, v in umbrales.items() if k != "global"}
            with open(CIUDADES_UMBRALES_FILE, 'w') as f:
                json.dump(umbrales_filtrados, f, indent=4)
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
        self.observacion_tab = ObservacionTab(self)  # Nueva pestaña
        
        # Agregar pestañas
        self.tabs.addTab(self.procesamiento_tab, "Procesamiento")
        self.tabs.addTab(self.ocupacion_tab, "Ocupación")
        self.tabs.addTab(self.observacion_tab, "Frecuencias en Observación")  # Nueva pestaña
        
        # Conectar señal de cambio de pestaña
        self.tabs.currentChanged.connect(self.cambiar_pestana)
        
        main_layout.addWidget(self.tabs)
        
        # Estado inicial
        self.current_tab = "procesamiento"
        
    def cambiar_pestana(self, index):
        """Manejar el cambio de pestaña"""
        if index == 0:
            self.current_tab = "procesamiento"
        elif index == 1:
            self.current_tab = "ocupacion"
        else:  # index == 2
            self.current_tab = "observacion"
            # Actualizar la lista de ciudades cuando se cambie a esta pestaña
            self.observacion_tab.cargar_ciudades()
    
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
        
        # Crear y configurar worker - pasar la referencia de la ventana principal
        self.worker = WorkerThread(fm_path, tv_path, output_path, mode, self)  # Pasar self como parent_window
        
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
        if not ciudades:  # Si no hay ciudades, salir temprano
            self.log_message("No se encontraron ciudades en el procesamiento", "ocupacion")
            return
        
        # Inicializar la lista desde el principio
        ciudades_normalizadas = []
        for ciudad in ciudades:
            ciudad_normalizada = normalizar_nombre_ciudad(ciudad)
            # Filtrar nombres vacíos y también evitar "GLOBAL" que viene del procesamiento
            if ciudad_normalizada and ciudad_normalizada.strip() and ciudad_normalizada != "GLOBAL":
                ciudades_normalizadas.append(ciudad_normalizada)
        
        # Verificar si después del filtrado quedan ciudades
        if not ciudades_normalizadas:
            self.log_message("Todas las ciudades fueron filtradas (posiblemente solo 'GLOBAL')", "ocupacion")
            return
        
        # Eliminar duplicados
        ciudades_normalizadas = list(set(ciudades_normalizadas))
        
        # Actualizar el diccionario de umbrales con las nuevas ciudades NORMALIZADAS
        for ciudad in ciudades_normalizadas:
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
        
        # Asegurarse de eliminar "global" si existe
        if "global" in self.ocupacion_tab.umbrales_ciudades:
            del self.ocupacion_tab.umbrales_ciudades["global"]
        
        # Guardar umbrales actualizados
        self.guardar_umbrales_ciudades(self.ocupacion_tab.umbrales_ciudades)
        
        # Actualizar combo box en la pestaña de ocupación con ciudades NORMALIZADAS
        self.ocupacion_tab.ciudad_combo.clear()
        
        # Ordenar alfabéticamente y agregar ciudades (excluyendo "global")
        ciudades_ordenadas = sorted([c for c in self.ocupacion_tab.umbrales_ciudades.keys() 
                                if c != "global" and c.strip()])
        
        for ciudad in ciudades_ordenadas:
            self.ocupacion_tab.ciudad_combo.addItem(ciudad)
        
        # Si no hay ciudades, agregar un mensaje
        if self.ocupacion_tab.ciudad_combo.count() == 0:
            self.ocupacion_tab.ciudad_combo.addItem("No hay ciudades configuradas")
            self.ocupacion_tab.ciudad_actual = ""
        else:
            # Seleccionar la primera ciudad por defecto
            self.ocupacion_tab.ciudad_actual = ciudades_ordenadas[0]
            self.ocupacion_tab.cargar_configuracion_ciudad()
        
        self.log_message(f"Lista de ciudades actualizada desde procesamiento: {len(ciudades_normalizadas)} ciudades encontradas", "ocupacion")
    
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
    
    def processing_finished(self, resultado, mode):
        """Manejar la finalización del procesamiento - MODIFICADA"""
        if mode == "procesamiento":
            tab = self.procesamiento_tab
            # Manejo normal para procesamiento (resultado es booleano)
            tab.start_btn.setEnabled(True)
            tab.stop_btn.setEnabled(False)
            
            if resultado:
                tab.status_label.setText("Procesamiento completado con éxito")
                self.log_message("Procesamiento completado exitosamente", mode)
            else:
                tab.status_label.setText("Procesamiento falló")
                self.log_message("Procesamiento falló", mode)
        
        else:  # modo == "ocupacion"
            tab = self.ocupacion_tab
            tab.start_btn.setEnabled(True)
            tab.stop_btn.setEnabled(False)
            
            # resultado ahora es un diccionario
            if isinstance(resultado, dict):
                if resultado.get("error"):
                    # Error en el procesamiento
                    tab.status_label.setText("Procesamiento falló")
                    error_msg = resultado.get("error", "Error desconocido")
                    self.log_message(f"Procesamiento falló: {error_msg}", mode)
                
                elif resultado.get("existen_problemas", False):
                    # Hay frecuencias problemáticas
                    tab.status_label.setText("Procesamiento completado - Verificar advertencias")
                    self.log_message("Procesamiento completado con frecuencias problemáticas", mode)
                    
                    # Mostrar advertencia
                    self.mostrar_advertencia_ocupacion_cero(resultado)
                else:
                    # Procesamiento exitoso sin problemas
                    tab.status_label.setText("Procesamiento completado con éxito")
                    self.log_message("Procesamiento completado exitosamente", mode)
            else:
                # Fallback para compatibilidad
                tab.status_label.setText("Procesamiento completado")
                self.log_message("Procesamiento completado", mode)
    def mostrar_advertencia_ocupacion_cero(self, resultado):
        """Mostrar diálogo de advertencia para frecuencias con ocupación 0%"""
        datos_problematicos = resultado.get("datos", {})
        archivos_generados = resultado.get("archivos_generados", [])
        
        # Preparar datos para el diálogo
        todas_frecuencias = []
        for tipo in ["FM", "TV"]:
            todas_frecuencias.extend(datos_problematicos.get(tipo, []))
        
        if not todas_frecuencias:
            return
        
        # Crear y mostrar diálogo
        dialog = AdvertenciaOcupacionCeroDialog(todas_frecuencias, self)
        resultado_dialogo = dialog.exec_()
        
        if resultado_dialogo == QDialog.Accepted:
            # Usuario eligió FINALIZAR
            self.log_message("Usuario decidió finalizar el procesamiento. Archivos guardados.", "ocupacion")
            QMessageBox.information(self, "Procesamiento Completado", 
                                "El procesamiento se ha completado. Los archivos se han guardado correctamente.")
        else:
            # Usuario eligió CANCELAR - eliminar archivos
            self.log_message("Usuario canceló el procesamiento. Eliminando archivos generados...", "ocupacion")
            archivos_eliminados = self.eliminar_archivos_generados(archivos_generados)
            
            if archivos_eliminados:
                self.log_message(f"Se eliminaron {archivos_eliminados} archivos generados", "ocupacion")
                QMessageBox.information(self, "Procesamiento Cancelado", 
                                    f"Se eliminaron {archivos_eliminados} archivos generados.")
            else:
                self.log_message("No se pudieron eliminar algunos archivos", "ocupacion")
                QMessageBox.warning(self, "Advertencia", 
                                "El procesamiento fue cancelado, pero algunos archivos no pudieron ser eliminados.")
    
    def eliminar_archivos_generados(self, archivos_generados):
        """Eliminar archivos generados durante el procesamiento"""
        eliminados_exitosos = 0
        
        for archivo in archivos_generados:
            try:
                if os.path.exists(archivo):
                    os.remove(archivo)
                    eliminados_exitosos += 1
                    self.log_message(f"Eliminado: {os.path.basename(archivo)}", "ocupacion")
            except Exception as e:
                self.log_message(f"Error eliminando {archivo}: {str(e)}", "ocupacion")
        
        # También intentar eliminar el archivo JSON de verificación
        json_files = ["AdvertenciaOcup.json", "VerificacionOcup.json"]
        for json_file in json_files:
            try:
                if os.path.exists(json_file):
                    os.remove(json_file)
                    self.log_message(f"Archivo {json_file} eliminado", "ocupacion")
            except Exception as e:
                self.log_message(f"Error eliminando {json_file}: {str(e)}", "ocupacion")
        
        return eliminados_exitosos
    
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