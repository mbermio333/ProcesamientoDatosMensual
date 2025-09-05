# -*- coding: utf-8 -*-
import sys
import os
import json
import subprocess
import platform
from PyQt5.QtWidgets import (QApplication, QMainWindow, QWidget, QVBoxLayout, 
                             QHBoxLayout, QPushButton, QLabel, QTextEdit, 
                             QFileDialog, QProgressBar, QMessageBox, QGroupBox,
                             QTabWidget)
from PyQt5.QtCore import Qt, QThread, pyqtSignal
from PyQt5.QtGui import QFont

# Añadir constantes para el archivo de configuración
CONFIG_FILE = "config.json"
DEFAULT_PATHS = {
    "fm_path": "MedicionesFmCSV",
    "tv_path": "MedicionesTvCSV", 
    "output_path": "ReportesUnificados",
    "ocupacion_output_path": "ReportesOcupacion"  # Nueva ruta para ocupación
}

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
        self.initUI()
        
    def initUI(self):
        layout = QVBoxLayout(self)
        
        # Grupo de configuración
        config_group = QGroupBox("Configuración de Directorios - Procesamiento")
        config_group.setStyleSheet("QGroupBox { font-weight: bold; }")
        config_layout = QVBoxLayout()
        
        # Rutas de entrada con mejor formato - usar valores de configuración
        paths = [
            ("Ruta FM:", self.parent.config.get("fm_path", "MedicionesFmCSV"), "fm_path_label"),
            ("Ruta TV:", self.parent.config.get("tv_path", "MedicionesTvCSV"), "tv_path_label"), 
            ("Ruta Salida:", self.parent.config.get("output_path", "ReportesUnificados"), "output_path_label")
        ]
        
        for label_text, default_path, attr_name in paths:
            path_layout = QHBoxLayout()
            
            # Crear etiqueta para el texto descriptivo
            label_desc = QLabel(label_text)
            label_desc.setMinimumWidth(80)
            
            # Crear etiqueta para la ruta (con texto truncado)
            label_ruta = QLabel(self.parent.truncar_texto(default_path))
            label_ruta.setStyleSheet("background-color: #e8e8e8; padding: 5px; border: 1px solid #ccc;")
            label_ruta.setMinimumWidth(300)
            label_ruta.setToolTip(default_path)
            setattr(self, attr_name, label_ruta)
            
            # Crear botón "..." para cambiar ruta
            btn_change = QPushButton("...")
            btn_change.setFixedSize(30, 30)
            btn_change.setStyleSheet("QPushButton { font-weight: bold; }")
            
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
        self.start_btn = QPushButton("Iniciar Procesamiento")
        self.start_btn.setStyleSheet("QPushButton { background-color: #4CAF50; color: white; font-weight: bold; padding: 10px; }")
        
        self.stop_btn = QPushButton("Detener")
        self.stop_btn.setStyleSheet("QPushButton { background-color: #f44336; color: white; font-weight: bold; padding: 10px; }")
        self.stop_btn.setEnabled(False)
        
        # Nuevo botón para abrir carpeta de salida
        self.open_output_btn = QPushButton("Abrir Carpeta de Salida")
        self.open_output_btn.setStyleSheet("QPushButton { background-color: #2196F3; color: white; font-weight: bold; padding: 10px; }")
        self.open_output_btn.clicked.connect(lambda: self.parent.open_output_folder("procesamiento"))
        
        self.start_btn.clicked.connect(lambda: self.parent.start_processing("procesamiento"))
        self.stop_btn.clicked.connect(self.parent.stop_processing)
        
        action_layout.addWidget(self.start_btn)
        action_layout.addWidget(self.stop_btn)
        action_layout.addWidget(self.open_output_btn)
        layout.addLayout(action_layout)
        
        # Barra de progreso
        self.progress_bar = QProgressBar()
        self.progress_bar.setStyleSheet("QProgressBar { height: 20px; }")
        layout.addWidget(self.progress_bar)
        
        # Área de log
        log_group = QGroupBox("Log de Actividad - Procesamiento")
        log_group.setStyleSheet("QGroupBox { font-weight: bold; }")
        log_layout = QVBoxLayout()
        self.log_text = QTextEdit()
        self.log_text.setReadOnly(True)
        self.log_text.setStyleSheet("font-family: monospace; font-size: 10pt;")
        log_layout.addWidget(self.log_text)
        log_group.setLayout(log_layout)
        layout.addWidget(log_group)
        
        # Estado
        self.status_label = QLabel("Listo para iniciar")
        self.status_label.setStyleSheet("background-color: #e0e0e0; padding: 5px; border: 1px solid #ccc;")
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
        self.initUI()
        
    def initUI(self):
        layout = QVBoxLayout(self)
        
        # Grupo de configuración
        config_group = QGroupBox("Configuración de Directorios - Ocupación")
        config_group.setStyleSheet("QGroupBox { font-weight: bold; }")
        config_layout = QVBoxLayout()
        
        # Rutas de entrada con mejor formato - usar valores de configuración
        paths = [
            ("Ruta FM:", self.parent.config.get("fm_path", "MedicionesFmCSV"), "fm_path_label_ocup"),
            ("Ruta TV:", self.parent.config.get("tv_path", "MedicionesTvCSV"), "tv_path_label_ocup"), 
            ("Ruta Salida:", self.parent.config.get("ocupacion_output_path", "ReportesOcupacion"), "output_path_label_ocup")
        ]
        
        for label_text, default_path, attr_name in paths:
            path_layout = QHBoxLayout()
            
            # Crear etiqueta para el texto descriptivo
            label_desc = QLabel(label_text)
            label_desc.setMinimumWidth(80)
            
            # Crear etiqueta para la ruta (con texto truncado)
            label_ruta = QLabel(self.parent.truncar_texto(default_path))
            label_ruta.setStyleSheet("background-color: #e8e8e8; padding: 5px; border: 1px solid #ccc;")
            label_ruta.setMinimumWidth(300)
            label_ruta.setToolTip(default_path)
            setattr(self, attr_name, label_ruta)
            
            # Crear botón "..." para cambiar ruta
            btn_change = QPushButton("...")
            btn_change.setFixedSize(30, 30)
            btn_change.setStyleSheet("QPushButton { font-weight: bold; }")
            
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
        layout.addWidget(config_group)
        
        # Botones de acción
        action_layout = QHBoxLayout()
        self.start_btn = QPushButton("Iniciar Análisis de Ocupación")
        self.start_btn.setStyleSheet("QPushButton { background-color: #4CAF50; color: white; font-weight: bold; padding: 10px; }")
        
        self.stop_btn = QPushButton("Detener")
        self.stop_btn.setStyleSheet("QPushButton { background-color: #f44336; color: white; font-weight: bold; padding: 10px; }")
        self.stop_btn.setEnabled(False)
        
        # Nuevo botón para abrir carpeta de salida
        self.open_output_btn = QPushButton("Abrir Carpeta de Salida")
        self.open_output_btn.setStyleSheet("QPushButton { background-color: #2196F3; color: white; font-weight: bold; padding: 10px; }")
        self.open_output_btn.clicked.connect(lambda: self.parent.open_output_folder("ocupacion"))
        
        self.start_btn.clicked.connect(lambda: self.parent.start_processing("ocupacion"))
        self.stop_btn.clicked.connect(self.parent.stop_processing)
        
        action_layout.addWidget(self.start_btn)
        action_layout.addWidget(self.stop_btn)
        action_layout.addWidget(self.open_output_btn)
        layout.addLayout(action_layout)
        
        # Barra de progreso
        self.progress_bar = QProgressBar()
        self.progress_bar.setStyleSheet("QProgressBar { height: 20px; }")
        layout.addWidget(self.progress_bar)
        
        # Área de log
        log_group = QGroupBox("Log de Actividad - Ocupación")
        log_group.setStyleSheet("QGroupBox { font-weight: bold; }")
        log_layout = QVBoxLayout()
        self.log_text = QTextEdit()
        self.log_text.setReadOnly(True)
        self.log_text.setStyleSheet("font-family: monospace; font-size: 10pt;")
        log_layout.addWidget(self.log_text)
        log_group.setLayout(log_layout)
        layout.addWidget(log_group)
        
        # Estado
        self.status_label = QLabel("Listo para iniciar")
        self.status_label.setStyleSheet("background-color: #e0e0e0; padding: 5px; border: 1px solid #ccc;")
        layout.addWidget(self.status_label)
        
        # Mensaje inicial
        self.log_text.append("Aplicación iniciada correctamente")
        self.log_text.append("1. Verifique las rutas de los directorios")
        self.log_text.append("2. Presione 'Iniciar Análisis de Ocupación' para comenzar")

class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.worker = None
        self.config = self.load_config()
        self.current_tab = None
        self.initUI()
        
    def load_config(self):
        """Cargar configuración desde archivo JSON"""
        if os.path.exists(CONFIG_FILE):
            try:
                with open(CONFIG_FILE, 'r') as f:
                    return json.load(f)
            except:
                return DEFAULT_PATHS.copy()
        return DEFAULT_PATHS.copy()
    
    def save_config(self):
        """Guardar configuración en archivo JSON"""
        try:
            with open(CONFIG_FILE, 'w') as f:
                json.dump(self.config, f)
        except Exception as e:
            self.log_message(f"Error guardando configuración: {str(e)}", "procesamiento")
    
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
                self.config["fm_path"] = new_path
            elif path_type == "tv":
                tab.tv_path_label.setText(truncated_text)
                tab.tv_path_label.setToolTip(new_path)
                self.config["tv_path"] = new_path
            elif path_type == "output":
                tab.output_path_label.setText(truncated_text)
                tab.output_path_label.setToolTip(new_path)
                self.config["output_path"] = new_path
        else:  # ocupacion
            if path_type == "fm":
                tab.fm_path_label_ocup.setText(truncated_text)
                tab.fm_path_label_ocup.setToolTip(new_path)
                self.config["fm_path"] = new_path
            elif path_type == "tv":
                tab.tv_path_label_ocup.setText(truncated_text)
                tab.tv_path_label_ocup.setToolTip(new_path)
                self.config["tv_path"] = new_path
            elif path_type == "ocupacion_output":
                tab.output_path_label_ocup.setText(truncated_text)
                tab.output_path_label_ocup.setToolTip(new_path)
                self.config["ocupacion_output_path"] = new_path
        
        self.save_config()
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