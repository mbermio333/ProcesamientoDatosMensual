# -*- coding: utf-8 -*-
import sys
import os
import json
import subprocess
import platform
from PyQt5.QtWidgets import (QApplication, QMainWindow, QWidget, QVBoxLayout, 
                             QHBoxLayout, QPushButton, QLabel, QTextEdit, 
                             QFileDialog, QProgressBar, QMessageBox, QGroupBox)
from PyQt5.QtCore import Qt, QThread, pyqtSignal
from PyQt5.QtGui import QFont

# Añadir constantes para el archivo de configuración
CONFIG_FILE = "config.json"
DEFAULT_PATHS = {
    "fm_path": "MedicionesFmCSV",
    "tv_path": "MedicionesTvCSV", 
    "output_path": "ReportesUnificados"
}

class WorkerThread(QThread):
    """Hilo para ejecutar el procesamiento en segundo plano"""
    progress_signal = pyqtSignal(int)
    log_signal = pyqtSignal(str)
    finished_signal = pyqtSignal(bool)
    
    def __init__(self, fm_path, tv_path, output_path):
        super().__init__()
        self.running = True
        self.fm_path = fm_path
        self.tv_path = tv_path
        self.output_path = output_path
        
    def run(self):
        try:
            # Importar y configurar el módulo principal
            import main
            main.ruta_fm = self.fm_path
            main.ruta_tv = self.tv_path
            main.ruta_salida = self.output_path
            
            # Ejecutar el procesamiento principal
            self.log_signal.emit("Iniciando procesamiento...")
            
            # Llamar a la función principal con nuestros callbacks
            resultado = main.procesar_datos(
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

class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.worker = None
        self.config = self.load_config()
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
            self.log_message(f"Error guardando configuración: {str(e)}")
    
    def truncar_texto(self, texto, max_caracteres=30):
        """Truncar texto largo para mostrar con puntos suspensivos"""
        if len(texto) > max_caracteres:
            return "..." + texto[-max_caracteres:]
        return texto
    
    def initUI(self):
        self.setWindowTitle("Sistema de Reportes Unificados - ARCOTEL")
        self.setGeometry(100, 100, 900, 700)  # Ventana un poco más grande
        
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
        
        # Grupo de configuración
        config_group = QGroupBox("Configuración de Directorios")
        config_group.setStyleSheet("QGroupBox { font-weight: bold; }")
        config_layout = QVBoxLayout()
        
        # Rutas de entrada con mejor formato - usar valores de configuración
        paths = [
            ("Ruta FM:", self.config.get("fm_path", "MedicionesFmCSV"), "fm_path_label."),
            ("Ruta TV:", self.config.get("tv_path", "MedicionesTvCSV"), "tv_path_label"), 
            ("Ruta Salida:", self.config.get("output_path", "ReportesUnificados"), "output_path_label")
        ]
        
        for label_text, default_path, attr_name in paths:
            path_layout = QHBoxLayout()
            
            # Crear etiqueta para el texto descriptivo
            label_desc = QLabel(label_text)
            label_desc.setMinimumWidth(80)  # Ancho fijo para alinear las etiquetas
            
            # Crear etiqueta para la ruta (con texto truncado)
            label_ruta = QLabel(self.truncar_texto(default_path))
            label_ruta.setStyleSheet("background-color: #e8e8e8; padding: 5px; border: 1px solid #ccc;")
            label_ruta.setMinimumWidth(300)  # Un poco más ancho para mostrar más texto
            label_ruta.setToolTip(default_path)  # Tooltip con la ruta completa
            setattr(self, attr_name, label_ruta)
            
            # Crear botón "..." para cambiar ruta
            btn_change = QPushButton("...")
            btn_change.setFixedSize(30, 30)  # Botón pequeño
            btn_change.setStyleSheet("QPushButton { font-weight: bold; }")
            
            # Conectar señal según el tipo de ruta
            if "fm" in attr_name:
                btn_change.clicked.connect(lambda: self.change_path("fm"))
                self.change_fm_btn = btn_change
            elif "tv" in attr_name:
                btn_change.clicked.connect(lambda: self.change_path("tv"))
                self.change_tv_btn = btn_change
            elif "output" in attr_name:
                btn_change.clicked.connect(lambda: self.change_path("output"))
                self.change_output_btn = btn_change
            
            path_layout.addWidget(label_desc)
            path_layout.addWidget(label_ruta)
            path_layout.addWidget(btn_change)
            path_layout.addStretch(1)  # Esto empujará todo a la izquierda
            config_layout.addLayout(path_layout)
        
        config_group.setLayout(config_layout)
        layout.addWidget(config_group)
        
        # Botones de acción (mantener el diseño original)
        action_layout = QHBoxLayout()
        self.start_btn = QPushButton("Iniciar Procesamiento")
        self.start_btn.setStyleSheet("QPushButton { background-color: #4CAF50; color: white; font-weight: bold; padding: 10px; }")
        
        self.stop_btn = QPushButton("Detener")
        self.stop_btn.setStyleSheet("QPushButton { background-color: #f44336; color: white; font-weight: bold; padding: 10px; }")
        self.stop_btn.setEnabled(False)
        
        # Nuevo botón para abrir carpeta de salida
        self.open_output_btn = QPushButton("Abrir Carpeta de Salida")
        self.open_output_btn.setStyleSheet("QPushButton { background-color: #2196F3; color: white; font-weight: bold; padding: 10px; }")
        self.open_output_btn.clicked.connect(self.open_output_folder)
        
        self.start_btn.clicked.connect(self.start_processing)
        self.stop_btn.clicked.connect(self.stop_processing)
        
        action_layout.addWidget(self.start_btn)
        action_layout.addWidget(self.stop_btn)
        action_layout.addWidget(self.open_output_btn)
        layout.addLayout(action_layout)
        
        # Barra de progreso
        self.progress_bar = QProgressBar()
        self.progress_bar.setStyleSheet("QProgressBar { height: 20px; }")
        layout.addWidget(self.progress_bar)
        
        # Área de log
        log_group = QGroupBox("Log de Actividad")
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
        self.log_message("Aplicación iniciada correctamente")
        self.log_message("1. Verifique las rutas de los directorios")
        self.log_message("2. Presione 'Iniciar Procesamiento' para comenzar")
    
    def showEvent(self, event):
        """Se ejecuta cuando la ventana se muestra"""
        super().showEvent(event)
        self.log_message("Ventana principal visible")
    
    def change_path(self, path_type):
        """Cambiar las rutas de los directorios"""
        dialog = QFileDialog()
        dialog.setFileMode(QFileDialog.Directory)
        
        if path_type == "fm":
            current_path = self.fm_path_label.text().replace("...", "")
            new_path = dialog.getExistingDirectory(self, "Seleccionar directorio FM", current_path)
            if new_path:
                truncated_text = self.truncar_texto(new_path)
                self.fm_path_label.setText(truncated_text)
                self.fm_path_label.setToolTip(new_path)  # Tooltip con la ruta completa
                self.config["fm_path"] = new_path
                self.save_config()
                self.log_message(f"Ruta FM cambiada a: {new_path}")
                
        elif path_type == "tv":
            current_path = self.tv_path_label.text().replace("...", "")
            new_path = dialog.getExistingDirectory(self, "Seleccionar directorio TV", current_path)
            if new_path:
                truncated_text = self.truncar_texto(new_path)
                self.tv_path_label.setText(truncated_text)
                self.tv_path_label.setToolTip(new_path)  # Tooltip con la ruta completa
                self.config["tv_path"] = new_path
                self.save_config()
                self.log_message(f"Ruta TV cambiada a: {new_path}")
                
        elif path_type == "output":
            current_path = self.output_path_label.text().replace("...", "")
            new_path = dialog.getExistingDirectory(self, "Seleccionar directorio de salida", current_path)
            if new_path:
                truncated_text = self.truncar_texto(new_path)
                self.output_path_label.setText(truncated_text)
                self.output_path_label.setToolTip(new_path)  # Tooltip con la ruta completa
                self.config["output_path"] = new_path
                self.save_config()
                self.log_message(f"Ruta de salida cambiada a: {new_path}")
    
    def open_output_folder(self):
        """Abrir la carpeta de salida en el explorador de archivos"""
        output_path = self.output_path_label.toolTip()  # Obtener la ruta completa del tooltip
        
        if not output_path:  # Si no hay tooltip, usar el texto (sin los puntos suspensivos)
            output_path = self.output_path_label.text().replace("...", "")
        
        if not os.path.exists(output_path):
            self.log_message(f"La carpeta de salida no existe: {output_path}")
            QMessageBox.warning(self, "Carpeta no encontrada", 
                               f"La carpeta de salida no existe:\n{output_path}")
            return
        
        try:
            # Abrir la carpeta según el sistema operativo
            if platform.system() == "Windows":
                os.startfile(output_path)
            elif platform.system() == "Darwin":  # macOS
                subprocess.Popen(["open", output_path])
            else:  # Linux
                subprocess.Popen(["xdg-open", output_path])
                
            self.log_message(f"Carpeta de salida abierta: {output_path}")
        except Exception as e:
            error_msg = f"No se pudo abrir la carpeta: {str(e)}"
            self.log_message(error_msg)
            QMessageBox.critical(self, "Error", error_msg)
    
    def start_processing(self):
        """Iniciar el procesamiento"""
        # Obtener las rutas completas de los tooltips
        fm_path = self.fm_path_label.toolTip() or self.fm_path_label.text().replace("...", "")
        tv_path = self.tv_path_label.toolTip() or self.tv_path_label.text().replace("...", "")
        output_path = self.output_path_label.toolTip() or self.output_path_label.text().replace("...", "")
        
        # Verificar que las rutas existan
        if not os.path.exists(fm_path):
            self.log_message(f"ERROR: La ruta FM no existe: {fm_path}")
            return
            
        if not os.path.exists(tv_path):
            self.log_message(f"ERROR: La ruta TV no existe: {tv_path}")
            return
        
        # Crear el hilo de trabajo con las rutas configuradas
        self.worker = WorkerThread(fm_path, tv_path, output_path)
        
        self.worker.progress_signal.connect(self.update_progress)
        self.worker.log_signal.connect(self.log_message)
        self.worker.finished_signal.connect(self.processing_finished)
        
        self.start_btn.setEnabled(False)
        self.stop_btn.setEnabled(True)
        self.worker.start()
        
        self.log_message("Procesamiento iniciado...")
    
    def stop_processing(self):
        """Detener el procesamiento"""
        if self.worker and self.worker.isRunning():
            self.worker.stop()
            self.worker.wait()
            self.log_message("Procesamiento detenido por el usuario")
        
        self.start_btn.setEnabled(True)
        self.stop_btn.setEnabled(False)
    
    def update_progress(self, value):
        """Actualizar la barra de progreso"""
        self.progress_bar.setValue(value)
        self.status_label.setText(f"Procesando... {value}%")
    
    def log_message(self, message):
        """Añadir mensaje al log"""
        from datetime import datetime
        timestamp = datetime.now().strftime("%H:%M:%S")
        self.log_text.append(f"[{timestamp}] {message}")
        # Auto-scroll to bottom
        scrollbar = self.log_text.verticalScrollBar()
        scrollbar.setValue(scrollbar.maximum())
    
    def processing_finished(self, success):
        """Procesamiento completado"""
        self.start_btn.setEnabled(True)
        self.stop_btn.setEnabled(False)
        
        if success:
            self.log_message("Procesamiento completado con éxito!")
            self.status_label.setText("Completado con éxito")
            QMessageBox.information(self, "Éxito", "El procesamiento se completó correctamente.")
        else:
            self.log_message("Error en el procesamiento.")
            self.status_label.setText("Error en el procesamiento")
            QMessageBox.warning(self, "Error", "Ocurrió un error durante el procesamiento.")
    
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