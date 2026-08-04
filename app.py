import sys
import os
import pandas as pd
from datetime import datetime
from PyQt5.QtWidgets import (QApplication, QMainWindow, QWidget, QVBoxLayout, 
                             QPushButton, QLabel, QFileDialog, QMessageBox, QTableWidget, QTableWidgetItem)
from PyQt5.QtCore import Qt

class ExcelBusinessEngine:
    def __init__(self, file_path):
        self.file_path = file_path
        self.df = None

    def cargar_datos(self, sheet_name=0):
        if self.file_path.endswith('.csv'):
            self.df = pd.read_csv(self.file_path)
        else:
            self.df = pd.read_excel(self.file_path, sheet_name=sheet_name)
        return self.df

    def limpiar_archivos_basura(self):
        if self.df is not None:
            self.df.dropna(how='all', inplace=True)
            self.df.dropna(axis=1, how='all', inplace=True)
            for col in self.df.select_dtypes(include=['object']).columns:
                self.df[col] = self.df[col].str.strip()
            self.df.drop_duplicates(inplace=True)
            self.df.reset_index(drop=True, inplace=True)
        return self.df

    def ordenar_datos(self, criterio='cronologico', columna_fecha_o_texto='Fecha'):
        if self.df is not None:
            if criterio == 'cronologico' and columna_fecha_o_texto in self.df.columns:
                self.df[columna_fecha_o_texto] = pd.to_datetime(self.df[columna_fecha_o_texto], errors='coerce')
                self.df.sort_values(by=columna_fecha_o_texto, inplace=True)
            elif criterio == 'alfabetico' and columna_fecha_o_texto in self.df.columns:
                self.df.sort_values(by=columna_fecha_o_texto, ascending=True, inplace=True)
            self.df.reset_index(drop=True, inplace=True)
        return self.df

    def calcular_impuestos_itbis(self, columna_precio, tasa_itbis=0.18):
        if self.df is not None and columna_precio in self.df.columns:
            self.df['ITBIS'] = self.df[columna_precio] * tasa_itbis
            self.df['Total_Con_ITBIS'] = self.df[columna_precio] + self.df['ITBIS']
        return self.df

    def guardar_resultado(self, output_path):
        if self.df is not None:
            self.df.to_excel(output_path, index=False)
            return True
        return False

class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.path_archivo = None
        self.initUI()

    def initUI(self):
        self.setWindowTitle("Bot Multifacético de Excel - Enterprise Edition")
        self.setGeometry(100, 100, 900, 600)

        central_widget = QWidget()
        self.setCentralWidget(central_widget)
        layout = QVBoxLayout()

        self.label_title = QLabel("🤖 Automatizador Inteligente y Camaleón de Excel")
        self.label_title.setStyleSheet("font-size: 16px; font-weight: bold; color: #2C3E50;")
        layout.addWidget(self.label_title)

        self.btn_cargar = QPushButton("📂 Cargar Archivo de Excel / Inventario")
        self.btn_cargar.setStyleSheet("background-color: #3498DB; color: white; padding: 10px; font-size: 14px;")
        self.btn_cargar.clicked.connect(self.cargar_archivo)
        layout.addWidget(self.btn_cargar)

        self.table_view = QTableWidget()
        layout.addWidget(self.table_view)

        self.btn_procesar = QPushButton("⚡ Ejecutar Limpieza, ITBIS y Ordenamiento")
        self.btn_procesar.setStyleSheet("background-color: #2ECC71; color: white; padding: 10px; font-size: 14px;")
        self.btn_procesar.clicked.connect(self.procesar_datos_excel)
        layout.addWidget(self.btn_procesar)

        central_widget.setLayout(layout)

    def cargar_archivo(self):
        options = QFileDialog.Options()
        file_name, _ = QFileDialog.getOpenFileName(self, "Seleccionar Archivo Excel", "", "Archivos Excel (*.xlsx *.xls *.csv)", options=options)
        if file_name:
            self.path_archivo = file_name
            QMessageBox.information(self, "Éxito", f"Archivo cargado correctamente: {file_name.split('/')[-1]}")

    def procesar_datos_excel(self):
        if not hasattr(self, 'path_archivo') or not self.path_archivo:
            QMessageBox.warning(self, "Advertencia", "Por favor carga un archivo primero.")
            return
        
        engine = ExcelBusinessEngine(self.path_archivo)
        engine.cargar_datos()
        engine.limpiar_archivos_basura()
        engine.ordenar_datos()
        
        # Guardar automáticamente un archivo procesado de salida
        output_file = "resultado_procesado.xlsx"
        engine.guardar_resultado(output_file)
        
        QMessageBox.information(self, "Proceso", f"¡Proceso exitoso! Archivo guardado como {output_file}")

if __name__ == '__main__':
    app = QApplication(sys.argv)
    ex = MainWindow()
    ex.show()
    sys.exit(app.exec_())
