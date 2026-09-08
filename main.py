import sys, os
def resource_path(relative_path):
    if hasattr(sys, "_MEIPASS"):
        base_path = sys._MEIPASS
    else:
        base_path = os.path.dirname(os.path.abspath(__file__))
    return os.path.join(base_path, relative_path)
import numpy as np
import menu_actions as ma
from openpyxl import Workbook
from PySide6.QtGui import QPixmap,QFont
from PySide6.QtCore import Qt
from PySide6.QtWidgets import (
    QApplication, QMainWindow, QWidget, QFileDialog,  
    QLabel, QLineEdit, QPushButton,
    QVBoxLayout, QHBoxLayout, QCheckBox,
    QTabWidget, QGroupBox, QComboBox, QMessageBox, QGridLayout, QFrame,QSlider,
)
from estimacion_pv import estimar_parametros_pv
from matplotlib.backends.backend_qtagg import FigureCanvasQTAgg as FigureCanvas
from matplotlib.figure import Figure

from mpl_toolkits.axes_grid1.inset_locator import zoomed_inset_axes, mark_inset


class VentanaPrincipal(QMainWindow):
    def __init__(self):
        super().__init__()

        self.setWindowTitle("Interfaz PV - Simulador de Sombreado")

        # ── Adaptar tamaño a la pantalla disponible ──────────────────────
        pantalla = QApplication.primaryScreen().availableGeometry()
        ancho_pantalla = pantalla.width()   # 1366
        alto_pantalla  = pantalla.height()  # 768

        # Escala basada en tu resolución original de diseño (1200x800)
        escala = min(ancho_pantalla / 1200, alto_pantalla / 800)

        fuente_label  = max(9,  int(13 * escala))
        fuente_linea  = max(8,  int(10 * escala))
        fuente_boton  = max(9,  int(12 * escala))
        fuente_grupo  = max(9,  int(10 * escala))
        fuente_titulo = max(11, int(16 * escala))
        fuente_titulo_seccion = max(28, int(32 * escala))

        self.setMinimumSize(900, 600)
        self.showMaximized()

        self.datos_paneles = {
            "Panel Multi-Crystalline MSX60": {
                "Isc (A):": 3.8,
                "Voc (V):": 21.1,
                "Vmp (V):": 17.1,
                "Imp (A):": 3.5,
                "Kv (V/°C):": -0.08,
                "Ki (A/°C):": 0.003,
                "T (°C):": 25,
                "Número de celdas :": 36,
                "Número de diodos bypass :": 2,
                "filas": 9,
                "columnas": 4, 
            },
            "Panel Mono-Crystalline SM55": {
                "Isc (A):": 3.45,
                "Voc (V):": 21.7,
                "Vmp (V):": 17.4,
                "Imp (A):": 3.15,
                "Kv (V/°C):": -0.075,
                "Ki (A/°C):": 0.0012,
                "T (°C):": 25,
                "Número de celdas :": 36,
                "Número de diodos bypass :": 3,
                "filas": 12,
                "columnas": 3, 
            },
            "Panel Thin Film ST40": {
                "Isc (A):": 2.68,
                "Voc (V):": 23.3,
                "Vmp (V):": 16.6,
                "Imp (A):": 2.41,
                "Kv (V/°C):": -0.100,
                "Ki (A/°C):":  0.00035,
                "T (°C):": 25,
                "Número de celdas :": 36,
                "Número de diodos bypass :": 1,
                "filas": 9,
                "columnas": 4, 
            },
            "Panel KGC200GT": {
                "Isc (A):": 8.21,
                "Voc (V):": 32.9,
                "Vmp (V):": 26.3,
                "Imp (A):": 7.61,
                "Kv (V/°C):": -0.124,
                "Ki (A/°C):": 0.00318,
                "T (°C):": 25,
                "Número de celdas :": 54,
                "Número de diodos bypass :": 1,
                "filas": 9,
                "columnas": 6, 
            },
            "Panel STP280": {
                "Isc (A):": 8.33,
                "Voc (V):": 44.8,
                "Vmp (V):": 35.2,
                "Imp (A):": 7.95,
                "Kv (V/°C):": -0.118,
                "Ki (A/°C):": 0.0047,
                "T (°C):": 25,
                "Número de celdas :": 72,
                "Número de diodos bypass :": 1,
                "filas": 6,
                "columnas": 12, 
            },
            "Panel SPR-315": {
                "Isc (A):": 6.14,
                "Voc (V):": 64.6,
                "Vmp (V):": 54.7,
                "Imp (A):": 5.76,
                "Kv (V/°C):": -0.123,
                "Ki (A/°C):": 0.00483,
                "T (°C):": 25,
                "Número de celdas :": 96,
                "Número de diodos bypass :": 3,
                "filas": 8,
                "columnas": 12, 
            }
            ,
             "Panel I-110/24": {
                "Isc (A):": 3.38,
                "Voc (V):": 43.2,
                "Vmp (V):": 34.8,
                "Imp (A):": 3.16,
                "Kv (V/°C):": -0.078,
                "Ki (A/°C):": 0.00429,
                "T (°C):": 25,
                "Número de celdas :": 72,
                "Número de diodos bypass :": 1,
                "filas": 12,
                "columnas": 6, 
            }
        }

        self.a1 = 1.0
        self.a2 = 2.0

        self.valores_default_rbd = {
            "aRBD": 0.0001036,
            "VRBD": -5.52726,
            "nRBD": 3.284628
        }

        self.filas_panel = 0
        self.columnas_panel = 0

        self.V_actual = None
        self.I_actual = None
        self.P_actual = None
        self.Vmpp_actual = None
        self.Impp_actual = None
        self.Pmpp_actual = None

        # Reemplaza la línea del scaled fijo
        ancho_img = int(ancho_pantalla * 0.25)   # 30% del ancho de pantalla
        alto_img  = int(alto_pantalla  * 0.20)   # 25% del alto de pantalla

        self.imagen_circuito = QLabel(self)
        pixmap = QPixmap(resource_path("assets/modelo_pv.png"))
        self.imagen_circuito.setPixmap(
            pixmap.scaled(ancho_img, alto_img, Qt.KeepAspectRatio, Qt.SmoothTransformation)
        )
        self.imagen_circuito.setAlignment(Qt.AlignCenter)

        # Variables para la matriz de irradiancia
        self.irradiance_matrix = None
        self.cell_labels = {}
        self.irradiance_step = 50
        self.max_irradiance = 1000.0
        self.min_irradiance = 0.0
        self.modelo_guardado = False

        self.temp_default_c = 25
        self.Tcell_K = self.temp_default_c + 273.15

        self.setStyleSheet(f"""
            QLabel {{ font-size: {fuente_label}px; }}
        
            QLineEdit {{
                font-size: {fuente_linea}px;
                padding: 3px;
                border: 1px solid #9c9c9c;
                border-radius: 4px;
            }}
        
            QPushButton {{
                font-size: {fuente_boton}px;
                padding: 4px 10px;
            }}
        
            QGroupBox {{
                font-size: {fuente_grupo}px;
                font-weight: bold;
                border: 1px solid #7a7a7a;
                border-radius: 10px;
                margin-top: 10px;
                background-color: #d5dded;
            }}
        
            QGroupBox::title {{
                font-size: {fuente_titulo}px;
                font-weight: bold;
                subcontrol-origin: margin;
                subcontrol-position: top center;
                padding: 0 20px;
            }}
            
            QGroupBox#seccionPrincipal::title {{
                color: red;
                padding-top: 30px;
                padding-left: 100px;
            }}
        """)

        # ================= MENÚ SUPERIOR =================
        menubar = self.menuBar()
        menu_archivo = menubar.addMenu("Menú")

        menu_archivo.addAction("Descargar parámetros panel (Excel)").triggered.connect(lambda: ma.descargar_parametros_panel(self))
        menu_archivo.addAction("Descargar parámetros estimados (Excel)").triggered.connect(lambda: ma.descargar_parametros_xlsx(self))
        menu_archivo.addAction("Guardar sombreado (Excel)").triggered.connect(lambda: ma.guardar_sombreado_excel(self))
        menu_archivo.addAction("Descargar curvas sombreado (PNG)").triggered.connect(lambda: ma.descargar_sombreado_png(self))

        menu_archivo.addSeparator()
        menu_archivo.addAction("Descargar todo").triggered.connect(lambda: ma.descargar_todo(self))
        menu_archivo.addSeparator()
        menu_archivo.addAction("Acerca de").triggered.connect(lambda: ma.mostrar_acerca_de(self))

        # ================= PESTAÑAS =================
        self.tabs = QTabWidget()
        self.tabs.addTab(self.tab_configuracion(), "Configuración")
        self.tabs.addTab(self.tab_sombreado(), "Sombreado")

        self.tabs.currentChanged.connect(self.validar_cambio_pestana)

        self.setCentralWidget(self.tabs)

    def tab_configuracion(self):
        widget = QWidget()
        layout_principal = QHBoxLayout()

        # ---------- PANEL IZQUIERDO ----------
        panel_izquierdo = QGroupBox()
        layout_izq = QVBoxLayout()
        panel_izquierdo.setLayout(layout_izq)

        modo_layout = QHBoxLayout()

        self.combo_modo = QComboBox()
        self.combo_modo.addItems([
            "Ingreso manual",
            "Seleccionar panel"
        ])
        self.combo_modo.currentTextChanged.connect(self.cambiar_modo)

        label_modo = QLabel("Modo de ingreso:")
        label_modo.setStyleSheet("font-weight: bold;")
        modo_layout.addWidget(label_modo)
        modo_layout.addWidget(self.combo_modo)

        layout_izq.addLayout(modo_layout)

        self.combo_paneles = QComboBox()
        self.combo_paneles.addItems(
            ["Seleccione un panel"] + list(self.datos_paneles.keys())
        )
        self.combo_paneles.currentTextChanged.connect(self.cargar_panel)
        self.combo_paneles.hide()

        layout_izq.addWidget(self.combo_paneles)

        titulo = QLabel("Parámetros del módulo fotovoltaico")
        titulo.setStyleSheet("font-weight: bold; font-size: 16px;")
        layout_izq.addWidget(titulo)

        self.inputs = {}

        parametros = [
            ("Isc (A):", "Corriente de cortocircuito (A)"),
            ("Voc (V):", "Voltaje en circuito abierto (V)"),
            ("Vmp (V):", "Voltaje en máxima potencia (V)"),
            ("Imp (A):", "Corriente en máxima potencia (A)"),
            ("Kv (V/°C):", "Coeficiente de voltaje"),
            ("Ki (A/°C):", "Coeficiente de corriente"),
            ("T (°C):", "Temperatura"),
            ("Número de celdas :", "Número de celdas del panel"),
            ("Número de diodos bypass :", "Diodos bypass")
        ]

        for nombre, ayuda in parametros:
            fila = QHBoxLayout()
            label = QLabel(nombre)
            textbox = QLineEdit()
            textbox.setMaximumWidth(120)
            textbox.setToolTip(ayuda)

            fila.addWidget(label)
            fila.addWidget(textbox)

            layout_izq.addLayout(fila)
            self.inputs[nombre] = textbox

        # a1
        fila = QHBoxLayout()
        label_a1 = QLabel("Factor de idealidad diodo 1 (a1):")
        valor_a1 = QLabel("1.0")

        valor_a1.setFixedWidth(120)   # para que quede alineado con los inputs

        fila.addWidget(label_a1)
        fila.addWidget(valor_a1)

        layout_izq.addLayout(fila)

        # a2
        fila = QHBoxLayout()
        label_a2 = QLabel("Factor de idealidad diodo 2 (a2):")
        valor_a2 = QLabel("2.0")

        valor_a2.setFixedWidth(120)

        fila.addWidget(label_a2)
        fila.addWidget(valor_a2)

        layout_izq.addLayout(fila)

        # ---------- BOTÓN ----------
        boton = QPushButton("Estimar parámetros")
        boton.clicked.connect(self.estimar_parametros)
        boton.setStyleSheet("""
            QPushButton {
                background-color:#919499;
                color: white;
                font-weight: bold;
                font-size: 14px;
                padding: 10px;
                border-radius: 8px;
            }
            QPushButton:hover {
                background-color: #050505;
            }
            QPushButton:pressed {
                background-color: #919499;
            }
        """)
        layout_izq.addWidget(boton)
        #layout_izq.addSpacing(10)

        titulo_imagen = QLabel("Modelo circuital de dos diodos")
        titulo_imagen.setAlignment(Qt.AlignCenter)
        titulo_imagen.setStyleSheet("font-weight: bold; font-size: 14px;")
        layout_izq.addWidget(titulo_imagen)

        #layout_izq.addSpacing(10)
        layout_izq.addWidget(self.imagen_circuito)
        self.imagen_circuito.setAlignment(Qt.AlignCenter)
        self.imagen_circuito.show()
        layout_izq.addStretch()

        panel_izquierdo.setLayout(layout_izq)

        # ---------- PANEL DERECHO ----------
        panel_derecho = QGroupBox()
        layout_der = QVBoxLayout()
        panel_derecho.setLayout(layout_der)

        self.resultado_widget = QWidget()
        self.resultado_layout = QVBoxLayout(self.resultado_widget)

        self.resultado_widget.hide()
        layout_der.addWidget(self.resultado_widget)
        layout_der.addStretch()

        # ---------- UNIR PANELES ----------
        layout_principal.addWidget(panel_izquierdo, 1)
        layout_principal.addWidget(panel_derecho, 1)

        widget.setLayout(layout_principal)
        return widget
    
    def cambiar_modo(self, modo):
        es_manual = modo == "Ingreso manual"

        self.combo_paneles.setVisible(not es_manual)

        for campo in self.inputs.values():
            campo.setReadOnly(not es_manual)
            if es_manual:
                campo.clear()

    def cargar_panel(self, nombre_panel):
        if nombre_panel not in self.datos_paneles:
            return

        datos = self.datos_paneles[nombre_panel]

        for k, v in datos.items():
            if k in self.inputs:
                self.inputs[k].setText(str(v))
                self.inputs[k].setReadOnly(True)

        # guardar internos
        self.filas_panel = datos["filas"]
        self.columnas_panel = datos["columnas"]

    def estimar_parametros(self):
        for nombre, campo in self.inputs.items():
            if campo.text().strip() == "":
                QMessageBox.warning(
                    self,
                    "Campos incompletos",
                    f"Por favor ingrese el valor de:\n{nombre}"
                )
                return
        
        try:
            self.Ns = int(self.inputs["Número de celdas :"].text())
            self.Nd = int(self.inputs["Número de diodos bypass :"].text())

            if self.Ns <= 0 or self.Nd <= 0:
                raise ValueError

        except ValueError:
            QMessageBox.warning(
                self,
                "Error",
                "Ingrese valores válidos para número de celdas y diodos."
            )
            return   
            
        Voc = float(self.inputs["Voc (V):"].text())
        Isc = float(self.inputs["Isc (A):"].text())
        Vmp = float(self.inputs["Vmp (V):"].text())
        Imp = float(self.inputs["Imp (A):"].text())
        Kv = float(self.inputs["Kv (V/°C):"].text())
        Ki = float(self.inputs["Ki (A/°C):"].text())
        Ns = int(self.inputs["Número de celdas :"].text())
        a1 = float(self.a1)
        a2 = float(self.a2)
        T = int(self.inputs["T (°C):"].text())

        resultados, datos_grafica = estimar_parametros_pv(Voc, Isc, Vmp, Imp, Kv, Ki, Ns, T, a1, a2)

        self.resultados = resultados  # guardar referencia global
        self.mostrar_resultados_editables(resultados)

        self.graficar_curvas(
        datos_grafica["V_2"],
        datos_grafica["I_final"],
        resultados["Voltaje (Vmpp)"],
        resultados["Corriente (Impp)"],
        datos_grafica["Vmp"],
        datos_grafica["Imp"],

        )

    def mostrar_resultados_editables(self, resultados):
        # ── Limpiar resultados anteriores correctamente ──────────────
        def limpiar_layout(layout):
            while layout.count():
                item = layout.takeAt(0)
                widget = item.widget()
                if widget:
                    widget.deleteLater()
                elif item.layout():
                    limpiar_layout(item.layout())

        limpiar_layout(self.resultado_layout)
        # ─────────────────────────────────────────────────────────────

        # ── Limpiar matriz de sombreado ──────────────────────────────
        if hasattr(self, 'grid_layout_celdas'):
            while self.grid_layout_celdas.count():
                item = self.grid_layout_celdas.takeAt(0)
                widget = item.widget()
                if widget:
                    widget.deleteLater()
            self.cell_labels = {}
            self.irradiance_matrix = None
        # ─────────────────────────────────────────────────────────────

        self.inputs_resultados = {}

        parametros_circuitales = {"Corriente de saturación diodo 1 (Io1)", "Corriente de saturación diodo 2 (Io2)", "Fotocorriente (Iph)", "Resistencia serie (Rs)", "Resistencia paralelo (Rp)"}
        

        titulo1 = QLabel("PARÁMETROS CIRCUITALES")
        titulo1.setStyleSheet("font-weight: bold; font-size: 14px;")
        self.resultado_layout.addWidget(titulo1)

        for k, v in resultados.items():
            if k in parametros_circuitales:
                fila = QHBoxLayout()
                label = QLabel(k)
                edit = QLineEdit(f"{v:.4e}")
                edit.setMaximumWidth(150)

                fila.addWidget(label)
                fila.addWidget(edit)

                self.resultado_layout.addLayout(fila)
                self.inputs_resultados[k] = edit
                # ===== CAMPOS MANUALES EXTRA =====

        campos_manuales = {
            "Coeficiente de ruptura inversa (aRBD)": "aRBD",
            "Tensión de ruptura inversa (VRBD)": "VRBD",
            "Exponente de ruptura inversa (nRBD)": "nRBD"
        }

        self.inputs_manuales = {}

        self.checkbox_editar_rbd = QCheckBox("Modificar parámetros de ruptura inversa")
        self.checkbox_editar_rbd.toggled.connect(self.toggle_parametros_rbd)

        # centrar checkbox
        fila_check = QHBoxLayout()
        fila_check.addStretch()
        fila_check.addWidget(self.checkbox_editar_rbd)
        fila_check.addStretch()

        self.resultado_layout.addLayout(fila_check)

        self.inputs_manuales = {}

        for texto_label, nombre_variable in campos_manuales.items():
        
            fila = QHBoxLayout()

            label = QLabel(texto_label)

            edit = QLineEdit()
            edit.setMaximumWidth(150)
            edit.setText(str(self.valores_default_rbd[nombre_variable]))
            edit.setEnabled(False)

            fila.addWidget(label)
            fila.addWidget(edit)

            self.resultado_layout.addLayout(fila)

            self.inputs_manuales[nombre_variable] = edit

        # TABLA DE RANGOS - rediseño con QGridLayout
        contenedor_tabla = QWidget()
        contenedor_tabla.setStyleSheet("""
            QWidget#contenedor_rangos {
                background-color: #3a6ea8;
                border-radius: 10px;
            }
        """)
        contenedor_tabla.setObjectName("contenedor_rangos")

        grid = QGridLayout(contenedor_tabla)
        grid.setSpacing(2)
        grid.setContentsMargins(8, 8, 8, 8)

        # Encabezados
        encabezados = ["Parámetro", "Rango recomendado"]
        for col, texto in enumerate(encabezados):
            lbl = QLabel(texto)
            lbl.setFont(QFont("Arial", 6, QFont.Bold))
            lbl.setStyleSheet("""
                QLabel {
                    background-color: #1e3f6e;
                    color: white;
                    padding: 8px 16px;
                    border-radius: 4px;
                    font-size: 13px;
                    font-weight: bold;
                }
            """)
            lbl.setAlignment(Qt.AlignCenter)
            lbl.setMinimumHeight(28)
            lbl.setMinimumWidth(160)
            grid.addWidget(lbl, 0, col)

        # Datos
        datos = [
            ("aRBD", "2e-3 → 3.5e-1"),
            ("VRBD", "-21.29 V → -10 V"),
            ("nRBD", "1 → 6")
        ]

        colores_fila = ["#5b8ec7", "#4a7ab5"]

        for fila, (p, r) in enumerate(datos):
            color = colores_fila[fila % 2]
            for col, texto in enumerate([p, r]):
                lbl = QLabel(texto)
                lbl.setFont(QFont("Arial", 6))
                lbl.setStyleSheet(f"""
                    QLabel {{
                        background-color: {color};
                        color: white;
                        padding: 6px 16px;
                        border-radius: 3px;
                        font-size: 12px;
                    }}
                """)
                lbl.setAlignment(Qt.AlignCenter)
                lbl.setMinimumHeight(28)
                lbl.setMinimumWidth(160)
                grid.addWidget(lbl, fila + 1, col)

        self.resultado_layout.addWidget(contenedor_tabla)
        self.tabla_rangos_widget = contenedor_tabla
        contenedor_tabla.hide()

        titulo2 = QLabel("VARIABLES DE MAXIMA POTENCIA")
        titulo2.setStyleSheet("font-weight: bold; font-size: 14px; margin-top: 10px;")
        self.resultado_layout.addWidget(titulo2)

        for k, v in resultados.items():
            if k not in parametros_circuitales:
                fila = QHBoxLayout()
                label = QLabel(k)
                edit = QLineEdit(f"{v:.4f}")
                edit.setMaximumWidth(150)

                fila.addWidget(label)
                fila.addWidget(edit)

                self.resultado_layout.addLayout(fila)
                self.inputs_resultados[k] = edit

        btn_guardar = QPushButton("Guardar cambios")
        btn_guardar.clicked.connect(self.guardar_resultados_editados)
        btn_guardar.setStyleSheet("""
            QPushButton {
                background-color:#919499;
                color: white;
                font-weight: bold;
                font-size: 14px;
                padding: 10px;
                border-radius: 8px;
            }
            QPushButton:hover {
                background-color: #050505;
            }
            QPushButton:pressed {
                background-color: #919499;
            }
        """)
        self.resultado_layout.addWidget(btn_guardar)

        self.resultado_widget.show()

    def toggle_parametros_rbd(self, checked):

        for edit in self.inputs_manuales.values():
            edit.setEnabled(checked)

        if checked:
            self.tabla_rangos_widget.show()
        else:
            self.tabla_rangos_widget.hide()

    def guardar_resultados_editados(self):
        for k, edit in self.inputs_resultados.items():
            try:
                valor = float(edit.text())
                self.resultados[k] = valor
            except ValueError:
                QMessageBox.warning(
                    self,
                    "Error",
                    f"Valor inválido para {k}"
                )
                return
            
        # ---- Validar campos manuales ----
        if self.checkbox_editar_rbd.isChecked():        

            self.aRBD = float(self.inputs_manuales["aRBD"].text())
            self.VRBD = float(self.inputs_manuales["VRBD"].text())
            self.nRBD = float(self.inputs_manuales["nRBD"].text())       

        else:       

            self.aRBD = self.valores_default_rbd["aRBD"]
            self.VRBD = self.valores_default_rbd["VRBD"]
            self.nRBD = self.valores_default_rbd["nRBD"]

        mapeo_parametros = {
            "Corriente de saturación diodo 1 (Io1)": "Isat1_T0",
            "Corriente de saturación diodo 2 (Io2)": "Isat2_T0",
            "Resistencia serie (Rs)": "Rs",
            "Resistencia paralelo (Rp)": "Rp"
        }
        
        for k, v in self.resultados.items():
            if k in mapeo_parametros:
                setattr(self, mapeo_parametros[k], v)

        #Generar matriz de sombreado
        isc = self.inputs["Isc (A):"].text()
        ki = self.inputs["Ki (A/°C):"].text()
        diodos = self.Nd 
        filas = self.filas_panel
        columnas = self.columnas_panel  
 
        # print(f"diodos: {diodos}")
        # print(f"isc: {isc}")
        # print(f"ki: {ki}")

        if columnas % diodos != 0:
            QMessageBox.warning(
                self,
                "Error",
                f"Las columnas ({columnas}) no son divisibles entre los diodos ({diodos})."
            )
            return

        segmento = columnas // diodos
        patron = [segmento] * diodos

        # print("Patrón generado:", patron)

        # print(f"filas: {filas}, columnas: {columnas}")
        # print(f"Rs: {self.Rs}, Rsh: {self.Rp}, Rp: {self.Rp} , Isat2_T0: {self.Isat2_T0}, ki: {ki}")
        # print(f"Isc: {isc}, Rs: {self.Rs}, Rp: {self.Rp} , Isat1_T0: {self.Isat1_T0}")
        # print(f"aRBD: {self.aRBD}, VRBD: {self.VRBD} , nRBD: {self.nRBD}")

        # CREAR MODELO PVMISMATCH
        from pvmismatch import pvconstants, pvcell, pvmodule, pvstring, pvsystem
        self.pvconst = pvconstants.PVconstants(npts=500) #200 puntos

        self.pvcell = pvcell.PVcell(
            Rs=self.Rs,
            Rsh=self.Rp,
            Isat1_T0=self.Isat1_T0,
            Isat2_T0=self.Isat2_T0,
            Isc0_T0 = float(isc),
            aRBD=self.aRBD,
            bRBD=0.0,
            VRBD=self.VRBD,
            nRBD=self.nRBD,
            Eg=1.1,
            alpha_Isc=float(ki),
            Tcell=self.Tcell_K,
            Ee=1.0,
            pvconst=self.pvconst
        )
    
        cell_pos = pvmodule.standard_cellpos_pat(filas, patron)
    
        self.pvmod = pvmodule.PVmodule(
            cell_pos,
            self.pvcell,
            self.pvconst
        )
    
        self.pvstr = pvstring.PVstring(
            numberMods=1,
            pvmods=self.pvmod,
            pvconst=self.pvconst
        )
    
        self.pvsys = pvsystem.PVsystem(
            numberStrs=1,
            pvstrs=self.pvstr
        )

        # Inicializar con irradiancia uniforme 1000 W/m2
        self.pvsys.setSuns({0: {0: [(1.0, ) * self.Ns, tuple(range(self.Ns))]}})
        self.pvsys.update()

        self.modelo_guardado = True
        self.simular_panel_sombreado()  
        self.actualizar_modelo_pvmismatch()      

        QMessageBox.information(
            self,
            "Éxito",
            "Resultados actualizados correctamente"
        )  

    def tab_sombreado(self):
        tab = QWidget()
        layout_principal = QHBoxLayout(tab)

        # --- PANEL IZQUIERDO: Entradas y Cuadrícula ---
        panel_izq = QGroupBox()
        titulo_sombreado = QLabel("Configuración de Sombreado")
        titulo_sombreado.setAlignment(Qt.AlignCenter)
        titulo_sombreado.setStyleSheet("""
            font-size: 20px;
            font-weight: bold;
        """)

        layout_izq = QVBoxLayout(panel_izq)
        layout_izq.addWidget(titulo_sombreado)


        # ── Control de temperatura del panel ──────────────────────────
        grupo_temperatura = QGroupBox()
        layout_temperatura = QVBoxLayout(grupo_temperatura)

        self.label_temperatura = QLabel(f"Temperatura: {self.temp_default_c} °C")
        self.label_temperatura.setAlignment(Qt.AlignCenter)
        self.label_temperatura.setStyleSheet("font-weight: bold;")

        self.slider_temperatura = QSlider(Qt.Horizontal)
        self.slider_temperatura.setMinimum(0)
        self.slider_temperatura.setMaximum(40)
        self.slider_temperatura.setValue(self.temp_default_c)
        self.slider_temperatura.setTickPosition(QSlider.TicksBelow)
        self.slider_temperatura.setTickInterval(5)
        self.slider_temperatura.valueChanged.connect(self.cambiar_temperatura)

        layout_temperatura.addWidget(self.label_temperatura)
        layout_temperatura.addWidget(self.slider_temperatura)

        layout_izq.addWidget(grupo_temperatura)
        # ─────────────────────────────────────────────────────────────

        # ── ScrollArea que envuelve la cuadrícula ─────────────────────
        from PySide6.QtWidgets import QScrollArea

        self.scroll_area = QScrollArea()
        self.scroll_area.setWidgetResizable(True)
        self.scroll_area.setStyleSheet("QScrollArea { border: none; background-color: #eeeeee; }")

        # --- CUADRO QUE ENCIERRA LAS CELDAS (QFrame) ---
        self.contenedor_grid = QFrame()
        self.contenedor_grid.setFrameShape(QFrame.StyledPanel)
        self.contenedor_grid.setStyleSheet("""
            QFrame {
                border: 2px solid #444444;
                border-radius: 8px;
                background-color: #eeeeee;
                padding: 10px;
            }
        """)

        # Layout interno para las celdas
        self.grid_layout_celdas = QGridLayout(self.contenedor_grid)
        self.grid_layout_celdas.setSpacing(4)
        self.grid_layout_celdas.setAlignment(Qt.AlignCenter)

        layout_izq.addWidget(self.contenedor_grid)
        layout_izq.addStretch()  # Empuja todo hacia arriba

        # --- PANEL DERECHO: Gráficas ---
        panel_der = QGroupBox()
        titulo_sombreado = QLabel("Visualización de Curvas")
        titulo_sombreado.setAlignment(Qt.AlignCenter)
        titulo_sombreado.setStyleSheet("""
            font-size: 20px;
            font-weight: bold;
        """)

        layout_der = QVBoxLayout(panel_der)
        layout_der.addWidget(titulo_sombreado)
        self.canvas = FigureCanvas(Figure(figsize=(7, 5)))
        layout_der.addWidget(self.canvas)

        layout_principal.addWidget(panel_izq, 2)
        layout_principal.addWidget(panel_der, 3)

        return tab

    def simular_panel_sombreado(self):

            # Verificar que existan Ns y Nd
            if not hasattr(self, "Ns") or not hasattr(self, "Nd"):
                QMessageBox.warning(
                    self,
                    "Error",
                    "Primero debe configurar el panel en la pestaña Configuración."
                )
                return
            
            # ── Limpiar la matriz anterior completamente ─────────────────
            while self.grid_layout_celdas.count():
                item = self.grid_layout_celdas.takeAt(0)
                widget = item.widget()
                if widget is not None:
                    widget.deleteLater()
            self.cell_labels = {}
            # ─────────────────────────────────────────────────────────────    

            filas = self.filas_panel
            columnas = self.columnas_panel

            self.irradiance_matrix = np.full((filas, columnas), 1000.0)
            self.cell_labels = {}

            for r in range(filas):
                for c in range(columnas):
                
                    label = QLabel("1000")
                    label.setProperty("row", r)
                    label.setProperty("col", c)
                    label.setAlignment(Qt.AlignCenter)   # ← solo esto, sin setFixedSize
                    label.installEventFilter(self)

                    self.grid_layout_celdas.addWidget(label, r, c)
                    self.cell_labels[(r, c)] = label
                    self._update_cell_appearance(label, 1000.0)

    def _update_cell_appearance(self, label, value):    
        # Lógica de color (Verde a Rojo)
        norm = (value - self.min_irradiance) / (self.max_irradiance - self.min_irradiance)
        r = int(255 * (1 - norm))
        g = int(255 * norm)
        
        # Color de texto (blanco si el fondo es muy oscuro, negro si es claro)
        texto_color = "white" if norm < 0.4 else "black"
        
        # Aplicamos TODO el estilo aquí. Esto sobreescribe cualquier conflicto previo.
        label.setStyleSheet(f"""
            QLabel {{
            background-color: rgb({r}, {g}, 0);
            color: {texto_color};
            border: 1px solid #444;
            font-weight: bold;
            font-size: 9px;
            border-radius: 2px;
            min-width: 50px;
            min-height: 20px;
            padding: 2px 4px;
        }}
        """)
        label.setText(str(int(value)))  

    def eventFilter(self, source, event):
        if event.type() == event.Type.Wheel and isinstance(source, QLabel):
            r = source.property("row")
            c = source.property("col")
            
            if r is not None and c is not None:
                delta = event.angleDelta().y()
                direction = 1 if delta > 0 else -1
                
                new_val = self.irradiance_matrix[r, c] + (self.irradiance_step * direction)
                new_val = np.clip(new_val, self.min_irradiance, self.max_irradiance)
                
                self.irradiance_matrix[r, c] = new_val
                self._update_cell_appearance(source, new_val)
                self.actualizar_modelo_pvmismatch()
                return True
        return super().eventFilter(source, event)

    def cambiar_temperatura(self, valor):
        self.Tcell_K = valor + 273.15
        self.label_temperatura.setText(f"Temperatura: {valor} °C")

        if self.modelo_guardado:
            self.actualizar_modelo_pvmismatch()

    def actualizar_modelo_pvmismatch(self):
    
        filas, columnas = self.irradiance_matrix.shape
    
        try:
            # ── Verificar si todas las celdas están a 1000 ───────────
            todas_a_1000 = np.all(self.irradiance_matrix == 1000.0)
    
            if todas_a_1000:
                # Resetear con formato de inicialización original
                self.pvsys.setSuns({0: {0: [(1.0,) * self.Ns, tuple(range(self.Ns))]}})
                self.pvsys.update()
            else:
                # Solo enviar celdas que NO están a 1000
                indices = []
                irradiancias = []
    
                for r in range(filas):
                    for c in range(columnas):
                        val = self.irradiance_matrix[r, c] / 1000.0
    
                        if val < 1.0:
                            val = max(val, 0.001)
    
                            if c % 2 == 0:
                                idx = c * filas + r
                            else:
                                idx = c * filas + (filas - 1 - r)
    
                            indices.append(idx)
                            irradiancias.append(val)
                            #print(f"Celda sombreada = fila:{r}, col:{c}, indice:{idx}, valor:{val:.3f}")
    
                #print(f"irradiancias: {tuple(irradiancias)}, indices: {tuple(indices)}")
                if len(indices) > 0:
                    self.pvsys.setSuns({
                        0: {0: [tuple(irradiancias), tuple(indices)]}
                    })
                else:
                    self.pvsys.setSuns({0: {0: [(1.0,) * self.Ns, tuple(range(self.Ns))]}})
    
            self.pvsys.update()
    
            # ── Aplicar la temperatura actual (tomada del slider) ────────
            self.pvsys.setTemps(self.Tcell_K)
            self.pvsys.update()
            # ───────────────────────────────────────────────────────────
    
            #print("Atributos pvsys:", [a for a in dir(self.pvsys) if not a.startswith('_')])
            #Obtener curvas
            P = self.pvsys.Psys
            #print(f"P: {P}")
            V = self.pvsys.Vsys
            #print(f"V: {V}")
            I = self.pvsys.Isys
            #print(f"I: {I}")
    
            # Filtrar valores negativos de potencia
            P = np.array(P).flatten()
            V = np.array(V).flatten()
    
            indP0 = P > 0.
            P = P[indP0]
            V = V[indP0]
            I = P / V  # Recalcular corriente a partir de P y V filtrados
    
            if len(V) > 0:
                # Punto de máxima potencia
                indm = np.argmax(P)
                Pmpp = P[indm]
                Vmpp = V[indm]
                Impp = I[indm]
                self.V_actual = V
                self.I_actual = I
                self.P_actual = P
                self.Vmpp_actual = Vmpp
                self.Impp_actual = Impp
                self.Pmpp_actual = Pmpp
                self.actualizar_graficas(V, I, P, Vmpp, Impp, Pmpp)
    
        except IndexError as e:
            print(f"pvmismatch no pudo calcular MPP: {e} — irradiancia muy baja")
        except Exception as e:
            print(f"Error en actualizar_modelo_pvmismatch: {e}")

    def actualizar_graficas(self, V, I, P, Vmpp, Impp, Pmpp):

        # Limpiar figura completa y redibujar 
        self.canvas.figure.clear()
        ax = self.canvas.figure.subplots(1, 2)

        # Curva I-V
        ax[0].plot(V, I, 'b-')
        ax[0].plot(Vmpp, Impp, 'r+', markersize=12, markeredgewidth=2,
                   label=f'MPP ({Vmpp:.2f} V, {Impp:.2f} A)')
        ax[0].set_xlabel('Voltaje (V)')
        ax[0].set_ylabel('Corriente (A)')
        ax[0].set_title('Curva I–V')
        ax[0].grid(True)
        ax[0].legend(fontsize=8)

        # Curva P-V
        ax[1].plot(V, P, 'r-')
        ax[1].plot(Vmpp, Pmpp, 'b+', markersize=12, markeredgewidth=2,
                   label=f'MPP ({Vmpp:.2f} V, {Pmpp:.2f} W)')
        ax[1].set_xlabel('Voltaje (V)')
        ax[1].set_ylabel('Potencia (W)')
        ax[1].set_title('Curva P–V')
        ax[1].grid(True)
        ax[1].legend(fontsize=8)

        self.canvas.figure.tight_layout()
        self.canvas.draw()

    def graficar_curvas(self, V_2, I_final, Vmax_2, Imax_2, Vmp, Imp):
        self.canvas.figure.clear()

        ax = self.canvas.figure.subplots(1, 2)

        # ===== CURVA I–V =====
        ax[0].plot(V_2, I_final, 'b-')
        ax[0].plot(Vmax_2, Imax_2, 'b+', markersize=10)

        ax[0].set_xlabel('Voltaje (V)')
        ax[0].set_ylabel('Corriente (A)')
        ax[0].set_title('Curva I–V')
        ax[0].grid(True)
        ax[0].legend()

        # ===== CURVA P–V =====
        P_final = V_2 * I_final

        ax[1].plot(V_2, P_final, 'r-')
        ax[1].plot(Vmax_2, P_final.max(), 'r+')

        ax[1].set_xlabel('Voltaje (V)')
        ax[1].set_ylabel('Potencia (W)')
        ax[1].set_title('Curva P–V')
        ax[1].grid(True)
        ax[1].legend()

        self.canvas.figure.tight_layout()
        self.canvas.draw()

    def validar_cambio_pestana(self, index):
     
        nombre_tab = self.tabs.tabText(index)
     
        if nombre_tab == "Sombreado":
         
            if not self.modelo_guardado:
             
                QMessageBox.warning(
                    self,
                    "Modelo no guardado",
                    "Debes guardar cambios antes de ir a Sombreado."
                )
     
                self.tabs.blockSignals(True)
                self.tabs.setCurrentIndex(0)
                self.tabs.blockSignals(False)
     
                return


# ================= MAIN =================
if __name__ == "__main__":
    app = QApplication(sys.argv)
    ventana = VentanaPrincipal()
    ventana.show()
    sys.exit(app.exec())

