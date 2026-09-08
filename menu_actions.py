# menu_actions.py
import sys, os

def resource_path(relative_path):
    if hasattr(sys, "_MEIPASS"):
        base_path = sys._MEIPASS
    else:
        base_path = os.path.dirname(os.path.abspath(__file__))
    return os.path.join(base_path, relative_path)
from openpyxl import Workbook
from PySide6.QtWidgets import QMessageBox, QFileDialog
from PySide6.QtCore import QUrl


def descargar_parametros_xlsx(ventana):
    if not hasattr(ventana, "resultados") or not ventana.resultados:
        QMessageBox.warning(ventana, "Sin datos", "Primero debe estimar los parámetros.")
        return

    ruta, _ = QFileDialog.getSaveFileName(
        ventana, "Guardar parámetros estimados", "parametros_estimados.xlsx", "Excel (*.xlsx)"
    )
    if not ruta:
        return

    wb = Workbook()
    ws = wb.active
    ws.title = "Parámetros Estimados"
    ws.append(["Parámetro", "Valor"])
    for k, v in ventana.resultados.items():
        ws.append([k, v])

    if hasattr(ventana, "aRBD"):
        ws.append([])
        ws.append(["--- Parámetros de ruptura inversa ---", ""])
        ws.append(["aRBD", ventana.aRBD])
        ws.append(["VRBD", ventana.VRBD])
        ws.append(["nRBD", ventana.nRBD])

    ws.column_dimensions["A"].width = 45
    ws.column_dimensions["B"].width = 20
    wb.save(ruta)
    QMessageBox.information(ventana, "Éxito", f"Archivo guardado en:\n{ruta}")


def guardar_sombreado_excel(ventana):
    if ventana.irradiance_matrix is None:
        QMessageBox.warning(ventana, "Sin datos", "No hay matriz de sombreado generada.")
        return

    ruta, _ = QFileDialog.getSaveFileName(
        ventana, "Guardar matriz de irradiancia", "sombreado.xlsx", "Excel (*.xlsx)"
    )
    if not ruta:
        return

    wb = Workbook()
    ws = wb.active
    ws.title = "Matriz de Irradiancia"

    filas, columnas = ventana.irradiance_matrix.shape
    for c in range(columnas):
        ws.cell(row=1, column=c + 2, value=f"Col {c}")
    for r in range(filas):
        ws.cell(row=r + 2, column=1, value=f"Fila {r}")
        for c in range(columnas):
            ws.cell(row=r + 2, column=c + 2, value=ventana.irradiance_matrix[r, c])

    fila_temp = filas + 3
    ws.cell(row=fila_temp, column=1, value="Temperatura del panel (°C):")
    ws.cell(row=fila_temp, column=2, value=round(ventana.Tcell_K - 273.15, 2))

    if ventana.V_actual is not None:
        ws2 = wb.create_sheet("Curva I-V P-V")
        ws2.cell(row=1, column=1, value="V (Voltios)")
        ws2.cell(row=1, column=2, value="I (Amperios)")
        ws2.cell(row=1, column=3, value="P (Watts)")
        for i in range(len(ventana.V_actual)):
            ws2.cell(row=i + 2, column=1, value=float(ventana.V_actual[i]))
            ws2.cell(row=i + 2, column=2, value=float(ventana.I_actual[i]))
            ws2.cell(row=i + 2, column=3, value=float(ventana.P_actual[i]))

    wb.save(ruta)
    QMessageBox.information(ventana, "Éxito", f"Matriz guardada en:\n{ruta}")

def descargar_sombreado_png(ventana):
    if not hasattr(ventana, "canvas") or ventana.canvas is None:
        QMessageBox.warning(ventana, "Sin datos", "No hay gráficas generadas.")
        return

    ruta, _ = QFileDialog.getSaveFileName(
        ventana, "Guardar gráficas de sombreado", "sombreado_curvas.png", "PNG (*.png)"
    )
    if not ruta:
        return

    ventana.canvas.figure.savefig(ruta, dpi=150, bbox_inches="tight")
    QMessageBox.information(ventana, "Éxito", f"Gráficas guardadas en:\n{ruta}")


def descargar_parametros_panel(ventana):
    campos = {k: v.text() for k, v in ventana.inputs.items() if v.text().strip() != ""}
    if not campos:
        QMessageBox.warning(ventana, "Sin datos", "No hay parámetros ingresados aún.")
        return

    ruta, _ = QFileDialog.getSaveFileName(
        ventana, "Guardar parámetros del panel", "parametros_panel.xlsx", "Excel (*.xlsx)"
    )
    if not ruta:
        return

    wb = Workbook()
    ws = wb.active
    ws.title = "Parámetros Panel"
    ws.append(["Parámetro", "Valor"])
    for k, v in campos.items():
        ws.append([k, v])
    ws.column_dimensions["A"].width = 35
    ws.column_dimensions["B"].width = 20
    wb.save(ruta)
    QMessageBox.information(ventana, "Éxito", f"Archivo guardado en:\n{ruta}")


def descargar_todo(ventana):
    tiene_panel   = any(v.text().strip() != "" for v in ventana.inputs.values())
    tiene_estim   = hasattr(ventana, "resultados") and ventana.resultados
    tiene_sombra  = ventana.irradiance_matrix is not None
    tiene_graficas = hasattr(ventana, "canvas") and ventana.canvas is not None

    if not tiene_panel and not tiene_estim:
        QMessageBox.warning(ventana, "Sin datos", "No hay datos suficientes para exportar.")
        return

    carpeta = QFileDialog.getExistingDirectory(ventana, "Seleccionar carpeta de destino")
    if not carpeta:
        return

    archivos_creados = []

    if tiene_panel:
        wb = Workbook()
        ws = wb.active
        ws.title = "Parámetros Panel"
        ws.append(["Parámetro", "Valor"])
        for k, v in ventana.inputs.items():
            if v.text().strip():
                ws.append([k, v.text()])
        ws.column_dimensions["A"].width = 35
        ws.column_dimensions["B"].width = 20
        wb.save(os.path.join(carpeta, "parametros_panel.xlsx"))
        archivos_creados.append("parametros_panel.xlsx")

    if tiene_estim:
        wb = Workbook()
        ws = wb.active
        ws.title = "Parámetros Estimados"
        ws.append(["Parámetro", "Valor"])
        for k, v in ventana.resultados.items():
            ws.append([k, v])
        if hasattr(ventana, "aRBD"):
            ws.append([])
            ws.append(["--- Ruptura inversa ---", ""])
            ws.append(["aRBD", ventana.aRBD])
            ws.append(["VRBD", ventana.VRBD])
            ws.append(["nRBD", ventana.nRBD])
        ws.column_dimensions["A"].width = 45
        ws.column_dimensions["B"].width = 20
        wb.save(os.path.join(carpeta, "parametros_estimados.xlsx"))
        archivos_creados.append("parametros_estimados.xlsx")

    if tiene_sombra:
        wb = Workbook()
        ws = wb.active
        ws.title = "Sombreado"
        filas, columnas = ventana.irradiance_matrix.shape
        for c in range(columnas):
            ws.cell(row=1, column=c + 2, value=f"Col {c}")
        for r in range(filas):
            ws.cell(row=r + 2, column=1, value=f"Fila {r}")
            for c in range(columnas):
                ws.cell(row=r + 2, column=c + 2, value=ventana.irradiance_matrix[r, c])
        wb.save(os.path.join(carpeta, "sombreado_matriz.xlsx"))
        archivos_creados.append("sombreado_matriz.xlsx")

    if tiene_graficas:
        ruta_png = os.path.join(carpeta, "curvas_IV_PV.png")
        ventana.canvas.figure.savefig(ruta_png, dpi=150, bbox_inches="tight")
        archivos_creados.append("curvas_IV_PV.png")

    resumen = "\n".join(f"  ✓ {f}" for f in archivos_creados)
    QMessageBox.information(ventana, "Éxito", f"Archivos guardados en:\n{carpeta}\n\n{resumen}")


def mostrar_acerca_de(ventana):
    from PySide6.QtWidgets import QDialog, QVBoxLayout, QTextBrowser, QPushButton

    dialog = QDialog(ventana)
    dialog.setWindowTitle("Acerca del simulador")
    dialog.setMinimumSize(750, 620)

    layout = QVBoxLayout(dialog)

    texto = QTextBrowser()
    texto.setOpenExternalLinks(True)
    texto.setStyleSheet("font-size: 13px; padding: 8px;")
    texto.setSearchPaths([resource_path("assets")])
    texto.setHtml("""
    <h2 style='color:#00000;'>Simulador de sombreado en módulos fotovoltaicos</h2>
    <p><i>Modelo de dos diodos con ruptura inversa (Bishop)</i></p>

    <h3 style='color:#00000;'>¿Qué hace este sistema?</h3>
    <p>Permite modelar el comportamiento eléctrico de un módulo fotovoltaico bajo condiciones 
    de sombreado parcial, usando el modelo de doble diodo con ruptura inversa según la formulación 
    de Bishop. Se estiman los parámetros del circuito equivalente a partir de los datos del 
    fabricante, se configuran patrones de irradiancia celda a celda, y se visualizan las curvas 
    I–V y P–V resultantes mediante la librería pvmismatch.</p>

    <h3 style='color:#00000;'>Flujo de uso</h3>
    <ol>
        <li><b>Seleccionar panel o ingresar datos:</b> Elija un panel del catálogo (MSX60, SM55, ST40, 
        KGC200GT, STP280, SPR-315) o ingrese manualmente los parámetros de la hoja de datos.</li>
        <li><b>Estimar parámetros:</b> Presione "Estimar parámetros" para obtener Iph, Io1, Io2, Rs y Rp.</li>
        <li><b>Revisar y ajustar:</b> Observe los parámetros estimados. Active el checkbox 
        "Modificar parámetros de ruptura inversa" para ajustar aRBD, VRBD y nRBD.</li>
        <li><b>Guardar cambios:</b> Construye el modelo pvmismatch con la matriz de celdas del panel 
        e irradiancia uniforme inicial de 1000 W/m².</li>
        <li><b>Configurar sombreado:</b> En la pestaña Sombreado, use la rueda del mouse sobre 
        cada celda para modificar su irradiancia en pasos de 50 W/m².</li>
        <li><b>Visualizar curvas:</b> Las curvas I–V y P–V se actualizan automáticamente.</li>
    </ol>

    <h3 style='color:#00000;'>Modelo de dos diodos</h3>
    <p>Representa la celda solar separando la recombinación en zona de carga espacial (diodo 2, 
    a₂ = 2) de la recombinación cuasi-neutral (diodo 1, a₁ = 1):</p>
    <p style='padding:8px; font-family:monospace; border-left:4px solid #378ADD;'>
    I = Iph − Io1·[exp((V+I·Rs)/(a1·Vt)) − 1] − Io2·[exp((V+I·Rs)/(a2·Vt)) − 1] − (V+I·Rs)/Rp
    </p>
    <p style='padding:8px; font-family:monospace; border-left:4px solid #378ADD;'>
    Vt = Ns · k · T / q
    </p>
                  
    <table width='100%' cellspacing='4'>
        <tr><td width='160'><b>Iph</b></td><td>Fotocorriente generada</td></tr>
        <tr><td><b>Io1, Io2</b></td><td>Corrientes de saturación diodos 1 y 2</td></tr>
        <tr><td><b>Rs</b></td><td>Resistencia serie</td></tr>
        <tr><td><b>Rp</b></td><td>Resistencia paralelo (shunt)</td></tr>
        <tr><td><b>a1, a2</b></td><td>Factores de idealidad (1.0 y 2.0)</td></tr>
        <tr><td><b>Vt</b></td><td>Voltaje térmico</td></tr>
        <tr><td><b>Ns</b></td><td>Número de celdas en serie</td></tr>
        <tr><td><b>k</b></td><td>Constante de Boltzmann (1.381×10⁻²³ J/K)</td></tr>
        <tr><td><b>q</b></td><td>Carga del electrón (1.602×10⁻¹⁹ C)</td></tr>
    </table>

    <h3 style='color:#00000;'>Modelo de Bishop — ruptura inversa</h3>
    <p>Extiende el modelo para incluir el efecto avalancha en polarización inversa, 
    crítico para modelar celdas sombreadas que operan en inversa:</p>
    <p style='padding:8px; font-family:monospace; border-left:4px solid #378ADD;'>
    I = Iph − Io1·[exp((V+I·Rs)/(a1·Vt))−1] − Io2·[exp((V+I·Rs)/(a2·Vt))−1]<br><br>
    &nbsp;&nbsp;&nbsp;&nbsp;− (V+I·Rs)/Rp · { 1 + aRBD · [1 − (V+I·Rs)/VRBD]<sup>−nRBD</sup> }
    </p>
    <br>
    <table width='100%' cellspacing='4'>
        <tr><td width='160'><b>aRBD</b></td><td>Coeficiente de avalancha (1×10⁻⁴ → 3.5×10⁻¹)</td></tr>
        <tr><td><b>VRBD</b></td><td>Tensión de ruptura inversa (−21.29 V → −10 V)</td></tr>
        <tr><td><b>nRBD</b></td><td>Exponente de ruptura (1 → 6)</td></tr>
    </table>
    <p>Valores por defecto: aRBD = 1.036×10⁻⁴, VRBD = −5.527 V, nRBD = 3.285.</p>
    
    <h4 style='color:#00000;'>Modelo circuital de Bishop</h4>
    <img src='modelobishop.png' width='600' style='display:block; margin:10px auto;'/>
                  
    <h3 style='color:#00000;'>Dependencia con temperatura e irradiancia</h3>
    
    <table width='100%' cellspacing='4'>
        <tr><td width='160'><b>Isc0</b></td><td>Corriente de cortocircuito a STC (25°C, 1000 W/m²)</td></tr>
        <tr><td><b>Ki</b></td><td>Coeficiente térmico de Isc (A/°C)</td></tr>
        <tr><td><b>Ee</b></td><td>Irradiancia efectiva normalizada (0 a 1)</td></tr>
        <tr><td><b>Eg</b></td><td>Energía de banda prohibida del Si (1.1 eV)</td></tr>
        <tr><td><b>T0</b></td><td>Temperatura STC = 298.15 K</td></tr>
    </table>

    <h3 style='color:#00000;'>Referencia</h3>
    <p>Bishop, J. W. (1988). <i>Computer simulation of the effects of electrical mismatches 
    in photovoltaic cell interconnection circuits.</i> Solar Cells, 25(1), 73–89.<br><br>
    Desarrollado con <b>PySide6</b> y <b>pvmismatch</b>.</p>
    """)

    btn = QPushButton("Cerrar")
    btn.setStyleSheet("""
        QPushButton {
            background-color: #919499; color: white; font-weight: bold;
            font-size: 13px; padding: 8px 20px; border-radius: 6px;
        }
        QPushButton:hover { background-color: #050505; }
    """)
    btn.clicked.connect(dialog.accept)

    layout.addWidget(texto)
    layout.addWidget(btn)
    dialog.exec()