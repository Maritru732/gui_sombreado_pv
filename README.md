# Simulador de Sombreado en Módulos Fotovoltaicos

Aplicación de escritorio (GUI) desarrollada en Python que modela el comportamiento eléctrico de un módulo fotovoltaico bajo condiciones de **sombreado parcial**, usando el **modelo de dos diodos con ruptura inversa (modelo de Bishop)**.

Permite estimar los parámetros circuitales de un panel a partir de su hoja de datos, configurar un patrón de irradiancia celda por celda, y visualizar en tiempo real las curvas resultantes **I–V** y **P–V**.

## Características

- Ingreso manual de parámetros del panel o selección desde un catálogo predefinido (MSX60, SM55, ST40, KGC200GT, STP280, SPR-315, I-110/24).
- Estimación automática de los parámetros del circuito equivalente de dos diodos (Iph, Io1, Io2, Rs, Rp) mediante Newton-Raphson.
- Edición manual de los parámetros estimados y de los parámetros de ruptura inversa (aRBD, VRBD, nRBD).
- Cuadrícula interactiva de celdas: la irradiancia de cada celda se ajusta con la rueda del mouse.
- Control de temperatura del panel mediante un deslizador.
- Simulación eléctrica del módulo con [`pvmismatch`](https://github.com/SunPower/PVMismatch), incluyendo el efecto de diodos de bypass.
- Visualización en vivo de las curvas I–V y P–V, con marcado del punto de máxima potencia (MPP).
- Exportación de resultados a Excel (`.xlsx`) y de las gráficas a PNG.

## Estructura del proyecto

```
gui_sompreado_pv/
├── main.py             # Ventana principal y lógica de la interfaz (PySide6)
├── menu_actions.py      # Acciones del menú: exportar a Excel/PNG, ventana "Acerca de"
├── estimacion_pv.py     # Estimación de parámetros del modelo de dos diodos
├── assets/               # Imágenes usadas en la interfaz (modelo circuital, modelo de Bishop)
├── SimuladorPV.spec     # Configuración de PyInstaller para generar el ejecutable
└── requirements.txt      # Dependencias de Python con versiones fijas
```

## Requisitos

- **Python 3.11** (probado con 3.11.9). No se garantiza compatibilidad con versiones anteriores a 3.10, ya que `numpy` y `scipy` en las versiones usadas aquí lo requieren.
- **Sistema operativo:** desarrollado y probado en Windows 11. Debería funcionar en Linux/macOS al ser Python + Qt puro, pero no ha sido verificado en esos entornos.
- Conexión gráfica (entorno de escritorio) para mostrar la interfaz Qt.

### Librerías utilizadas

| Librería | Versión usada | Uso en el proyecto |
|---|---|---|
| [PySide6](https://pypi.org/project/PySide6/) | 6.10.1 | Interfaz gráfica (ventanas, pestañas, widgets, menús) |
| [numpy](https://pypi.org/project/numpy/) | 2.4.2 | Cálculo numérico y manejo de la matriz de irradiancia |
| [scipy](https://pypi.org/project/scipy/) | 1.17.0 | Dependencia numérica requerida por `pvmismatch` |
| [matplotlib](https://pypi.org/project/matplotlib/) | 3.10.8 | Gráficas de las curvas I–V y P–V embebidas en la interfaz |
| [openpyxl](https://pypi.org/project/openpyxl/) | 3.1.5 | Generación de archivos Excel (`.xlsx`) con los resultados |
| [pvmismatch](https://pypi.org/project/pvmismatch/) | 4.1 | Motor de simulación eléctrica del módulo fotovoltaico (mismatch por sombreado) |

Todas las versiones exactas quedan fijadas en [`requirements.txt`](requirements.txt).

## Instalación

1. **Clonar el repositorio**

   ```bash
   git clone https://github.com/Maritru732/gui_sombreado_pv.git
   cd gui_sombreado_pv
   ```

2. **Crear y activar un entorno virtual** (recomendado)

   En Windows (PowerShell):

   ```powershell
   python -m venv venv
   .\venv\Scripts\Activate.ps1
   ```

   En Linux/macOS:

   ```bash
   python3 -m venv venv
   source venv/bin/activate
   ```

3. **Instalar las dependencias**

   ```bash
   pip install -r requirements.txt
   ```

## Ejecución

Con el entorno virtual activado:

```bash
python main.py
```

Se abrirá la ventana principal con dos pestañas:

1. **Configuración:** seleccione un panel del catálogo o ingrese sus parámetros manualmente, presione **"Estimar parámetros"**, revise/ajuste los resultados y presione **"Guardar cambios"** para construir el modelo eléctrico.
2. **Sombreado:** use la rueda del mouse sobre cada celda de la cuadrícula para modificar su irradiancia (pasos de 50 W/m²) y observe cómo cambian las curvas I–V y P–V en tiempo real. También puede ajustar la temperatura del panel con el deslizador.

Desde el menú **"Menú"** puede exportar los parámetros del panel, los parámetros estimados, la matriz de sombreado y las curvas, en Excel o PNG.

## Generar el ejecutable (opcional)

El proyecto incluye [`SimuladorPV.spec`](SimuladorPV.spec) para empaquetar la aplicación con [PyInstaller](https://pyinstaller.org/):

```bash
pip install pyinstaller==6.21.0
pyinstaller SimuladorPV.spec
```

El ejecutable se genera en `dist/SimuladorPV/`. Esta carpeta y `build/` no se versionan en el repositorio (ver [`.gitignore`](.gitignore)).

## Fundamento teórico

El simulador se basa en el modelo de dos diodos:

```
I = Iph − Io1·[exp((V+I·Rs)/(a1·Vt)) − 1] − Io2·[exp((V+I·Rs)/(a2·Vt)) − 1] − (V+I·Rs)/Rp
```

extendido con el modelo de **Bishop** para representar el efecto de avalancha en polarización inversa (crítico para celdas sombreadas):

```
I = Iph − Io1·[exp((V+I·Rs)/(a1·Vt))−1] − Io2·[exp((V+I·Rs)/(a2·Vt))−1]
    − (V+I·Rs)/Rp · { 1 + aRBD · [1 − (V+I·Rs)/VRBD]^(−nRBD) }
```

**Referencia:** Bishop, J. W. (1988). *Computer simulation of the effects of electrical mismatches in photovoltaic cell interconnection circuits.* Solar Cells, 25(1), 73–89.

Más detalle sobre el flujo de uso y las variables del modelo está disponible dentro de la propia aplicación, en **Menú → Acerca de**.
