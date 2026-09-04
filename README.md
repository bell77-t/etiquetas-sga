# 🌱 SGA Label Studio — Visor y Generador de Etiquetas Fitosanitarias

> **Sistema Web Local para la visualización, búsqueda en tiempo real, desglose volumétrico por tanques de 1,000 L y generación de etiquetas de seguridad SGA (Sistema Globalmente Armonizado) para mezclas agronómicas.**

---

## 📋 Tabla de Contenidos
- [Características Principales](#-características-principales)
- [Estructura del Proyecto y Archivos de Entrada](#-estructura-del-proyecto-y-archivos-de-entrada)
- [Arquitectura Técnica](#-arquitectura-técnica)
- [Instalación y Puesta en Marcha](#-instalación-y-puesta-en-marcha)
- [Guía de Uso](#-guía-de-uso)
- [Opciones de Impresión y Fotocopiado](#-opciones-de-impresión-y-fotocopiado)
- [🚀 ¿Qué le falta a la aplicación para ser perfecta? (Roadmap y Mejoras)](#-qué-le-falta-a-la-aplicación-para-ser-perfecta-roadmap-y-mejoras)
- [Licencia](#-licencia)

---

## ✨ Características Principales

1. **Lectura No Bloqueante en Tiempo Real**:
   - Lee `aplicacion.xlsm` y `Base.xlsx` directamente en memoria (`io.BytesIO`) con permisos de lectura compartida.
   - Puedes **editar o tener abiertos los libros en Microsoft Excel** y la aplicación nunca se bloqueará ni generará errores de archivo en uso.

2. **Desglose Automático de Tanques de 1,000 Litros y Colitas**:
   - Para aplicaciones con volumen $\ge 1,000\text{ L}$, calcula $N = \lfloor \text{Volumen} / 1000 \rfloor$ tanques completos de $1,000\text{ L}$ con dosificación proporcional exacta ($1,000 \times \text{Dosis}$).
   - Si queda remanente ($\text{Volumen} \bmod 1000 > 0$), genera automáticamente la etiqueta residual ("Colita") con los litros y cantidad restantes calculados.
   - Para aplicaciones $< 1,000\text{ L}$, genera una etiqueta única con el volumen exacto.

3. **Réplica Visual Exacta de la Plantilla Física de Excel (`A2:F12`)**:
   - Rejilla estructurada con bordes negros sólidos (`1.5px solid black`) y tipografía `Arial`.
   - **Bloque Superior Izquierdo**: Nombre del producto, Bloque/Sector, Fecha, Reentrada, Unidad, Cantidad y Volumen de Tanque.
   - **Bloque Superior Derecho**: Matriz $2 \times 2$ de pictogramas SGA con la palabra de advertencia destacada (`PELIGRO` o `ATENCIÓN`).
   - **Bloque Inferior**: Frases H (Peligros) y Frases P (Consejos de Prudencia) completas, visibles sin recortes y con resaltado nítido de códigos normativos.

4. **Servidor y Resolución Flexible de Pictogramas**:
   - Mapeo automático de códigos GHS (`GHS01` a `GHS09`) e imágenes `.jpg` desde la carpeta `Picto/`.
   - Manejo inteligente de fallbacks: si un producto no tiene imagen asignada o dice `SIN FOTO` en `Base.xlsx`, renderiza un recuadro limpio con el texto `SIN IMAGEN` sin romper el diseño.

5. **Multi-impresión, Copias Múltiples y Ahorro de Papel**:
   - Permite imprimir **$N$ copias consecutivas** de una etiqueta individual o de lotes seleccionados.
   - Modos de disposición: **1 Columna Continua** (fiel a Excel), **2 Columnas** (para fotocopias masivas de 4 a 6 por hoja), y **1 por Hoja** (para rollos de etiquetas térmicas adhesivas).

---

## 📂 Estructura del Proyecto y Archivos de Entrada

```plaintext
etiquetas_app/
├── app.py                     # Servidor backend FastAPI y rutas REST / estáticas
├── data_loader.py             # Motor de lectura de Excel, cruce SGA y cálculo de tanques
├── test_app.py                # Suite de pruebas unitarias automatizadas
├── run_app.bat                # Lanzador con doble clic para Windows
├── .gitignore                 # Filtro para ignorar temporales de Excel y Python
├── README.md                  # Documentación completa del proyecto
├── Picto/                     # Carpeta local con imágenes de pictogramas GHS (.jpg)
│   ├── GHS01.jpg ... GHS09.jpg
│   └── SIN FOTO.jpg
└── static/                    # Frontend Web
    ├── index.html             # Interfaz interactiva HTML5 con TailwindCSS y Lucide
    ├── styles.css             # Estilos de la plantilla Excel y reglas @media print
    ├── app.js                 # Lógica cliente: filtros, copias, visor e impresión
    └── picto/                 # Espejo estático de pictogramas
```

### Archivos Excel leídos por la aplicación:
- **`aplicacion.xlsm`** (Hoja `Data`): Programación semanal con `Fecha`, `Cultivo`, `Bloque`, `Litros Agua`, `Dosis gr ó cc x Litro`, `PRODUCTO`, `Reentrada`, `Categoria`, etc.
- **`Base.xlsx`** (Hoja `SGA`): Maestro de productos químicos con `Codigo`, `Nombre del Producto`, `U/m`, `FRASES H`, `FRASES P`, `PALABRA DE ADVERTENCIA`, `PIG1`, `PIG2`, `PIG3`, `PIG4`.
- **`Picto/`**: Directorio con los rombos de seguridad en formato JPG.

---

## 🛠️ Arquitectura Técnica

- **Backend**: Python 3.10+ | [FastAPI](https://fastapi.tiangolo.com/) | [Uvicorn](https://www.uvicorn.org/) | [openpyxl](https://openpyxl.readthedocs.io/)
- **Frontend**: HTML5 Semántico | [Tailwind CSS](https://tailwindcss.com/) | [Lucide Icons](https://lucide.dev/) | Vanilla JavaScript ES6+
- **Motor de Impresión**: Reglas CSS avanzadas `@media print` (`page-break-after: always;`, `break-inside: avoid;`, `print-color-adjust: exact;`).

---

## 🚀 Instalación y Puesta en Marcha

### Requisitos Previos:
- Python 3.9 o superior instalado en el sistema.

### 1. Clonar el repositorio
```bash
git clone https://github.com/bell77-t/etiquetas-sga.git
cd etiquetas-sga
```

### 2. Instalar dependencias
```bash
pip install fastapi uvicorn openpyxl pandas
```

### 3. Ejecutar la aplicación

**En Windows (Modo Rápido):**
Haz doble clic en el archivo **`run_app.bat`** (o en **`ABRIR_ETIQUETAS.bat`**).

**Desde la terminal:**
```bash
python -m uvicorn app:app --host 127.0.0.1 --port 8000
```
Abre tu navegador en: 👉 **[http://localhost:8000](http://localhost:8000)**

---

## 📖 Guía de Uso

1. **Filtro de Fechas**: Selecciona cualquier fecha en las insignias superiores para ver solo los lotes programados de ese día.
2. **Buscador en Tiempo Real**: Filtra por producto fitosanitario, número de bloque o sector.
3. **Navegador de Tanques**: En el panel derecho de vista previa, usa los botones `<` / `>` o el menú desplegable para inspeccionar el tanque $1, 2, \dots, N$ o la "Colita".
4. **Número de Copias**: Elige cuántas copias repetidas de cada etiqueta deseas imprimir antes de pulsar el botón de impresión.
5. **Recargar Excel**: Si modificas las fechas, dosis o productos en tu archivo Excel, haz clic en **"Recargar Excel"** para actualizar la aplicación al instante sin reiniciar.

---

## 🖨️ Opciones de Impresión y Fotocopiado

| Modo de Disposición | Uso Recomendado | Características |
| :--- | :--- | :--- |
| **📜 1 Columna Continua** | Fichas bibliográficas / Fotocopias | Imprime las etiquetas verticalmente una tras otra separadas por 6 mm (idéntico a la hoja `imprimir` de Excel). |
| **📑 2 Columnas** | Máximo Ahorro de Papel | Organiza 2 columnas paralelas en hoja Carta/A4 (entran de **4 a 6 etiquetas por página** para fotocopiar y recortar). |
| **📄 1 por Hoja** | Impresoras Térmicas | Salto de página estricto por cada etiqueta individual (rollos adhesivos de 100 mm). |
| **📑 2 por Hoja** | Formato Media Carta | 2 etiquetas por hoja con salto de página par. |

---

## 🚀 ¿Qué le falta a la aplicación para ser perfecta? (Roadmap y Mejoras)

Para llevar este sistema al siguiente nivel de madurez operativa y convertirlo en una solución agronómica de clase empresarial, se recomiendan las siguientes mejoras:

### 1. 🤖 Autocompletado e Inferencia Inteligente de Pictogramas
- **Problema actual**: En `Base.xlsx`, el 72.8% de los productos tienen escrito `SIN FOTO`.
- **Solución propuesta**: Implementar un recomendador inteligente que, al detectar un producto sin pictograma, analice las Frases H (ej. si contiene *H318 Provoca lesiones oculares graves*, sugerir automáticamente `GHS05 Corrosión`; si contiene *H410 Tóxico para la vida acuática*, sugerir `GHS09 Medio ambiente`).

### 2. ✏️ Editor Rápido Bidireccional desde la Web
- Permitir que el usuario edite o asigne pictogramas, frases H/P o dosis faltantes directamente desde una ventana emergente en la web y guardar los cambios de vuelta en el archivo `Base.xlsx` sin necesidad de abrir Microsoft Excel manualmente.

### 3. 📱 Generación de Códigos QR Dinámicos en la Etiqueta
- Incorporar en una celda de la etiqueta un código QR escaneable que redirija al operario o aplicador en campo a la **Ficha de Datos de Seguridad (FDS / HDS)** en PDF o a la guía de primeros auxilios del producto.

### 4. 📄 Exportación Directa a PDF Vectorial Multietiqueta
- Añadir un botón de **"Descargar PDF"** generado en el backend (vía `WeasyPrint` o `ReportLab`) con plantillas estándar de etiquetas adhesivas comerciales (ej. marcas Avery, Formacol o rollos térmicos Zebra/Zijiang).

### 5. 📊 Registro de Auditoría y Trazabilidad (Certificaciones GlobalGAP / ICA)
- Base de datos SQLite local para guardar el historial de órdenes impresas:
  - *¿Quién generó la mezcla?*
  - *¿A qué hora se imprimieron las etiquetas?*
  - *¿Cuántos tanques fueron despachados a cada sector?*

### 6. 📶 Acceso en Red Local para Dispositivos Móviles (PWA)
- Permitir que los operarios desde una tablet o smartphone en la caseta de mezclas o en el tractor se conecten a `http://IP-LOCAL:8000` con interfaz optimizada para pantallas táctiles y modo offline.

---

## 📄 Licencia
Proyecto desarrollado para la gestión fitosanitaria y cumplimiento de normas SGA / GHS. Uso interno y distribución bajo estándares de la organización.
