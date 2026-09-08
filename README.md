# 🌿 SGA Label Studio Enterprise
### Sistema Inteligente de Dosificación Fitosanitaria, Plantillas Físicas SGA y Trazabilidad ICA / GlobalGAP

![Python](https://img.shields.io/badge/Python-3.10%2B-blue?logo=python&logoColor=white)
![FastAPI](https://img.shields.io/badge/FastAPI-0.100%2B-009688?logo=fastapi&logoColor=white)
![GHS / SGA](https://img.shields.io/badge/Normativa-SGA%20%2F%20GHS%20Dec.%201496-dc2626)
![ICA Colombia](https://img.shields.io/badge/Cumplimiento-ICA%20Res.%2007723-059669)
![GlobalGAP](https://img.shields.io/badge/Auditor%C3%ADa-GlobalGAP%20v6-16a34a)
![ReportLab](https://img.shields.io/badge/PDF-Vectorial%20ReportLab-orange)
![SQLite](https://img.shields.io/badge/Database-SQLite%20Audit-003B57?logo=sqlite&logoColor=white)
![License](https://img.shields.io/badge/Licencia-Uso%20Agron%C3%B3mico%20Empresarial-slate)

---

## 📋 Tabla de Contenidos

1. [Descripción General](#-descripción-general)
2. [Valor Operativo y Cumplimiento Normativo](#-valor-operativo-y-cumplimiento-normativo)
3. [Arquitectura del Sistema y Flujo de Datos](#-arquitectura-del-sistema-y-flujo-de-datos)
4. [Estructura y Fidelidad de la Plantilla Física (A2:F12)](#-estructura-y-fidelidad-de-la-plantilla-física-a2f12)
5. [Módulos y Capacidades Implementadas](#-módulos-y-capacidades-implementadas)
   - [Motor de Desglose de Tanques de 1,000 L y Colitas](#1-motor-de-desglose-de-tanques-de-1000-l-y-colitas)
   - [Motor de Inferencia Inteligente GHS](#2-motor-de-inferencia-inteligente-ghs)
   - [Editor Rápido Bidireccional de Excel](#3-editor-rápido-bidireccional-de-excel)
   - [Exportación Directa a PDF Vectorial](#4-exportación-directa-a-pdf-vectorial)
   - [Historial de Auditoría y Trazabilidad (ICA / GlobalGAP)](#5-historial-de-auditoría-y-trazabilidad-ica--globalgap)
   - [Conectividad en Red Local WiFi (PWA)](#6-conectividad-en-red-local-wifi-pwa)
6. [Guía Interactiva de Botones y Acciones](#-guía-interactiva-de-botones-y-acciones)
7. [Referencia de API REST](#-referencia-de-api-rest)
8. [Modos de Impresión y Disposición en Papel](#-modos-de-impresión-y-disposición-en-papel)
9. [Instalación, Configuración y Puesta en Marcha](#-instalación-configuración-y-puesta-en-marcha)
10. [Batería de Pruebas Automatizadas](#-batería-de-pruebas-automatizadas)

---

## 📖 Descripción General

**SGA Label Studio Enterprise** es una plataforma web técnica de alta precisión diseñada para estaciones de mezclas agronómicas y viveros comerciales. Permite digitalizar, desglosar volumétricamente y generar las etiquetas físicas de seguridad del **Sistema Globalmente Armonizado (SGA / GHS)** a partir de las planificaciones semanales registradas en archivos de Microsoft Excel (`aplicacion.xlsm` y `Base_Actualizada.xlsx`).

El sistema resuelve tres desafíos críticos en campo:
1. **Eliminación del error humano de cálculo** en tanques múltiples y residuos ("colitas") para aplicaciones fitosanitarias de gran escala (ej. lotes de más de 100,000 litros).
2. **Fidelidad absoluta a la plantilla física oficial** (`A2:F12`), compatible con impresión en hojas continuas, fotocopias de alto rendimiento y rollos térmicos de etiquetas adhesivas.
3. **Trazabilidad inmutable de mezclas químicas** para auditorías fitosanitarias exigidas por el **ICA** (Instituto Colombiano Agropecuario) y la certificación internacional **GlobalGAP**.

---

## ⚖️ Valor Operativo y Cumplimiento Normativo

El uso de productos fitosanitarios y agroquímicos está estrictamente regulado en Colombia e internacionalmente:

- **Decreto 1496 de 2018 (Ministerio de Trabajo / Salud)**: Adopta el Sistema Globalmente Armonizado de Clasificación y Etiquetado de Productos Químicos (SGA). Exige la identificación visible de peligros (diamantes rojos con pictogramas GHS), palabras de advertencia (`PELIGRO` o `ATENCIÓN`), indicaciones de peligro (Frases H) y consejos de prudencia (Frases P).
- **Resolución ICA 07723 de 2020**: Reglamenta las Buenas Prácticas Agrícolas (BPA) y exige registros detallados de cada aplicación de plaguicidas (producto, dosis, cultivo, fecha, número de lote, operario y periodo de reentrada).
- **Norma GlobalGAP IFA (Integrated Farm Assurance) v6**: Demanda trazabilidad documental completa desde la formulación de la mezcla en caseta hasta la aplicación en el bloque, incluyendo justificación técnica y registro de despacho por tanque.

---

## 🏗️ Arquitectura del Sistema y Flujo de Datos

El sistema opera bajo una arquitectura desacoplada y reactiva en Python con FastAPI y Vanilla JavaScript ES6+:

```mermaid
graph TD
    subgraph Archivos de Entrada (Excel)
        A1[aplicacion.xlsm<br/>Hojas: Data, ALZ, R-S-L]
        A2[Base_Actualizada.xlsx<br/>Hoja: SGA - 228 Productos]
        A3[static/picto/<br/>GHS01.jpg a GHS09.jpg]
    end

    subgraph Backend FastAPI
        B1[DataManager / io.BytesIO<br/>Lectura Compartida No Bloqueante]
        B2[GHSEngine<br/>Inferencia NLP Frases H]
        B3[TankLabelGenerator<br/>Desglose 1,000L + Colitas]
        B4[ExcelWriter<br/>Persistencia Bidireccional + .bak]
        B5[AuditLogger<br/>SQLite var/audit.db]
        B6[PDFGenerator<br/>ReportLab Carta / Térmico]
    end

    subgraph Frontend Web SaaS
        C1[Barra de Filtros y Búsqueda<br/>Spotlight Search + Chips Fechas]
        C2[Selector de Programas<br/>Data, ALZ, R-S-L]
        C3[Lista de Aplicaciones<br/>Selección Múltiple y Lotes]
        C4[Visor de Plantilla Excel A2:F12<br/>Efecto Papel Físico]
        C5[Modales Interactivos<br/>Editor SGA, Auditoría, PDF, Red Local]
    end

    A1 -->|Lectura en Memoria| B1
    A2 -->|Lectura en Memoria| B1
    A3 -->|Imágenes| B1
    B1 --> B2
    B1 --> B3
    B2 --> C1
    B3 --> C3
    B3 --> C4
    C5 -->|Actualizar Ficha| B4
    B4 -->|Escribe y Respalda| A2
    C4 -->|Imprimir / Exportar| B5
    C5 -->|Generar PDF| B6
```

---

## 🏷️ Estructura y Fidelidad de la Plantilla Física (A2:F12)

La etiqueta generada en pantalla, en impresión del navegador y en PDF vectorial es una **réplica exacta al 100%** de la plantilla física de Excel:

| Fila Excel | Columnas A - D | Columnas E - F |
| :--- | :--- | :--- |
| **Fila 1 (A2:F2)** | **PRODUCTO**: Nombre comercial en mayúsculas (negrita, 12 pt). | **PALABRA DE ADVERTENCIA**: `PELIGRO` (rojo) o `ATENCIÓN` (ámbar). |
| **Fila 2 (A3:F3)** | **BLOQUE**: Sector/bloque agronómico.<br/>**FECHA APLICACIÓN**: Fecha programada en formato `DD/MM/AAAA`. | **PICTOGRAMA 1 y PICTOGRAMA 2**: Rombos GHS normativos (dimensión fija de 52&times;52 px con borde negro sólido). |
| **Fila 3 (A4:F4)** | **REENTRADA**: Intervalo de seguridad en horas (ej. `12`).<br/>**UNIDAD**: Unidad de medida del insumo (`LITRO`, `KILO`, `CC`). | *(Abarcado por el rowspan de los pictogramas superiores)* |
| **Fila 4 (A5:F5)** | **CATEGORIA**: Categoría toxicológica (`I`, `II`, `III`, `IV`).<br/>**CANTIDAD**: Cantidad exacta calculada para este tanque específico ($1,000 \times \text{Dosis}$). | *(Abarcado por el rowspan de los pictogramas superiores)* |
| **Fila 5 (A6:F6)** | **VOL. TANQUE**: Volumen de este tanque (ej. `1000 L` o volumen remanente).<br/>**ETIQUETA**: Identificador (`Tanque 1`, `Tanque 2`, o `Colita`). | **PICTOGRAMA 3 y PICTOGRAMA 4**: Rombos GHS normativos secundarios (o rombo punteado en blanco si no aplica). |
| **Filas 6-7 (A7:D8)** | Celdas espaciadoras limpias originales sin saturaciones gráficas. | *(Abarcado por el rowspan de los pictogramas inferiores)* |
| **Fila 8 (A9:F9)** | **FRASE H**: Encabezado de sección destacado en gris institucional (`#f1f5f9`). | *(Abarca las 6 columnas A-F)* |
| **Fila 9 (A10:F10)** | **Texto completo de Indicaciones de Peligro (Frases H)**: Renderizado sin cortes, con ajuste automático de línea y códigos normativos legibles. | *(Abarca las 6 columnas A-F)* |
| **Fila 10 (A11:F11)** | **FRASE P**: Encabezado de sección para Consejos de Prudencia. | *(Abarca las 6 columnas A-F)* |
| **Fila 11 (A12:F12)** | **Texto completo de Consejos de Prudencia (Frases P)**: Instrucciones obligatorias de manejo, EPP, almacenamiento y disposición final de envases. | *(Abarca las 6 columnas A-F)* |

---

## ⚙️ Módulos y Capacidades Implementadas

### 1. Motor de Desglose de Tanques de 1,000 L y Colitas
Implementado en `backend/data_loader.py` (`generate_tank_labels`):
- Para lotes mayores o iguales a 1,000 L:
  $$\text{Tanques Completos} = \left\lfloor \frac{\text{Litros Totales}}{1000} \right\rfloor$$
  Cada tanque recibe exactamente $1,000\text{ L}$ y una cantidad a dosificar de $1,000 \times \text{Dosis}$.
- Si existe residuo ($\text{Litros Totales} \bmod 1000 > 0$), se genera automáticamente una etiqueta denominada **"Colita"** con el volumen remanente exacto y la dosis proporcional calculada.
- Para aplicaciones menores a 1,000 L, genera una etiqueta única con el volumen exacto configurado.

### 2. Motor de Inferencia Inteligente GHS
Implementado en `backend/ghs_engine.py`:
- Analiza mediante expresiones regulares las Frases H del producto químico.
- Rescata automáticamente productos que en el catálogo decían `SIN FOTO` y les asigna sus pictogramas oficiales:
  - $H318 \rightarrow$ `GHS05` (Corrosión / Lesiones oculares graves).
  - $H410, H400 \rightarrow$ `GHS09` (Medio ambiente / Toxicidad acuática).
  - $H300, H310, H330 \rightarrow$ `GHS06` (Toxicidad aguda severa).
  - $H315, H319, H335 \rightarrow$ `GHS07` (Irritación cutánea / ocular).
  - $H350, H360 \rightarrow$ `GHS08` (Peligro para la salud / Carcinogenicidad).
- Muestra el distintivo visual `🤖 Sugerido GHS` en la tarjeta y permite al usuario confirmar o editar los pictogramas en 1 clic.

### 3. Editor Rápido Bidireccional de Excel
Implementado en `backend/excel_writer.py` y endpoint `POST /api/products/update`:
- Permite hacer clic en **"Editar SGA"** sobre cualquier producto en pantalla.
- Abre una ventana con selector interactivo de los 9 diamantes GHS, modificador de Palabra de Advertencia (`PELIGRO`, `ATENCIÓN`), Unidad de Medida (`LITRO`, `KILO`, `GRAMO`, `CC`) y edición de Frases H/P.
- Al guardar:
  1. Crea una copia de seguridad automática (`Base_Actualizada.xlsx.bak`).
  2. Escribe los datos en el archivo Excel conservando fórmulas y estilos de celda.
  3. Registra el evento en la auditoría SQLite.
  4. Recarga los datos en memoria en caliente sin necesidad de reiniciar la aplicación.

### 4. Exportación Directa a PDF Vectorial
Implementado en `backend/pdf_generator.py` con **ReportLab**:
- **Hoja Carta / A4**: 2 etiquetas por página con línea de corte segmentada, listas para imprimir en papel adhesivo comercial.
- **Rollo Térmico Continuo (100 &times; 150 mm / 4" &times; 6")**: Formato estándar para impresoras industriales térmicas directas (Zebra ZD220/ZD421, TSC, Xprinter, Honeywell).
- Generación de PDF tanto individual como en lote de múltiples aplicaciones seleccionadas.

### 5. Historial de Auditoría y Trazabilidad (ICA / GlobalGAP)
Implementado en `backend/audit_logger.py` (`var/audit.db`):
- Base de datos relacional SQLite que registra automáticamente:
  - Marca de tiempo exacta (`AAAA-MM-DD HH:MM:SS`).
  - Nombre del operario responsable de la mezcla.
  - Acción realizada (`IMPRESION_WEB`, `DESCARGA_PDF`, `EDICION_PRODUCTO`).
  - Programa de cultivo, producto fitosanitario, sector/bloque, dosis, volumen y copias impresas.
- Visor con buscador en tiempo real dentro del modal **"Auditoría ICA"**.
- Botón **"Exportar CSV (Auditoría ICA)"** que descarga la sábana de datos exigida por los auditores fitosanitarios.

### 6. Conectividad en Red Local WiFi (PWA)
Implementado en `app.py` y `manifest.json`:
- El servidor Uvicorn se enlaza a `0.0.0.0:8000` detectando la dirección IP local de la máquina (ej. `http://10.10.12.192:8000`).
- Modal **"Móvil / Red"** que genera un código QR escaneable con la cámara de cualquier smartphone o tablet para operar desde la caseta de preparación de mezclas o en el tractor.
- Soporte Progressive Web App (PWA) para instalación directa en la pantalla de inicio del dispositivo móvil.

---

## 🎯 Guía Interactiva de Botones y Acciones

Cada botón en la interfaz posee una acción real, verificada y con respuesta visual inmediata:

| Botón / Control | Ubicación | Acción Real Ejecutada |
| :--- | :--- | :--- |
| **`Móvil / Red`** | Barra Superior | Abre el modal de conexión LAN, consulta `/api/network-info`, muestra la IP local y el código QR de conexión. |
| **`Copiar Enlace`** | Modal Red | Copia la URL de acceso local al portapapeles con confirmación visual (toast) y soporte de fallback. |
| **`Auditoría ICA`** | Barra Superior | Abre el modal de trazabilidad, consulta `/api/audit/logs` y carga el historial en tabla paginada. |
| **`Exportar CSV`** | Modal Auditoría | Descarga directa del archivo `auditoria_fitosanitaria_sga.csv` con todos los despachos registrados. |
| **`Refrescar Auditoría`** | Modal Auditoría | Recarga el registro de eventos en tiempo real con animación de rotación en el icono. |
| **`Recargar`** | Barra Superior | Llama a `POST /api/reload`, lee nuevamente los archivos Excel en disco y refresca la interfaz sin F5. |
| **`Configuración (📁)`** | Barra Superior | Abre el modal de configuración de rutas Excel y preferencias de impresión, guardando cambios en `localStorage` y en el servidor. |
| **Pestañas de Programa** | Panel Izquierdo | Alterna entre las hojas de Excel (`Data`, `ALZ`, `R-S-L`), actualizando los lotes y el catálogo en tiempo real. |
| **Chips de Fecha** | Panel Izquierdo | Filtra los lotes programados para el día seleccionado y actualiza los contadores de tanques. |
| **`Mostrar todas`** | Panel Izquierdo | Limpia el filtro por fecha y muestra la programación semanal completa. |
| **`Buscador Inteligente`** | Panel Izquierdo | Filtrado instantáneo por producto, código ICA, bloque o dosis con atajos de teclado (`/` para enfocar, `Esc` para limpiar). |
| **`Seleccionar todos`** | Panel Izquierdo | Marca o desmarca todas las casillas de la lista y despliega los botones de impresión en lote. |
| **`Ver Plantilla Excel`** | Tarjeta de Lote | Selecciona la aplicación, desplaza la pantalla suavemente hacia la vista previa y resalta la etiqueta en verde esmeralda. |
| **`Editar SGA`** | Vista Previa | Abre el editor modal para actualizar pictogramas, advertencia, unidad y frases H/P de ese producto en `Base_Actualizada.xlsx`. |
| **`PDF` (Individual)** | Vista Previa | Abre el modal para descargar la etiqueta activa en formato vectorial (Carta o Rollo Térmico). |
| **`PDF` (En Lote)** | Barra de Lotes | Abre el modal para exportar todas las aplicaciones seleccionadas en un solo archivo PDF consolidado. |
| **Controles `−` y `+`** | Vista Previa | Ajustan el número de copias físicas a imprimir (1 a 500 copias por tanque). |
| **`Imprimir (N)`** | Vista Previa | Compila el documento de impresión exacto con formato A2:F12 y abre el cuadro de diálogo de la impresora (`window.print()`). |
| **Selectores `<` y `>`** | Vista Previa | Navegan entre los distintos tanques del lote (Tanque 1, Tanque 2, Colita), deshabilitándose en los extremos. |

---

## 📡 Referencia de API REST

| Método | Endpoint | Parámetros / Cuerpo | Descripción |
| :---: | :--- | :--- | :--- |
| `GET` | `/api/summary` | Ninguno | Retorna estadísticas generales, programa activo, fechas y totales. |
| `GET` | `/api/applications` | `date` *(opcional)* | Retorna las aplicaciones fitosanitarias desglosadas en tanques. |
| `POST` | `/api/set-program` | `{"program": "ALZ"}` | Cambia la hoja activa de `aplicacion.xlsm` (`Data`, `ALZ`, `R-S-L`). |
| `POST` | `/api/reload` | Ninguno | Relee los libros Excel en caliente sin reiniciar el servidor. |
| `POST` | `/api/products/update` | JSON con código, frases H/P, pictogramas | Actualiza un producto en `Base_Actualizada.xlsx` con copia de seguridad. |
| `POST` | `/api/export-pdf` | JSON con IDs, layout (`letter`/`thermal`), copias | Genera y descarga el archivo PDF vectorial con ReportLab. |
| `GET` | `/api/audit/logs` | `search`, `limit`, `offset` | Retorna los registros de auditoría en formato JSON. |
| `GET` | `/api/audit/export` | Ninguno | Descarga el archivo CSV oficial para inspecciones del ICA. |
| `GET` | `/api/network-info` | Ninguno | Informa la IP local de la máquina y la URL del código QR de conexión. |
| `GET` | `/ficha/{codigo}` | `codigo` en ruta | Visualiza la Ficha Digital de Emergencia FDS / Primeros Auxilios en móvil. |
| `GET` | `/picto/{filename}` | Nombre de imagen en ruta | Entrega los rombos oficiales GHS en formato JPG. |

---

## 🖨️ Modos de Impresión y Disposición en Papel

El sistema integra reglas `@media print` de alta precisión para diversos escenarios de producción:

1. **📜 1 Columna Continua (Fiel a Excel)**:
   - Imprime las etiquetas verticalmente separadas por 6 mm.
   - Idéntico a la hoja `imprimir` original de Excel. Ideal para fichas bibliográficas en cartulina o papel bond.
2. **📑 2 Columnas (Máximo Ahorro de Papel)**:
   - Organiza 2 columnas paralelas en hoja Carta / A4.
   - Entran de **4 a 6 etiquetas completas por hoja**, perfecto para fotocopias masivas y recorte manual.
3. **📄 1 Etiqueta por Hoja (Rollo Térmico Continuo)**:
   - Formato ajustado con salto de página estricto (`page-break-after: always`).
   - Diseñado para rollos adhesivos de 100 mm en impresoras térmicas de transferencia directa.
4. **📑 2 Etiquetas por Hoja (Formato Media Carta)**:
   - Formato con salto de página par para imprentas o guillotinas de medio pliego.

---

## 🚀 Instalación, Configuración y Puesta en Marcha

### Requisitos del Sistema
- Python 3.9 o superior (probado en Python 3.10, 3.11, 3.12 y 3.13).
- Windows 10/11, Linux o macOS.
- Navegador moderno (Google Chrome, Microsoft Edge, Firefox, Safari).

### Datos de entrada
El catálogo incluido vive en `data/catalog/`. Guarde `aplicacion.xlsm` en la carpeta operativa de cada finca y defina `SGA_DATA_DIR` con esa ruta antes de iniciar el servidor. Si la carpeta operativa también contiene `Picto/`, sus imágenes tienen prioridad sobre las incluidas en `static/picto/`.

En PowerShell:
```powershell
$env:SGA_DATA_DIR = 'C:\ruta\a\datos-operativos'
```

### Paso 1: Iniciar el servidor
**En Windows (Doble Clic):**
Ejecuta el archivo **`run_app.bat`** (o **`ABRIR_ETIQUETAS.bat`**).

**Desde la Terminal:**
```bash
python -m uvicorn app:app --host 0.0.0.0 --port 8000
```

Abre tu navegador en:
👉 **[http://localhost:8000](http://localhost:8000)** (o la IP local indicada en la consola para dispositivos móviles).

---

## 🧪 Batería de Pruebas Automatizadas

El proyecto cuenta con una suite integral de pruebas unitarias en `tests/test_app.py`:

```bash
python -m unittest tests.test_app
```

### Resultados de Verificación:
```text
..........
----------------------------------------------------------------------
Ran 10 tests in 5.780s

OK
[TEST PASS] Endpoints avanzados validados (Red Local, QR PNG, Ficha HTML, Auditoría, PDF Export).
[TEST PASS] Registro y trazabilidad en SQLite verificado.
[TEST PASS] Endpoints FastAPI para cambio de programa validados con TestClient.
[TEST PASS] Inferencia GHS verificada: H318->GHS05, H410->GHS09: ['GHS05', 'GHS09']
[TEST PASS] Generación de PDF vectorial (Carta y Rollo Térmico) verificada.
[TEST PASS] Pictogramas para ABAMECAL resueltos: ['GHS06.jpg', 'GHS07.jpg', 'GHS09.jpg']
[TEST PASS] Soporte multi-programa verificado (Data: 90, ALZ: 124, R-S-L: 134).
[TEST PASS] Catálogo cargado con 90 apps y 90 etiquetas.
[TEST PASS] Regla de 162,600 L desglose exacto (162 tanques 1,000L + 1 colita 600L).
[TEST PASS] Regla < 1,000 L genera etiqueta única con volumen exacto.
```

---

## 📄 Licencia y Derechos
Desarrollado para la gestión agronómica, inocuidad química y cumplimiento de normativas fitosanitarias internacionales. Uso bajo estándares de la organización.
