# Monitor de Impresoras

Aplicación de escritorio para el monitoreo centralizado de impresoras de red, gestión de consumibles (tóner, unidad imagen), control de stock en depósito y estadísticas de consumo por sucursal.

## Características

- **Monitoreo automático** — Consulta HTTP simultánea a todas las impresoras activas, con detección de niveles bajos/medios de tóner, unidad de imagen y kit de mantenimiento. Monitoreo automático programable con intervalos configurables.
- **Catálogo de impresoras** — CRUD completo con campos: IP, modelo, sucursal, nombre, número de serie, ubicación. Vista con 8 columnas, filtros por modelo/sucursal/estado, ordenamiento y exportación a Excel.
- **Historial de monitoreos** — Visualización histórica con filtros por sucursal, modelo, IP y nivel de alerta. Ordenamiento por columna, paginación (200 registros) y modo vista árbol (fecha → registros). Doble clic para ver gráfico de tendencia.
- **Envío de insumos** — Registro de envíos de tóner y unidad de imagen a cada sucursal, con descuento automático del stock. Anulación y edición de envíos con ajuste de stock.
- **Stock de depósito** — Gestión de inventario con alertas de stock crítico/bajo. Entradas, salidas y ajustes. Exportación a Excel. Paginación en historial de movimientos.
- **Estadísticas de consumo** — Gráfico de barras apiladas por sucursal con filtros por fecha y tipo de insumo. Tabla resumen con ordenamiento y exportación. Detalle de envíos por sucursal al seleccionar una fila.
- **Configuración** — Ventana con pestañas para base de datos compartida (ruta de red), notificaciones por email (SMTP con STARTTLS) y monitoreo (umbrales, hilos simultáneos, intervalo por defecto).
- **Notificaciones por email** — Alertas automáticas cuando se detectan impresoras con nivel bajo durante el monitoreo automático.

## Requisitos

- Python 3.13 o superior
- Red local con acceso HTTP a las impresoras (interfaz web de estado)
- Impresoras compatibles: aquellas que exponen la página `/cgi-bin/dynamic/printer/PrinterStatus.html` con datos de consumibles en formato HTML

### Dependencias

Ver `requirements.txt`:

```
requests
beautifulsoup4
openpyxl
matplotlib
numpy
```

## Instalación

1. Clonar el repositorio:
   ```bash
   git clone https://github.com/edfsosa/impresoras.git
   cd impresoras
   ```

2. Crear y activar un entorno virtual:
   ```bash
   python -m venv .venv
   .venv\Scripts\activate  # Windows
   ```

3. Instalar dependencias:
   ```bash
   pip install -r requirements.txt
   ```

4. Ejecutar:
   ```bash
   python impresoras.py
   ```

La base de datos SQLite (`impresoras.db`) se crea automáticamente en el directorio de la aplicación al iniciar.

## Configuración

### Base de datos compartida

Se puede configurar una ruta de red para compartir la base de datos entre varios equipos. Ir a **Configuración → Base de datos** y seleccionar la ruta (ej: `\\servidor\share\impresoras.db`). Los cambios se aplican al reiniciar la aplicación.

### Umbrales de alerta

En la pantalla principal, ajustar los porcentajes de nivel bajo y medio. Se guardan automáticamente en `config.json`.

### Notificaciones por email

En **Configuración → Correo electrónico**, habilitar las notificaciones y completar:
- Remitente (dirección SMTP)
- Contraseña
- Destinatarios (separados por coma)
- Servidor SMTP y puerto (por defecto `smtp.office365.com:587` con STARTTLS)

Usar el botón **Probar envío** para verificar la configuración.

### Monitoreo automático

Activar el check **Auto** en la pantalla principal y seleccionar el intervalo. El intervalo por defecto se puede configurar en **Configuración → Monitoreo**.

## Uso

1. Agregar impresoras desde el botón **Impresoras** (catálogo).
2. Opcionalmente, registrar stock inicial de insumos desde **Stock Depósito**.
3. Ejecutar **Iniciar Monitoreo** para consultar todas las impresoras.
4. Los resultados se muestran en la tabla principal con códigos de color:
   - **Rojo**: nivel bajo (por debajo del umbral)
   - **Amarillo**: nivel medio
   - **Gris**: sin datos (impresora no respondió)
5. Usar los filtros (búsqueda, sucursal, modelo, solo alertas) para segmentar la vista.
6. Exportar los resultados a Excel con el botón **Exportar Excel**.
7. Hacer clic derecho sobre una fila para ver el gráfico histórico, copiar la IP o abrir la impresora en el catálogo.

### Atajos de teclado

| Tecla | Acción |
|-------|--------|
| `F5` o `Ctrl+R` | Iniciar monitoreo |
| `Ctrl+F` | Enfocar búsqueda |
| `Ctrl+E` | Exportar a Excel |
| `Ctrl+G` | Ver gráfico de la fila seleccionada |
| `Escape` | Cancelar monitoreo en curso |

## Estructura del proyecto

```
impresoras/
├── impresoras.py        # Aplicación principal (~4000 líneas)
├── config.json          # Configuración persistente (umbrales, email, DB)
├── requirements.txt     # Dependencias Python
├── impresoras.db        # Base de datos SQLite (se crea al iniciar)
├── errores.log          # Registro de errores
├── MonitorImpresoras.spec  # Archivo PyInstaller para compilar .exe
├── build/               # Directorio de compilación PyInstaller
└── dist/                # Ejecutable compilado
```

### Base de datos

Tablas principales:
- `impresoras` — Catálogo de impresoras
- `monitoreos` — Lecturas históricas de consumibles
- `envios` — Registro de envíos de insumos a sucursales
- `stock_deposito` — Inventario actual de insumos en depósito
- `movimientos_stock` — Auditoría de entradas, salidas y ajustes de stock
- `modelos` — Modelos de impresoras normalizados
