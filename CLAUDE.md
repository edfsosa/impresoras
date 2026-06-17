# Monitor de Impresoras — CLAUDE.md

## Convenciones de código

- **No agregar comentarios** en el código fuente a menos que el usuario lo pida explícitamente.
- Nombres de funciones en español con snake_case: `db_impresoras_todas`, `abrir_historial`, `_cargar_stock`.
- Constantes en MAYÚSCULA: `BG_MAIN`, `COLOR_BAJO`, `TIPOS_INSUMO`, `FONT_UI`.
- Funciones privadas (internas de una ventana) con prefijo `_`: `_dibujar`, `_cargar_detalle`.
- Variables tkinter con prefijo según tipo: `var_*` (StringVar/BooleanVar), `btn_*`, `entry_*`, `tree_*`, `combo_*`, `lbl_*`, `frame_*`, `spin_*`.

## Proyecto

- Aplicación Python 3.13 + Tkinter para monitoreo de impresoras de red.
- BD SQLite con WAL. Migraciones automáticas en `init_db()`.
- `init_db()` se llama al importar el módulo (línea ~4185), lo que impide importar directamente desde scripts de prueba.
- Archivo principal: `impresoras.py` (~4200 líneas). No modularizar a menos que el usuario lo pida.

## Base de datos

- `DB_PATH` global, configurable desde `config.json` clave `db_path`.
- Usar siempre `with db_connect() as conn:` para acceso a BD (context manager con commit/rollback automático).
- Funciones DB con prefijo `db_*` en el mismo archivo.
- `sqlite3.Row` para acceso por nombre de columna.

## Tablas principales

| Tabla | Propósito |
|-------|-----------|
| `impresoras` | Catálogo de impresoras (ip, modelo, sucursal, nombre, sn, activa, ubicacion, modelo_id) |
| `monitoreos` | Lecturas históricas de consumibles (fecha, ip, toner, unidad_imagen, kit_mantenimiento) |
| `envios` | Envíos de insumos a sucursales (fecha, sucursal, ip, tipo_insumo, modelo_impresora, cantidad, anulado) |
| `stock_deposito` | Inventario de insumos (tipo_insumo, modelo_impresora, cantidad, stock_minimo) |
| `movimientos_stock` | Auditoría de movimientos (fecha, tipo, tipo_insumo, modelo_impresora, cantidad, observacion, envio_id) |
| `modelos` | Modelos normalizados (id, nombre) |

## Tuplas de datos

- `envios`: tuplas de 8 elementos `(id, fecha_str, sucursal, ip, tipo, modelo, cantidad, anulado)`.
- `filas_tabla` (monitoreo): tuplas de 8 elementos `(sucursal, ip, modelo, ult_monitoreo, toner_str, unidad_str, kit_str, tag)` donde tag es `"bajo"`, `"medio"`, `"sin_datos"` o `""`.

## Testing

- Tests unitarios en `C:\Users\Amp51463\AppData\Local\Temp\opencode\test_envios.py`
- Ejecutar con: `python -X utf8 test_envios.py`
- Usa copia aislada de funciones DB (no modifica la BD real).
- Los tests cubren: registrar envío, agregar stock, anular envío, re-anulación, editar cantidad, filtros por estado.

## UI

- Ventanas secundarias con `tk.Toplevel()`, modales con `grab_set()`.
- Botones estilizados con `_estilo_btn(btn, primario=True/False)`.
- No usar `ttk.Notebook` sin tener `import tkinter.ttk as ttk`.
- Filtros automáticos con `trace_add("write", lambda *_: funcion)`.
- Ordenamiento por columna con función `ordenar_por_columna(ctx, col_idx)`.
- Paginación client-side con corte de lista `data[offset:offset+page_size]`.
- Exportar Excel con `openpyxl`, patrón: `Workbook()`, encabezados con `PatternFill`, guardar con `wb.save()`.

## Ventanas principales

| Función | Ventana |
|---------|---------|
| `crear_interfaz()` | Pantalla principal (Tk root) |
| `abrir_catalogo_impresoras(seleccionar_ip=None)` | Catálogo de impresoras |
| `abrir_historial()` | Historial de monitoreos |
| `abrir_envio_insumos()` | Envío de insumos |
| `abrir_stock_deposito()` | Stock de depósito |
| `abrir_estadisticas_consumo()` | Estadísticas de consumo |
| `abrir_configuracion()` | Configuración (BD, email, monitoreo) |

## Estructura del proyecto

```
C:\impresoras\
├── impresoras.py           # Aplicación principal
├── config.json             # Configuración persistente
├── requirements.txt        # Dependencias
├── MonitorImpresoras.spec  # PyInstaller spec
├── MonitorImpresoras.bat   # Launcher (python impresoras.py)
├── README.md               # Documentación de usuario
├── DEPLOY.md               # Documentación de deploy/operación
├── CLAUDE.md               # Este archivo
├── .opencode/
│   └── opencode.json       # Configuración de opencode
├── dist/
│   └── MonitorImpresoras.exe  # Ejecutable compilado
├── build/                  # Directorio de compilación
└── .venv/                  # Entorno virtual
```
