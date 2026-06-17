# Deploy y operación

## Arquitectura

```
┌────────────────────────────────────────────────────┐
│                  PC de desarrollo                    │
│  (tu máquina — C:\impresoras\)                      │
│                                                     │
│  • Código fuente + git                              │
│  • Entorno virtual (.venv)                          │
│  • Ejecuta: python impresoras.py                    │
│  • Compila el .exe para distribuir                  │
└───────────────────────┬────────────────────────────┘
                        │ git push / pull
                        ▼
              ┌─────────────────┐
              │   GitHub repo   │
              │ edfsosa/impresoras │
              └────────┬────────┘
                       │
          ┌────────────┴────────────┐
          ▼                         ▼
┌──────────────────┐   ┌──────────────────────┐
│ PC usuarios final │   │ PC usuarios final    │
│ (sin Python)      │   │ (sin Python)         │
│                   │   │                      │
│ • MonitorImp.exe  │   │ • MonitorImp.exe     │
│ • Misma BD en red │   │ • Misma BD en red    │
└──────────────────┘   └──────────────────────┘
          │                         │
          └──────────┬──────────────┘
                     ▼
        ┌────────────────────────┐
        │  Base de datos común   │
        │  \\servidor\share\     │
        │  impresoras.db         │
        └────────────────────────┘
```

### Roles

| Equipo | Rol | Tiene Python | Usa git |
|--------|-----|-------------|---------|
| Tu PC | Desarrollo y distribución | Sí | Sí |
| PC usuarios finales | Solo ejecutar la app | No | No |

### Base de datos compartida

Todas las PCs apuntan a la misma base de datos en una carpeta de red. Esto permite que cualquier usuario vea los mismos datos (impresoras, monitoreos, stock, envíos). La ruta se configura en **Configuración → Base de datos** y se guarda en `config.json`.

---

## Setup inicial en tu PC (ya está hecho)

```powershell
git clone https://github.com/edfsosa/impresoras.git
cd C:\impresoras
python -m venv .venv
.venv\Scripts\activate
pip install -r requirements.txt
pip install pyinstaller   # solo si vas a compilar .exe
```

---

## Uso diario (tu PC)

Ejecutar la app desde el código fuente (sin compilar):

```powershell
cd C:\impresoras
.venv\Scripts\activate
python impresoras.py
```

### Acceso directo recomendado

1. Crear un archivo `MonitorImpresoras.bat` en `C:\impresoras\` con:
   ```batch
   @echo off
   cd /d C:\impresoras
   .venv\Scripts\python.exe impresoras.py
   ```
2. Crear un acceso directo a ese `.bat` en el escritorio.

### Ventajas de ejecutar como .py

- Los cambios de `git pull` se reflejan al instante (sin recompilar)
- Podés ver la salida de la consola (errores, logs)

---

## Actualizar tu PC con nuevos cambios

```powershell
cd C:\impresoras
git pull
.venv\Scripts\activate
pip install -r requirements.txt   # por si cambiaron las dependencias
```

Listo. La próxima vez que abras la app ya tiene los cambios.

---

## Distribuir a usuarios finales (sin Python)

### Compilar el .exe

Después de actualizar el código y probar que funciona:

```powershell
cd C:\impresoras
.venv\Scripts\activate
pyinstaller MonitorImpresoras.spec
```

Esto genera `C:\impresoras\dist\MonitorImpresoras.exe`.

### Copiar a cada PC de usuario

1. Crear una carpeta compartida o usar un USB
2. Copiar todo `C:\impresoras\dist\` a la PC destino (ej: `C:\Program Files\MonitorImpresoras\`)
3. Opcionalmente crear un acceso directo en el escritorio apuntando al `.exe`
4. Configurar la ruta de base de datos compartida desde la ventana Configuración (si no se usó el mismo `config.json`)

> **Nota:** Si la PC destino ya tiene una versión anterior, solo reemplazar el `.exe` (los datos están en la BD compartida, no en el .exe).

---

## Compilación con PyInstaller

El archivo `MonitorImpresoras.spec` ya está configurado para incluir las dependencias necesarias. Si necesitás crear el spec desde cero:

```powershell
pyinstaller --onefile --windowed --name MonitorImpresoras ^
  --add-data "config.json;." ^
  impresoras.py
```

El flag `--windowed` evita que se abra una ventana de consola al ejecutar el `.exe`.

---

## Estructura de carpetas en PC de usuario final

```
C:\MonitorImpresoras\          (o donde se haya copiado)
├── MonitorImpresoras.exe      # Aplicación compilada
├── config.json                # Configuración (se crea al abrir)
├── impresoras.db              # Base de datos local (si no usa BD compartida)
└── errores.log                # Registro de errores
```

Si usa base de datos compartida, `config.json` tendrá la ruta de red y `impresoras.db` local puede no existir o ser otra.

---

## Monitoreo automático 24/7

Actualmente la app se usa bajo demanda (se abre, se consulta, se cierra). Si en el futuro se desea monitoreo automático continuo:

1. Dedicar una PC (puede ser un mini PC, notebook vieja, o servidor) que quede encendida 24/7
2. Instalar el `.exe` en esa máquina
3. Configurar la base de datos compartida
4. Activar el check **Auto** y seleccionar el intervalo deseado
5. La PC debe permanecer encendida y con la app abierta

> No hay forma de que el monitoreo automático funcione si la PC se apaga. Es una aplicación de escritorio, no un servicio de Windows.
