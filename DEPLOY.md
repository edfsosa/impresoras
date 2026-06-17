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
│  \\MXL8372J8P\impresoras\  │
│  impresoras.db         │
        └────────────────────────┘
```

### Roles

| Equipo | Rol | Tiene Python | Usa git |
|--------|-----|-------------|---------|
| Tu PC | Desarrollo y distribución | Sí | Sí |
| PC usuarios finales | Solo ejecutar la app | No | No |

### Base de datos compartida

Todas las PCs apuntan a la misma base de datos en una carpeta de red. Esto permite que cualquier usuario vea los mismos datos (impresoras, monitoreos, stock, envíos).

| Recurso | Ruta |
|---------|------|
| Base de datos | `\\MXL8372J8P\impresoras\impresoras.db` |
| Ejecutable | `\\MXL8372J8P\impresoras\dist\MonitorImpresoras.exe` |

La ruta de la BD se configura en **Configuración → Base de datos** y se guarda en `config.json`.

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

La estrategia recomendada es dejar el `.exe` en un recurso compartido de red para que todos los usuarios accedan a la misma versión desde cualquier PC.

### Recursos compartidos

| Recurso | Ruta |
|---------|------|
| Carpeta compartida (app) | `\\MXL8372J8P\impresoras` |
| Base de datos compartida | `\\MXL8372J8P\impresoras\impresoras.db` |
| Ejecutable | `\\MXL8372J8P\impresoras\dist\MonitorImpresoras.exe` |

### Paso 1: Compilar el .exe

Después de actualizar el código y probar que funciona:

```powershell
cd C:\impresoras
.venv\Scripts\activate
pyinstaller MonitorImpresoras.spec
```

Esto genera `C:\impresoras\dist\MonitorImpresoras.exe`.

### Paso 2: Copiar a la carpeta compartida

```powershell
Copy-Item "C:\impresoras\dist\MonitorImpresoras.exe" "\\MXL8372J8P\impresoras\dist\" -Force
```

### Paso 3: Que los usuarios creen un acceso directo

En cada PC de usuario:

1. Abrir `\\MXL8372J8P\impresoras\dist\`
2. Clic derecho en `MonitorImpresoras.exe` → **Enviar a → Escritorio (acceso directo)**
3. Abrir la aplicación desde el acceso directo
4. Ir a **Configuración → Base de datos** y verificar que apunte a `\\MXL8372J8P\impresoras\impresoras.db`

### Actualizar a nueva versión

Cuando haya cambios:

```powershell
cd C:\impresoras
git pull
.venv\Scripts\activate
pip install -r requirements.txt
pyinstaller MonitorImpresoras.spec
Copy-Item "C:\impresosas\dist\MonitorImpresoras.exe" "\\MXL8372J8P\impresoras\dist\" -Force
```

Los usuarios no necesitan hacer nada: la próxima vez que abran el acceso directo, Windows cargará la nueva versión desde la red automáticamente.

### Script de actualización automática

Se incluye `actualizar.bat` en la raíz del proyecto. Al ejecutarlo, hace todo en un solo paso:

```batch
actualizar.bat
# 1. git pull
# 2. pip install -r requirements.txt
# 3. pyinstaller MonitorImpresoras.spec
# 4. Copy-Item a \\MXL8372J8P\impresoras\dist\
```

Usar este script para actualizar la app y distribuirla a los usuarios finales.


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

En el recurso compartido de red:

```
\\MXL8372J8P\impresoras\
├── config.json                   # Configuración (db_path, email, umbrales)
├── impresoras.db                 # Base de datos compartida
├── errores.log                   # Registro de errores
├── dist\
│   └── MonitorImpresoras.exe     # Aplicación compilada
└── (otros archivos de desarrollo)
    ├── impresoras.py
    ├── MonitorImpresoras.bat
    └── ...
```

La carpeta compartida contiene tanto el `.exe` como la base de datos. Los usuarios solo necesitan el acceso directo al `.exe` en `dist\`.

Desde la PC de desarrollo, se accede a los mismos archivos para compilar y actualizar.

---

## Backup de la base de datos

La base de datos compartida (`\\MXL8372J8P\impresoras\impresoras.db`) contiene todos los datos del sistema. Se recomienda:

### Backup manual

```powershell
# Copiar la BD a una carpeta de backups con fecha
Copy-Item "\\MXL8372J8P\impresoras\impresoras.db" "D:\backups\impresoras_$(Get-Date -Format yyyyMMdd).db"
```

### Backup automático (programado)

Crear una tarea en el **Programador de tareas de Windows**:

1. Abrir **taskschd.msc**
2. Crear tarea básica: "Backup Monitor Impresoras"
3. Disparador: **Diario** a las 12:00
4. Acción: **Iniciar un programa**
   - Programa: `powershell.exe`
   - Argumentos:
     ```powershell
     Copy-Item "\\MXL8372J8P\impresoras\impresoras.db" "D:\backups\impresoras_$(Get-Date -Format yyyyMMdd_HHmmss).db" -Force
     ```

> La BD usa WAL (Write-Ahead Logging), lo que permite copiarla mientras la aplicación está en uso sin riesgo de corrupción.

---

## Monitoreo automático 24/7

Actualmente la app se usa bajo demanda (se abre, se consulta, se cierra). Si en el futuro se desea monitoreo automático continuo:

1. Dedicar una PC (puede ser un mini PC, notebook vieja, o servidor) que quede encendida 24/7
2. Instalar el `.exe` en esa máquina
3. Configurar la base de datos compartida
4. Activar el check **Auto** y seleccionar el intervalo deseado
5. La PC debe permanecer encendida y con la app abierta

> No hay forma de que el monitoreo automático funcione si la PC se apaga. Es una aplicación de escritorio, no un servicio de Windows.
