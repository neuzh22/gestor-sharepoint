<div align="center">

# 📦 SP-Migrar

### Descarga carpetas de SharePoint Online a tu NAS desde la terminal

Eliges las carpetas y el destino, y lo dejas corriendo toda la noche: si se cae la red, el NAS o la sesión, espera a que vuelvan y continúa.

![Python](https://img.shields.io/badge/Python-3.10%2B-3776AB?logo=python&logoColor=white)
![SharePoint](https://img.shields.io/badge/SharePoint-Online-0078D4?logo=microsoftsharepoint&logoColor=white)
![Microsoft Graph](https://img.shields.io/badge/Microsoft%20Graph-API-5E5E5E?logo=microsoft&logoColor=white)
![Windows](https://img.shields.io/badge/Windows-10%20%7C%2011-0078D6?logo=windows&logoColor=white)

[Características](#-características) ·
[Instalación](#-instalación) ·
[Uso](#-uso) ·
[Manejo de errores](#%EF%B8%8F-manejo-de-errores) ·
[Configuración](#%EF%B8%8F-configuración) ·
[Problemas comunes](#-problemas-comunes)

</div>

---

## 👀 Así se ve

```text
==============================================================================
 ESTÁS EN:  Ventas  >  Documentos     (12 archivo(s) aquí)
==============================================================================
    1) Clientes      (18.40 GB, 1203 elementos)
    2) Facturas      (1.20 GB, 50 elementos)
    3) Proyectos     (4.77 GB, 312 elementos)

  ¿Qué quieres hacer?
    Escribe un NÚMERO        -> abrir esa carpeta y ver lo que tiene
    Escribe D y el número    -> DESCARGAR esa carpeta completa   (ej: D 3   o varias: D 1,4)
    Escribe D sola           -> DESCARGAR la carpeta donde estás ahora
    Escribe A                -> volver Atrás
    Escribe S                -> Salir sin descargar

  Opción: D 1,3
```

```text
17:25:07 INFO    OK  Clientes/2024/Contrato_ACME.pdf  [2.31 MB en 0.6s]
17:25:08 INFO    OK  Clientes/2024/Propuesta.docx  [412.20 KB en 0.3s]
17:26:07 INFO    [ESTADO] OK 1840 | omitidos 0 | errores 0 | en cola 16 | 9.71 GB a 42.30 MB/s | 4m 02s
02:13:44 WARNING Sin conexión a internet (ConnectionError). Esperando a que vuelva la red...
02:31:10 INFO    Conexión restablecida tras 17m 26s. Continuando.
```

## ✨ Características

| | |
|---|---|
| 🖥️ **100% terminal** | El inicio de sesión es con código de dispositivo de Microsoft, así que no hace falta navegador en el equipo. |
| 📂 **Eliges qué copiar** | Recorres bibliotecas y carpetas, ves su tamaño y descargas una o varias. |
| 🎯 **Eliges a dónde** | Puede ser un NAS (`\\NAS\Respaldos`), una unidad de red (`Z:\`) o un disco local. |
| 🔁 **Reanudable** | Si se corta, continúa donde quedó, incluso a mitad de un archivo grande. |
| 🛡️ **A prueba de fallos** | Aguanta caídas de red, del NAS, de la sesión, la limitación de SharePoint y el disco lleno. |
| ⚡ **Rápido** | Hace varias descargas a la vez (4 por defecto, configurable). |
| 🌙 **Desatendido** | Evita que el PC se suspenda y hace pasadas automáticas de reintento. |
| 📊 **Trazable** | Guarda un log detallado y deja un reporte CSV de errores que se abre en Excel. |
| 🗂️ **Fiel al original** | Respeta la estructura de carpetas y las fechas de modificación, y admite rutas de más de 260 caracteres. |

## 🔄 Cómo funciona

```mermaid
flowchart LR
    A[🔐 Inicio de sesión<br/>código de dispositivo] --> B[🌐 URL del sitio]
    B --> C[📂 Elegir carpetas]
    C --> D[🎯 Elegir destino]
    D --> E[⬇️ Descarga en paralelo]
    E --> F{¿Errores?}
    F -- No --> G[✅ Listo]
    F -- Sí --> H[🔁 Pasadas de reintento]
    H --> I[📄 Reporte CSV]
    E -. corte / Ctrl+C .-> J[(💾 Estado SQLite)]
    J -. reanudar .-> E
```

## 📥 Instalación

### Opción A: ejecutable, sin instalar nada (recomendado)

1. Copia **`SP-Migrar.exe`** a una carpeta del equipo, por ejemplo `C:\SP-Migrar\`.
2. Haz doble clic. Listo.

No necesita Python. La configuración, la sesión, los trabajos y los logs se guardan **en la misma carpeta del `.exe`**, así que elige una carpeta donde tengas permiso de escritura (no `C:\Program Files`).

> 🛠️ Para generar el `.exe` tú mismo, ejecuta **`compilar.bat`** en un equipo con Python. El resultado queda en `dist\SP-Migrar.exe`.

> ⚠️ Windows SmartScreen puede advertir *"Windows protegió su PC"*, porque el ejecutable no está firmado. Pulsa **Más información → Ejecutar de todas formas**.

### Opción B: con Python

**Requisitos:** Windows 10/11, [Python 3.10+](https://www.python.org/downloads/) (marca *"Add Python to PATH"* al instalar) y una cuenta de Microsoft 365 con acceso al sitio.

```powershell
git clone https://github.com/neuzh22/gestor-sharepoint.git
cd gestor-sharepoint
pip install -r requirements.txt
```

> 💡 También puedes hacer doble clic en **`iniciar.bat`**: instala lo que falte y abre el menú.

## 🚀 Uso

### 1. Abrir el programa

Haz doble clic en **`SP-Migrar.exe`**, o con Python:

```powershell
python sp_migrar.py        # o doble clic en iniciar.bat
```

```text
==================================================
  SP-MIGRAR  ·  SharePoint -> NAS
==================================================
  1) DESCARGAR carpetas de SharePoint
  2) Continuar una descarga que se interrumpió
  3) Reintentar solo los fallidos de un trabajo
  4) Ver estado de un trabajo
  5) Cerrar sesión de Microsoft
  0) Salir
```

### 2. Iniciar sesión

La primera vez te muestra un código. Abre **https://microsoft.com/devicelogin**, pégalo y entra con tu cuenta. La sesión queda guardada y se renueva sola.

### 3. Escribir la URL del sitio

```text
https://empresa.sharepoint.com/sites/Ventas
```

### 4. Elegir las carpetas

| Escribes | Qué pasa |
|:---:|---|
| `3` | Abre la carpeta 3 para ver lo que tiene |
| `D 3` | **Descarga** la carpeta 3 completa (varias: `D 1,4` o `D 2-5`) |
| `D` | **Descarga** la carpeta donde estás |
| `A` | Vuelve atrás |
| `S` | Sale sin descargar |

### 5. Elegir el destino y confirmar

```text
  Ruta destino: \\192.168.1.50\Respaldos\SharePoint
  Espacio libre en destino: 3.20 TB  |  Tamaño a copiar: 22.37 GB
```

Lo descargado queda así:

```text
\\192.168.1.50\Respaldos\SharePoint\
└── Ventas\                ← sitio
    └── Documentos\        ← biblioteca
        ├── Clientes\
        └── Proyectos\
```

### 6. Si se interrumpe

Usa la opción **2** del menú, o ejecútalo directamente:

```powershell
SP-Migrar.exe --trabajo NOMBRE_DEL_TRABAJO
SP-Migrar.exe --trabajo NOMBRE_DEL_TRABAJO --solo-fallidos
SP-Migrar.exe --trabajo NOMBRE_DEL_TRABAJO --hilos 8
# (con Python: python sp_migrar.py --trabajo ...)
```

Lo que ya está copiado y no cambió se **omite**. Solo baja lo nuevo, lo modificado o lo que falló.

## 🛡️ Manejo de errores

| Situación | Qué hace |
|---|---|
| 🌐 Se cae internet | Pausa todo y revisa la conexión cada 5–60 s. Continúa sola cuando vuelve, sin límite de tiempo. |
| ✂️ Descarga cortada a la mitad | Reanuda desde el byte exacto donde quedó (archivo temporal `.sppart`). |
| 🐢 SharePoint limita (429/503) | Respeta el `Retry-After` y pausa todos los hilos. |
| 💥 Errores 5xx | Reintenta, esperando cada vez más entre intentos. |
| 🔑 Token vencido | Lo renueva solo. Si la sesión fue revocada, muestra un nuevo código y espera. |
| 🗄️ El NAS no responde | Pausa hasta que el destino vuelva a estar accesible. |
| 💾 Disco lleno | Avisa y reintenta cada 5 minutos. |
| ❌ Un archivo falla siempre | Lo registra y sigue con el resto. Al final hace 3 pasadas de reintento solo con los fallidos. |
| ⛔ Ctrl+C | Se detiene de forma segura y luego continúa donde quedó. |
| 🔤 Nombres inválidos en Windows | Los corrige (`a:b?.txt` → `a_b_.txt`). |
| 🖱️ Clic en la consola | Desactiva la "edición rápida", que congela el programa con un clic. |
| 😴 Suspensión del PC | La bloquea mientras copia. |

> ℹ️ Es normal que algunos documentos de Office pesen un poco distinto que en SharePoint, porque SharePoint les agrega metadatos al descargarlos. El log lo indica y no lo cuenta como error.

**Códigos de salida:** `0` todo OK · `2` quedaron errores · `130` detenido por el usuario.

## ⚙️ Configuración

La primera vez se crea `config.json` (puedes partir de [`config.example.json`](config.example.json)). Ese archivo **no se sube a git**.

| Clave | Por defecto | Descripción |
|---|:---:|---|
| `hilos` | `4` | Descargas simultáneas. Súbelo a 6–8 si tu enlace es rápido y bájalo si aparecen muchos 429. |
| `reintentos_archivo` | `5` | Intentos por archivo dentro de una pasada. |
| `reintentos_peticion` | `8` | Reintentos por petición ante errores 5xx o de red. |
| `pasadas_reintento` | `3` | Pasadas finales solo con los fallidos. |
| `espera_entre_pasadas_seg` | `300` | Espera antes de cada pasada de reintento. |
| `timeout_lectura_seg` | `120` | Segundos sin recibir datos antes de dar la conexión por colgada. |
| `intervalo_estado_seg` | `60` | Cada cuánto imprime la línea `[ESTADO]`. |
| `incluir_carpeta_sitio` | `true` | Crea una carpeta con el nombre del sitio dentro del destino. |
| `evitar_suspension` | `true` | Impide que Windows se suspenda durante la copia. |
| `tenant` · `client_id` · `client_secret` | — | Autenticación (ver abajo). |

### 🔐 Autenticación

Por defecto usa la app pública de Microsoft **"Microsoft Graph Command Line Tools"** con tu usuario, que solo pide permisos de lectura (`Sites.Read.All` y `Files.Read.All`).

<details>
<summary><b>Mi organización bloquea esa app / quiero usar una app propia</b></summary>

1. En **Entra ID → Registros de aplicaciones**, crea una app.
2. En **Autenticación**, activa *"Allow public client flows"*.
3. En **Permisos de API**, agrega los permisos delegados de Microsoft Graph `Sites.Read.All` y `Files.Read.All`.
4. En `config.json`, pon `client_id` (el de tu app) y `tenant` (tu dominio o el ID del tenant).

</details>

<details>
<summary><b>Correrlo sin usuario (ideal para tareas programadas)</b></summary>

1. Crea una app con el permiso **de aplicación** `Sites.Read.All` y dale consentimiento de administrador.
2. Crea un secreto de cliente.
3. En `config.json`, pon `tenant`, `client_id` y `client_secret`.

Así la copia no depende de la sesión de ninguna persona.

</details>

### ⏰ Tarea programada (opcional)

En el Programador de tareas de Windows crea una acción con:

```text
Programa:   C:\SP-Migrar\SP-Migrar.exe
Argumentos: --trabajo NOMBRE_DEL_TRABAJO
```

Usa una **ruta UNC** como destino, porque las unidades mapeadas (`Z:`) no existen dentro de las tareas programadas. Además, inicia sesión al menos una vez de forma interactiva antes.

## 📁 Estructura del proyecto

```text
├── sp_migrar.py          # Programa principal
├── iniciar.bat           # Lanzador con doble clic (versión Python)
├── compilar.bat          # Genera dist\SP-Migrar.exe (no requiere Python para usarlo)
├── requirements.txt      # Dependencias (msal, requests)
├── config.example.json   # Plantilla de configuración
│
│   Generados al usarlo (ignorados por git):
├── config.json           # Tu configuración
├── token_cache.bin       # 🔒 Tu sesión de Microsoft. NO compartir.
├── trabajos/             # Definición (.json) y estado (.db) de cada descarga
├── logs/                 # Log detallado de cada ejecución
└── reportes/             # CSV con lo que quedó con error
```

## 🩺 Problemas comunes

<details>
<summary><b>"AADSTS65001" o "necesita aprobación del administrador" al iniciar sesión</b></summary>

Tu tenant exige que un administrador apruebe la app. Pídele que dé consentimiento a *Microsoft Graph Command Line Tools*, o usa una [app propia](#-autenticación).
</details>

<details>
<summary><b>No aparece la biblioteca o la carpeta que busco</b></summary>

Solo se muestra lo que tu usuario puede ver. Revisa tus permisos en el sitio y comprueba que la URL sea la del sitio (`/sites/Nombre`).
</details>

<details>
<summary><b>Muchos avisos "SharePoint pide bajar el ritmo"</b></summary>

Es la limitación de Microsoft. El programa ya espera solo, pero si pasa mucho baja `hilos` a 2 o 3.
</details>

<details>
<summary><b>El programa parece congelado</b></summary>

Revisa el título de la ventana: se actualiza cada 5 segundos con el progreso. Si estás descargando un archivo muy grande, la línea `[ESTADO]` aparece cada minuto.
</details>

<details>
<summary><b>¿Cómo veo qué falló?</b></summary>

Usa la opción **4) Ver estado** del menú, o abre el CSV más reciente en `reportes/`.
</details>
