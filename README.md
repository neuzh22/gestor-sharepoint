<div align="center">

# ðŸ“¦ SP-Migrar

### Descarga carpetas de SharePoint Online a tu NAS desde la terminal

Eliges las carpetas y el destino, y lo dejas corriendo toda la noche: si se cae la red, el NAS o la sesiÃ³n, espera a que vuelvan y continÃºa.

![Python](https://img.shields.io/badge/Python-3.10%2B-3776AB?logo=python&logoColor=white)
![SharePoint](https://img.shields.io/badge/SharePoint-Online-0078D4?logo=microsoftsharepoint&logoColor=white)
![Microsoft Graph](https://img.shields.io/badge/Microsoft%20Graph-API-5E5E5E?logo=microsoft&logoColor=white)
![Windows](https://img.shields.io/badge/Windows-10%20%7C%2011-0078D6?logo=windows&logoColor=white)

[CaracterÃ­sticas](#-caracterÃ­sticas) Â·
[InstalaciÃ³n](#-instalaciÃ³n) Â·
[Uso](#-uso) Â·
[Manejo de errores](#%EF%B8%8F-manejo-de-errores) Â·
[ConfiguraciÃ³n](#%EF%B8%8F-configuraciÃ³n) Â·
[Problemas comunes](#-problemas-comunes)

</div>

---

## ðŸ‘€ AsÃ­ se ve

```text
==============================================================================
 ESTÃS EN:  Ventas  >  Documentos     (12 archivo(s) aquÃ­)
==============================================================================
    1) Clientes      (18.40 GB, 1203 elementos)
    2) Facturas      (1.20 GB, 50 elementos)
    3) Proyectos     (4.77 GB, 312 elementos)

  Â¿QuÃ© quieres hacer?
    Escribe un NÃšMERO        -> abrir esa carpeta y ver lo que tiene
    Escribe D y el nÃºmero    -> DESCARGAR esa carpeta completa   (ej: D 3   o varias: D 1,4)
    Escribe D sola           -> DESCARGAR la carpeta donde estÃ¡s ahora
    Escribe A                -> volver AtrÃ¡s
    Escribe S                -> Salir sin descargar

  OpciÃ³n: D 1,3
```

```text
17:25:07 INFO    OK  Clientes/2024/Contrato_ACME.pdf  [2.31 MB en 0.6s]
17:25:08 INFO    OK  Clientes/2024/Propuesta.docx  [412.20 KB en 0.3s]
17:26:07 INFO    [ESTADO] OK 1840 | omitidos 0 | errores 0 | en cola 16 | 9.71 GB a 42.30 MB/s | 4m 02s
02:13:44 WARNING Sin conexiÃ³n a internet (ConnectionError). Esperando a que vuelva la red...
02:31:10 INFO    ConexiÃ³n restablecida tras 17m 26s. Continuando.
```

## âœ¨ CaracterÃ­sticas

| | |
|---|---|
| ðŸ–¥ï¸ **100% terminal** | El inicio de sesiÃ³n es con cÃ³digo de dispositivo de Microsoft, asÃ­ que no hace falta navegador en el equipo. |
| ðŸ“‚ **Eliges quÃ© copiar** | Recorres bibliotecas y carpetas, ves su tamaÃ±o y descargas una o varias. |
| ðŸŽ¯ **Eliges a dÃ³nde** | Puede ser un NAS (`\\NAS\Respaldos`), una unidad de red (`Z:\`) o un disco local. |
| ðŸ” **Reanudable** | Si se corta, continÃºa donde quedÃ³, incluso a mitad de un archivo grande. |
| ðŸ›¡ï¸ **A prueba de fallos** | Aguanta caÃ­das de red, del NAS, de la sesiÃ³n, la limitaciÃ³n de SharePoint y el disco lleno. |
| âš¡ **RÃ¡pido** | Hace varias descargas a la vez (4 por defecto, configurable). |
| ðŸŒ™ **Desatendido** | Evita que el PC se suspenda y hace pasadas automÃ¡ticas de reintento. |
| ðŸ“Š **Trazable** | Guarda un log detallado y deja un reporte CSV de errores que se abre en Excel. |
| ðŸ—‚ï¸ **Fiel al original** | Respeta la estructura de carpetas y las fechas de modificaciÃ³n, y admite rutas de mÃ¡s de 260 caracteres. |
| ðŸ–¼ï¸ **Comprime imÃ¡genes** | Recomprime JPG, PNG y WebP con FFmpeg al 65% de calidad mientras descarga, y tambiÃ©n las que ya bajaste. |

## ðŸ”„ CÃ³mo funciona

```mermaid
flowchart LR
    A[ðŸ” Inicio de sesiÃ³n<br/>cÃ³digo de dispositivo] --> B[ðŸŒ URL del sitio]
    B --> C[ðŸ“‚ Elegir carpetas]
    C --> D[ðŸŽ¯ Elegir destino]
    D --> E[â¬‡ï¸ Descarga en paralelo]
    E --> K[ðŸ–¼ï¸ Comprime imÃ¡genes<br/>FFmpeg 65%]
    K --> F{Â¿Errores?}
    F -- No --> G[âœ… Listo]
    F -- SÃ­ --> H[ðŸ” Pasadas de reintento]
    H --> I[ðŸ“„ Reporte CSV]
    E -. corte / Ctrl+C .-> J[(ðŸ’¾ Estado SQLite)]
    J -. reanudar .-> E
```

## ðŸ“¥ InstalaciÃ³n

### OpciÃ³n A: ejecutable, sin instalar nada (recomendado)

1. Copia **`SP-Migrar.exe`** y **`ffmpeg.exe`** a una misma carpeta del equipo, por ejemplo `C:\SP-Migrar\`.
2. Haz doble clic en `SP-Migrar.exe`. Listo.

No necesita Python. `ffmpeg.exe` solo hace falta para comprimir imÃ¡genes: sin Ã©l, todo lo demÃ¡s funciona igual. La configuraciÃ³n, la sesiÃ³n, los trabajos y los logs se guardan **en la misma carpeta del `.exe`**, asÃ­ que elige una carpeta donde tengas permiso de escritura (no `C:\Program Files`).

> ðŸ› ï¸ Para generar el `.exe` tÃº mismo, ejecuta **`compilar.bat`** en un equipo con Python. El resultado queda en `dist\`. Si pusiste FFmpeg en `ffmpeg\ffmpeg.exe` (ver [CompresiÃ³n de imÃ¡genes](#%EF%B8%8F-compresiÃ³n-de-imÃ¡genes)), tambiÃ©n se copia ahÃ­.

> âš ï¸ Windows SmartScreen puede advertir *"Windows protegiÃ³ su PC"*, porque el ejecutable no estÃ¡ firmado. Pulsa **MÃ¡s informaciÃ³n â†’ Ejecutar de todas formas**.

### OpciÃ³n B: con Python

**Requisitos:** Windows 10/11, [Python 3.10+](https://www.python.org/downloads/) (marca *"Add Python to PATH"* al instalar) y una cuenta de Microsoft 365 con acceso al sitio.

```powershell
git clone https://github.com/neuzh22/gestor-sharepoint.git
cd gestor-sharepoint
pip install -r requirements.txt
```

> ðŸ’¡ TambiÃ©n puedes hacer doble clic en **`iniciar.bat`**: instala lo que falte y abre el menÃº.

## ðŸš€ Uso

### 1. Abrir el programa

Haz doble clic en **`SP-Migrar.exe`**, o con Python:

```powershell
python sp_migrar.py        # o doble clic en iniciar.bat
```

```text
==================================================
  SP-MIGRAR  Â·  SharePoint -> NAS
==================================================
  1) DESCARGAR carpetas de SharePoint
  2) Continuar una descarga que se interrumpiÃ³
  3) Reintentar solo los fallidos de un trabajo
  4) Ver estado de un trabajo
  5) Comprimir imÃ¡genes ya descargadas (FFmpeg)
  6) Cerrar sesiÃ³n de Microsoft
  0) Salir
```

### 2. Iniciar sesiÃ³n

La primera vez te muestra un cÃ³digo. Abre **https://microsoft.com/devicelogin**, pÃ©galo y entra con tu cuenta. La sesiÃ³n queda guardada y se renueva sola.

### 3. Escribir la URL del sitio

```text
https://empresa.sharepoint.com/sites/Ventas
```

### 4. Elegir las carpetas

| Escribes | QuÃ© pasa |
|:---:|---|
| `3` | Abre la carpeta 3 para ver lo que tiene |
| `D 3` | **Descarga** la carpeta 3 completa (varias: `D 1,4` o `D 2-5`) |
| `D` | **Descarga** la carpeta donde estÃ¡s |
| `A` | Vuelve atrÃ¡s |
| `S` | Sale sin descargar |

### 5. (Opcional) Descargar solo ciertas subcarpetas, por ejemplo el aÃ±o 2026

Si tus carpetas tienen muchos niveles y solo quieres las de un nombre concreto, escrÃ­belo cuando lo pregunte. Se buscan en **todos los niveles**:

```text
  Nombre de carpeta (Enter = descargar TODO): 2026

  BÃºsqueda terminada: 214 carpetas revisadas.
  Se encontraron 87 carpeta(s), 35.20 GB en total:
    - Documentos/Clientes/ACME/CamiÃ³n/2026       (1.20 GB)
    - Documentos/Clientes/ACME/Contenedor/2026   (840.00 MB)
    - Documentos/Clientes/BETA/2026              (2.10 GB)
    ...
  Â¿Descargar solo estas carpetas? (S/n):
```

- Se descarga cada carpeta `2026` **completa**, con sus subcarpetas, y se respeta su ruta (`Clientes\ACME\CamiÃ³n\2026\...`).
- No se crean carpetas vacÃ­as de clientes que no tengan 2026, ni se bajan archivos sueltos de otros niveles.
- El nombre debe ser **exacto**, sin distinguir mayÃºsculas: `Notas 2026` no cuenta. Puedes poner varios (`2025, 2026`) o usar comodines (`2026*`).

### 6. Elegir el destino y confirmar

```text
  Ruta destino: \\192.168.1.50\Respaldos\SharePoint
  Espacio libre en destino: 3.20 TB  |  TamaÃ±o a copiar: 22.37 GB
```

Lo descargado queda asÃ­:

```text
\\192.168.1.50\Respaldos\SharePoint\
â””â”€â”€ Ventas\                â† sitio
    â””â”€â”€ Documentos\        â† biblioteca
        â”œâ”€â”€ Clientes\
        â””â”€â”€ Proyectos\
```

### 7. Si se interrumpe

Usa la opciÃ³n **2** del menÃº, o ejecÃºtalo directamente:

```powershell
SP-Migrar.exe --trabajo NOMBRE_DEL_TRABAJO
SP-Migrar.exe --trabajo NOMBRE_DEL_TRABAJO --solo-fallidos
SP-Migrar.exe --trabajo NOMBRE_DEL_TRABAJO --hilos 8
# (con Python: python sp_migrar.py --trabajo ...)
```

Lo que ya estÃ¡ copiado y no cambiÃ³ se **omite**. Solo baja lo nuevo, lo modificado o lo que fallÃ³.

## ðŸ›¡ï¸ Manejo de errores

| SituaciÃ³n | QuÃ© hace |
|---|---|
| ðŸŒ Se cae internet | Pausa todo y revisa la conexiÃ³n cada 5â€“60 s. ContinÃºa sola cuando vuelve, sin lÃ­mite de tiempo. |
| âœ‚ï¸ Descarga cortada a la mitad | Reanuda desde el byte exacto donde quedÃ³ (archivo temporal `.sppart`). |
| ðŸ¢ SharePoint limita (429/503) | Respeta el `Retry-After` y pausa todos los hilos. |
| ðŸ’¥ Errores 5xx | Reintenta, esperando cada vez mÃ¡s entre intentos. |
| ðŸ”‘ Token vencido | Lo renueva solo. Si la sesiÃ³n fue revocada, muestra un nuevo cÃ³digo y espera. |
| ðŸ—„ï¸ El NAS no responde | Pausa hasta que el destino vuelva a estar accesible. |
| ðŸ’¾ Disco lleno | Avisa y reintenta cada 5 minutos. |
| âŒ Un archivo falla siempre | Lo registra y sigue con el resto. Al final hace 3 pasadas de reintento solo con los fallidos. |
| â›” Ctrl+C | Se detiene de forma segura y luego continÃºa donde quedÃ³. |
| ðŸ”¤ Nombres invÃ¡lidos en Windows | Los corrige (`a:b?.txt` â†’ `a_b_.txt`). |
| ðŸ–±ï¸ Clic en la consola | Desactiva la "ediciÃ³n rÃ¡pida", que congela el programa con un clic. |
| ðŸ˜´ SuspensiÃ³n del PC | La bloquea mientras copia. |

> â„¹ï¸ Es normal que algunos documentos de Office pesen un poco distinto que en SharePoint, porque SharePoint les agrega metadatos al descargarlos. El log lo indica y no lo cuenta como error.

**CÃ³digos de salida:** `0` todo OK Â· `2` quedaron errores Â· `130` detenido por el usuario.

## ðŸ–¼ï¸ CompresiÃ³n de imÃ¡genes

Al crear una descarga, el programa pregunta **"Â¿Comprimir las imÃ¡genes?"**. Si respondes que sÃ­, cada imagen se recomprime con FFmpeg justo despuÃ©s de descargarla:

| Formato | QuÃ© se hace |
|---|---|
| `.jpg` `.jpeg` `.jfif` | Se recomprime al **65% de calidad**. Suele ahorrar entre 40 y 60%. |
| `.webp` | Se recomprime al **65% de calidad**. |
| `.png` | Se optimiza **sin pÃ©rdida**, porque PNG no tiene calidad. Ahorra poco. |
| Otros (`.gif`, `.heic`, `.tif`, documentosâ€¦) | Se copian tal cual. |

```text
OK  Fotos/playa.jpg  [99.56 KB en 0.9s]  Â· comprimida 99.56 KB -> 42.54 KB (-58%)
```

**Reglas de seguridad:**
- Las imÃ¡genes conservan su **nombre y extensiÃ³n**, asÃ­ que no se rompen enlaces.
- La imagen comprimida solo reemplaza a la original si **realmente pesa menos**.
- Si FFmpeg falla con una imagen (daÃ±ada, CMYKâ€¦), **se conserva la original** y la copia sigue sin marcarse como error.
- Los WebP y PNG **animados** no se tocan, para no perder la animaciÃ³n.
- Las fotos giradas por EXIF (las de celular) **se guardan ya derechas**.
- Al relanzar el trabajo, las imÃ¡genes ya comprimidas **no se vuelven a descargar**.

> âš ï¸ El JPEG recomprimido **pierde los metadatos EXIF** (fecha de captura, cÃ¡mara, GPS). La fecha de modificaciÃ³n del archivo sÃ­ se conserva. Si necesitas esos datos, no actives la compresiÃ³n.

**Â¿Ya descargaste imÃ¡genes sin comprimir?** La opciÃ³n **5) Comprimir imÃ¡genes ya descargadas** del menÃº las comprime directamente en el destino, sin volver a conectarse a SharePoint.

**Obtener FFmpeg:** descarga *ffmpeg-release-essentials.zip* de [gyan.dev](https://www.gyan.dev/ffmpeg/builds/), saca `bin\ffmpeg.exe` y ponlo en cualquiera de estos lugares (se busca en este orden):

1. La ruta indicada en `ffmpeg_ruta` de `config.json`.
2. Junto a `SP-Migrar.exe` / `sp_migrar.py`.
3. `ffmpeg\ffmpeg.exe` o `ffmpeg\bin\ffmpeg.exe` dentro de esa carpeta.
4. Cualquier carpeta del `PATH`.

## âš™ï¸ ConfiguraciÃ³n

La primera vez se crea `config.json` (puedes partir de [`config.example.json`](config.example.json)). Ese archivo **no se sube a git**.

| Clave | Por defecto | DescripciÃ³n |
|---|:---:|---|
| `hilos` | `4` | Descargas simultÃ¡neas. SÃºbelo a 6â€“8 si tu enlace es rÃ¡pido y bÃ¡jalo si aparecen muchos 429. |
| `reintentos_archivo` | `5` | Intentos por archivo dentro de una pasada. |
| `reintentos_peticion` | `8` | Reintentos por peticiÃ³n ante errores 5xx o de red. |
| `pasadas_reintento` | `3` | Pasadas finales solo con los fallidos. |
| `espera_entre_pasadas_seg` | `300` | Espera antes de cada pasada de reintento. |
| `timeout_lectura_seg` | `120` | Segundos sin recibir datos antes de dar la conexiÃ³n por colgada. |
| `intervalo_estado_seg` | `60` | Cada cuÃ¡nto imprime la lÃ­nea `[ESTADO]`. |
| `incluir_carpeta_sitio` | `true` | Crea una carpeta con el nombre del sitio dentro del destino. |
| `evitar_suspension` | `true` | Impide que Windows se suspenda durante la copia. |
| `comprimir_imagenes` | `true` | Respuesta propuesta al crear una descarga. TambiÃ©n se aplica a trabajos creados antes de esta funciÃ³n. |
| `calidad_imagenes` | `65` | Calidad de JPG/WebP (1â€“100). Menos calidad, archivos mÃ¡s pequeÃ±os. |
| `ffmpeg_ruta` | `""` | Ruta a `ffmpeg.exe` si no estÃ¡ junto al programa ni en el `PATH`. |
| `tenant` Â· `client_id` Â· `client_secret` | â€” | AutenticaciÃ³n (ver abajo). |

### ðŸ” AutenticaciÃ³n

Por defecto usa la app pÃºblica de Microsoft **"Microsoft Graph Command Line Tools"** con tu usuario, que solo pide permisos de lectura (`Sites.Read.All` y `Files.Read.All`).

<details>
<summary><b>Mi organizaciÃ³n bloquea esa app / quiero usar una app propia</b></summary>

1. En **Entra ID â†’ Registros de aplicaciones**, crea una app.
2. En **AutenticaciÃ³n**, activa *"Allow public client flows"*.
3. En **Permisos de API**, agrega los permisos delegados de Microsoft Graph `Sites.Read.All` y `Files.Read.All`.
4. En `config.json`, pon `client_id` (el de tu app) y `tenant` (tu dominio o el ID del tenant).

</details>

<details>
<summary><b>Correrlo sin usuario (ideal para tareas programadas)</b></summary>

1. Crea una app con el permiso **de aplicaciÃ³n** `Sites.Read.All` y dale consentimiento de administrador.
2. Crea un secreto de cliente.
3. En `config.json`, pon `tenant`, `client_id` y `client_secret`.

AsÃ­ la copia no depende de la sesiÃ³n de ninguna persona.

</details>

### â° Tarea programada (opcional)

En el Programador de tareas de Windows crea una acciÃ³n con:

```text
Programa:   C:\SP-Migrar\SP-Migrar.exe
Argumentos: --trabajo NOMBRE_DEL_TRABAJO
```

Usa una **ruta UNC** como destino, porque las unidades mapeadas (`Z:`) no existen dentro de las tareas programadas. AdemÃ¡s, inicia sesiÃ³n al menos una vez de forma interactiva antes.

## ðŸ“ Estructura del proyecto

```text
â”œâ”€â”€ sp_migrar.py          # Programa principal
â”œâ”€â”€ iniciar.bat           # Lanzador con doble clic (versiÃ³n Python)
â”œâ”€â”€ compilar.bat          # Genera dist\SP-Migrar.exe (no requiere Python para usarlo)
â”œâ”€â”€ requirements.txt      # Dependencias (msal, requests)
â”œâ”€â”€ config.example.json   # Plantilla de configuraciÃ³n
â”‚
â”‚   Ignorados por git:
â”œâ”€â”€ ffmpeg/ffmpeg.exe     # FFmpeg para comprimir imÃ¡genes (100 MB, se descarga aparte)
â”œâ”€â”€ config.json           # Tu configuraciÃ³n
â”œâ”€â”€ token_cache.bin       # ðŸ”’ Tu sesiÃ³n de Microsoft. NO compartir.
â”œâ”€â”€ trabajos/             # DefiniciÃ³n (.json) y estado (.db) de cada descarga
â”œâ”€â”€ logs/                 # Log detallado de cada ejecuciÃ³n
â””â”€â”€ reportes/             # CSV con lo que quedÃ³ con error
```

## ðŸ©º Problemas comunes

<details>
<summary><b>"AADSTS65001" o "necesita aprobaciÃ³n del administrador" al iniciar sesiÃ³n</b></summary>

Tu tenant exige que un administrador apruebe la app. PÃ­dele que dÃ© consentimiento a *Microsoft Graph Command Line Tools*, o usa una [app propia](#-autenticaciÃ³n).
</details>

<details>
<summary><b>No aparece la biblioteca o la carpeta que busco</b></summary>

Solo se muestra lo que tu usuario puede ver. Revisa tus permisos en el sitio y comprueba que la URL sea la del sitio (`/sites/Nombre`).
</details>

<details>
<summary><b>Muchos avisos "SharePoint pide bajar el ritmo"</b></summary>

Es la limitaciÃ³n de Microsoft. El programa ya espera solo, pero si pasa mucho baja `hilos` a 2 o 3.
</details>

<details>
<summary><b>El programa parece congelado</b></summary>

Revisa el tÃ­tulo de la ventana: se actualiza cada 5 segundos con el progreso. Si estÃ¡s descargando un archivo muy grande, la lÃ­nea `[ESTADO]` aparece cada minuto.
</details>

<details>
<summary><b>Â¿CÃ³mo veo quÃ© fallÃ³?</b></summary>

Usa la opciÃ³n **4) Ver estado** del menÃº, o abre el CSV mÃ¡s reciente en `reportes/`.
</details>
