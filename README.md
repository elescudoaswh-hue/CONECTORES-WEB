# Analizador de Conectores Yazaki · Web 0.7.4

Versión web del programa **Crear Conectores y sus Colores**. Incluye el diseño de escritorio con modos claro y oscuro, las pestañas de resultados y el motor Python de la versión 0.7.4.

## Uso

1. Selecciona el **Initial Report** en formato `.xlsx` y, si tienes una, la **base de familias** `.json`. Una coincidencia completa y única aplica la familia automáticamente. También puedes elegir una familia o asignar colores a mano.
2. Selecciona la **carpeta DWG** y espera a que aparezca **Carpeta lista** con el total de dibujos. Se admiten subcarpetas y archivos `.dwg` o `.dxf`. Si el navegador lo pide, concede permiso de lectura o confirma **Cargar**. Puedes analizar el reporte antes de elegir una carpeta.
3. Revisa **Colores unidos**, por ejemplo `ROSA+VERDE;CAFE;AZUL`, y pulsa **Analizar y contar**. En **Validación** aparecen los conectores faltantes, los duplicados, las cavidades bloqueadas y las variantes.
4. En **Conectores**, marca las posiciones que quieres generar. Usa **Guardar base JSON** para descargar la asignación y el conteo. Usa **Generar DXF** para procesar la selección y descargar el dibujo cuando termine.

El análisis de Excel, la base JSON y la búsqueda de dibujos se ejecutan en el navegador. **AutoCAD de esta PC** entrega el Initial Report y los dibujos seleccionados al lanzador dentro de la misma computadora. **Generador en red** los envía al Windows conectado mediante el túnel. El generador retira los archivos de entrada tras el trabajo y conserva el resultado para descarga durante dos horas mientras se mantiene abierto. La carga por generación admite hasta 94 MiB de archivos. Para cargas mayores, selecciona los conectores en varios lotes.

El navegador recuerda el tema, la familia, el nombre de salida y los últimos archivos de Excel y JSON cuando permite almacenamiento local. Selecciona de nuevo la carpeta DWG al reabrir la aplicación. El nombre de **Salida DXF** se usa para la descarga a través del navegador.

## Generar con AutoCAD de esta PC

En cada equipo Windows de 64 bits con AutoCAD, sigue estos pasos con tu cuenta de Windows.

1. En la web, elige **AutoCAD de esta PC** y pulsa **Configurar**. Copia la dirección que aparece y descarga el ZIP. [Lanzador local para Windows](https://raw.githubusercontent.com/elescudoaswh-hue/CONECTORES-WEB/main/web/descargas/Conectores-Generador-Local-Windows.zip)
2. Extrae todos los archivos. Ejecuta **`INSTALAR_GENERADOR_LOCAL.bat`** y pega la dirección. El lanzador incluye su propio Python y queda instalado en el perfil del usuario. Puedes instalarlo con una cuenta estándar de Windows.
3. Abre la web publicada por HTTPS en Chrome o Edge. Usa **Generar DXF**. Si la web pide iniciar el lanzador, pulsa **Abrir lanzador** y permite la apertura y la conexión local cuando el navegador lo solicite. AutoCAD se abre en el escritorio al recibir el trabajo. Deja abierta la ventana del lanzador hasta descargar el archivo.

El lanzador acepta trabajos de la dirección HTTPS que indicaste en la instalación. Para cambiarla, cierra el lanzador o el navegador si usas la extensión. Ejecuta de nuevo el instalador e introduce la nueva dirección.

### Activar el lanzador con la extensión de Chrome o Edge

1. Instala o actualiza el lanzador de Windows desde **AutoCAD de esta PC → Configurar**. La versión 1.1 registra la conexión con la extensión en tu perfil de Windows.
2. Descarga y extrae el ZIP de la extensión. En `chrome://extensions` o `edge://extensions`, activa **Modo desarrollador**, pulsa **Cargar descomprimida** y elige la carpeta que contiene `manifest.json`. [Extensión de Conectores](https://raw.githubusercontent.com/elescudoaswh-hue/CONECTORES-WEB/main/web/descargas/Conectores-AutoCAD-Chrome-Edge.zip)
3. Abre la web de Conectores. Pulsa el icono de la extensión, elige **Vincular este sitio** y confirma el permiso para esa dirección.
4. Pulsa **Generar DXF** en la web. La extensión inicia el lanzador en segundo plano y AutoCAD se abre al recibir el trabajo. Si el navegador te pide acceso a la red local, concédelo para comunicar la web con el lanzador de tu PC.

La extensión recibe la orden de iniciar el generador. El navegador entrega los archivos seleccionados al generador local y recibe el DXF desde la misma PC.

## Publicar en Cloudflare Workers

1. En **Workers & Pages**, crea un Worker con **Import a repository** y elige `elescudoaswh-hue/CONECTORES-WEB`. Usa `main` como rama de producción y deja el directorio raíz del repositorio.
2. Usa **conectores-web** como nombre del Worker y `npx wrangler deploy` como comando de despliegue. Deja vacío el comando de compilación. Wrangler ya ejecuta `npm run build` y publica `dist/` mediante la configuración del repositorio.
3. El repositorio selecciona Node.js 24 con `.node-version`. Inicia el despliegue y abre la URL que te dé Cloudflare. Podrás analizar Excel y JSON desde ese momento. [Integración de Git con Workers](https://developers.cloudflare.com/workers/ci-cd/builds/).

### Conectar el generador DXF

Para usar un equipo Windows compartido, selecciona **Generador en red** en la web y configura ese equipo con estos pasos.

1. En un equipo **Windows** con **AutoCAD** y **Python 3.12 de 64 bits**, abre AutoCAD y ejecuta `INICIAR_GENERADOR.bat` desde la misma sesión de Windows. El lanzador prepara Python y solicita el token del servicio. Usa un token aleatorio que puedas volver a introducir cuando reinicies el generador.
2. Configura un **Cloudflare Tunnel** en ese equipo y dirige su nombre de host, por ejemplo `dxf.tu-dominio.com`, al servicio local `http://127.0.0.1:8765`. Mantén el generador iniciado en la sesión de Windows para que pueda controlar AutoCAD. [Cloudflare Tunnel](https://developers.cloudflare.com/cloudflare-one/networks/connectors/cloudflare-tunnel/).
3. En el Worker de Cloudflare, abre **Settings → Variables and Secrets**. Agrega `DXF_API_ORIGIN` con la URL HTTPS del túnel y **`DXF_API_TOKEN` como secreto**, con el mismo valor que introdujiste en Windows. Para un túnel protegido con una política de servicio de Cloudflare Access, agrega también `CF_ACCESS_CLIENT_ID` y `CF_ACCESS_CLIENT_SECRET` como secretos. [Secretos de Workers](https://developers.cloudflare.com/workers/configuration/secrets/), [tokens de servicio](https://developers.cloudflare.com/cloudflare-one/access-controls/service-credentials/service-tokens/).
4. Abre la web y prueba **Generar DXF** con un reporte y un dibujo de referencia de tu equipo. El trabajo avanza en segundo plano y muestra enlaces de descarga al terminar.

Para un despliegue interno, restringe el acceso al sitio con Cloudflare Access. Aplica la política a todos los dominios habilitados del Worker o desactiva las URL alternativas en **Domains & Routes**. [Rutas workers.dev](https://developers.cloudflare.com/workers/configuration/routing/workers-dev/).

## Ejecutar y verificar localmente

```sh
npm ci
npm run dev
```

Wrangler muestra la URL local en la terminal. Para probar también el DXF en ese Windows, copia `.dev.vars.example` a `.dev.vars`, completa el mismo token e inicia `INICIAR_GENERADOR.bat`.

```sh
npm run check:cloudflare
python -m pip install -r requirements.txt
python -m unittest discover -s tests -v
```

El primer comando prepara el paquete de Cloudflare y valida la configuración en modo de prueba. Las pruebas de Python cubren la lectura de reportes, colores, accesorios, bloqueos, selección y el flujo de descarga con un generador de prueba.

El ZIP del lanzador se genera con `python scripts/package_local_launcher.py Conectores-Generador-Local-Windows.zip --cache .cache --go ruta/al/compilador/go`. El empaquetado compila el proceso de conexión de la extensión, usa el intérprete portátil de Python para Windows y comprueba los hashes de las dependencias.
