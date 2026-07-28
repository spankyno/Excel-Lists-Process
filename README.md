# Excel Lists Process

Herramienta web gratuita para **combinar y sumar**, o **comparar en paralelo**, listas de Excel (`.xlsx`, `.xls`) y CSV directamente en el navegador. Sin backend: ningún archivo se sube a ningún servidor.

Sitio 100% estático (HTML + CSS + JS vanilla), pensado para cargar y responder lo más rápido posible.

## Procesos

1. **Consolidar y sumar** — El Campo Clave y el Campo a Calcular se eligen una sola vez y deben existir (con el mismo nombre) en todos los archivos cargados. Si coinciden, sus valores del Campo a Calcular se suman en una fila; el resultado conserva las columnas comunes a todos los archivos más una columna nueva con la suma. Si no hay ninguna columna común, la app avisa y no permite continuar.
2. **Comparar en paralelo** — El Campo Clave se elige una sola vez y también debe ser común a todos los archivos. Cuando coincide, las columnas completas del Archivo 1, luego del Archivo 2 y sucesivos se combinan en una sola fila; si una clave solo existe en un archivo, el resto de columnas queda vacío.

Ambos procesos permiten reordenar el resultado por hasta **dos criterios** (campo + dirección ascendente/descendente), y exportar a `.xlsx`, `.csv` o copiar al portapapeles.

## Estructura del repositorio

```
.
├── index.html                  # Herramienta principal
├── about.html                  # Página "Acerca de"
├── robots.txt
├── sitemap.xml
├── site.webmanifest
├── _headers                    # Cabeceras de caché/seguridad (Cloudflare Pages)
├── og-image.png                 # ⚠️ Añádela tú mismo (1200×630) antes de desplegar
├── assets/
│   ├── css/
│   │   ├── styles.css          # Fuente legible
│   │   └── styles.min.css      # Producción (referenciada en el HTML)
│   ├── js/
│   │   ├── app.js              # Fuente legible
│   │   └── app.min.js          # Producción (referenciada en index.html)
│   ├── vendor/
│   │   └── xlsx.full.min.js    # SheetJS vendorizado (se carga de forma diferida)
│   └── icons/
│       ├── favicon.svg / favicon.ico / apple-touch-icon.png / icon-192.png / icon-512.png
└── package.json                 # Scripts para regenerar los assets minificados
```

## Rendimiento

- Sin frameworks de frontend: HTML/CSS/JS vanilla.
- SheetJS (la única dependencia pesada) se carga de forma **diferida**: no bloquea la carga inicial, solo se descarga cuando el usuario carga o pega un archivo.
- CSS y JS de producción minificados (`*.min.*`), fuente legible conservada aparte.
- Cabeceras de caché inmutable de un año para `/assets/*` e iconos vía `_headers` (Cloudflare Pages).
- Fuentes vía Google Fonts con `preconnect` y `display=swap`.

Para regenerar los `.min.*` tras editar `styles.css` o `app.js`:

```bash
npm install
npm run build
```

## Desplegar en Cloudflare Pages (vía GitHub)

1. Sube este repositorio a GitHub.
2. En Cloudflare Pages: **Create a project → Connect to Git** y selecciona el repositorio.
3. Configuración de build:
   - **Framework preset:** None
   - **Build command:** (vacío)
   - **Build output directory:** `/`
4. Añade tu propio `og-image.png` (1200×630 px) en la raíz del proyecto antes de desplegar — se referencia en las etiquetas Open Graph / Twitter Card de `index.html` y `about.html`.
5. El dominio `*.pages.dev` que te asigne Cloudflare debe coincidir con las URLs usadas en `index.html`, `about.html`, `robots.txt` y `sitemap.xml` (`https://excel-lists-process.pages.dev/`). Si usas otro subdominio o dominio propio, actualiza esas referencias (canonical, Open Graph, JSON-LD y sitemap).
6. La verificación de Google Search Console ya está incluida como `<meta name="google-site-verification">` en `index.html`.

## Autor

**Aitor Sánchez Gutiérrez**
Blog: https://aitorsanchez.pages.dev/ · Más apps: https://aitorhub.vercel.app/ · Contacto: https://aitor-blog-contacto.vercel.app/

© 2026 Aitor Sánchez Gutiérrez — Reservados todos los derechos.
