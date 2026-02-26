# Visor de Excel en el navegador (HTML + JS)

Aplicación web estática que permite cargar un archivo de Excel (`.xlsx` / `.xls`) desde el sistema local, listar sus hojas, mostrar la primera hoja visible como tabla HTML y navegar entre hojas con un selector. Incluye utilidades de formateo, estandarización de textos y copia rápida de celdas.

## Características

- **Carga local de Excel** mediante `<input type="file">` y `FileReader` (ArrayBuffer).
- **Parseo de libro/hojas** con [SheetJS (XLSX) 0.18.5] vía CDN (`xlsx.full.min.js`).
- **Respeto de hojas ocultas**: se listan en el selector como “(oculta)” y se deshabilitan; solo se renderizan las visibles.
- **Selector de hojas**: muestra/oculta el contenedor de cada hoja sin recargar.
- **Render a tabla HTML** usando `XLSX.utils.sheet_to_html` y posterior post-procesamiento del DOM.
- **Estilos semánticos por filas**: si la primera celda contiene `language`, `geoloc`, `preview copy`, o `extended copy`, la fila recibe clases especiales (`.language`, `.geoloc`, `.preview`, `.extended`) con fondo azul, texto blanco y centrado.
- **Normalización de formato inline**: convierte `<b>/<i>/<u>/<del>` (y estilos inline en `<span>`) a etiquetas semánticas (`<strong>`, `<i>`, `<u>`, `<del>`) preservando espacios y saltos de línea de forma visual. *(Marcado como “FUNCIONA A MEDIAS” en el código)*.
- **Estandarización de textos**: reemplazos automáticos (p. ej. `all inclusive` → `All Inclusive`, `all-inclusive` → `All-Inclusive`, `todo incluido` → `Todo Incluido`, y `Grand -` → `Grand –`).
- **Casillas para booleanos**: convierte `true/1/yes/si/on/x` en checkbox marcado y `false` en checkbox deshabilitado; centra la celda.
- **Resaltado de metacampos**: celdas con términos como `bw`, `tw`, `from`, `to`, `headline`, `description`, `title`, `type`, `disclaimer`, `tyc`, `block`, `bo dates` se sombrean en gris y sin interacción.
- **Copia rápida de celdas**: click para seleccionar/copiar el contenido (`document.execCommand('copy')`), con notificación `✓ Copiado`. (Aplica a celdas no vacías; prioriza textos largos > 16 y no numéricos para cursor tipo “copy”).

## Estructura

- `index.html`: HTML, estilos y lógica JS en un solo archivo, incluye el CDN de SheetJS (`xlsx.full.min.js`).

## Requisitos

- **Navegador moderno** con soporte para `FileReader`, `ArrayBuffer`, `DOMParser` y `querySelector`.
- Conexión a Internet para cargar el **CDN de SheetJS** utilizado en el `<script src="https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.18.5/xlsx.full.min.js"></script>`. *(Si deseas uso 100% offline, descarga ese archivo y sirvelo localmente.)*

## Uso

1. Abre `index.html` en tu navegador.
2. Haz clic en **“Elegir archivo”** y selecciona un `.xlsx` o `.xls`.
3. El selector mostrará las hojas del libro. Las hojas ocultas aparecerán como “(oculta)” y estarán deshabilitadas.
4. Navega entre hojas desde el selector; solo la hoja seleccionada queda visible.
5. Haz clic en una celda de texto para **copiar** su contenido; verás una notificación “✓ Copiado”.

## Detalles de implementación

- **Render de tabla**: `XLSX.utils.sheet_to_html(hoja, { editable: false, header: '', footer: '' })`.
- **Normalización de formato**:
  - Función `normalizarEtiquetasYBR` que encadena `agregarEtiquetasDeFormato` y `remplazarTextos`.
  - `agregarEtiquetasDeFormato` detecta `<b>`, `<i>`, `<u>`, `<del>` y estilos en `<span>` para envolver con etiquetas semánticas, intentando **mover espacios** fuera/dentro del wrapper para conservar la apariencia (función `moverBlancos`). *(El propio código anota que “FUNCIONA A MEDIAS”.)*
- **Marcado/estilo de filas**: inspección de la **primera celda** (lowercased, `includes`) para asignar clases.
- **Conversión de booleanos**: mapea varias cadenas (“true/1/yes/si/on/x”) a `<input type="checkbox" checked>`; “false” a `<input type="checkbox" disabled>`.
- **Resaltado de metacampos**: términos clave provocan `background-color: #e6e6e6; pointer-events: none;`.
- **Copia al portapapeles**: selección con `Range` + `document.execCommand('copy')` y UI mínima de confirmación.

## Limitaciones y notas

- **Normalización de etiquetas**: el comentario “FUNCIONA A MEDIAS :/” sugiere casos borde con espacios y `<br>` que podrían requerir refinamiento adicional.
- **Booleans**: actualmente solo reconoce valores exactos (con `trim().toLowerCase()`) definidos en el código; otros equivalentes (`sí` con tilde, `true/false` capitalizados en celdas con formato, etc.) pueden no detectarse.
- **Compatibilidad de portapapeles**: `document.execCommand('copy')` funciona ampliamente, pero `navigator.clipboard.writeText` sería el método moderno (requiere HTTPS/origen seguro). El proyecto usa el método clásico por simplicidad.
- **Hojas ocultas**: se listan deshabilitadas en el selector y **no** se renderizan; si necesitas inspeccionarlas, habría que permitir el render bajo confirmación.

## Roadmap (ideas)

- [ ] Hacer el **normalizador** más robusto (manejo de espacios alrededor de `<br>`, nested tags y spans con estilos combinados).
- [ ] Exportar la tabla procesada a un nuevo **.xlsx**.
- [ ] Parametrizar los **reemplazos de texto** y la lista de **metacampos** vía UI.

## Cómo desarrollar

Este proyecto es **estático** (sin build). Basta con un servidor local opcional para evitar restricciones de algunos navegadores al abrir archivos locales:
