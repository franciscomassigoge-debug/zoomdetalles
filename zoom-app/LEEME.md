# Zoom Agricultura — Detalles Administración

App instalable (PWA) para generar el detalle de trabajos realizados a partir
del Excel "Registro Zoom", con ajuste de precios, exportación de Excel
actualizado e historial compartido entre Francisco y Tomás.

## Qué incluye esta carpeta

- `index.html`, `app.js`, `store.js`, `styles.css` — la app.
- `manifest.json`, `service-worker.js` — lo que la hace instalable (PWA).
- `firebase-config.js` — acá se pega la configuración de Firebase para que
  el historial y los precios se compartan entre los dos usuarios.
- `assets/` — los dos logos, los íconos de la app y el Excel base por defecto.

## Paso 1 — Publicar la app (para que sea instalable)

Una PWA solo se puede "instalar" (ícono en el teléfono/compu) si se abre
desde una URL real, no con doble clic en el archivo. La forma más simple y
gratuita es **GitHub Pages**:

1. Crear una cuenta en https://github.com si no tenés una.
2. Crear un repositorio nuevo, por ejemplo `zoom-detalles` (puede ser
   privado o público, ambos casos funcionan con GitHub Pages).
3. Subir todo el contenido de esta carpeta (`zoom-app/`) a ese repositorio
   — se puede arrastrar los archivos directamente desde la web de GitHub
   ("Add file" → "Upload files").
4. Ir a "Settings" → "Pages" del repositorio → en "Source" elegir la rama
   principal (`main`) y carpeta `/ (root)` → Guardar.
5. GitHub va a dar una URL tipo `https://tu-usuario.github.io/zoom-detalles/`.
   Esa es la dirección que van a usar Francisco y Tomás.

Alternativa igual de simple: **Netlify** (https://app.netlify.com/drop) —
se arrastra la carpeta entera a la página y da una URL al instante, sin
necesidad de cuenta de GitHub.

## Paso 2 — Activar el historial y precios compartidos (Firebase)

Sin este paso, la app funciona igual pero cada dispositivo guarda su propio
historial y precios (modo local). Para que Francisco y Tomás vean lo mismo:

Seguí las instrucciones detalladas dentro de `firebase-config.js` (están
comentadas al principio del archivo). En resumen: crear un proyecto gratis
en https://console.firebase.google.com, activar Firestore Database, copiar
la configuración que te da Firebase y pegarla en `firebase-config.js`.
Después volvés a subir ese archivo actualizado al repositorio (Paso 1).

## Paso 3 — Instalar la app en el teléfono y la compu

Una vez publicada la URL (Paso 1):

- **Celular (Android/Chrome):** abrir la URL → menú (⋮) → "Instalar app" o
  "Agregar a pantalla de inicio".
- **Celular (iPhone/Safari):** abrir la URL → botón compartir → "Agregar a
  pantalla de inicio".
- **Computadora (Chrome/Edge):** abrir la URL → ícono de instalación en la
  barra de direcciones (o menú → "Instalar Zoom Agricultura…").

## Cómo se usa

1. La app carga por defecto el Excel base. Si tenés una versión más nueva
   del "Registro Zoom", subila con el botón "Subir Excel actualizado" —
   si no subís nada, sigue usando la base.
2. Elegí fecha desde/hasta, cliente(s) y establecimiento(s) (dejar vacío
   = todos) y tocá "Buscar trabajos". La app busca combinado en Facturación
   1 (trabajos por hectárea) y Facturación 2 (trabajos por día).
3. Si hace falta, ajustá el precio por hectárea/día o por km directamente
   en la tabla — los totales se recalculan solos.
4. "Guardar precios ajustados" los deja disponibles para la próxima vez que
   aparezca ese cliente + tipo de trabajo (compartido si configuraste
   Firebase).
5. "Generar PDF" descarga el detalle con el logo de Zoom y queda registrado
   en el Historial.
6. "Descargar Excel actualizado" baja una copia del Registro Zoom con los
   precios ajustados ya reflejados en las filas correspondientes.

## Notas técnicas

- La app funciona 100% en el navegador, sin servidor propio (salvo
  Firestore para lo compartido).
- Los datos del Excel se leen de las solapas "Trabajos" (por hectárea) y
  "Trabajos 2" (por día) — las solapas "Facturacion"/"Facturacion 2" son
  solo una vista con fórmulas de esas mismas filas, así que no hace falta
  tocarlas.
- El Excel base embebido es el que subiste (`Registro Zoom 25-26.xlsx`).
  Para reemplazarlo por defecto, hay que subir un nuevo archivo a
  `assets/registro-zoom-base.xlsx` en el repositorio.
