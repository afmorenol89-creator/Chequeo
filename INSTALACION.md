# Chequeo Reproductivo — instalación

Son tres pasos. El primero es obligatorio, los otros dos se pueden hacer después.

1. Publicar la aplicación en una dirección de internet (una sola vez, ~10 minutos).
2. Conectar la hoja de Google (una sola vez, ~10 minutos).
3. Instalar la app en la tablet (2 minutos, con wifi).

---

## Paso 1 — Publicar la aplicación

La app necesita una dirección web fija. No sirve guardar el archivo suelto en la tablet: cuando Chrome abre un archivo local bloquea el almacenamiento del navegador, y el chequeo se perdería al cerrar la pestaña.

### Opción A — GitHub Pages (gratis y permanente, es la que recomiendo)

1. Crea una cuenta en https://github.com si no tienes.
2. Botón **New repository**. Nombre: `chequeo`. Márcalo **Public**. Crear.
3. En el repositorio vacío: **uploading an existing file**. Arrastra estos seis archivos:
   - `index.html`
   - `sw.js`
   - `manifest.webmanifest`
   - `xlsx.full.min.js`
   - `icon-192.png`
   - `icon-512.png`
4. Abajo, **Commit changes**.
5. Pestaña **Settings** → menú lateral **Pages** → en *Branch* elige `main` y carpeta `/ (root)` → **Save**.
6. Espera un par de minutos y recarga esa página. Aparece la dirección:
   `https://TUUSUARIO.github.io/chequeo/`

Esa es la dirección de la app. Guárdala.

El repositorio es público, pero solo contiene el programa. Los datos del hato nunca quedan ahí: viven en la tablet y en tu hoja de Google privada.

Para actualizar la app más adelante: subes el archivo nuevo al mismo repositorio y listo. En la tablet se actualiza sola la próxima vez que abra con internet.

### Opción B — Netlify Drop (más rápido, 2 minutos)

Entra a https://app.netlify.com/drop y arrastra la carpeta completa. Te da una dirección de inmediato. Crea una cuenta gratis para que la dirección no se venza.

---

## Paso 2 — Conectar la hoja de Google

Esto es lo que guarda el histórico de todos los chequeos.

1. Crea una hoja nueva en https://sheets.new y ponle nombre, por ejemplo *Chequeos Reproductivos La Morelia*.
2. Menú **Extensiones** → **Apps Script**.
3. Borra el contenido que aparece y pega todo el contenido del archivo `apps-script.gs`.
4. Guarda (ícono del disquete).
5. Botón azul **Implementar** → **Nueva implementación**.
6. En el engranaje elige **Aplicación web**. Configura:
   - Ejecutar como: **Yo**
   - Quién tiene acceso: **Cualquier usuario**
7. **Implementar**. Google pide autorizar: acepta. Si sale la advertencia de "app no verificada", entra en *Configuración avanzada* → *Ir a (nombre del proyecto)*. Es tu propio script, es seguro.
8. Copia la **URL de la aplicación web**. Termina en `/exec`.
9. Abre la app en la tablet o en el computador, entra al engranaje ⚙ de arriba a la derecha, pega esa dirección y toca **Guardar**. Luego **Probar conexión**: en la hoja debe aparecer una fila de prueba.

La hoja se organiza sola en dos pestañas:

- **Chequeos**: una fila por animal, de todos los chequeos, uno debajo del otro. Ese es el histórico.
- **Resumen**: una fila por chequeo con los totales.

Si un chequeo se envía dos veces (porque se reintentó al recuperar señal), reemplaza sus filas anteriores en vez de duplicarlas.

**Importante:** cada vez que cambies el código del Apps Script tienes que hacer *Implementar → Administrar implementaciones → editar → Versión nueva*, si no, sigue corriendo la versión vieja.

---

## Paso 3 — Instalar en la tablet

Con wifi, una sola vez:

1. Abre Chrome en la tablet y entra a la dirección del Paso 1.
2. Menú de los tres puntos → **Instalar aplicación** o **Agregar a pantalla de inicio**.
3. Queda un ícono verde en el escritorio. De ahí en adelante se abre como cualquier app, sin barra de navegador.

La primera carga baja alrededor de 1 MB (la librería que lee Excel). Hazlo con wifi. Después de eso la app abre y funciona completa sin señal.

**No borres los datos de navegación de Chrome** en esa tablet: ahí es donde viven los chequeos que todavía no se han enviado.

---

## Cómo se usa el día del chequeo

1. **Nuevo chequeo** → elegir el Excel que exporta DairyPlan.
2. La app aplica los cinco parámetros y muestra cuántos animales quedan. Se pueden ajustar los días o apagar un parámetro con el interruptor.
3. **Empezar chequeo**. La lista queda ordenada descendente por estado: VACÍA, PRONT, PREÑ, INSEM, ALTA, ABORT.
4. En el corral, por cada animal se toca **Preñada** o **Vacía** y se escribe comentario si hace falta. Se guarda solo en cada toque; no hay botón de guardar y no se pierde nada al cerrar.
5. Si llega una vaca que no estaba en la lista: **Agregar animal que no está en la lista**.
6. Al terminar: **Guardar y enviar**. Si no hay señal queda pendiente y se manda solo cuando la tablet vuelva a tener internet. El botón **Excel** descarga el archivo con todas las columnas originales más el resultado y el comentario.

El archivo se puede cargar tanto en la oficina como en la tablet; funciona igual en las dos.

### Preparar el chequeo desde el computador

Si prefieres armar tú la lista y que el encargado solo reciba el trabajo listo:

1. En el computador, **Nuevo chequeo** → cargas el Excel → ajustas los parámetros.
2. En vez de "Empezar chequeo", tocas **📤 Enviar a la tablet**. La lista viaja a la hoja de Google.
3. En la tablet, con internet, en la pantalla de inicio aparece arriba un aviso verde: *Chequeo preparado · 43 animales*, con botón **Descargar**.
4. Al descargarla, la lista queda guardada en la tablet y el aviso desaparece. De ahí en adelante todo funciona sin señal, igual que siempre.

La lista llega con todas las columnas originales del archivo, así que "Ver todos los datos" y el Excel final salen completos. Las listas preparadas que nadie recoge se borran solas a los 30 días.

Este paso sí necesita internet en los dos aparatos, porque la lista viaja por la hoja. Si ese día no hay señal, carga el Excel directamente en la tablet.

---

## Cosas que conviene saber

- **El chequeo no se pierde si se cierra la app.** Todo queda en la memoria de la tablet apenas se toca un botón. La pantalla de inicio muestra el historial con un punto de color: verde enviado, amarillo pendiente de enviar, gris sin empezar.
- **Estados sin regla.** Al cargar el archivo, si aparecen estados que no encajan en ningún parámetro (por ejemplo SECA), la app lo dice con el conteo. Antes esos animales desaparecían en silencio.
- **Tildes y variantes.** VACIA, VACÍA, PREÑ, PREÑADA, ABORT y ABORTO se reconocen igual.
- **Los ceros funcionan.** Si escribes 0 en un parámetro, vale 0. La versión anterior lo cambiaba por el valor por defecto sin avisar.
- **El Excel exportado sale con DEL y DUS como números**, así que se puede ordenar y filtrar en Excel.
- **Respaldo.** Aunque esté la hoja de Google, no está de más descargar el Excel al final de cada chequeo y guardarlo en Drive.

## Si algo falla

| Problema | Qué hacer |
|---|---|
| "Faltan columnas en el archivo" | El listado de DairyPlan debe traer STATUS, DEL y DUS, y los encabezados deben estar en la primera fila. Quita filas de título encima. |
| El botón de enviar no hace nada | Falta pegar la dirección de la hoja en el engranaje ⚙. |
| Dice que no responde la hoja | El Apps Script debe estar implementado con acceso para *cualquier usuario*. Revisa que la dirección termine en `/exec`. |
| La app no abre sin internet | Se abrió desde el navegador y no desde el ícono instalado, o la primera visita no alcanzó a descargar todo. Ábrela una vez con wifi. |
| Quedaron chequeos sin enviar | Engranaje ⚙ → **Enviar chequeos pendientes**. |
