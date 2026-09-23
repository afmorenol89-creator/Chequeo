# Especificación: App de la Finca La Morelia — herramienta "Lotes de ordeño"

Este archivo va en la raíz del repositorio `Chequeo`. Claude Code debe leerlo antes de trabajar en la herramienta de lotes.

## 1. Objetivo

Convertir la app actual (chequeo reproductivo) en la **app de la finca**: un menú de inicio con varias herramientas. La segunda herramienta es **Lotes de ordeño**:

- Andrés carga desde un Excel la lista de animales en ordeño.
- Los trabajadores, desde la tablet, asignan cada animal a un lote (1, 2 o 3) y pueden cambiarlo de lote.
- Cualquier trabajador puede **ingresar animales nuevos** (número y nombre) y **sacar animales** de ordeño (secado, venta, muerte, otro).
- Todo funciona **sin internet** y se sincroniza con Google Sheets cuando hay señal.
- La usan campesinos en una tablet Android: botones grandes, pocas pantallas, confirmación en toda acción importante.

## 2. Estructura del repositorio

```
Chequeo/
├── index.html          ← menú de inicio (botones: Chequeo reproductivo, Lotes de ordeño)
├── sw.js               ← service worker; cachea TODAS las herramientas
├── manifest.json
├── comun/
│   ├── estilos.css     ← estilos compartidos (botones grandes, colores)
│   ├── almacen.js      ← guardado local (IndexedDB) + cola de pendientes
│   ├── sync.js         ← envío/lectura contra el Apps Script
│   └── usuario.js      ← selección de "¿quién eres?" (se guarda en el dispositivo)
├── chequeo/            ← la app actual, movida aquí SIN cambiar su funcionamiento
├── lotes/
│   ├── index.html
│   └── lotes.js
└── apps-script/        ← copia de referencia del código de Google (se pega a mano)
```

## 3. Datos en Google Sheets (misma hoja que el chequeo)

### Pestaña `Animales_Ordeno`
| Columna | Descripción |
|---|---|
| numero | Número del animal (entero, clave única) |
| nombre | Nombre (puede estar vacío). Sin espacios al inicio ni al final |
| estado | `activo` o `fuera` |
| origen | `excel` o `manual` |
| fecha_entrada | Fecha y hora en que entró a ordeño |
| ingresado_por | Nombre del trabajador (si origen = manual) o "Excel" |
| fecha_salida | Vacío si está activo |
| motivo_salida | `Secado`, `Venta`, `Muerte` u `Otro: <texto>` |
| sacado_por | Quién lo sacó |

Nunca se borran filas. Un animal que sale queda con `estado = fuera`.

### Pestaña `Movimientos`
Registro que **solo crece**. El lote actual de cada animal es su último movimiento.

| Columna | Descripción |
|---|---|
| id | Identificador único generado en la tablet (evita duplicados al reintentar) |
| fecha_hora | Momento en que se hizo en la tablet (no cuando se sincronizó) |
| numero | Animal |
| lote_anterior | Vacío si no tenía |
| lote_nuevo | Lote asignado |
| usuario | Quién lo hizo |
| dispositivo | Identificador de la tablet |

Si dos tablets mueven el mismo animal, **gana el movimiento con fecha_hora más reciente**.

### Pestaña `Config_Lotes`
Lista de lotes disponibles (inicialmente 1, 2, 3). Agregar un lote aquí lo hace aparecer en la app sin tocar código.

### Pestaña `Trabajadores`
Lista de nombres para el selector "¿quién eres?".

## 4. Formato del Excel de entrada

- Una hoja, dos columnas con encabezado: `Número de animal` y `Nombre`.
- Los nombres traen espacios de relleno al final: **aplicar trim** a número y nombre.
- Validar: números enteros, sin repetidos. Si hay error, mostrar la fila y no cargar nada.
- Leer con SheetJS desde cdnjs.

## 5. Pantallas

### 5.1 Inicio de la herramienta
- Si el dispositivo no tiene usuario, pedir "¿Quién eres?" (lista de `Trabajadores`) antes de continuar.
- Indicador de sincronización siempre visible: "Todo al día" / "3 cambios pendientes por subir" / "Sin conexión".

### 5.2 Lista principal
- Arriba, bloque destacado **Sin lote (N)** si hay animales sin asignar.
- Debajo, los animales agrupados por lote, con contador por lote.
- Buscador por número o nombre.
- Cada animal se muestra como "16 · Vanidosa".

### 5.3 Ficha del animal (al tocarlo)
- Botones grandes con los lotes. Si ya tenía lote: confirmar "¿Mover 16 Vanidosa del Lote 1 al Lote 2?".
- Botón **Sacar de ordeño** → motivo con botones grandes: **Secado · Venta · Muerte · Otro** (Otro abre un campo de texto) → confirmar "¿Sacar a 16 Vanidosa por secado?".
- Historial corto: últimos movimientos del animal.

### 5.4 Modo rápido
- Elegir primero el lote destino y luego tocar varios animales. Cada toque genera un movimiento. Botón "Terminar".

### 5.5 Nuevo animal (disponible para todos)
- Campos: número (obligatorio), nombre (opcional), lote (opcional).
- Si el número existe y está **activo**: avisar y no duplicar.
- Si el número existe y está **fuera**: ofrecer **Reactivar** (conserva su historial) en vez de crear otro.

### 5.6 Salidas recientes
- Lista de animales sacados en los últimos 30 días, con motivo y quién lo hizo, y un botón **Reactivar** para corregir errores. Como todos pueden sacar animales, esta pantalla es la red de seguridad.

### 5.7 Cargar Excel
- Seleccionar el archivo → pantalla de revisión **antes** de guardar:
  - **Nuevos en el Excel** → entran activos, sin lote.
  - **Activos que no vienen en el Excel** → NO salen solos. Por cada uno: escoger motivo de salida o "Mantener".
  - **Ingresados a mano que no están en el Excel** → se muestran aparte con quién y cuándo los ingresó; confirmar si se quedan.
  - **Sin cambio** → solo el conteo.
- Regla: **la app nunca saca un animal sin una decisión explícita con motivo.**

## 6. Funcionamiento sin internet

- Al abrir con señal, descargar animales, movimientos recientes, lotes y trabajadores, y guardarlos en IndexedDB.
- Toda acción se guarda primero en local y entra a una cola de pendientes.
- La cola se envía al recuperar señal y también con un botón "Sincronizar ahora".
- Cada envío lleva su `id`; el Apps Script ignora ids ya recibidos, así que reintentar es seguro.

## 7. Apps Script

Ampliar el script existente (sin romper lo del chequeo) con:

- `GET ?accion=lotes_estado` → animales activos, lote actual de cada uno, lotes, trabajadores, salidas recientes.
- `POST accion=movimientos` → lista de movimientos (asignar / mover).
- `POST accion=animal_nuevo`, `accion=animal_salida`, `accion=animal_reactivar`.
- `POST accion=cargar_excel` → recibe las decisiones ya tomadas en la pantalla de revisión.

El código del Apps Script se guarda en `apps-script/` como referencia, pero **se pega a mano en Google**; no se publica solo.

## 8. Reglas para Claude Code

- Cada vez que cambie cualquier archivo de la app, **subir la versión del caché en `sw.js`**.
- No cambiar el funcionamiento del chequeo reproductivo al mover su carpeta.
- Trabajar por fases, probar cada una y subirla antes de seguir.
- Textos de la app en español, sencillos.

## 9. Fases de construcción

1. Mover el chequeo a `chequeo/`, crear el menú de inicio y `comun/`. Verificar que el chequeo funcione igual, también sin internet.
2. Crear las pestañas en Google Sheets y las acciones nuevas del Apps Script.
3. Pantalla de carga del Excel con su revisión.
4. Lista principal, ficha del animal y modo rápido.
5. Nuevo animal, sacar de ordeño y salidas recientes.
6. Prueba completa sin internet en la tablet y actualización de `CLAUDE.md`.

## 10. Lista de prueba

- [ ] El chequeo reproductivo sigue funcionando desde el menú.
- [ ] Cargar el Excel de 91 animales: todos entran sin lote y sin espacios en el nombre.
- [ ] Asignar, mover y usar el modo rápido sin internet; al volver la señal, todo sube sin duplicados.
- [ ] Ingresar un número que ya existe activo → aviso, sin duplicar.
- [ ] Sacar un animal por secado y reactivarlo desde Salidas recientes.
- [ ] Cargar un Excel al que le falta un animal → la app pide decisión, no lo saca sola.
- [ ] Dos tablets mueven el mismo animal → queda el movimiento más reciente.
