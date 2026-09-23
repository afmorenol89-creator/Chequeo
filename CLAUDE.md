# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## What this is

A Spanish-language PWA ("Chequeo Reproductivo") for recording pregnancy checks (preñada/vacía/aborto)
on a dairy farm's cattle herd, usable offline in the corral on a tablet. There is no build system,
no package manager, and no test framework — this is hand-written static HTML/CSS/JS plus a Google
Apps Script backend, deployed as-is.

- `index.html` — the entire app: markup, CSS, and JS in one file.
- `sw.js` — service worker (offline caching for the tablet).
- `manifest.webmanifest` — PWA install metadata.
- `xlsx.full.min.js` — vendored SheetJS, used to read/write Excel files client-side.
- `apps-script.gs` — Google Apps Script Web App source. **Not deployed by this repo** — see below.
- `INSTALACION.md` — end-user (Spanish) setup guide; read it before changing the deploy/setup flow.

## Commands

There is nothing to install, build, lint, or test. To work on the app:

- **Preview**: open `index.html` directly in a browser (or serve the folder statically). Note that
  `file://` origins block IndexedDB in some sandboxed browser contexts — if `abrirDB()` throws
  `SecurityError`, that's an environment limitation, not a bug; verify logic by calling functions
  from the devtools console instead (see Verification below).
- **Deploy the web app**: `git push origin main`. GitHub Pages serves directly from the repo root
  on `main` — there is no CI/build step. Live at the URL under the repo's GitHub Pages settings.
- **Deploy the Sheets backend**: pasting `apps-script.gs` into Google's Apps Script editor and
  redeploying is a *manual, separate* action the user does in their Google account — `git push`
  does **not** propagate this file anywhere. Whenever `apps-script.gs` changes, tell the user to
  paste it in and do Implementar → Administrar implementaciones → Versión nueva → Implementar.
- **Verification**: with no test suite, changes are verified by loading `index.html` in a browser
  and driving it from the console — inject fake `sesion`/`sesiones` objects, stub `confirm`,
  `toast`, `pedir` (the fetch-to-Apps-Script helper) and `dbBorrar`/`cargarSesiones` to isolate
  logic from real network calls and from IndexedDB.

## Architecture

### Client (`index.html`)

No framework: a hand-rolled SPA. `<section id="v-*">` blocks are toggled via the `.oculto` class by
`ver(nombreVista)`; `pila` is the back-navigation stack (`irAtras()` pops it). Global mutable state:
`sesion` (the currently open chequeo, if any), `sesiones` (cached list of all saved chequeos),
`config` (`urlSheet`, `operario`), `filtro`/`vista`/`pila`.

Data lives in IndexedDB (`chequeo-repro` DB, stores `sesiones` and `config`) — persistent per device,
never synced except through the explicit "send" flow below. Data model:

- **sesión**: `{id, creado, nombre, archivo, heads, colId/colStatus/colDel/colDus, operario, animales,
  enviado, enviadoEn, reabierto, reabiertoEn}`. `enviado === true` means the session is **locked**
  (read-only in modo campo) because it's already on the Google Sheet — editing is blocked both in
  the UI (`disabled`/`readonly`) and defensively inside `marcar()`/`anotar()`/`agregarAnimal()`.
  `reabrirSesion()` is the only way to unlock it; it sets `reabierto = true` permanently (this flag
  is never cleared, even after the session is re-sent and re-locked).
- **animal**: `{id, num, estado, del, dus, motivos, datos, resultado, comentario, agregado, hora}`.
  `resultado` is `null | 'PREÑADA' | 'VACIA' | 'ABORTO'`. The Aborto button only renders when
  `esPrenada(a)` (i.e. `norm(a.estado)` starts with `pren`) — matching how `coincideEstado()` already
  does "starts with" state matching elsewhere.
- `norm()` strips accents and normalizes case everywhere state strings are compared, so
  `VACIA`/`VACÍA`, `PREÑ`/`PREÑADA`, etc. are treated as equal.

**Import flow**: an exported DairyPlan Excel is read with SheetJS; `STATUS`/`DEL`/`DUS` columns are
located by header name (`norm()`-insensitive), then `PARAMS_DEF` (5 rules: state prefix + a DEL/DUS
threshold) decides which animals need checking. This filtering logic is the core business rule of
the app — read it before changing which animals show up in a chequeo.

**Two ways a chequeo reaches the tablet**: load the Excel directly on the tablet, or load it on a
computer and use "Enviar a la tablet" (`preparar_`/`pendientes_`/`traer_`/`recogido_` in
`apps-script.gs`), which stages it through the Sheet so the tablet can pick it up later.

**Export/send**: `filasExport(s)` and `payload(s)` both take an *optional* session argument
(defaulting to the global `sesion`) — this lets read-only views (like the chequeo detail screen)
export/inspect a session other than the one open in modo campo, without touching modo campo's state.

### Service worker (`sw.js`)

Stale-while-revalidate cache of the app shell. **The `CACHE` constant (`chequeo-vN`) must be bumped
on every deploy that changes any cached file**, or devices that already installed the PWA won't pick
up the update (they'll keep serving the old cached version indefinitely).

### Google Sheets backend (`apps-script.gs`)

A single `doPost(e)` action dispatcher (`datos.accion`) standing in for a REST API, writing to named
sheet tabs it creates/repairs on demand via `hoja_()`:

- **Chequeos** — one row per animal, historical, across all sessions. Its header row is **not**
  hardcoded in the script — it's whatever `columnas` array the client sends in `payload()`
  (`index.html`), and `hoja_()` rewrites the header to match if it drifts.
- **Resumen** — one row per session with totals (counts by `RESULTADO`).
- **Preparados** / **Preparados_Index** — the "prepare on computer, pick up on tablet" handoff queue
  (auto-expires after 30 days via `limpiarViejos_`).
- **Papelera** — soft-delete destination for `borrar_`: rows are moved here (with a `BORRADO EL`
  timestamp) and removed from Chequeos/Resumen, never hard-deleted.

**Column-ordering convention**: because `Chequeos`' columns are driven by the client and old rows
are never rewritten, any new column must be appended at the **end** of the relevant `columnas`/
header array — never inserted in the middle — or every historical row silently shifts under the
wrong header. This has been the rule for every column added so far (`ABORTOS` in Resumen,
`REABIERTO`/`AGREGADO EN CAMPO` in Chequeos).

## Conventions

- All UI copy, code comments, and commit messages are in Spanish — match that.
- No new dependencies/build tooling without a strong reason — this project is intentionally a
  single static `index.html` plus a single Apps Script file, deployable by copy-paste.
