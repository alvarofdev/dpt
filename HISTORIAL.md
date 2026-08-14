# Historial de Modificaciones — Sistema de Inventario

---

## 2026-07-14

### inventario_android.html

#### Validaciones de codigo de articulo
- Se agrega validacion en `agregarLote()`: si el codigo empieza con `DEP` (mayusculas o minusculas) se rechaza con mensaje "Codigo de UBICACION escaneado".
- Se agrega validacion en `agregarLote()`: si el codigo tiene mas de 9 caracteres se rechaza con mensaje "Codigo invalido: debe tener 9 digitos o menos (tiene N)." — reemplaza la validacion anterior que solo rechazaba longitud exacta de 10.

#### Feature: listar lotes ya registrados al abrir una ubicacion
- Nueva funcion `mostrarLotesExistentes(ub)`: filtra `getData().registros` por `id_ubicacion` y muestra la lista en un panel verde en la pantalla de rollos.
- El panel aparece automaticamente al confirmar la ubicacion (`confirmarUbicacion`) y al recuperar sesion guardada (`goRollosConSesion`).
- Si la ubicacion no tiene rollos registrados, el panel permanece oculto.

#### Boton eliminar en lista de lotes ya registrados
- Cada item del panel verde (lotes existentes) tiene un boton rojo "✕ Eliminar".
- Nueva funcion `eliminarLoteExistente(idLote)`: pide confirmacion y elimina el registro de `localStorage`, luego refresca la lista.

#### Boton eliminar en tanda en curso
- `slistHTML()` actualizado: cada rollo de la tanda actual muestra un boton rojo "✕".
- Nueva funcion `eliminarDeTanda(idx)`: elimina el rollo de `S.pending` y de `S.ses`, actualiza contador y lista, guarda sesion. Sin confirmacion (los items de la tanda no estan guardados aun).

#### Nombre del archivo JSON exportado
- El nombre del archivo descargado ahora incluye fecha y hora completa con minutos y segundos.
- Formato: `inventario_export_YYYY-MM-DD_HH-MM-SS.json`
- Ejemplo: `inventario_export_2026-07-14_10-30-45.json`

---

## 2026-07-14 (continuacion)

### inventario.py + inventario_android.html

#### Formato de ubicacion ampliado (D077, D099, D100, DEP1, DEP2, DEV1)
- La validacion de ubicacion ahora acepta cualquier codigo que empiece con `"D"` en lugar de solo `"DEP1"` o `"DEP2"`.
- **inventario.py** — 3 lugares actualizados:
  - `RegistrarRolloScreen` paso 1: `startswith(("DEP1","DEP2"))` → `startswith("D")`
  - `NuevoInventarioScreen` paso 2: idem
  - `NuevaUbicacionScreen` campo deposito: `not in ("DEP1","DEP2")` → `not startswith("D")`
- **inventario_android.html** — 3 lugares actualizados:
  - `confirmarUbicacion()`: `startsWith('DEP1') || startsWith('DEP2')` → `startsWith('D')`
  - `agregarLote()` deteccion de ubicacion escaneada como articulo: `startsWith('DEP')` → `startsWith('D')`
  - Placeholder del input de ubicacion actualizado con los nuevos formatos

---

## 2026-07-14 (continuacion)

### generar_tutorial_resumido.py (nuevo) + Tutorial_Resumido_Tablet.docx (generado)

#### Tutorial resumido para operarios
- Nuevo script `generar_tutorial_resumido.py` que genera `Tutorial_Resumido_Tablet.docx`.
- Contenido: 6 secciones — Acceso, Registrar Rollo, Inventario Mensual, Exportar a PC, Tabla de codigos validos, Errores frecuentes.
- Incorpora todas las funcionalidades nuevas: lista de lotes existentes al abrir ubicacion, botones eliminar en tanda y en lista existente, validaciones de codigo (≤9 chars, no empieza con D), nombre del JSON con hora.
- Formato compacto pensado para imprimir en 1-2 hojas A4.

---

## 2026-08-13

### PRD.md y HISTORIAL.md

#### Simplificacion de correcciones durante la carga
- Se elimina la pregunta pendiente sobre quién puede corregir durante la carga de rollos y la autorización especial asociada.
- Se establece que las correcciones están permitidas mediante el permiso operativo normal del usuario autenticado, con observación obligatoria y auditoría.
- Se simplifica la matriz de permisos, reglas, requisitos, criterios de aceptación, riesgos y decisiones pendientes para no agregar aprobaciones, workflows ni permisos complejos no solicitados.
- Se mantienen como pendientes reales la matriz de transiciones de estados, el formato exacto de códigos y la pregunta 12 sobre columnas y mapeo del Excel.

### PRD.md y HISTORIAL.md

#### Reglas confirmadas sobre ubicaciones, códigos, movimientos y correcciones
- Se elimina el concepto de ubicación genérica: toda ubicación es física y se identifica por depósito, pasillo, columna y nivel; las ubicaciones no se pueden desactivar.
- Se documentan los ejemplos `FJSB754-4U` para rollo y `DEP1 C2 46 03` para ubicación. Se proponen normalización, estructura, referencias, unicidad y mensajes por campo, dejando pendientes las longitudes, rangos y separadores definitivos.
- `expedicion` queda como estado separado de la ubicación. Se mantienen los estados confirmados y pendiente la matriz de transiciones.
- Se confirman los tipos de movimiento `alta`, `traslado` y `baja`. Se permiten correcciones con observación obligatoria y usuario autenticado normal; no se prevé un permiso especial para corregir durante la carga de rollos.
- La pregunta 12 sobre columnas del Excel real y su mapeo queda explícitamente en stand-by.
- Se actualizan resumen, alcance, dominio, flujos, importación, reglas, permisos, requisitos, criterios de aceptación, riesgos y preguntas abiertas del PRD.

### PRD.md

#### Decisiones de dominio sobre rollos, estados e inventario
- Se define que la cantidad existente en DPT incluye todos los rollos registrados, sin filtrar por estado; el estado se muestra como dato adicional.
- Se documenta que un rollo puede cambiar de ubicación múltiples veces, incluido el reacomodamiento para aprovechar una ubicación con lugar libre, y que el cambio de ubicación puede ocurrir sin cambio de estado.
- Se confirman como estados válidos: `disponible`, `reservado`, `picking`, `expedicion`, `despachado`, `devuelto` y `bloqueado`.
- Se actualizan el resumen ejecutivo, objetivos e indicadores, conceptos, flujos, consulta de inventario, reglas, requisitos, criterios de aceptación, riesgos y preguntas del PRD.
- Se mantiene pendiente la matriz de transiciones válidas entre estados y no se agregan reglas no confirmadas sobre la relación entre estados y ubicaciones.

### PRD.md

#### Migracion del producto a sistema web
- Se reemplaza el PRD anterior, orientado a CLI/offline, por un PRD para una aplicacion web con React, Express y PostgreSQL.
- Se define Excel como unico mecanismo de carga o migracion inicial y se elimina JSON del alcance.
- Se incorporan consulta de ubicacion actual, cantidades en DPT, usuarios, roles y permisos.
- Se modelan rollo, ubicacion fisica, ubicacion generica, movimiento y estado, con trazabilidad completa de movimientos y soporte para expedicion.
- Se agregan flujos, reglas de negocio, requisitos funcionales y no funcionales, criterios de aceptacion, riesgos y preguntas abiertas.
- Los formatos de codigos, catalogos y reglas no confirmados quedan marcados como configurables o pendientes.

#### Exportacion operativa a Excel
- Se agrega la exportacion operativa desde PostgreSQL a Excel, diferenciada de la importacion inicial.
- Se definen permisos, filtros, alcance, datos minimos, fecha/hora y usuario de generacion, version de formato y auditoria.
- Se aclara que Excel no es la fuente de verdad y que editar una exportacion no modifica PostgreSQL ni genera movimientos; cualquier reimportacion futura debera ser un flujo separado y autorizado.
- Se incorporan requisitos funcionales, criterios de aceptacion, riesgos y preguntas abiertas asociados a la exportacion.

---

## 2026-08-13 (continuacion)

### openspec/config.yaml, .atl/skill-registry.md y Engram

#### Inicializacion del contexto SDD
- Se inicializa SDD en modo hibrido, con artefactos OpenSpec y persistencia Engram.
- Se documentan el stack actual Python/Textual/openpyxl, el stack objetivo React/Express/PostgreSQL y las convenciones del proyecto.
- Se detecta pytest como runner para pruebas unitarias e integracion y se activa strict TDD por ausencia de una configuracion explicita.
- Se actualiza el registro de skills y se persiste el contexto del proyecto, las capacidades de testing y el registro en Engram.

---

### Base minima de migracion web TypeScript

#### Nueva estructura frontend/backend y PostgreSQL
- Se crean `frontend/package.json`, `frontend/package-lock.json`, `frontend/tsconfig.json`, `frontend/tsconfig.app.json`, `frontend/tsconfig.node.json`, `frontend/vite.config.ts`, `frontend/.env.example`, `frontend/index.html`, `frontend/src/main.tsx`, `frontend/src/App.tsx`, `frontend/src/vite-env.d.ts` y `frontend/src/styles.css` con una app React + Vite + TypeScript y pantalla de estado en espanol.
- Se crean `backend/package.json`, `backend/package-lock.json`, `backend/tsconfig.json`, `backend/.env.example` y `backend/src/index.ts` con Express + TypeScript, CORS configurable, JSON parsing, pool `pg` y `GET /api/health` tolerante a PostgreSQL no disponible.
- Se crean `docker-compose.yml` con PostgreSQL, volumen persistente, healthcheck y defaults compatibles, `.gitignore` para artefactos Node y `README.md` con instrucciones de instalacion y ejecucion.
- No se modifican `PRD.md`, la aplicacion Python/HTML legacy, assets generados ni tests existentes.

---
