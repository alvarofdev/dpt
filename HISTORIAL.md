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
