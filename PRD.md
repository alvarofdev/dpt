# PRD — Sistema de Control de Almacén

## 1. Problema

El sistema de escáner de código de barras del almacén dejó de funcionar (sin red, sin wifi). Se necesita una solución offline rápida para que la producción no se detenga.

## 2. Objetivo

Script CLI que permita a los operarios registrar la ubicación de rollos de tela en un almacén con ~12,000 productos, usando tablet + escáner Bluetooth. El sistema debe compilarce en un ejecutable fácil de distribuir en tablets.

## 3. Alcance

### IN

- Registro de ubicación de rollos (escaneo de ubicación + escaneo de rollos)
- Validación de unicidad de códigos de rollo
- Validación de formato de códigos (ubicación y rollo)
- Bloqueo de re-ubicación (un rollo no se puede mover)
- Fecha/hora de cada registro para trazabilidad
- Exportación a Excel (mismo formato actual)
- **Compilación a ejecutable**: empaquetado con PyInstaller para fácil distribución en tablets
- Tests automatizados

### FUERA DE ALCANCE

- Búsqueda de rollos
- Conteo por ubicación
- Reportes/resúmenes
- Base de datos de productos
- Interfaz gráfica (se mantiene CLI táctil)

## 4. Estructura del almacén

```
DEP1A14702
│    │  │  │
│    │  │  └─ Nivel (2 chars)
│    │  └─── Columna (2 chars, par/impar)
│    └───── Pasillo (2 chars)
└────────── Depósito (4 chars)
```

## 5. Flujo del usuario

```
1. Ingresar nombre de usuario
2. [LOOP] Escanear ubicación (10 chars)
   │  → Validar formato
   │  → Registrar en Excel
   └─→ [LOOP] Escanear rollo (7 chars)
         │  → Validar formato
         │  → Validar unicidad (no repetido en sesión)
         │  → Validar que no exista ya en el sistema
         │  → Registrar con fecha/hora
         └─→ Si ingresa 0, volver a paso 2
```

## 6. Datos por registro

### Tabla `ubicacion`

| Columna  | Tipo   | Descripción         |
|----------|--------|---------------------|
| id       | string | Código completo (10 chars) |
| depo     | string | Depósito            |
| pasillo  | string | Pasillo             |
| columna  | string | Columna             |
| nivel    | string | Nivel               |

### Tabla `articulos`

| Columna      | Tipo     | Descripción                    |
|--------------|----------|--------------------------------|
| id_usuario   | string   | Quién registró                 |
| id_ubicacion | string   | Ubicación del rollo            |
| id_lote      | string   | Código del rollo (7 chars)     |
| fecha_hora   | datetime | Fecha y hora del registro      |

## 7. Validaciones

- **Ubicación**: exactamente 10 caracteres
- **Rollo**: exactamente 7 caracteres
- **Unicidad de rollo**: no puede escanear el mismo código dos veces en la misma sesión
- **Re-ubicación**: si el rollo ya existe en Excel, rechazar con mensaje

## 8. Criterios de aceptación

1. El operario puede ingresar su nombre al inicio
2. Puede escanear ubicaciones y rollos de forma continua
3. Si escanea un rollo duplicado en la sesión, se rechaza
4. Si el rollo ya existe en el sistema (Excel), se rechaza
5. Cada registro tiene fecha/hora
6. El Excel se genera con el mismo formato actual
7. El script se compila a ejecutable (.exe) para distribución en tablets
8. Los tests pasan al 100%

## 9. Compilación

El script se empaqueta con **PyInstaller** en un ejecutable standalone:

```bash
pyinstaller --onefile main.py --name "ControlAlmacen"
```

Cada tablet recibe el ejecutable. No necesita Python instalado.

> **NOTA**: La gestión de espacios/zonas por tablet se implementará cuando se conozcan los nombres de los pasillos.
