# PRD — Sistema Web de Inventario y Trazabilidad de Rollos

Este producto reemplaza el flujo offline por una aplicación web centralizada para consultar y operar el inventario de DPT. El producto conservará la gestión de rollos y ubicaciones, incorporará usuarios y permisos, y registrará la historia completa de cada rollo. La implementación prevista es React en frontend, Express en backend y PostgreSQL como base de datos.

## Resumen ejecutivo

| Decisión | Definición |
|---|---|
| Producto | Aplicación web online para inventario de rollos en DPT |
| Persistencia | PostgreSQL como fuente central de verdad |
| Frontend | React |
| Backend | Express |
| Carga inicial | Importación de Excel validada; no se utilizará JSON como mecanismo de carga o migración |
| Exportación operativa | Exportación de datos vigentes desde PostgreSQL a Excel, según permisos, filtros y alcance del usuario |
| Trazabilidad | Historial inmutable de movimientos por rollo, con usuario autenticado normal, fecha/hora, estado y ubicación física de origen/destino |
| Cantidad existente en DPT | Incluye todos los rollos registrados en DPT, sin filtrar por estado; el estado se muestra como dato adicional |
| Ubicación | Siempre física; se identifica por depósito, pasillo, columna y nivel |
| Ejemplos de códigos | Rollo: `FJSB754-4U`; ubicación: `DEP1 C2 46 03` |
| Códigos | Se validan estructura, referencias y unicidad; longitudes y representación definitiva quedan pendientes donde no fueron confirmadas |

## 1. Problema

El sistema actual funciona de manera offline y distribuye la información entre archivos y dispositivos. Esto dificulta:

- Consultar en un único lugar la ubicación actual de un rollo.
- Conocer cantidades existentes en DPT con información consistente.
- Coordinar operaciones de múltiples usuarios.
- Reconstruir todos los cambios que atravesó un rollo.
- Separar el estado `expedicion` de la ubicación física del rollo.
- Administrar accesos y responsabilidades sobre los datos.

La migración debe mantener la operación de inventario, pero cambiar el modelo de persistencia: el estado actual se consulta desde una base central y la historia se conserva como trazabilidad, no como una simple sobrescritura.

## 2. Objetivos

### Objetivos del producto

- Centralizar el inventario de DPT en una aplicación web.
- Permitir buscar un rollo y consultar su ubicación actual, estado y datos relevantes.
- Permitir consultar la cantidad existente en DPT contando todos los rollos, sin excluirlos por estado, y mostrar el estado como dato adicional.
- Registrar movimientos de rollos con trazabilidad completa.
- Representar ubicaciones físicas identificadas por depósito, pasillo, columna y nivel.
- Gestionar usuarios, roles y permisos.
- Migrar los datos iniciales desde Excel mediante un proceso controlado y auditable.
- Permitir exportar datos operativos a Excel para que otros operarios puedan analizarlos o trabajarlos fuera del sistema, sin convertir el archivo en fuente de verdad.
- Reducir errores operativos mediante validaciones y confirmaciones claras.

### Indicadores de resultado

- Cada rollo registrado tiene un único estado y una ubicación física actual identificable.
- La cantidad existente en DPT cuenta todos los rollos registrados, sin filtrar por estado.
- Cada cambio aceptado de ubicación o estado genera un movimiento asociado al usuario y a una fecha/hora.
- Las consultas de inventario reflejan la información de PostgreSQL, sin depender de archivos locales.
- Las importaciones informan errores por fila y no incorporan datos inválidos silenciosamente.

## 3. Alcance

### Incluido

- Inicio de sesión y gestión de usuarios.
- Roles y permisos configurables según las responsabilidades aprobadas.
- Alta, consulta y actualización controlada de rollos.
- Consulta de ubicación actual y estado de cada rollo.
- Registro de movimientos entre ubicaciones y cambios de estado.
- Ubicaciones físicas identificadas por depósito, pasillo, columna y nivel; `expedicion` se modela únicamente como estado.
- Consultas y conteos del inventario existente en DPT.
- Historial de movimientos por rollo.
- Importación inicial y eventualmente controlada desde Excel.
- Exportación operativa a Excel con filtros, alcance y permisos.
- Validaciones de códigos, integridad, duplicados, referencias y transiciones de estado.
- Auditoría de acciones relevantes.

### No incluido en esta versión

- Uso como aplicación offline o sincronización entre dispositivos sin conexión.
- Persistencia o migración mediante JSON.
- Definición de un formato de códigos que aún no haya sido confirmado por el negocio.
- Automatización de expedición, transporte o integración con sistemas externos no especificados.
- Gestión completa de compras, ventas, producción, facturación o planificación logística.
- Aplicación móvil nativa; el alcance inicial es una aplicación web responsive.
- Reportes avanzados y tableros analíticos no priorizados.

## 4. Usuarios y roles

Los nombres de roles son una propuesta inicial. La matriz final debe validarse con el responsable del negocio.

| Rol | Responsabilidad principal | Capacidades esperadas |
|---|---|---|
| Administrador | Configuración y control del sistema | Gestionar usuarios, roles, catálogos, permisos e importaciones; consultar y auditar toda la información |
| Supervisor | Control operativo del inventario | Consultar inventario, ejecutar movimientos operativos, revisar errores y trazabilidad |
| Operador | Registro diario de operaciones | Consultar lo necesario para operar y registrar movimientos autorizados |
| Consulta | Lectura de información | Consultar inventario, ubicación y estado sin modificar datos |

El producto debe permitir desactivar usuarios sin borrar su historial. Toda acción de negocio debe quedar asociada a la identidad autenticada que la ejecutó.

## 5. Conceptos de dominio

### Rollo

Unidad inventariable identificada por un código de rollo. Tiene una ubicación física actual, un estado actual y una colección ordenada de movimientos. Puede cambiar de ubicación múltiples veces, incluso para reacomodarse y aprovechar una ubicación con lugar libre. Un ejemplo válido es `FJSB754-4U`: se espera una estructura normalizada en mayúsculas, compatible con letras, dígitos y guion. La longitud exacta, la cantidad de bloques y si existen excepciones requieren confirmación.

### Ubicación

Lugar físico donde puede encontrarse un rollo. No existen ubicaciones genéricas. Toda ubicación se identifica mediante cuatro componentes: depósito, pasillo, columna y nivel. El ejemplo `DEP1 C2 46 03` documenta esos componentes en ese orden; el uso de espacios, guiones u otro separador y la representación canónica del código quedan pendientes de confirmación.

Cada ubicación debe tener una combinación única de depósito, pasillo, columna y nivel, referencias existentes y un código normalizado. Las ubicaciones no se pueden desactivar. No se deben inventar longitudes, rangos o separadores definitivos a partir del ejemplo.

### Movimiento

Evento que registra una operación sobre el rollo. Los tipos confirmados son `alta`, `traslado` y `baja`. Un traslado registra el cambio de ubicación, de estado o ambos; una corrección permitida debe conservar el motivo y una observación obligatoria. Como mínimo debe conservar:

- Rollo afectado.
- Fecha y hora del evento.
- Usuario que lo ejecutó.
- Ubicación de origen, si existe.
- Ubicación de destino, si existe.
- Estado anterior y estado posterior, si cambia.
- Tipo de movimiento: `alta`, `traslado` o `baja`.
- Motivo y observación; la observación es obligatoria para correcciones.

Los movimientos aceptados son históricos y no deben editarse para corregir el inventario actual. Se permiten correcciones dentro de la operación normal, requieren observación obligatoria y auditoría, y quedan asociadas al usuario autenticado que las ejecuta. No existe un permiso especial ni una lógica particular para corregir mientras se cargan rollos.

### Estado

Situación operativa del rollo. Los estados válidos confirmados son: `disponible`, `reservado`, `picking`, `expedicion`, `despachado`, `devuelto` y `bloqueado`. El estado actual forma parte del modelo del rollo y se conserva junto con su ubicación física actual; ambos deben mostrarse en las consultas. La matriz de transiciones válidas entre estados queda pendiente de definición. `expedicion` es un estado, no una ubicación.

## 6. Flujos principales

### 6.1 Ingreso y consulta

1. El usuario inicia sesión.
2. El sistema aplica sus permisos.
3. El usuario busca un rollo por su código normalizado u otros filtros autorizados.
4. El sistema muestra estado actual, ubicación actual, fecha del último movimiento y datos disponibles.
5. El usuario puede abrir la trazabilidad completa si su rol lo permite.

### 6.2 Registro de movimiento

1. El usuario selecciona o escanea un rollo.
2. El sistema muestra el estado y la ubicación actuales.
3. El usuario selecciona el destino físico, el nuevo estado y el motivo que correspondan. El destino puede cambiar aunque el estado permanezca igual.
4. El sistema valida permisos, existencia de referencias, transición de estado cuando corresponda y ausencia de conflicto de concurrencia.
5. El usuario confirma la operación.
6. El sistema registra el movimiento y actualiza el estado derivado del rollo en una operación consistente.
7. El sistema muestra el resultado y deja disponible el nuevo historial.

No se permite registrar una ubicación genérica ni usar `expedicion` como ubicación. El movimiento siempre debe referenciar una ubicación física válida; `expedicion` se registra separadamente como estado cuando corresponda.

### 6.3 Consulta de inventario

1. El usuario abre la consulta de inventario.
2. Puede filtrar los resultados por estado, depósito, pasillo, columna, nivel u otros atributos disponibles.
3. El sistema muestra como cantidad existente en DPT todos los rollos registrados, sin filtrar por estado.
4. El estado y la ubicación se muestran como datos adicionales y los resultados pueden desglosarse mediante los filtros autorizados.

### 6.4 Importación inicial desde Excel

1. Un usuario autorizado descarga o consulta la plantilla vigente.
2. Carga un archivo Excel.
3. El sistema valida estructura de columnas, códigos, campos obligatorios, referencias, duplicados y consistencia de estados/ubicaciones.
4. El sistema muestra un resumen con filas válidas, filas rechazadas y advertencias.
5. El usuario autorizado confirma la importación.
6. El sistema incorpora solo los datos aprobados y registra quién, cuándo y qué archivo se procesó.
7. El sistema deja disponible un informe de resultados y errores.

La importación inicial debe poder establecer el inventario base. Si se requiere importar historial previo, se debe definir explícitamente la estructura de movimientos y cómo se distinguirá de los movimientos generados en el nuevo sistema.

### 6.5 Exportación operativa a Excel

1. Un usuario autenticado solicita una exportación desde una consulta o vista habilitada.
2. El sistema aplica los permisos del usuario, el alcance autorizado y los filtros seleccionados.
3. El sistema muestra el alcance y un resumen de los datos que serán exportados antes de generar el archivo, cuando el volumen o la operación lo requieran.
4. El usuario confirma la generación.
5. El sistema genera un archivo Excel con los datos vigentes consultados desde PostgreSQL.
6. El archivo incluye fecha y hora de generación, usuario que lo generó, filtros y alcance aplicados, además de los datos mínimos definidos para la operación.
7. El sistema registra la solicitud, el resultado, el usuario, la fecha/hora, los filtros, el alcance y la referencia del archivo generado.

La exportación es una copia de trabajo o consulta. Editar el archivo no modifica PostgreSQL ni genera movimientos. Si en el futuro se habilitara una reimportación o actualización desde Excel, deberá existir un flujo separado, explícito, validado y autorizado; nunca se debe interpretar automáticamente una exportación editada como una orden de actualización.

## 7. Importación Excel

- Excel es el único formato de archivo requerido para la carga o migración inicial.
- La plantilla debe documentar nombres de columnas, obligatoriedad, tipos, valores permitidos y ejemplos aprobados para códigos de rollo y ubicación.
- El sistema no debe aceptar que el usuario defina silenciosamente columnas equivalentes o mapeos ambiguos.
- La validación debe detectar códigos de rollo duplicados, ubicaciones inexistentes o duplicadas, estados no permitidos, campos vacíos y conflictos con datos ya cargados.
- Los códigos deben normalizarse para validar mayúsculas/minúsculas y luego comprobarse por campo con mensajes accionables. El código de rollo debe ser compatible con letras, dígitos y guion; la ubicación debe contener cuatro componentes: depósito, pasillo, columna y nivel. Las longitudes, rangos y separadores definitivos requieren confirmación.
- La importación debe ejecutarse en modo previsualización antes de persistir cambios.
- Los errores deben indicar fila, columna, valor recibido y causa corregible.
- Debe existir un registro de auditoría de cada intento y de cada importación confirmada.
- La política para reintentos, rollback y cargas parciales debe definirse antes de producción.
- No se requiere exportar datos a JSON ni usar JSON como formato intermedio de negocio.

### 7.1 Exportación operativa a Excel

- La exportación debe leer siempre el estado vigente desde PostgreSQL; Excel no es la fuente central de verdad.
- Solo pueden exportar quienes tengan el permiso correspondiente. El alcance puede restringirse por rol, ubicación, estado, turno u otra dimensión definida por el negocio.
- La exportación debe respetar los filtros seleccionados y no incluir registros fuera del alcance autorizado, aunque existan en PostgreSQL.
- Como mínimo, el detalle de rollos debe incluir código de rollo, estado actual, ubicación física actual, fecha/hora del último movimiento y los atributos operativos aprobados.
- Si se exporta trazabilidad, cada movimiento debe incluir como mínimo rollo, fecha/hora, usuario, origen, destino, estado anterior, estado posterior, tipo o motivo y observación o referencia cuando corresponda.
- El archivo y/o una hoja de metadatos debe informar fecha/hora de generación, usuario generador, filtros, alcance aplicado y versión de la plantilla o formato.
- El sistema debe advertir que el archivo es una instantánea operativa y que su edición no actualiza el sistema ni reemplaza el registro de movimientos.
- La exportación no debe generar movimientos ni alterar datos de inventario; solo debe dejar la auditoría de la solicitud y del archivo producido.
- Los errores de generación deben informarse sin presentar como válido un archivo incompleto, y deben quedar registrados para su seguimiento.

## 8. Reglas de negocio

1. Un rollo no puede tener más de una ubicación física actual efectiva.
2. Un rollo no puede tener más de un estado actual efectivo.
3. Todo movimiento confirmado debe conservar origen, destino, usuario y fecha/hora según corresponda.
4. No se debe sobrescribir ni eliminar el historial normal de movimientos.
5. No se puede mover un rollo inexistente; las ubicaciones no se desactivan.
6. No se puede registrar un destino que no exista o cuya combinación de depósito, pasillo, columna y nivel no sea válida.
7. Las transiciones de estado deben validarse contra la matriz de reglas aprobada, cuya definición queda pendiente.
8. Un cambio de ubicación puede registrarse sin cambio de estado.
9. Un rollo puede cambiar de ubicación múltiples veces; cada cambio confirmado debe quedar trazado, incluido un reacomodamiento para aprovechar una ubicación con lugar libre.
10. Toda ubicación debe cumplir los cuatro componentes físicos; `expedicion` no es una ubicación.
11. Las operaciones concurrentes deben impedir que dos usuarios confirmen movimientos incompatibles sobre el mismo rollo.
12. Los códigos, motivos y catálogos no definidos en este documento deben ser configurables o quedar pendientes de decisión; no deben codificarse como supuestos.
13. Las acciones administrativas requieren permisos según la matriz; las correcciones usan el permiso operativo normal del usuario autenticado y siempre deben dejar motivo y observación obligatoria.
14. Los códigos de rollo y ubicación deben normalizarse, validarse contra su estructura y referencias existentes, y ser únicos dentro de su entidad.
15. Los errores de código deben indicar el campo, el valor recibido y la causa corregible; no se deben fijar longitudes no confirmadas como regla definitiva.

## 9. Permisos

La autorización debe evaluarse en backend; ocultar controles en React no constituye una protección suficiente.

| Capacidad | Administrador | Supervisor | Operador | Consulta |
|---|---:|---:|---:|---:|
| Consultar inventario | Sí | Sí | Sí | Sí |
| Consultar trazabilidad | Sí | Sí | Según autorización | Según autorización |
| Registrar movimientos | Sí | Sí | Según alcance operativo | No |
| Ejecutar ajustes/correcciones | Sí | Sí | Sí | No |
| Importar Excel | Sí | Según autorización | No | No |
| Exportar inventario a Excel | Sí | Según autorización | Según autorización | Según autorización |
| Gestionar usuarios y roles | Sí | No | No | No |
| Gestionar catálogos | Sí | Según autorización | No | No |

La matriz es inicial y debe completarse solo con los permisos operativos que el negocio confirme. Las correcciones no tienen un permiso especial ni un flujo de aprobación: siguen el permiso normal para registrar movimientos y siempre requieren observación obligatoria. Toda acción queda asociada al usuario autenticado que la ejecuta.

## 10. Requisitos funcionales

- **RF-01**: El sistema debe autenticar usuarios y aplicar roles y permisos.
- **RF-02**: El sistema debe permitir gestionar usuarios sin eliminar la auditoría histórica.
- **RF-03**: El sistema debe permitir buscar rollos y consultar su ubicación y estado actuales.
- **RF-04**: El sistema debe permitir consultar la cantidad existente en DPT contando todos los rollos, sin filtrar por estado, y mostrar estado, ubicación y filtros de desglose autorizados.
- **RF-05**: El sistema debe registrar movimientos con trazabilidad completa.
- **RF-06**: El sistema debe soportar únicamente ubicaciones físicas identificadas por depósito, pasillo, columna y nivel, que no se puedan desactivar.
- **RF-07**: El sistema debe registrar y consultar el estado actual del rollo como parte de su modelo, separado de la ubicación física.
- **RF-08**: El sistema debe validar la matriz de transiciones aprobada cuando cambie el estado, además de códigos, duplicados, referencias y concurrencia, y debe permitir traslados sin cambio de estado.
- **RF-09**: El sistema debe importar datos iniciales desde Excel con previsualización, errores por fila y confirmación explícita.
- **RF-10**: El sistema debe auditar operaciones administrativas, importaciones, ajustes y movimientos.
- **RF-11**: El sistema debe mostrar la historia ordenada de movimientos de un rollo.
- **RF-12**: El sistema debe permitir configurar o administrar los catálogos aprobados sin fijar códigos no confirmados.
- **RF-13**: El sistema debe permitir exportar a Excel los datos operativos autorizados, aplicando permisos, filtros y alcance, e incluyendo metadatos de generación.
- **RF-14**: El sistema debe registrar movimientos de tipo `alta`, `traslado` y `baja`.
- **RF-15**: El sistema debe permitir correcciones mediante el permiso operativo normal del usuario autenticado, exigir una observación obligatoria y conservar la auditoría de la acción.
- **RF-16**: El sistema debe validar códigos de rollo y ubicación con normalización, estructura, referencias, unicidad y mensajes por campo.

## 11. Requisitos no funcionales

- **RNF-01 Seguridad**: autenticación segura, autorización en backend, contraseñas protegidas y sesiones con expiración según política definida.
- **RNF-02 Integridad**: actualización del estado actual y registro del movimiento en una transacción consistente.
- **RNF-03 Auditoría**: conservar actor, fecha/hora, acción, resultado y referencia de los datos afectados.
- **RNF-04 Concurrencia**: evitar movimientos perdidos o estados divergentes cuando operen varios usuarios.
- **RNF-05 Disponibilidad**: el sistema debe estar preparado para operación online y manejo explícito de errores de red; no se promete funcionamiento offline en esta versión.
- **RNF-06 Rendimiento**: las consultas habituales de rollo y ubicación deben responder dentro del objetivo que se acuerde para la operación real.
- **RNF-07 Escalabilidad**: la solución debe soportar el volumen actual y el crecimiento esperado de rollos, movimientos y usuarios.
- **RNF-08 Usabilidad**: interfaz responsive, clara para operación con teclado o escáner y con mensajes accionables.
- **RNF-09 Observabilidad**: registrar errores técnicos y eventos relevantes sin exponer datos sensibles.
- **RNF-10 Recuperación**: definir copias de seguridad, restauración y retención de PostgreSQL antes de producción.
- **RNF-11 Mantenibilidad**: separar frontend React, API Express y persistencia PostgreSQL mediante contratos documentados.
- **RNF-12 Privacidad**: limitar la exposición de datos de usuarios y aplicar el principio de mínimo privilegio.

## 12. Criterios de aceptación

- [ ] Un usuario autenticado solo puede ejecutar acciones permitidas por su rol.
- [ ] Un usuario autorizado puede buscar un rollo y ver su ubicación y estado actuales.
- [ ] Una consulta muestra como cantidad existente en DPT todos los rollos registrados, sin filtrar por estado, y permite ver estado y ubicación como datos adicionales.
- [ ] Cada movimiento confirmado registra rollo, usuario, fecha/hora, origen, destino, estado anterior y posterior cuando corresponda.
- [ ] Se puede reconstruir cronológicamente toda la trayectoria de un rollo, incluyendo los cambios de estado como `expedicion`, sin tratarlo como ubicación.
- [ ] El sistema permite registrar un cambio de ubicación sin cambio de estado y conserva ambos valores correctamente.
- [ ] Se puede reconstruir una secuencia de múltiples cambios de ubicación del mismo rollo, incluido un reacomodamiento por disponibilidad de espacio.
- [ ] El sistema rechaza movimientos con rollo, destino físico, estado o transición inválidos, según las reglas aprobadas, y explica el motivo por campo.
- [ ] El sistema acepta únicamente ubicaciones físicas con depósito, pasillo, columna y nivel válidos; no ofrece ubicaciones genéricas ni permite desactivar ubicaciones.
- [ ] El sistema valida códigos de rollo y ubicación con normalización, unicidad y referencias existentes, sin imponer longitudes o separadores aún no confirmados.
- [ ] Los movimientos se clasifican como `alta`, `traslado` o `baja`; una corrección permitida por el permiso operativo normal exige observación obligatoria, deja auditoría y queda asociada al usuario autenticado.
- [ ] Dos operaciones incompatibles sobre el mismo rollo no pueden confirmarse como si fueran válidas.
- [ ] Un Excel inválido no modifica la base de datos y muestra errores por fila y columna.
- [ ] Una importación válida requiere confirmación y deja auditoría del usuario, fecha, archivo y resultado.
- [ ] Un usuario sin permiso de exportación no puede generar un Excel operativo.
- [ ] Una exportación autorizada solo incluye los datos dentro del alcance y filtros aplicados, sin exponer registros fuera de ese alcance.
- [ ] El Excel exportado incluye como mínimo código de rollo, estado, ubicación, fecha/hora del último movimiento y los atributos operativos aprobados; si incluye trazabilidad, contiene los datos mínimos de cada movimiento.
- [ ] El archivo exportado informa fecha/hora de generación, usuario generador, filtros, alcance y versión de formato o plantilla.
- [ ] La edición de un Excel exportado no modifica PostgreSQL, no genera movimientos y muestra una advertencia clara sobre su carácter de copia de trabajo.
- [ ] Una exportación queda registrada con usuario, fecha/hora, filtros, alcance, resultado y referencia del archivo generado.
- [ ] La migración inicial no depende de JSON ni requiere archivos locales para consultar el inventario.
- [ ] Los códigos y catálogos no confirmados pueden configurarse o aparecen explícitamente como pendientes, sin validaciones inventadas.
- [ ] Las pruebas cubren permisos, importación, movimientos, estados, ubicaciones físicas, validaciones de códigos, concurrencia y trazabilidad.

## 13. Riesgos y mitigaciones

| Riesgo | Impacto | Mitigación |
|---|---|---|
| Excel de origen inconsistente | Importación incompleta o inventario incorrecto | Plantilla, previsualización, validación por fila, informe de errores y respaldo previo |
| Códigos o estados no definidos | Reglas rígidas e incompatibles con la operación | Catálogos configurables y decisiones pendientes explícitas |
| Conteo de inventario interpretado como filtro por estado | Cantidad de DPT incorrecta para la operación | Definir el conteo como todos los rollos registrados y mostrar el estado solo como dato adicional |
| Matriz de estados incompleta | Movimientos rechazados o aceptados incorrectamente | Mantener la matriz de transiciones pendiente y no inventar reglas de transición |
| Modelo incorrecto de expedición | Pérdida de ubicación física o historial | Mantener `expedicion` como estado separado y exigir ubicación física en cada rollo |
| Formato de códigos incompleto | Rechazos incorrectos o duplicados | Normalizar mayúsculas/minúsculas, validar estructura y referencias, mostrar errores por campo y mantener pendientes las longitudes y separadores |
| Movimientos concurrentes | Estado actual divergente | Transacciones, control de concurrencia y pruebas multiusuario |
| Permisos mal definidos | Cambios no autorizados o complejidad innecesaria | Mantener una matriz simple, aplicar el permiso operativo normal en backend y asociar cada acción al usuario autenticado |
| Exportación fuera de alcance o con datos sensibles | Exposición de información o uso operativo incorrecto | Permisos en backend, filtros aplicados en servidor, alcance explícito, metadatos y auditoría de cada generación |
| Edición de un Excel exportado interpretada como actualización | Divergencia entre el archivo y el inventario central | Advertencia visible, exportación de solo lectura respecto de PostgreSQL y flujo de reimportación separado si se aprobara |
| Migración sin historial | No se puede reconstruir el pasado anterior | Decidir si el Excel contiene solo estado inicial o también movimientos históricos |
| Dependencia de conectividad | Operación detenida ante fallas online | Definir disponibilidad objetivo, monitoreo y procedimiento operativo ante caída |
| Crecimiento de historial | Consultas lentas o costos crecientes | Índices, paginación, retención y métricas definidas antes de escalar |

## 14. Preguntas resueltas

- **Cantidad existente en DPT**: incluye todos los rollos registrados, sin filtrar por estado; el estado se muestra como dato adicional.
- **Cambios de ubicación**: un rollo puede cambiar de ubicación múltiples veces, incluido el reacomodamiento para aprovechar una ubicación con lugar libre; el cambio puede ocurrir sin cambio de estado.
- **Estados válidos del rollo**: `disponible`, `reservado`, `picking`, `expedicion`, `despachado`, `devuelto` y `bloqueado`.

## 15. Preguntas pendientes

- ¿Qué longitudes, rangos y separadores definitivos deben aplicarse a los componentes de ubicación y al código de rollo? El ejemplo de ubicación es `DEP1 C2 46 03` y el de rollo `FJSB754-4U`; la estructura base queda documentada, pero estos detalles requieren confirmación.
- ¿Cuál es la matriz de transiciones válidas entre estados?
- ¿El Excel inicial representa solo el estado actual o incluye movimientos históricos?
- **Pregunta 12 (stand-by)**: ¿Qué columnas contiene el Excel real y cuál es el mapeo aprobado a las entidades del sistema? Mantener abierta hasta recibir el archivo real; no cerrar ni eliminar.
- ¿Qué permisos necesitan los operadores por ubicación, turno o tipo de operación?
- ¿Qué roles pueden exportar, qué datos mínimos deben incluirse y qué límites de volumen o frecuencia deben aplicarse?
- ¿Debe existir en el futuro un flujo explícito de reimportación de un Excel exportado editado, con qué validaciones y autorizaciones?
- ¿Debe existir integración con escáneres, ERP, producción, expedición u otro sistema externo?
- ¿Cuál es el volumen esperado de rollos, movimientos, usuarios y consultas simultáneas?
- ¿Cuáles son los objetivos de disponibilidad, rendimiento, respaldo y recuperación?
- ¿Qué política de retención y acceso debe aplicarse a la auditoría?

## 16. Decisiones pendientes para pasar a diseño

1. Definir el modelo de estados y las transiciones válidas.
2. Confirmar longitudes, rangos, separadores y representación canónica de ubicaciones físicas y códigos de rollo.
3. Obtener y validar el Excel fuente para cerrar la plantilla de importación y resolver la pregunta 12, actualmente en stand-by.
4. Definir si se migra historial o únicamente el estado inicial.
5. Confirmar la matriz básica de roles y permisos operativos; las correcciones usan el permiso normal de movimientos, requieren observación obligatoria y no tienen aprobación o permiso especial.
6. Aprobar permisos, alcance, filtros, datos mínimos y formato de las exportaciones operativas.
7. Establecer objetivos de seguridad, disponibilidad, rendimiento y recuperación.
