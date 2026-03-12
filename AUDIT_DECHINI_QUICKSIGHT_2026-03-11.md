# Auditoria Dechini Quicksigth

Fecha: 2026-03-11

## Alcance

Revision tecnica de `Dechini Quicksigth` sobre:

- integridad de datos y persistencia
- funcionalidad real vs funcionalidad aparente
- visualizacion y consistencia UX
- brechas para llegar a un BI Studio de nivel alto

Esta auditoria se basa en lectura de codigo del repo. No incluye una prueba manual completa en navegador.

## Hallazgos

### Critico

1. Fuente de verdad de guardado inconsistente entre `localStorage` e `IndexedDB`.
   - `saveState(true)` persiste via `storageAdapter.setJson(...)` en [app.js:19648](app.js#L19648).
   - El adapter hibrido escribe primero a `localStorage` y luego dispara la escritura a `IndexedDB` sin esperar resultado en [midp-storage.js:257](midp-storage.js#L257) y [midp-storage.js:265](midp-storage.js#L265).
   - En el arranque, la hidratacion prioriza `IndexedDB` sobre `localStorage` en [midp-storage.js:273](midp-storage.js#L273) y [midp-storage.js:281](midp-storage.js#L281).
   - Riesgo: si `IndexedDB` queda atrasado o la escritura async no termina, el siguiente arranque puede restaurar una version vieja del estado y pisar la mas nueva.

2. El panel de configuracion visible no coincide siempre con el comportamiento de `Agregar visual`.
   - Cuando hay un visual seleccionado, los controles QuickSight editan ese visual y ya no actualizan `project.quickSightConfig` en [app.js:1004](app.js#L1004) y [app.js:1027](app.js#L1027).
   - Sin embargo, `Agregar visual` sigue creando desde `project.quickSightConfig` en [app.js:1570](app.js#L1570) y [app.js:20610](app.js#L20610).
   - Riesgo: el usuario cree que esta "configurando el siguiente visual", pero en realidad esta editando el actual; luego el nuevo visual nace con otra configuracion.

### Alto

3. La UI expone acciones de estudio que hoy no estan implementadas.
   - La vista muestra `Configuracion`, `Publicar`, `Deshacer`, `Rehacer` y boton de nueva hoja en [index.html:979](index.html#L979), [index.html:1034](index.html#L1034), [index.html:1047](index.html#L1047), [index.html:1055](index.html#L1055) y [index.html:1208](index.html#L1208).
   - El registro de elementos QuickSight solo recoge controles operativos reales como `qsAddVisualButton` y `qsRefreshButton` en [app.js:576](app.js#L576), sin wiring equivalente para esas acciones.
   - Impacto: el producto aparenta una madurez mayor a la real y erosiona confianza.

4. El modelo de "hojas" es solo cosmetico.
   - La UI presenta `Hojas`, `Hoja 1` y un boton `+` en [index.html:1027](index.html#L1027), [index.html:1207](index.html#L1207) y [index.html:1208](index.html#L1208).
   - El titulo se fija siempre como `Hoja 1` o `Hoja 1 | proyecto` en [app.js:8121](app.js#L8121) y [app.js:8241](app.js#L8241).
   - El estado persistido sigue siendo un solo arreglo `quickSightVisuals` en [app.js:22016](app.js#L22016) y [app.js:22017](app.js#L22017), sin coleccion de worksheets.
   - Impacto: la shell parece multipagina, pero el dominio sigue siendo de hoja unica.

5. `scatter` y `bubble` estan a medio modelar en QuickSight.
   - El motor de snapshot si contempla `optionalMetrics` y produce `xValue` en [app.js:15403](app.js#L15403) y [app.js:15461](app.js#L15461).
   - Pero la sincronizacion QuickSight solo conserva `dimensions`, `metrics` y `breakdownDimension` en [app.js:20484](app.js#L20484) y [app.js:20502](app.js#L20502).
   - La UI sigue anunciando una sola metrica y el boton adicional es placeholder en [index.html:1146](index.html#L1146) y [app.js:1118](app.js#L1118).
   - Impacto: esos tipos existen visualmente, pero no tienen todo el modelado de datos que necesitan para competir con un BI serio.

### Medio

6. Inconsistencia entre capacidades del renderer y capacidades expuestas por la UI.
   - QuickSight solo habilita configuracion de leyenda para `pie` y `donut` via `supportsLegend` en [app.js:10030](app.js#L10030) y [app.js:10046](app.js#L10046).
   - El panel de propiedades depende de esa bandera en [app.js:8682](app.js#L8682).
   - El adapter de `ECharts` si sabe resolver leyendas para mas tipos en [bi-echarts-adapter.js:151](bi-echarts-adapter.js#L151).
   - Impacto: el motor ya puede mas de lo que la UI permite configurar.

7. El selector principal de tipos no refleja todo lo ya soportado.
   - La grilla visual QuickSight no incluye `pareto` en [index.html:1125](index.html#L1125).
   - El selector oculto y el modelo interno si lo incluyen en [index.html:1168](index.html#L1168) y [app.js:7808](app.js#L7808).
   - Impacto: capacidad real desaprovechada y experiencia inconsistente.

8. El render de QuickSight sigue siendo de pizarra completa.
   - `renderQuickSightPanel(...)` recompone `innerHTML` de la superficie y vuelve a montar cards en [app.js:8295](app.js#L8295), [app.js:8304](app.js#L8304) y [app.js:8377](app.js#L8377).
   - Luego vuelve a reconstruir snapshots y renderizar cada visual en [app.js:8382](app.js#L8382) y [app.js:8393](app.js#L8393).
   - Impacto: con muchos visuales, filtros o interacciones, el costo de repintado va a crecer rapido.

9. La historia de exportacion de QuickSight sigue incompleta.
   - Hoy existe export CSV del drawer de drill-through en [index.html:1404](index.html#L1404) y [app.js:4406](app.js#L4406).
   - No existe export de hoja/dashboard QuickSight equivalente a BI (`PNG`, `PDF` o export por visual) pese a que BI si tiene esa capacidad en [index.html:245](index.html#L245) y [app.js:17187](app.js#L17187).
   - Impacto: QuickSight ya parece estudio principal, pero sigue sin salida de artefactos de nivel ejecutivo.

10. Algunas acciones visibles son mas limitadas de lo que su etiqueta sugiere.
   - `Agregar una medida` solo muestra un mensaje de limitacion en [app.js:1118](app.js#L1118).
   - `Agregar una dimension` no agrega otra dimension; solo limpia `GRUPO/COLOR` en [app.js:1122](app.js#L1122).
   - Impacto: la interfaz promete un builder de roles mas flexible del que realmente existe.

### Bajo

11. La capa visual actual funciona, pero la arquitectura CSS esta muy fragmentada.
   - El refresh visual premium vive en [styles.css:4425](styles.css#L4425).
   - Luego hay overrides globales de esquinas rectas en [styles.css:4611](styles.css#L4611) y aplanado visual en [styles.css:4634](styles.css#L4634).
   - Impacto: el resultado visual puede verse bien, pero la cascada se vuelve fragil y costosa de mantener.

12. La identidad de naming sigue con deuda.
   - La UI y varios mensajes siguen usando `Quicksigth` en vez de `QuickSight`, por ejemplo en [app.js:1643](app.js#L1643) y [app.js:1119](app.js#L1119).
   - No rompe el sistema, pero si complica busquedas, documentacion y profesionalismo del modulo.

## Lo que ya esta bien

- Hay base funcional real: canvas libre, drag/resize, propiedades por visual, cross-filter, drill-down y drill-through.
- La migracion a `ECharts` ya arranco con fallback seguro por visual, lo cual baja riesgo de reescritura total.
- El desacople reciente de persistencia y consultas BI va en la direccion correcta.

## Lo que falta para un BI Studio fuerte

- historial real de `undo/redo`
- hojas multiples reales
- exportacion de hoja/dashboard QuickSight
- semantic roles mas completos por tipo visual
- alineacion entre acciones visibles y acciones reales
- render parcial por visual
- versionado y publicacion reales

## Plan propuesto

### Fase 1. Integridad y confianza

1. Corregir el adapter hibrido para que exista una sola fuente de verdad.
   - O `IndexedDB` pasa a ser primario y se espera la escritura.
   - O la hidratacion compara timestamps/version y no prioriza ciegamente `IndexedDB`.
2. Separar con claridad `defaults del proyecto` vs `configuracion del visual activo`.
   - `Agregar visual` debe crear desde un preset explicito o desde "duplicar configuracion actual".
3. Ocultar o implementar las acciones fantasma.
   - `Deshacer`, `Rehacer`, `Publicar`, `Configuracion`, `Nueva hoja`.

### Fase 2. Completitud funcional de QuickSight

1. Introducir `dataRoles` completos por visual en QuickSight.
   - `optionalMetrics`
   - `dateDimension`
   - roles adicionales por tipo si aplican
2. Alinear la UI de tipos con el motor.
   - agregar `pareto` a la grilla
   - aclarar `radar` si sigue en fallback canvas
3. Corregir etiquetas de acciones builder para que describan lo que de verdad hacen.

### Fase 3. Super BI UX

1. Implementar hojas reales:
   - `quickSightSheets[]`
   - `activeSheetId`
   - visuales por hoja
2. Agregar export de hoja QuickSight:
   - PNG
   - PDF
   - CSV por visual
3. Construir `undo/redo` con snapshots compactos o command log.

### Fase 4. Rendimiento y mantenibilidad

1. Migrar de rerender total a render parcial por visual.
2. Consolidar tema QuickSight en un solo bloque CSS con tokens.
3. Reutilizar mejor la capa `ECharts` y definir una matriz clara:
   - soportado en ECharts
   - soportado solo en canvas
   - soportado en ambos

## Prioridad recomendada

1. Integridad de guardado
2. Coherencia entre panel visible y `Agregar visual`
3. Acciones fantasma / hoja falsa
4. Roles de datos incompletos para `scatter` y `bubble`
5. Exportacion y rendimiento

## Riesgos si no se corrige

- perdida silenciosa o rollback de estado al reabrir la app
- usuarios creando visuales con configuracion distinta a la que creen ver
- percepcion de producto "demo" por botones sin efecto
- deuda tecnica creciente en CSS y render
- limite funcional temprano cuando entren datasets reales desde backend/MySQL

## Seguimiento incremental

Cambios ya aplicados despues de esta auditoria:

- se redujo el riesgo de rollback en hidratacion hibrida favoreciendo `localStorage` cuando difiere de `IndexedDB`
- `Agregar visual` en QuickSight ya reutiliza la configuracion visible del visual activo
- `Automatico` en limites de eje ya no cae a `0` en el renderer `ECharts`
- `Configuracion` ahora abre Propiedades en QuickSight
- `Publicar` ahora exporta la hoja QuickSight a PNG local
- `Undo/Redo` ya funciona a nivel de hoja QuickSight con botones y atajos
- la grilla principal de tipos QuickSight ya incluye `Pareto`
- `scatter` y `bubble` ya pueden definir `EJE X` con metrica opcional desde panel, drag/drop y propiedades
- la seccion `Leyenda` ya no aparece por tipo fijo; ahora sigue el motor real resuelto del visual
- `Agregar una medida` y `Agregar una dimension` ya no quedan como click vacio; ahora responden segun el tipo y el contexto del visual
- QuickSight ya entra en modo mas compacto, agrega `Visualizacion` para ver solo la pizarra y cada visual puede maximizarse/minimizarse sin alterar su layout guardado
- el layout de QuickSight ya contiene bien la altura del shell y el panel de propiedades vuelve a tener scroll interno usable
- se corrigio el traslape entre vistas: los tabs ocultos vuelven a ganar siempre frente a estilos especificos como `#tab-quicksight`, evitando que QuickSight se monte sobre `Listas`, `MIDP`, `Control de Flujos de Revision` o `Dechini BI`
- se corrigio un bug de integridad CSS en QuickSight: `.hidden` quedaba anulado por reglas como `.qs-role-block`, haciendo visible el bloque `EJE X` opcional incluso en graficos que no lo soportan
- el renderer `ECharts` de QuickSight ya limpia `optionalMetric` cuando el tipo no lo soporta y el grafico de barras ahora separa mejor categorias vs `GRUPO/COLOR`, evitando mezclar ambos en una sola etiqueta del eje X
- el bloque legacy `Metrica eje X` quedo oculto de forma robusta con `hidden` real en DOM, para que ya no vuelva a filtrarse visualmente en tipos como `bar`
- la leyenda de `ECharts` en QuickSight ya respeta `legendPosition` tambien en graficos cartesianos y el `grid` ahora reserva espacio coherente para la leyenda, evitando compresion rara del chart en la franja superior
- se corrigio una falla de integridad en `GRUPO/COLOR` y `Metrica eje X`: los valores vacios ya no son reemplazados por fallbacks viejos del proyecto o `dataRoles`, asi que ahora si puede limpiarse la agrupacion secundaria de un visual

Riesgos residuales tras este bloque:

- `Undo/Redo` es local en memoria; no sobrevive a una recarga completa
- la exportacion PNG de QuickSight depende del render actual de la hoja en DOM
- multi-hoja sigue pendiente a nivel de modelo, aunque ahora ya no queda como click muerto silencioso
- `dateDimension` todavia no esta expuesto en la UX de QuickSight para tipos temporales futuros
- el modo `Visualizacion` y el maximizado por chart quedaron validados por sintaxis, pero aun no por prueba manual real en navegador
- si el usuario configura demasiadas series via `GRUPO/COLOR` en barras, la lectura visual seguira degradandose; eso ya es mas una decision de diseno/UX que un fallo duro del renderer
