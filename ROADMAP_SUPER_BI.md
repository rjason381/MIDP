# MIDP Super BI Blueprint

## Estado

Este documento guarda la ruta objetivo del nuevo modulo BI para MIDP.
La idea central es:

- construir primero una experiencia local de alto nivel
- separar por completo UI, modelo semantico y motor de consultas
- dejar lista la migracion futura a backend en Cloud Run + Cloud SQL MySQL

No se debe volver a acoplar el frontend a una fuente puntual de datos.

## Progreso en repo

Estado actual del trabajo ya aplicado en este repositorio:

- [x] Auditoria tecnica QuickSight persistida en `AUDIT_DECHINI_QUICKSIGHT_2026-03-11.md`
- [x] Blueprint maestro persistido en repo
- [x] Inicio del Hito A sobre la base actual
- [x] Autosave real con debounce y persistencia efectiva
- [x] Fix de referencia de visual para drag de labels en QuickSight
- [x] Busqueda global coherente en render QuickSight
- [x] Introduccion de `analytics-core.js` como base para query contracts y local adapter
- [x] Introduccion de `midp-storage.js` para desacoplar persistencia local de `localStorage`
- [x] Base hibrida `localStorage + IndexedDB` con hidratacion inicial del estado
- [x] Introduccion de `analytics-local-data.js` para armado y filtrado local del dataset BI
- [x] Introduccion de `analytics-project-query.js` para catalogo BI y query service por proyecto
- [x] Centralizacion de consultas BI repetidas hacia `queryBiProjectRows(...)`
- [x] Primer refresh visual de BI Studio y QuickSight sobre la shell actual
- [x] Interacciones QuickSight reales con filtro cruzado por click sobre datos
- [x] Drill-down QuickSight por visual con jerarquia persistente, breadcrumb y retroceso
- [x] Drill-through QuickSight con drawer de detalle real, export CSV y acceso a seguimiento
- [x] Primer adapter local a ECharts en QuickSight con fallback seguro al renderer canvas
- [x] Adapter ECharts ampliado en QuickSight para `treemap`, `funnel`, `scatter` y `bubble`
- [x] Adapter ECharts ampliado en QuickSight para `waterfall` y `pareto`, manteniendo `radar` en fallback canvas
- [x] `Dechini Quicksigth` ya permite elegir motor por visual (`Auto`, `ECharts`, `Canvas`) y suma `combo` en ECharts
- [x] Fix de integridad en hidratacion hibrida: preferencia por `localStorage` cuando difiere de `IndexedDB`
- [x] Fix de `Automatico` en limites de eje para renderer `ECharts`, con UI mas clara en BI/QuickSight
- [x] Fix de coherencia en QuickSight: `Agregar visual` ahora reutiliza la configuracion visible del visual activo
- [x] QuickSight ahora tiene `Undo/Redo` local por hoja y atajos `Ctrl/Cmd+Z`, `Ctrl+Y`, `Cmd+Shift+Z`
- [x] Botones fantasmas reducidos en QuickSight: `Configuracion` abre Propiedades y `Publicar` exporta la hoja a PNG
- [x] Acciones de multi-hoja quedan explicitadas como pendientes, evitando clicks muertos
- [x] QuickSight ya soporta `EJE X` opcional para `scatter/bubble`, con drag/drop y persistencia por visual
- [x] La leyenda y propiedades QuickSight ahora se muestran segun el motor real resuelto (`ECharts` vs `Canvas`)
- [x] Los controles `Agregar medida` y `Agregar una dimension` en QuickSight ya son contextuales y no quedan como placeholder engañoso
- [x] QuickSight ahora oculta chrome inutil, suma modo `Visualizacion` y permite maximizar/minimizar cada chart
- [ ] Extraer mas uso del BI/QuickSight hacia adapters y contratos separados
- [ ] Migrar persistencia principal desde `localStorage` a `IndexedDB`
- [ ] Reconstruir BI Shell premium sobre stack nuevo

## Vision de producto

MIDP debe evolucionar hacia un BI Studio propio con estas capacidades:

- editor visual de dashboards con pizarra libre
- datasets reutilizables
- dimensiones, metricas y campos calculados
- filtros globales y por visual
- drill-down y drill-through
- tablas, pivot y scorecards de nivel enterprise
- versionado, vistas guardadas y plantillas
- modo presentacion y exportaciones
- integracion con modelos ACC/RVT y dashboards sobre propiedades BIM
- visor de modelo embebido con seleccion sincronizada contra visuales

La experiencia debe sentirse como un producto BI serio, no como una suma de widgets.

## Principios de arquitectura

1. Local-first ahora, backend-first despues
2. El frontend no depende de MySQL ni de Cloud SQL
3. Todo acceso a datos pasa por un Query Adapter
4. Los dashboards deben funcionar igual con datos locales o remotos
5. El motor visual no debe contener logica de negocio
6. El modelo semantico debe ser estable y versionable

## Stack objetivo

### Fase local

- Frontend: React + TypeScript + Vite
- Estado UI: Zustand
- Validacion: Zod
- Persistencia local: IndexedDB con Dexie
- Motor de consultas local: DuckDB-Wasm
- Graficos: Apache ECharts
- Tablas y pivot:
  - preferido: AG Grid Enterprise
  - alternativo: TanStack Table

### Fase backend

- API: Node.js + TypeScript
- Runtime: Cloud Run
- Base de datos: Cloud SQL MySQL
- Secretos: Secret Manager
- Observabilidad: Cloud Logging + metricas de aplicacion

## Integracion ACC / RVT

Objetivo:

- vincular modelos RVT publicados en ACC
- extraer propiedades del modelo via APS Model Properties
- registrar cada modelo como fuente independiente dentro de QuickSight
- mostrar el nombre real del modelo en `Conjunto de datos`
- permitir dashboards por modelo sin mezclar todo en una sola vista global
- incorporar un visor embebido para navegar el modelo y cruzarlo con filtros BI

Lineamientos:

- no exponer una vista compuesta tipo `TODAS LAS VISTAS` para modelos BIM
- cada modelo debe entrar como dataset/fuente propia
- los filtros globales pueden cruzar multiples fuentes, pero la seleccion de dataset debe seguir siendo limpia y explicita
- el visor debe poder seleccionar elementos y reflejar esa seleccion en tablas y graficos
- los graficos deben poder disparar aislamiento, enfoque o seleccion en el visor

## Capas del sistema

### 1. BI Shell

Responsable de:

- navegacion entre dashboards
- toolbar
- command palette
- paneles laterales
- modo presentacion
- estado de seleccion y layout

### 2. Semantic Layer

Responsable de:

- datasets
- fields
- dimensions
- measures
- calculated fields
- formatos
- relaciones entre entidades

### 3. Query Adapter

Contrato unico para cualquier fuente de datos.

Implementaciones planeadas:

- LocalQueryAdapter
- ApiQueryAdapter

### 4. Visual Registry

Responsable de:

- registrar tipos de visual
- validar configuracion por visual
- mapear QueryResult -> ECharts/Grid config

### 5. Dashboard Canvas

Responsable de:

- layout libre
- snap
- guides
- rulers
- seleccion multiple
- align/distribute
- group/ungroup
- lock/unlock
- z-index
- undo/redo

### 6. Persistence Layer

Responsable de:

- dashboards
- widgets
- snapshots
- drafts
- saved views
- version history

## Contratos base

Estos contratos deben existir antes de migrar el modulo actual:

```ts
type DatasetId = string;
type DashboardId = string;
type WidgetId = string;

type FieldKind = "dimension" | "measure" | "date" | "calculated";

interface FieldDef {
  id: string;
  datasetId: DatasetId;
  name: string;
  label: string;
  kind: FieldKind;
  dataType: "string" | "number" | "date" | "boolean";
  expression?: string;
  format?: string;
}

interface QueryRequest {
  datasetId: DatasetId;
  fields: string[];
  filters: unknown[];
  orderBy: unknown[];
  limit?: number;
  groupBy?: string[];
}

interface QueryResult {
  columns: Array<{ id: string; label: string; dataType: string }>;
  rows: Array<Record<string, unknown>>;
  stats?: {
    rowCount: number;
    executionMs: number;
  };
}

interface DashboardWidget {
  id: WidgetId;
  type: string;
  title: string;
  layout: { x: number; y: number; w: number; h: number };
  query: QueryRequest;
  style: Record<string, unknown>;
  config: Record<string, unknown>;
}
```

## Funcionalidades objetivo

### Canvas y UX

- pizarra libre con snap y guides
- zoom fit all / fit width / custom
- duplicar, bloquear, ocultar, agrupar
- auto layout por secciones
- panel de propiedades contextual
- temas y presets visuales
- command palette operativa

### Datos y analitica

- browser de datasets y campos
- medidas agregadas
- campos calculados
- top N
- sort avanzado
- agrupacion por tiempo
- filtros globales
- filtros por widget
- cross-filter
- drill-down
- drill-through

### Visuales

- bar, line, area, combo
- pie, donut, treemap, funnel
- scatter, bubble, radar, pareto
- scorecard
- table
- pivot
- timeline
- sankey

### Enterprise

- versionado de dashboards
- saved views
- templates
- export PNG / SVG / PDF / CSV / XLSX
- modo presentacion
- permisos por dashboard
- auditoria de cambios

## Fases del roadmap

### Fase 0. Estabilizar el modulo actual

Objetivo: dejar la base actual util mientras se construye la nueva.

Tareas minimas:

- arreglar guardado real y autosave
- arreglar drag de labels en QuickSight
- separar config global vs config del visual activo
- reducir dependencias entre render, estado y dataset
- documentar bugs y deuda tecnica actual

### Fase 1. Extraer contratos y modelo semantico

Objetivo: romper el monolito actual sin cambiar aun la UI final.

Tareas:

- definir contratos TS
- crear mapeadores desde estado actual MIDP
- aislar Query Adapter local
- normalizar datasets, fields y metrics

### Fase 2. Nuevo BI Shell

Objetivo: redisenar la experiencia visual completa.

Tareas:

- layout 3 paneles
- toolbar superior
- command palette
- panel de datasets
- inspector de propiedades
- dashboard list + tabs

### Fase 3. Nuevo motor visual

Objetivo: reemplazar el renderer manual por ECharts y grid dedicado.

Tareas:

- renderer por tipo visual
- registry de widgets
- adaptadores ECharts
- exportaciones
- temas
- interacciones comunes

### Fase 4. Motor local de consultas

Objetivo: soportar BI de verdad en modo local.

Tareas:

- integrar DuckDB-Wasm
- cargar datasets desde snapshots MIDP o archivos
- ejecutar agregaciones SQL locales
- soportar filtros, joins y campos calculados

### Fase 5. Persistencia local seria

Objetivo: abandonar localStorage como nucleo principal.

Tareas:

- IndexedDB
- drafts
- autosave real
- historial local
- snapshots

### Fase 6. API backend

Objetivo: dejar el frontend listo para operar remoto.

Tareas:

- ApiQueryAdapter
- endpoints de schema
- endpoints de query
- endpoints de dashboards
- auth y permisos

### Fase 7. Cloud Run + Cloud SQL

Objetivo: mover la fuente de verdad a Google Cloud.

Tareas:

- backend Node/TS en Cloud Run
- conexion a Cloud SQL MySQL
- secrets en Secret Manager
- pooling de conexiones
- observabilidad
- migracion gradual por adapter

## Orden recomendado de entrega

1. Estabilizar QuickSight/BI actual
2. Definir contratos de datos
3. Construir BI Shell nuevo
4. Integrar ECharts
5. Integrar Grid/Pivot
6. Integrar DuckDB local
7. Migrar persistencia a IndexedDB
8. Exponer API backend
9. Conectar Cloud SQL MySQL

## Decision sobre grids

### Opcion A: AG Grid Enterprise

Elegir si se prioriza:

- velocidad de entrega
- pivot enterprise
- agrupacion fuerte
- tool panels maduros

### Opcion B: TanStack Table

Elegir si se prioriza:

- menor costo
- mayor control
- UI totalmente propia

Costo: exige mucho mas trabajo para llegar a experiencia BI premium.

## Hito inmediato dentro de este repo

Antes de migrar a un stack nuevo, el siguiente hito recomendado es:

### Hito A. Hardening del modulo actual

Entregables:

- fix de autosave real
- fix de label drag QuickSight
- fix de busqueda global coherente en QuickSight
- separar estado de plantilla vs estado de visual activo
- inventario de funciones reutilizables para el futuro Query Adapter

Este hito permite seguir operando el modulo actual mientras se prepara la reconstruccion.

## Regla de migracion

No hacer una migracion Big Bang.

La regla es:

- mantener el modulo actual funcionando
- extraer capacidades reutilizables
- crear el nuevo BI Studio en paralelo
- mover primero el renderer
- mover despues el query layer
- mover al final la fuente remota

## Meta final

MIDP debe terminar con un BI Studio propio que:

- se vea al nivel de una herramienta premium
- funcione localmente sin depender del backend
- cambie a Cloud SQL sin rehacer el frontend
- soporte dashboards complejos, tablas, pivots, filtros e interacciones reales
