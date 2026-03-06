# Dechini BI - Roadmap de Configuracion (persistente)

## Estado actual (aplicado)
- [x] Alcance de configuracion visual: global del proyecto vs widget seleccionado.
- [x] Seleccion de widget en pizarra (resaltado visual).
- [x] Herencia visual: widget usa global cuando no tiene override.
- [x] Presets visuales integrados: Corporativo, Minimal, Contraste, Presentacion.
- [x] Guardado de presets personalizados por proyecto.
- [x] Configuracion avanzada:
  - [x] Leyenda: mostrar, posicion, maximo de items.
  - [x] Rejilla: mostrar, opacidad, guiones.
  - [x] Labels de datos: mostrar/ocultar.
  - [x] Ejes: titulos X/Y, escala lineal/logaritmica, min/max manual.
  - [x] Formato numerico: auto, numero, porcentaje, moneda S/, horas.
  - [x] Prefijo y sufijo.
  - [x] Tipografia base: tamano.
  - [x] Series: grosor de linea, suavizado, opacidad de area, ancho de barra.
  - [x] Tooltip: modo completo/compacto.
- [x] Exportaciones:
  - [x] Exportar widget a PNG.
  - [x] Exportar pizarra completa a PNG.
  - [x] Exportar CSV de detalle y widget.

## Pendiente (siguiente fase)
- [x] Tipografia avanzada por capas (fuente distinta para titulo/ejes/labels/tooltip).
- [x] Estilos por serie (linea punteada/continua, marcadores por tipo).
- [x] Stacking real (normal y 100%) para tipos que aplique.
- [x] Filtro cruzado avanzado:
  - [x] Ctrl+click multi-seleccion.
  - [x] Alcance configurable (misma fuente/todos).
- [x] Ayudas UX en cada control (tooltips de configuracion).
- [x] Exportar pizarra con mayor fidelidad de tablas (render completo de celdas y KPI).
- [ ] Optimizacion de rendimiento: rerender de solo widget afectado.

## Mejoras aplicadas extra (fase actual)
- [x] Auto organizar widgets y generar dashboard base automaticamente.
- [x] Nuevos tipos de grafico: Radar y Pareto.
- [x] Clonar/Bloquear widget (botones y atajos Ctrl+D / L).
- [x] Orden configurable por widget (valor asc/desc, etiqueta A-Z/Z-A).
- [x] Linea objetivo configurable (valor, etiqueta y color) en graficos cartesianos.
- [x] Edicion rapida del widget seleccionado desde Propiedades ("Actualizar seleccionado").
- [x] Exportacion PNG en alta definicion (escala 2x) para widget y pizarra.
- [x] Command Palette BI (Ctrl+K) con comandos operativos y apertura de paneles.
- [x] Insights accionables (tarjetas de recomendacion con filtros directos).
- [x] Pizarra con rejilla configurable + snap real en mover/redimensionar/teclado.
- [x] Paneles flotantes BI estables (sin ocupar espacio del lienzo).

## Plan maestro UX (siguiente iteracion)
- [ ] Optimizacion de rendimiento: rerender solo widget afectado (sin redibujar toda la pizarra).
- [ ] Modo Presentacion: ocultar controles, fullscreen y navegacion por paginas.
- [ ] Historico de versiones de dashboard por proyecto (snapshot + rollback).
- [ ] Widgets calculados con formulas (builder tipo `fx`: IF, SUMIF, progreso ponderado).
- [ ] Drill-down multinivel en widgets (ej. Disciplina -> Sistema -> Paquete -> Entregable).
- [ ] Biblioteca de plantillas (PMO, Produccion, Calidad, Control semanal).
- [ ] Colaboracion avanzada: bloqueo optimista por widget y auditoria de cambios por usuario.

## Plan Looker Parity (avance actual)
- [x] Tipos agregados: Time series, Scatter, Bubble, Bullet, Boxplot, Candlestick, Pivot, Sankey, Timeline.
- [x] Contrato `dataRoles` por widget (dimension, metrica, optional metric, breakdown y fecha).
- [x] Contrato `chartConfig` por widget con normalizacion segura.
- [x] Persistencia `labelLayoutV2` por `rowKey` (compatible con offsets legacy).
- [x] Hover/cross-filter para tipos nuevos (canvas y vistas custom).
- [x] Panel de propiedades contextual por tipo (controles ocultos cuando no aplican).
- [ ] Persistencia de labels movibles v2 en todos los tipos (completado en line/time series/scatter/bubble; faltan tipos restantes).
- [ ] Render parcial por widget (dirty render) para performance en pizarras grandes.

## Notas tecnicas
- Persistencia en `project.biConfig.visual`, `project.biConfig.visualPresets` y `widget.visualOverride`.
- Seleccion activa de widget en memoria: `selectedBiWidgetId`.
- Alcance activo de configuracion en memoria: `biVisualScopeMode`.
