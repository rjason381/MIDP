# Analisis del MIDP EIMI

Archivo fuente: `C:\Users\user\Downloads\EIMI-AITEC-DS-000-ZZZ-MIDP-ZZZ-ZZZ-001.xlsx`

## Resumen

- Entregables extraidos: 493
- Hoja usada: `MIDP_EIMI`
- Proyecto generado: `MIDP - EIMI Real`
- Estado importable limpio: `eimi_midp_clean_import_state.json`
- Proyecto suelto para merge: `eimi_midp_project_only.json`

## Integridad de fuente

- La hoja `MIDP_EIMI` tiene una cabecera multinivel entre filas 10 y 13.
- Los datos utiles empiezan en la fila 14 y llegan hasta la fila 506.
- La hoja no trae cronograma usable para `startDate/endDate` en el modelo MIDP actual.
- La hoja tampoco trae un peso de avance comparable con `baseUnits`; por eso se carg? `baseUnits = 1` por entregable para habilitar agregaciones reales por conteo.
- `realProgress` se dej? vacio para no inventar avance que el archivo no aporta de forma consistente.

## Perfil de datos

- Creadores unicos: 8
- Disciplinas unicas: 16
- Sistemas unicos: 26
- Tipos unicos: 8
- Niveles unicos: 21

### Top creadores
- AITEC: 336
- ARTADI: 86
- LUMI: 16
- FALEN: 15
- PLED: 12
- FNOX: 11
- FB: 9
- LAND: 8

### Top disciplinas
- ELE: 119
- ARQ: 90
- EST: 53
- MEC: 51
- SAN: 39
- DACI: 25
- ACI: 23
- SEG: 19
- COM: 16
- INT: 16

### Top tipos
- PLN: 446
- MOD: 15
- MD: 9
- ANEX: 9
- ET: 7
- MC: 5
- PEB: 1
- LEV: 1

## Visuales precargados

### BI
- Entregables por Disciplina
- Entregables por Creador
- Entregables por Tipo

### Dechini Quicksigth
- Entregables por Disciplina
- Entregables por Creador
- Tipos por Disciplina

## Observaciones

- `Estado` en el Excel aparece mezclado entre `100%` y `1`; no es una base suficientemente confiable para mapear progreso real sin una regla funcional adicional.
- `Revision` viene practicamente constante en `1`, asi que no aporta valor analitico fuerte por ahora.
- Los sistemas genericos como `ZZZ` o `DT` aparecen bajo varias disciplinas; el proyecto importado asigna una disciplina dominante por compatibilidad con la UI actual de MIDP.
