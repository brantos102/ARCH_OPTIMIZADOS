# ARCH_OPTIMIZADOS

Repositorio de los archivos optimizados a código Excel (VBA).

| Proyecto | Descripción |
|---|---|
| [`IMP_STOCK_DETALLADO/`](IMP_STOCK_DETALLADO/) | Impresión de etiquetas en lote desde la hoja `Consolidado` hacia una Zebra ZD421 (respeta el autofiltro, permite elegir copias por etiqueta y envía lotes de 100 etiquetas por trabajo de impresión) |

Cada proyecto incluye:

- `src/` — módulos `.bas` listos para importar en el editor de VBA (`Alt + F11`).
- `src/_original/` — el código tal como estaba en el libro, como respaldo.
- `docs/` — análisis del código original y documentación funcional.
