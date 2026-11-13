# CLAUDE.md — Liquidación de Corte (lives-liquidacion-corte)

Las reglas generales de trabajo están en el repo privado `lives-instrucciones` (CLAUDE.md). Este archivo solo describe esta herramienta.

## Qué hace
`index.html` — "Liquidación de Corte": muestra el reporte del sistema "RESUMEN LIQUIDACION CORTE - ORDENES DE CORTE TERMINADAS" y calcula solo cuántas líneas ya fueron liquidadas y cuántas faltan.
- Una línea = una orden de corte con una tela (una OC se repite si lleva más de una tela).
- Regla (decisión de Christian, 13/11/2026; reemplaza las celdas W44, W45 y W46 que él calculaba a mano en el Excel):
  - cuenta la línea que tiene dato en "% DESBALANCE" de CORTE PIEZA;
  - liquidada = ese % es menor a 1 (100%); el resto queda "Por liquidar";
  - % liquidado = líneas liquidadas ÷ líneas que cuentan.
  - La regla vive en la vista `publico_liqc_lineas` (columnas `cuenta` y `liquidada`), no en la pantalla.
- KPIs: % liquidado, líneas del reporte (y órdenes de corte), liquidadas, por liquidar.
- Gráfico "% liquidado por mes": una línea por año, con la fecha de la orden de corte. Siempre recorre todos los años (no sigue el filtro de año ni de mes; sí módulo, grupo y estilo). Usa Chart.js y el bloque `lsAnualJs` (copiado de Check List de Orden y Limpieza).
- Paneles: avance por módulo y por grupo, cada uno con los 15 que tienen más líneas por liquidar (el Excel los trae todos). Tabla con cada línea (botones Todas / Por liquidar / Liquidadas, que solo filtran la tabla). Regla 500 sin filtros.
- Filtros: buscar estilo (solo estilo), año, mes, módulo y grupo. Por defecto se ve todo lo cargado.
- Excel LIVES (Resumen igual a la pantalla y Detalle) e Imprimir. Solo lectura: no tiene formulario.
- "Merma corte" llega vacía en el reporte: se guarda pero no se muestra.

## Cómo se actualiza (decisión de Christian, 13/11/2026)
Nadie sube archivos desde la pantalla. Christian descarga el reporte del sistema a Descargas (`liquidacion de corte.xlsx`) y le pide a Claude que lo actualice. Claude:
1. copia el Excel a `AD\datos\Corte\Liquidacion` (nunca a GitHub);
2. corre `carga/cargar_liquidacion.py "<ruta>"` para revisar (se detiene si el reporte cambió de forma; solo filas visibles);
3. muestra el resumen y, con el visto bueno, corre lo mismo con `--enviar` (acción `cargar`).
Cada carga reemplaza las líneas que ya había con fecha dentro del rango del reporte y las de la misma OC + tela + color + partida. Una línea "por liquidar" de un reporte viejo solo cambia si se vuelve a descargar un reporte que la incluya.

## Base de datos (Supabase lives-datos)
- Tablas: `liqc_cargas` (un registro por reporte cargado), `liqc_lineas` (las líneas), `liqc_parametros` (`registro_libre`).
- Vista: `publico_liqc_lineas`. Funciones: `liqc_api` (puerta única: `leer`, `puede`, `cargar`) y `liqc_puede`.
- Códigos (regla "Todo codificado"): tela, color y módulo se guardan con su código del sistema (`tela_cod`, `color_cod`, `modulo_cod`) separado del nombre. En pantalla se ve el nombre; el código de tela sí se muestra.
- No usa Storage ni guarda nada en el navegador.

## Quién puede cargar
`liqc_puede`: durante las pruebas responde sí (`liqc_parametros.registro_libre = si`); Christian siempre puede; Rodrigo solo ve. Ver es libre.

## A Sistemas
- `liqc_puede` es la pieza que reemplazan con su seguridad; en SQL Server se exige usuario para abrir y leer.
- El reporte debería salir directo de su sistema (es un reporte del SIGE), sin Excel de por medio.
- Exclusivo de Postgres: jsonb, security definer, RLS, `jsonb_to_recordset`.
