# CLAUDE.md — Liquidación de Corte (lives-liquidacion-corte)

Las reglas generales de trabajo están en el repo privado `lives-instrucciones` (CLAUDE.md). Este archivo solo describe esta herramienta.

## Qué hace
`index.html` — "Liquidación de Corte": muestra el reporte del sistema "RESUMEN LIQUIDACION CORTE - ORDENES DE CORTE TERMINADAS" y calcula solo cuántas líneas ya fueron liquidadas y cuántas faltan.
- Lo que se cuenta es cada TELA de cada orden de corte (una OC se repite si lleva más de una tela). En pantalla y en el Excel se llama "telas", nunca "líneas" (decisión de Christian, 13/11/2026: "líneas del reporte" no se entendía). Verificado con el reporte 2025-2026: 4 746 telas en 3 837 OC y 1 635 OP; no coincide con la OP ni con la OC, y el reporte no trae PO. En este archivo y en la base (`liqc_lineas`) "línea" sigue siendo el nombre interno.
- Regla (decisión de Christian, 13/11/2026; reemplaza las celdas W44, W45 y W46 que él calculaba a mano en el Excel):
  - cuenta la línea que tiene dato en "% DESBALANCE" de CORTE PIEZA;
  - liquidada = ese % es menor a 1 (100%); el resto queda "Por liquidar";
  - % liquidado = líneas liquidadas ÷ líneas que cuentan.
  - La regla vive en la vista `publico_liqc_lineas` (columnas `cuenta` y `liquidada`), no en la pantalla.
- KPIs: % liquidado, telas de órdenes de corte (y en cuántas OC), liquidadas, por liquidar. Al presionar la tarjeta "Liquidadas" o "Por liquidar", la lista de abajo queda solo con esas y la pantalla baja a la lista; al volver a presionarla se quita (pedido de Christian, 13/11/2026). Es el mismo filtro que los botones Todas / Por liquidar / Liquidadas de la lista: no cambia los KPI ni los gráficos.
- Gráfico "% liquidado por mes": una línea por año, con la fecha de la orden de corte. Siempre recorre todos los años (no sigue el filtro de año ni de mes; sí módulo, grupo y estilo). Usa Chart.js y el bloque `lsAnualJs` (copiado de Check List de Orden y Limpieza). Leyenda arriba a la izquierda, con "Año 2026", "Año 2025" (pedido de Christian, 13/11/2026: la anterior, chica y a la derecha, no se notaba); el plugin `liqLegAire` deja espacio bajo la leyenda para que no tape las etiquetas.
- Paneles: avance por módulo y por grupo, cada uno con los 15 que tienen más líneas por liquidar (el Excel los trae todos).
- Líneas del reporte: en TARJETAS por defecto, con botón para elegir Tarjetas o Tabla (pedido de Christian, 13/11/2026); la elección se recuerda en ese equipo (`localStorage` `lives_liqc_vista`). Botones Todas / Por liquidar / Liquidadas, que solo filtran las líneas. Regla 500 sin filtros. En celular solo la tabla se pliega tras "Ver tabla"; las tarjetas se ven directo.
- Filtros: buscar estilo (solo estilo), año, mes, módulo y grupo. Estilo y grupo son cuadros donde se escribe y van saliendo las opciones (`<datalist>`), y también se puede abrir la lista y elegir (pedido de Christian, 13/11/2026); el grupo filtra por lo escrito (contiene). Por defecto se ve todo lo cargado. La barra lleva `data-ls-fijo` para que lo escrito en los filtros no cuente como trabajo sin guardar. Botón "Limpiar" en la barra (pedido de Christian, 13/11/2026): quita de una vez estilo, año, mes, módulo, grupo y el filtro Liquidadas / Por liquidar; sale apagado cuando no hay nada que limpiar.
- Excel LIVES (Resumen igual a la pantalla y Detalle) e Imprimir. Solo lectura: no tiene formulario.
- "Merma corte" llega vacía en el reporte: se guarda pero no se muestra.
- Datos cargados al 13/11/2026: 9 495 telas de 2023 a 2026 (reporte 01/01/2023–13/12/2024 y reporte 01/01/2025–13/11/2026). Del 13 al 31/12/2024 no hay nada cargado.

## Cómo se actualiza (decisión de Christian, 13/11/2026)
Nadie sube archivos desde la pantalla. Christian descarga el reporte del sistema a Descargas (`liquidacion de corte.xlsx`) y le pide a Claude que lo actualice. Claude:
1. copia el Excel a `AD\datos\Corte\Liquidacion` (nunca a GitHub);
2. corre `carga/cargar_liquidacion.py "<ruta>"` para revisar (se detiene si el reporte cambió de forma; solo filas visibles); Solo se saltan las filas sin OC: las que traen el grupo en blanco SÍ entran (corregido el 13/11/2026; en el avance por grupo salen como "Sin dato").
3. muestra el resumen y, con el visto bueno, corre lo mismo con `--enviar` (acción `cargar`).
Alertas LIVES (desde el 17/11/2026, registro `Oc04`): al cargar, `liqc_api` compara cada tela con lo que había ANTES de reemplazar (`opcambios_liq_comparar_`) y anota en `opcambios_liq_historial` las telas que quedaron liquidadas, las que volvieron a "por liquidar" y las que cambiaron su % de desbalance de pieza. La respuesta de `cargar` trae `cambios` (cuántas alertas) y `cambios_error` (si la comparación falla, la carga igual se hace); el programa lo dice al terminar. Para cargar un reporte histórico sin llenar la pantalla de alertas: `--enviar --sin-alertas`. Detalle en el CLAUDE.md de `lives-alertas-op`.

Cada carga reemplaza las líneas que ya había con fecha dentro del rango del reporte y las de la misma OC + tela + color + partida. Una línea "por liquidar" de un reporte viejo solo cambia si se vuelve a descargar un reporte que la incluya.

Siempre el año completo (decisión de Christian, 17/11/2026): el reporte se descarga del 1 de enero del año en curso a la fecha, nunca solo el mes. Con solo el mes no se borra nada, pero lo que quedó "Por liquidar" en meses anteriores se queda congelado aunque después lo liquiden, y el % liquidado sale más bajo que la realidad.
- Antes de enviar, Claude confirma que el rango que muestra el programa empieza el 1 de enero del año en curso; si trae solo un mes u otro rango, avisa antes de cargar.
- En enero y febrero conviene descargar desde el 1 de enero del año anterior, para que diciembre no quede congelado.
- Los años anteriores solo se ponen al día si Christian descarga un reporte que los incluya (al 17/11/2026 tenían por liquidar: 240 telas de 2025 y 318 de 2023-2024).

## Base de datos (Supabase lives-datos)
- Tablas: `liqc_cargas` (un registro por reporte cargado), `liqc_lineas` (las líneas), `liqc_parametros` (`registro_libre`).
- Vista: `publico_liqc_lineas`. Funciones: `liqc_api` (puerta única: `leer`, `puede`, `cargar`) y `liqc_puede`. Desde el 13/11/2026 (Lq02) `liqc_api` devuelve `json`, no `jsonb`: con 9 495 líneas el jsonb tardaba hasta 9 s y la base cortaba la consulta; con json la lectura queda en unos 2 s. Si la base sigue creciendo, lo siguiente es que `leer` entregue menos columnas o por año.
- Códigos (regla "Todo codificado"): tela, color y módulo se guardan con su código del sistema (`tela_cod`, `color_cod`, `modulo_cod`) separado del nombre. En pantalla se ve el nombre; el código de tela sí se muestra.
- No usa Storage. En el navegador solo guarda la vista elegida (tarjetas o tabla).

## Quién puede cargar
`liqc_puede`: durante las pruebas responde sí (`liqc_parametros.registro_libre = si`); Christian siempre puede; Rodrigo solo ve. Ver es libre.

## A Sistemas
- `liqc_puede` es la pieza que reemplazan con su seguridad; en SQL Server se exige usuario para abrir y leer.
- El reporte debería salir directo de su sistema (es un reporte del SIGE), sin Excel de por medio.
- Exclusivo de Postgres: jsonb, security definer, RLS, `jsonb_to_recordset`.
