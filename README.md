# Control y análisis de compras y facturas

Proyecto de portafolio desarrollado en Excel con Power Query y tablas dinámicas. Consolida archivos mensuales, normaliza códigos de proveedor, cruza maestros y clasifica registros que requieren revisión.

Todos los datos son simulados. La empresa Comercial Horizonte y sus proveedores son ficticios. Los montos son netos en pesos chilenos, sin IVA; no representan pagos, pérdidas confirmadas ni costo de ventas.

## Contenido

```text
Proyecto_control_ facturas/
├── Control_Facturas..xlsx
├── LEEME.txt
├── Facturas/
│   ├── Facturas_2026_07.csv
│   ├── Facturas_2026_08.csv
│   └── Facturas_2026_09.csv
└── Maestros/
    ├── Proveedores.csv
    └── Ordenes_Compra.csv
```

El nombre del libro se conserva tal como está en el proyecto original, incluidos los dos puntos antes de xlsx.

[Descargar el libro Excel](Control_Facturas..xlsx) · [Ver las facturas de ejemplo](Facturas/) · [Ver los maestros](Maestros/)

## Qué demuestra

- Consolidación de archivos de una carpeta con Power Query.
- Tratamiento de tipos de datos y normalización de códigos mediante recorte de espacios y conversión a mayúsculas.
- Cruces con maestros de proveedores y órdenes de compra mediante combinaciones externas izquierdas.
- Comparación entre montos facturados y autorizados.
- Identificación de OC ausentes o inexistentes y proveedores no registrados.
- Agrupación de registros para identificar posibles duplicados.
- Tablas dinámicas e indicadores de resumen basados en los resultados del control.
- Conservación de inconsistencias del origen para su análisis, sin eliminarlas automáticamente.

## Flujo del análisis

CSV mensuales → Facturas_Consolidadas → Conteo_documentos → Control_Facturas → Tablas dinámicas → RESUMEN.

Los maestros Proveedores y Ordenes_Compra enriquecen la consolidación. La consulta Facturas_Julio conserva el trabajo inicial de un único mes. INTRUCCIONES explica las reglas y el procedimiento de actualización.

Cada fila representa un registro recibido de un documento completo. ID_Registro identifica la fila importada; no garantiza que el documento sea único. En esta práctica se usa proveedor + tipo de documento + folio + fecha como clave de agrupación, conforme a INTRUCCIONES. Los folios del conjunto simulado se repiten entre meses, por lo que excluir la fecha cambiaría la regla y los resultados. Esta regla debe revisarse para un caso real.

## Cómo abrir y actualizar

Se necesita Excel de escritorio con soporte de Power Query. Los CSV usan UTF-8, separador punto y coma y fechas de factura día/mes/año. No basta con previsualizar el archivo en GitHub para ejecutar las consultas.

1. Descarga el repositorio completo y conserva las carpetas Facturas y Maestros junto al libro.
2. Abre Control_Facturas..xlsx en Excel. Los resultados de la última actualización están guardados en el archivo.
3. Antes de actualizar, ve a Datos → Obtener datos → Configuración de origen de datos y cambia los orígenes que apuntan al escritorio del autor por las ubicaciones de tu copia: carpeta Facturas, CSV de julio, los dos maestros y LEEME.txt. Si tu versión no permite cambiar un origen desde ese panel, ajusta únicamente su paso Origen en el editor de Power Query. No elimines ni reordenes los pasos de transformación.
4. Comprueba la interpretación de Fecha_Emision como día/mes/año antes de cargar. En equipos con otra configuración regional, revisa la configuración regional de conversión de fechas.
5. Actualiza las consultas y espera a que terminen. Después actualiza las tres tablas dinámicas.
6. Revisa RESUMEN y contrasta los resultados con la tabla final. El texto de hallazgos es manual y debe revisarse después de cada actualización.
7. Conserva en Facturas únicamente los CSV que quieras consolidar. La consulta actual recorre también subcarpetas y no restringe explícitamente la extensión; documentos de otro tipo o copias de CSV dentro de esa carpeta pueden provocar errores o duplicar la importación.

El proyecto conserva rutas absolutas de la máquina del autor. La actualización en otro equipo requiere los ajustes anteriores; no se presenta como una solución portátil sin configuración.

## Resultados del conjunto de ejemplo

| Indicador | Resultado |
|---|---:|
| Registros recibidos | 37 |
| Total neto de los registros recibidos | $6.307.500 |
| Registros con alertas | 7 |
| Porcentaje de registros con alertas | 18,92 % |
| Registros sin alertas | 30 |
| Diferencias netas contra OC | $40.000 |
| Monto de registros sin OC válida | $335.000 |
| Registros marcados como posible duplicado | 2 |
| Documento potencialmente repetido | 1 |
| Importe repetido potencial | $102.500 |

El total neto incluye los registros recibidos, incluso el posible duplicado. No es una cifra de gasto depurada ni de pagos. Los indicadores de alertas pueden solaparse y no deben sumarse como una pérdida total.

## Validación y límites

Se realizaron 69 comprobaciones independientes: correspondencia fila a fila con los CSV, cruces con maestros, indicadores, agrupación de duplicados, conciliación de las tres tablas dinámicas y ausencia de errores Excel guardados. No se modificaron los archivos originales ni las consultas.

La revisión comprobó el código M y los resultados guardados. No ejecutó el motor de Power Query ni una actualización nativa de Excel. La apertura, actualización y recálculo en Excel de escritorio requieren una comprobación adicional en el equipo que lo vaya a usar.

Se conservan estos límites de la versión actual:

- El indicador RESUMEN!C8 divide entre dos el importe de los registros duplicados. Coincide con el ejemplo de dos copias iguales; no generaliza a tres o más copias. La columna Monto_Duplicado_Potencial de Control_Facturas ya calcula otro agregado, bajo el supuesto documentado de importes iguales.
- RESUMEN!C6 suma diferencias con signo. En el ejemplo ambas diferencias son positivas. Si hubiera facturas por debajo del autorizado, estas compensarían los excesos; el resultado sería diferencia neta, no el total de excesos positivos.
- Algunas fórmulas GETPIVOTDATA dependen de categorías concretas de las tablas dinámicas. Si una categoría desaparece o se oculta, pueden devolver un error de referencia. No hay errores guardados actualmente.
- El texto de hallazgos de RESUMEN!B11 es manual, conforme a INTRUCCIONES. No se actualiza automáticamente con los datos.

Estas limitaciones no alteran los resultados comprobados del conjunto de ejemplo. Deben resolverse o mantenerse explícitas antes de presentar el proyecto como un control automático generalizable.

Referencia: [documentación de Microsoft sobre GETPIVOTDATA](https://support.microsoft.com/en-us/excel/functions/getpivotdata-function).
