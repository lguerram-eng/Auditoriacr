# Manual de usuario — Muenra Vouching

Muenra Vouching ayuda a verificar que cada partida contabilizada esté respaldada por documentos soporte, con evidencia trazable por campo y página.

## 1. Roles

| Rol | Puede |
|---|---|
| **Administrador** | Todo, incluidos usuarios, auditoría global, eliminación de proyectos y retención |
| **Auditor** | Crear y configurar proyectos, importar, cargar, procesar, revisar y exportar |
| **Revisor** | Consultar, revisar (campos y partidas) y exportar |
| **Consulta** | Consultar y exportar |

Cada usuario ve únicamente los proyectos de los que es miembro (el administrador ve todos).

## 2. Inicio de sesión

Ingrese su correo y contraseña. Tras 5 intentos fallidos la cuenta se bloquea 15 minutos. Use **¿Olvidó su contraseña?** para recibir un enlace de un solo uso (válido 30 minutos). Las contraseñas requieren 10 caracteres con mayúscula, minúscula, número y símbolo; puede cambiarla en **Configuración → Mi cuenta**.

## 3. Crear proyecto

**Crear proyecto**: código único, nombre, cliente, NIT del cliente (permite distinguir facturas de venta y de compra), periodo, retención y miembros. Marque **Autorizar IA** solo si la política del encargo lo permite. Los documentos nunca se usan para entrenamiento.

## 4. Importar el Excel de referencia

1. **Importar Excel** → arrastre el archivo (`.xlsx`, `.xls` o `.csv`).
2. **Validar integridad**: revisa sin guardar y muestra el informe:
   - hojas encontradas y columnas mapeadas;
   - **errores bloqueantes**: columnas obligatorias faltantes, `ID_MUESTRA` vacío o duplicado, valores no numéricos, totales de control que no concilian;
   - **advertencias**: fechas inválidas, valores vacíos, DV del NIT incorrecto, partidas repetidas (NIT + número + valor), valores cero o negativos, tipos documentales no reconocidos;
   - conciliación: registros leídos/válidos y suma de `VALOR_ESPERADO` contra `CONTROL_TOTAL_REGISTROS` / `CONTROL_TOTAL_VALOR`.
3. **Importar**. Una nueva importación reemplaza la población vigente (queda en el historial).

### Plantilla simple (recomendada)

Use **Descargar plantilla simple**: una sola hoja `CARGA`, una fila por documento y válida para cualquier tipo documental. Obligatorias: `ID`, `TIPO_DOCUMENTO`, `NUMERO_DOCUMENTO`, `FECHA_DOCUMENTO`, `NOMBRE_TERCERO`, `NIT_IDENTIFICACION` y `VALOR_TOTAL`. Opcionales: `CONCEPTO`, `SUBTOTAL`, `IVA`, `RETENCIONES` (se comparan como referencia, sin cambiar el estado), `MONEDA`, `CONTRATO_OC` (prefijo `OC`/`ORDEN` = orden de compra; en otro caso, contrato), `ARCHIVO_ESPERADO`, `TOLERANCIA_VALOR` (tolerancia propia de la fila: 0,10 o 10 %) y `OBSERVACIONES`. Borre las filas amarillas de ejemplo antes de importar.

### Columnas de REFERENCIA_VOUCHING (plantilla completa)

| Columna | Obligatoria | Ejemplo |
|---|---|---|
| ID_MUESTRA | Sí | M001 |
| TIPO_DOCUMENTO | Sí | FACTURA, FE, RC, CE, OC, CONTRATO, CUENTA_COBRO, NC, ND… |
| NUMERO_DOCUMENTO | Sí | FV-1001 |
| FECHA | Sí | 2026-03-02 o 02/03/2026 |
| TERCERO | Sí | Proveedor Andino S.A.S. |
| NIT | Sí | 900.123.456-8 |
| VALOR_ESPERADO | Sí | 1000000 |
| MONEDA | No | COP (predeterminada) |
| CONTRATO / ORDEN_COMPRA | No | CT-2026-015 / OC-7788 |
| CONCEPTO / CENTRO_COSTO / CUENTA_CONTABLE | No | |
| ARCHIVO_SOPORTE | No | Nombre(s) de archivo esperados, separados por `;` |

Las demás hojas (CONFIGURACION, ENTIDADES_ALIAS, TIPOS_DOCUMENTO, CAMPOS_EXTRACCION, RESULTADO_ESPERADO) son opcionales; ver el [README](../README.md#formato-del-excel-de-referencia) y el ejemplo `demo/referencia_vouching_demo.xlsx`.

## 5. Cargar documentos

**Cargar documentos** → arrastre archivos o use **Seleccionar carpeta**. Formatos: PDF (texto o escaneado), XML de facturación electrónica, PNG, JPG/JPEG, TIFF y DOCX; máximo 50 MB por archivo (configurable). Cada archivo se valida (tipo real, contenido activo, antivirus), se cifra y se le calcula el SHA-256. Los rechazados se listan con su motivo. Un archivo idéntico a otro ya cargado se marca como posible duplicado.

## 6. Procesamiento

La pantalla **Procesamiento** muestra el avance en tiempo real: documentos pendientes, en proceso, procesados, ilegibles y con error, el método usado (XML, PDF_TEXTO, OCR, DOCX), el tipo documental detectado y la confianza OCR. Al terminar la lectura, el motor de vouching se ejecuta automáticamente. Botones: **Procesar pendientes**, **Reintentar errores**, **Ejecutar vouching** (por ejemplo, tras cambiar la configuración) y **Reprocesar** por documento.

## 7. Tablero principal

Indicadores: total de partidas y documentos, partidas con y sin soporte, coincidencias completas y con tolerancia, excepciones, pendientes de revisión, documentos ilegibles, valor de la población, valor soportado, valor con diferencias, porcentaje de cobertura (valor soportado / población) y de avance (lectura + revisión). Incluye la distribución por estado y la **conciliación** antes/después: conteo de partidas, suma de valores, documentos sin pérdida, ausencia de duplicación de valor y, si existe, precisión contra RESULTADO_ESPERADO.

## 8. Resultados

Tabla con filtros por estado (chips), búsqueda y ordenamiento por columna. Haga clic en una partida para abrir el visor.

### Estados

| Estado | Significado |
|---|---|
| COINCIDE | Soporte relacionado; valor idéntico y demás criterios conformes |
| COINCIDE CON TOLERANCIA | Diferencia de valor dentro de la tolerancia (10 % por defecto) |
| EXCEPCIÓN | Diferencia por investigar: valor fuera de tolerancia, NIT contradictorio, fecha fuera de tolerancia o moneda distinta |
| REVISIÓN MANUAL | Requiere criterio del auditor: valor esperado cero, valor no extraído, confianza OCR baja, valor por IA, DV distinto, tercero no confirmado |
| SIN SOPORTE | Ningún documento alcanzó el umbral de relación |
| SOPORTE NO REFERENCIADO | Documento sin partida relacionada |
| POSIBLE DUPLICADO | Archivo repetido, misma factura cargada en dos archivos del mismo formato, o partida repetida en la población |
| DOCUMENTO ILEGIBLE | Sin texto suficiente tras OCR |
| ERROR DE LECTURA | Falla técnica al leer el archivo |
| PENDIENTE | Aún no procesado |

Una **excepción no es una conclusión de fraude**: es una diferencia que debe investigarse.

## 9. Visor documental

- **Izquierda**: el documento (página renderizada) con las regiones de cada campo resaltadas; el campo activo se marca en rojo. Para XML y DOCX se muestra el texto con la línea de evidencia resaltada. Pestañas para el soporte principal (★), la representación gráfica (⎘) y los complementarios (＋). **Descargar original** descarga el archivo descifrado (queda en la auditoría).
- **Derecha**:
  - **Comparación de valor**: esperado, extraído, diferencia absoluta y porcentual, tolerancia aplicada y resultado.
  - **Por qué se relacionó este documento**: cada criterio con prioridad, peso, valor esperado, valor extraído y puntaje. Haga clic en un criterio para ir a la página y resaltar el texto.
  - **Campos extraídos y evidencia**: valor, página, método, confianza y texto de evidencia. **Aceptar**, **Corregir** (comentario obligatorio) o **Rechazar** (comentario obligatorio). El valor original siempre se conserva y el motor se recalcula.
  - **Decisión del revisor**: aprobar, rechazar o cambiar el estado (comentario obligatorio), agregar comentarios y **relacionar documentos manualmente** como principal o complementario.
  - **Historial de la partida** con usuario, fecha y hora.

## 10. Panel de excepciones

Agrupa EXCEPCIÓN, REVISIÓN MANUAL, POSIBLE DUPLICADO, SIN SOPORTE, ILEGIBLE y ERROR con sus motivos y valor; lista además los documentos no referenciados o con incidencias. **Revisar** abre el visor.

## 11. Revisión y aprobación

Tres vistas: excepciones sin decisión, coincidencias por aprobar (con aprobación en lote) y revisadas. Cada decisión registra usuario, fecha, comentario y estado final; el estado automático se conserva.

## 12. Configuración

Tolerancia de valor, política para valor esperado cero, umbral de similitud de nombres (85 %), tolerancia de días, puntaje mínimo de relación, confianza mínima OCR, máximo de soportes combinables, pesos por criterio, autorización de IA y retención. Guarde y pulse **Ejecutar vouching** para aplicar.

## 13. Historial de auditoría

**Accesos y cambios** (inicios de sesión, consultas, cargas, importaciones, procesamientos, cambios de configuración, exportaciones, descargas) e **Historial de revisiones** (correcciones y decisiones). El administrador dispone además del historial global.

## 14. Exportaciones

Seleccione las hojas o **Exportar todo**: resumen ejecutivo, resultado completo, coincidencias, excepciones, partidas sin soporte, soportes no referenciados, evidencias por campo, historial de revisiones, parámetros utilizados y registro del modelo/OCR/reglas. Cada fila del resultado incluye ID, archivo, hash, tipo, página de evidencia, número/fecha/tercero/NIT/valor extraídos, valor esperado, diferencias, similitud, coincidencias de NIT y número, diferencia en días, confianza OCR, estado, motivo, texto de evidencia, revisor y fecha de revisión.

## 15. Usuarios y permisos (administrador)

Crear usuarios con contraseña temporal (se exige cambio), cambiar rol, activar/desactivar y restablecer contraseñas. La matriz de permisos se muestra en la misma pantalla.
