# Decisiones técnicas — Muenra Vouching

Formato: contexto → decisión → consecuencias.

### DT-01. Python + FastAPI para el backend
El núcleo del producto es OCR, PDF, XML y Excel; Python tiene las bibliotecas más maduras (Tesseract, pdfplumber, pypdfium2, lxml, openpyxl, RapidFuzz). FastAPI aporta validación con Pydantic y documentación OpenAPI automática. *Consecuencia:* un solo lenguaje para API, trabajador y pruebas.

### DT-02. Frontend sin compilación (ES Modules nativos)
Las firmas de auditoría suelen operar en redes restringidas y con revisiones de seguridad estrictas. Una SPA con módulos nativos, sin CDN ni dependencias npm en tiempo de ejecución, permite una CSP `script-src 'self'`, despliegue como archivos estáticos y auditoría sencilla del código. *Consecuencia:* sin TypeScript ni framework; se compensa con utilidades propias (`ui.js`) que construyen el DOM sin `innerHTML`. Migrar a un framework es posible conservando la API.

### DT-03. PostgreSQL como base y como cola
Se evitó Redis/RabbitMQ: la cola es una tabla `jobs` con `FOR UPDATE SKIP LOCKED`, lo que permite varios trabajadores, consultar el avance por proyecto y conservar el historial de errores sin otro servicio. SQLite se usa en desarrollo y pruebas (la suite completa también corre en PostgreSQL). *Consecuencia:* para volúmenes muy altos podría migrarse a un broker dedicado sin cambiar `run_job`.

### DT-04. Prioridad determinística antes que IA
XML → PDF con texto → OCR → IA. Los valores determinísticos nunca son sobrescritos por la IA; la IA solo completa campos faltantes, debe citar texto literal (verificado contra el documento) y sus valores se envían a revisión humana. Desactivada por defecto y sujeta a autorización por proyecto. *Consecuencia:* resultados reproducibles y auditables; la IA es un apoyo, no una caja negra.

### DT-05. Puntaje explicable por criterio
Cada relación guarda una tabla con criterio, prioridad, peso, valor esperado, valor extraído, puntaje, detalle, página y evidencia. El puntaje total es un promedio ponderado de los criterios evaluables, penalizado sin identificador fuerte. *Consecuencia:* el auditor ve por qué se relacionó cada soporte; pesos y umbrales son configurables por proyecto.

### DT-06. El NIT prevalece sobre el nombre
Una contradicción de NIT limita el puntaje por debajo del umbral salvo que el número documental o el archivo coincidan; en ese caso se relaciona pero como EXCEPCIÓN. *Consecuencia:* nunca se aprueba por similitud de nombre si el NIT lo contradice.

### DT-07. Asignación voraz con fases explícitas
Se prefirió un algoritmo voraz por puntaje con fases (manual → 1:1 → compartido → suma → complementario) a una optimización global (p. ej. húngaro): es explicable, determinístico y maneja relaciones 1:N y N:1 que el emparejamiento bipartito no cubre. La suma de varios soportes explora combinaciones acotadas. *Consecuencia:* casos extremos pueden requerir relación manual (documentado en restricciones).

### DT-08. Valor original inmutable
Las correcciones se guardan en `corrected_value` con usuario, fecha y comentario, y en `review_events`. El motor usa el valor efectivo y se re-ejecuta en segundo plano. Las decisiones del revisor sobreviven a nuevas ejecuciones (`manual_status` separado de `auto_status`).

### DT-09. Cifrado de archivos en la aplicación
Los documentos se cifran con Fernet (AES-128-CBC + HMAC-SHA256) antes de escribirse, y se verifica el SHA-256 del original al leerlos. *Consecuencia:* una copia del volumen sin la clave es ilegible; la clave debe custodiarse y respaldarse aparte.

### DT-10. Coordenadas normalizadas
Las regiones se guardan como fracciones (0-1) del ancho y alto de la página, independientes de la resolución de renderizado. El visor superpone los resaltados sobre una imagen PNG generada por la API, lo que evita cargar PDF.js desde Internet.

### DT-11. Excel como contrato de entrada
El Excel de referencia define población, parámetros, alias, tipos, campos y el resultado esperado, de modo que el encargo completo es reproducible desde un archivo; la hoja `RESULTADO_ESPERADO` permite medir la precisión del motor en cada ejecución.

### DT-12. Exportación protegida
Toda celda de texto que comience con `=`, `+`, `-`, `@`, tabulación o retorno se prefija con `'` para impedir inyección de fórmulas; se retiran caracteres de control ilegales.
