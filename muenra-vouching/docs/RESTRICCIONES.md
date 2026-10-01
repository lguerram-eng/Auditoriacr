# Restricciones conocidas — Muenra Vouching

## Extracción y OCR

- **Extracción por reglas.** Los campos de PDF/imagen se obtienen con etiquetas y expresiones regulares adaptadas a documentos colombianos (español). Formatos atípicos (etiquetas inusuales, varias facturas por archivo, tablas complejas) pueden requerir revisión manual o la IA opcional.
- **Un documento por archivo.** Un PDF que contiene varias facturas se trata como un solo documento (toma el primer número y el total más relevante). Divida los PDF compuestos antes de cargarlos.
- **Calidad del OCR.** Tesseract funciona bien con escaneos de 200-300 DPI. Fotografías inclinadas, con sombras, manuscritas o de muy baja resolución pueden producir confianza baja (se envían a REVISIÓN MANUAL) o quedar como DOCUMENTO ILEGIBLE. No se corrige la inclinación (deskew) automáticamente.
- **Coordenadas.** PDF con texto y OCR guardan la región exacta de cada línea. XML y DOCX no tienen coordenadas: el visor muestra el texto y resalta la línea de evidencia.
- **Tablas.** Se detectan tablas en PDF con texto (pdfplumber), XML (líneas de factura) y DOCX; las tablas en imágenes escaneadas no se estructuran.
- **DOCX.** Se leen párrafos y tablas; no se procesan imágenes incrustadas ni encabezados/pies de página.
- **XML.** Se soporta UBL 2.1 (Invoice, CreditNote, DebitNote y AttachedDocument de la DIAN). Otros esquemas XML se leen como `OTRO` sin campos estructurados. No se valida la firma digital ni el CUFE contra la DIAN.
- **Clasificación factura de venta / compra.** Requiere el NIT del cliente en el proyecto o en `CONFIGURACION` (`NIT_CLIENTE`); sin él, las facturas se clasifican como de compra.

## Motor de coincidencias

- La combinación de varios soportes para alcanzar un valor explora hasta `MAX_SOPORTES_SUMA` (4) documentos entre los 12 mejores candidatos del mismo tercero; combinaciones mayores requieren relación manual.
- El emparejamiento de pares usa bloqueo (descarta pares sin ninguna señal común). Para poblaciones de decenas de miles de partidas el tiempo crece en forma aproximadamente cuadrática; se recomienda dividir por periodos o proveedores.
- No hay conversión de moneda: una moneda distinta genera EXCEPCIÓN.
- Los alias se aplican tal como se cargan en `ENTIDADES_ALIAS`; no se consultan registros externos (RUES/DIAN).

## IA

- Opcional y desactivada por defecto. Requiere conexión a la API de Anthropic y autorización por proyecto.
- Solo completa campos faltantes; los valores sin cita literal verificable se descartan; los valores obtenidos por IA siempre se envían a revisión humana.
- Textos de más de 150.000 caracteres se envían truncados (queda advertencia en el documento).

## Seguridad e infraestructura

- **Cifrado en reposo de la base de datos.** Los documentos se cifran en la aplicación (Fernet). La base PostgreSQL contiene texto extraído y debe ubicarse sobre un volumen cifrado (LUKS, cifrado del proveedor de nube, etc.).
- **Rotación de la clave de cifrado.** No hay herramienta de recifrado automático; rotar la clave exige descifrar y volver a cifrar los archivos.
- **Sesiones.** Los JWT no se revocan individualmente antes de expirar (60 minutos por defecto); desactivar un usuario impide nuevas solicitudes inmediatamente porque cada solicitud verifica que esté activo.
- **Antivirus.** Sin ClamAV configurado solo se aplican las verificaciones internas (firma, contenido activo, EICAR); en producción se recomienda `MUENRA_ANTIVIRUS_REQUIRED=true`.
- **Autenticación.** No incluye inicio de sesión único (SAML/OIDC) ni segundo factor; se recomienda publicar detrás de un proxy con SSO si se requiere.
- **Idioma.** La interfaz está en español; el diccionario en inglés (`frontend/js/i18n.js`) cubre la navegación y está preparado para completarse.
- **Docker.** El repositorio se verificó ejecutando los mismos comandos de los contenedores (migración, API y trabajador separados sobre PostgreSQL 16 en modo producción); el entorno donde se construyó no permitía ejecutar el demonio de Docker, por lo que la construcción de las imágenes debe validarse en su infraestructura.
