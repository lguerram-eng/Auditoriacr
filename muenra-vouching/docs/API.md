# Documentación de la API — Muenra Vouching

La API REST se documenta automáticamente con OpenAPI:

- Swagger UI: `/api/docs` · ReDoc: `/api/redoc` · Esquema: `/api/openapi.json` (copia en [openapi.json](openapi.json)).

## Autenticación

```bash
curl -s -X POST http://localhost:8080/api/auth/login -H 'Content-Type: application/json' \
     -d '{"email":"auditor@muenra.local","password":"Demo#Muenra2026"}'
# → {"access_token": "...", "token_type": "bearer", "expires_in": 3600, "user": {...}, "permisos": [...]}
```

Envíe `Authorization: Bearer <token>` en cada solicitud. Errores: `401` sesión inválida, `403` sin permiso, `404` recurso inexistente o de un proyecto ajeno, `422` validación, `423` cuenta bloqueada.

## Flujo típico

```bash
TOKEN=...; P=1
# 1. Validar e importar el Excel
curl -H "Authorization: Bearer $TOKEN" -F file=@referencia.xlsx "http://localhost:8080/api/projects/$P/imports?dry_run=true"
curl -H "Authorization: Bearer $TOKEN" -F file=@referencia.xlsx  http://localhost:8080/api/projects/$P/imports
# 2. Cargar documentos (se encolan automáticamente)
curl -H "Authorization: Bearer $TOKEN" -F files=@factura.pdf -F files=@factura.xml http://localhost:8080/api/projects/$P/documents
# 3. Seguir el avance
curl -H "Authorization: Bearer $TOKEN" http://localhost:8080/api/projects/$P/jobs
# 4. Resultados, conciliación y exportación
curl -H "Authorization: Bearer $TOKEN" "http://localhost:8080/api/projects/$P/results?status=EXCEPCIÓN"
curl -H "Authorization: Bearer $TOKEN" http://localhost:8080/api/projects/$P/reconciliation
curl -H "Authorization: Bearer $TOKEN" -o resultado.xlsx http://localhost:8080/api/projects/$P/export
```

## Endpoints

| Método | Ruta | Permiso | Descripción |
|---|---|---|---|
| GET | `/api/audit` | auditoria.global | Registro de auditoría global (admin) |
| POST | `/api/auth/login` | público | Inicia sesión; devuelve JWT y permisos |
| POST | `/api/auth/logout` | autenticado | Registra el cierre de sesión |
| GET | `/api/auth/me` | autenticado | Usuario actual y permisos |
| POST | `/api/auth/password/change` | autenticado | Cambia la contraseña propia |
| POST | `/api/auth/password/forgot` | público | Solicita enlace de recuperación (respuesta genérica) |
| POST | `/api/auth/password/reset` | público | Restablece con token de un solo uso |
| GET | `/api/dashboard` | ver | Indicadores de todos los proyectos accesibles |
| GET | `/api/documents/{document_id}` | ver | Detalle: páginas, líneas con regiones, campos con evidencia, tablas, relaciones |
| DELETE | `/api/documents/{document_id}` | cargar | Detalle: páginas, líneas con regiones, campos con evidencia, tablas, relaciones |
| GET | `/api/documents/{document_id}/file` | ver | Descarga el original descifrado |
| GET | `/api/documents/{document_id}/pages/{page}/image` | ver | Página renderizada en PNG (PDF e imágenes) |
| POST | `/api/documents/{document_id}/reprocess` | procesar | Vuelve a procesar un documento |
| GET | `/api/export/sheets` | ver | Hojas exportables |
| POST | `/api/fields/{field_id}/review` | revisar | ACEPTAR / RECHAZAR / CORREGIR un campo (original inmutable) |
| GET | `/api/health` | público | Estado del servicio, OCR e IA |
| DELETE | `/api/links/{link_id}` | revisar | Retira una relación manual |
| GET | `/api/projects` | ver | Lista proyectos accesibles / crea proyecto |
| POST | `/api/projects` | proyecto.crear | Lista proyectos accesibles / crea proyecto |
| POST | `/api/projects/retention/apply` | proyecto.eliminar | Elimina proyectos cerrados con retención vencida |
| GET | `/api/projects/{project_id}` | ver | Consulta / actualiza / elimina (con archivos) un proyecto |
| PATCH | `/api/projects/{project_id}` | proyecto.editar | Consulta / actualiza / elimina (con archivos) un proyecto |
| DELETE | `/api/projects/{project_id}` | proyecto.eliminar | Consulta / actualiza / elimina (con archivos) un proyecto |
| GET | `/api/projects/{project_id}/audit` | auditoria.proyecto | Registro de auditoría del proyecto |
| GET | `/api/projects/{project_id}/dashboard` | ver | Indicadores del tablero del proyecto |
| POST | `/api/projects/{project_id}/documents` | cargar | Carga masiva (multipart `files`, `process`) / lista documentos (`status`) |
| GET | `/api/projects/{project_id}/documents` | ver | Carga masiva (multipart `files`, `process`) / lista documentos (`status`) |
| GET | `/api/projects/{project_id}/export` | exportar | Excel (`sheets=resumen,resultado,…`) |
| POST | `/api/projects/{project_id}/imports` | importar | Importa Excel (`dry_run` para solo validar) / lista importaciones |
| GET | `/api/projects/{project_id}/imports` | ver | Importa Excel (`dry_run` para solo validar) / lista importaciones |
| GET | `/api/projects/{project_id}/jobs` | ver | Estado de la cola y de los documentos |
| POST | `/api/projects/{project_id}/match` | procesar | Encola el motor de vouching |
| POST | `/api/projects/{project_id}/process` | procesar | Encola pendientes (`reprocess_errors`) y el motor |
| GET | `/api/projects/{project_id}/reconciliation` | ver | Conciliación de conteos y valores |
| GET | `/api/projects/{project_id}/references` | ver | Partidas importadas |
| GET | `/api/projects/{project_id}/results` | ver | Resultados (`status` múltiple, `q`, `exceptions_only`) |
| GET | `/api/projects/{project_id}/review-history` | ver | Historial de revisiones |
| GET | `/api/projects/{project_id}/settings` | ver | Parámetros del motor (GET / PUT) |
| PUT | `/api/projects/{project_id}/settings` | configurar | Parámetros del motor (GET / PUT) |
| GET | `/api/results/{result_id}` | ver | Detalle con explicación por criterio, historial y candidatos |
| POST | `/api/results/{result_id}/comment` | revisar | Comentario del revisor |
| POST | `/api/results/{result_id}/links` | revisar | Relación manual documento ↔ partida |
| POST | `/api/results/{result_id}/review` | revisar | Decisión APROBADO / RECHAZADO y estado manual |
| GET | `/api/users` | ver | Lista (admin: completa) / crea usuarios |
| POST | `/api/users` | usuarios | Lista (admin: completa) / crea usuarios |
| GET | `/api/users/roles` | ver | Roles y matriz de permisos |
| PATCH | `/api/users/{user_id}` | usuarios | Actualiza nombre, rol, estado o contraseña |

## Estados y códigos

- Estado técnico del documento: `PENDIENTE`, `PROCESANDO`, `PROCESADO`, `ILEGIBLE`, `ERROR`.
- Estado de vouching: `COINCIDE`, `COINCIDE CON TOLERANCIA`, `EXCEPCIÓN`, `REVISIÓN MANUAL`, `SIN SOPORTE`, `SOPORTE NO REFERENCIADO`, `POSIBLE DUPLICADO`, `DOCUMENTO ILEGIBLE`, `ERROR DE LECTURA`, `PENDIENTE`.
- Roles de relación: `PRINCIPAL`, `REPRESENTACION_GRAFICA`, `COMPLEMENTARIO`.
- Trabajos: `PROCESAR_DOCUMENTO`, `EJECUTAR_VOUCHING` con estados `EN_COLA`, `EN_PROCESO`, `COMPLETADO`, `FALLIDO`.

## Ejemplo de explicación por criterio (`GET /api/results/{id}`)

```json
{
 "estado": "EXCEPCIÓN",
 "motivos": [
  "Valor: Diferencia 12,00 % supera la tolerancia 10,00 %"
 ],
 "comparacion_valor": {
  "valor_esperado": "1.000.000",
  "valor_extraido": "1.120.000",
  "diferencia_absoluta": "120.000",
  "diferencia_porcentual": 0.12,
  "tolerancia": 0.1,
  "resultado": "FUERA_TOLERANCIA"
 },
 "soportes_principales": [
  {
   "archivo": "M004_factura_TC-4001.pdf",
   "puntaje_total": 0.8761,
   "criterios": [
    {
     "criterio": "NIT / identificación",
     "prioridad": "Muy alta",
     "peso": 30,
     "puntaje": 1.0,
     "esperado": "860.456.789-1",
     "extraido": "860.456.789-1",
     "detalle": "NIT base coincide (emisor); DV: COINCIDE",
     "pagina": 1,
     "evidencia": "NIT: 860.456.789-1"
    },
    {
     "criterio": "Valor",
     "prioridad": "Alta",
     "peso": 20,
     "puntaje": 0.352,
     "esperado": "1.000.000",
     "extraido": "1.120.000",
     "detalle": "Diferencia 12,00 % supera la tolerancia 10,00 %",
     "pagina": 1,
     "evidencia": "TOTAL A PAGAR $ 1.120.000"
    }
   ]
  }
 ]
}
```
