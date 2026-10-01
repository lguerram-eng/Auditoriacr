# Manual de instalación — Muenra Vouching

## 1. Requisitos

| Componente | Versión mínima | Notas |
|---|---|---|
| Docker Engine + Compose v2 | 24 / 2.20 | Opción recomendada |
| CPU / RAM | 2 vCPU / 4 GB | El OCR es intensivo en CPU; escale el servicio `worker` |
| Disco | 20 GB + documentos | Los documentos se guardan cifrados en el volumen `documentos` |
| Python (sin Docker) | 3.11 | Más Tesseract OCR 5 con idioma `spa` |

## 2. Instalación con Docker Compose

```bash
git clone <repositorio> && cd muenra-vouching
cp .env.example .env
```

Genere y complete los secretos en `.env`:

```bash
python3 -c "import secrets; print('MUENRA_SECRET_KEY=' + secrets.token_urlsafe(64))"
python3 -c "from cryptography.fernet import Fernet; print('MUENRA_STORAGE_ENCRYPTION_KEY=' + Fernet.generate_key().decode())"
```

Defina también `POSTGRES_PASSWORD` y `MUENRA_BOOTSTRAP_ADMIN_PASSWORD` (mínimo 10 caracteres, mayúscula, minúscula, número y símbolo). **Respalde la clave de cifrado**: sin ella los documentos no se pueden descifrar.

```bash
docker compose up -d --build
docker compose ps            # db, api (healthy), worker, frontend
```

- Interfaz: `http://localhost:8080` (puerto configurable con `HTTP_PORT`).
- API y documentación interactiva: `http://localhost:8080/api/docs`.
- La API aplica las migraciones Alembic al arrancar (`alembic upgrade head`).
- Ingrese con `MUENRA_BOOTSTRAP_ADMIN_EMAIL` / `MUENRA_BOOTSTRAP_ADMIN_PASSWORD`.

### Datos de demostración (ficticios)

```bash
docker compose exec api python -m app.seed --forzar
```

Crea los usuarios `admin@`, `auditor@`, `revisor@` y `consulta@muenra.local` (contraseña `Demo#Muenra2026`) y el proyecto `DEMO-2026` con 17 partidas y 21 documentos. Úselo solo en instalaciones de prueba.

### Antivirus

```bash
echo "MUENRA_CLAMAV_HOST=clamav" >> .env
echo "MUENRA_ANTIVIRUS_REQUIRED=true" >> .env     # rechaza archivos si ClamAV no responde
docker compose --profile antivirus up -d
```

La primera descarga de firmas de ClamAV tarda algunos minutos.

### Escalar el procesamiento

```bash
WORKER_REPLICAS=4 docker compose up -d --scale worker=4
```

Los trabajadores comparten la cola en PostgreSQL (`FOR UPDATE SKIP LOCKED`), sin duplicar trabajos.

### TLS (cifrado en tránsito)

Termine TLS en un balanceador o en Nginx: monte los certificados en el contenedor `frontend` y descomente el bloque `listen 443` de `frontend/nginx.conf`. En producción la API envía `Strict-Transport-Security`.

### IA opcional

`MUENRA_LLM_ENABLED=true` y `ANTHROPIC_API_KEY=...` en `.env`; luego autorice la IA en cada proyecto (Configuración → Privacidad). La IA solo completa campos faltantes y cada valor debe citar texto literal del documento.

### Correo (recuperación de contraseña)

Configure `MUENRA_SMTP_*` y `MUENRA_PUBLIC_URL`. Sin SMTP, en modo `development` el enlace de recuperación se muestra en pantalla; en `production` no se revela.

## 3. Instalación local sin Docker (desarrollo)

```bash
# Ubuntu/Debian
sudo apt install tesseract-ocr tesseract-ocr-spa
# macOS
brew install tesseract tesseract-lang

cd muenra-vouching/backend
python3.11 -m venv .venv && source .venv/bin/activate
pip install -r requirements-dev.txt
python -m app.seed                      # SQLite en backend/data/, usuarios y proyecto de demostración
uvicorn app.main:app --reload           # http://localhost:8000
```

Por defecto: `MUENRA_ENV=development`, SQLite, trabajador en hilo (`MUENRA_INLINE_WORKER=true`) y la interfaz servida por la propia API desde `../frontend`.

Para usar PostgreSQL localmente: `export MUENRA_DATABASE_URL=postgresql+psycopg://usuario:clave@localhost:5432/muenra`, `alembic upgrade head` y, en otra terminal, `python -m app.worker` con `MUENRA_INLINE_WORKER=false`.

## 4. Pruebas

```bash
cd backend
pytest                               # SQLite temporal
MUENRA_TEST_DATABASE_URL=postgresql+psycopg://usuario@localhost:5432/muenra_test pytest
```

Las pruebas de OCR se omiten automáticamente si Tesseract no está instalado.

## 5. Regenerar los datos ficticios

```bash
cd backend && python -m demo.generate_demo ../demo
```

## 6. Operación

| Tarea | Comando |
|---|---|
| Respaldo de la base | `docker compose exec db pg_dump -U muenra muenra > respaldo.sql` |
| Respaldo de documentos | Copie el volumen `documentos` (cifrado) **y** la clave `MUENRA_STORAGE_ENCRYPTION_KEY` por separado |
| Nueva migración | `docker compose exec api alembic upgrade head` (automático al reiniciar) |
| Registros | `docker compose logs -f api worker` (datos sensibles enmascarados) |
| Retención | Administrador: `POST /api/projects/retention/apply` elimina proyectos cerrados con plazo vencido |

## 7. Variables de entorno

Todas llevan el prefijo `MUENRA_` y están documentadas en [`.env.example`](../.env.example). Las más importantes:

| Variable | Predeterminado | Descripción |
|---|---|---|
| `MUENRA_ENV` | `development` | `production` exige secretos y oculta tokens de desarrollo |
| `MUENRA_DATABASE_URL` | SQLite local | URL SQLAlchemy |
| `MUENRA_SECRET_KEY` | — | Firma de sesiones (obligatoria en producción) |
| `MUENRA_STORAGE_ENCRYPTION_KEY` | — | Clave Fernet de cifrado en reposo (obligatoria en producción) |
| `MUENRA_MAX_FILE_MB` | 50 | Tamaño máximo por archivo |
| `MUENRA_OCR_LANGUAGES` | `spa+eng` | Idiomas Tesseract |
| `MUENRA_INLINE_WORKER` | `true` | Trabajador en hilo (desarrollo) |
| `MUENRA_LLM_ENABLED` | `false` | Habilita la IA (además requiere autorización por proyecto) |
