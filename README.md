# Lead Instant Response Bot

Servicio webhook en Python (FastAPI) para **agencias de marketing, inmobiliarias y negocios con formularios web / Meta Ads**. Recibe cada lead, valida su correo, evita duplicados con SQLite, lo califica con OpenAI, envía respuesta inmediata por correo SMTP y exporta de forma segura a Excel (`.xlsx`).

## Características de confiabilidad
- **Autenticación por header `X-Webhook-Secret`** para bloquear envíos no autorizados.
- **Deduplicación en SQLite** por `(email, source)`: reintentos de formularios no envían correos repetidos al cliente ni duplican filas.
- **Fallback ante fallos de IA/SMTP**: nunca pierde un lead si OpenAI o el servidor SMTP fallan temporalmente.
- **Escritura Excel thread-safe y sanitizada** contra *Excel formula injection*.

## Instalación y uso
```bash
uv sync
cp .env.example .env
uv run uvicorn app:app --reload --port 8001
```

## Pruebas automatizadas
```bash
uv run pytest
```
