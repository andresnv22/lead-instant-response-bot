# Lead Instant Response Bot (Web → Respuesta Inmediata → Excel/CRM)

Bot en Python (FastAPI) pensado para **agencias de marketing, inmobiliarias y negocios que pautan en Meta/Google Ads**. Recibe un lead desde un formulario web, lo califica, prepara una respuesta y lo guarda en Excel (`.xlsx`).

Por defecto `DEMO_MODE=true`: la respuesta por email se simula. Para enviarla de verdad, configura `DEMO_MODE=false`, `SMTP_USER` y `SMTP_PASS`. Si falta la configuración o SMTP falla, la petición devuelve error; no informa falsamente que el correo salió. El endpoint recibe leads enviados a él, pero no está conectado automáticamente con Meta/Google Ads.

## Instalación rápida
```bash
python -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt
cp .env.example .env
uvicorn app:app --reload --port 8001
```

## Probar en modo demo
```bash
curl -X POST http://127.0.0.1:8001/webhook/lead \
  -H "Content-Type: application/json" \
  -d '{"name": "Carlos Ruiz", "email": "carlos@empresa.com", "phone": "+573009998877", "source": "Meta Ads", "message": "Hola, quiero cotizar urgente una página web"}'
```
Esto responderá con el `score` (`ALTO`), el mensaje de respuesta inmediata y creará/actualizará `leads_calificados.xlsx`.
