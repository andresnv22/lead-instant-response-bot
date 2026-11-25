# Lead Instant Response Bot (Web → Respuesta Inmediata → Excel/CRM)

Bot en Python (FastAPI) pensado para **agencias de marketing, inmobiliarias y negocios que pautan en Meta/Google Ads**. Recibe un lead desde un formulario web, lo califica con IA en segundos, le envía respuesta inmediata y lo guarda en Excel (`.xlsx`).

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
