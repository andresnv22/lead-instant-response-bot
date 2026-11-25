import json
import os
import smtplib
import ssl
from datetime import datetime
from email.message import EmailMessage
from pathlib import Path

from dotenv import load_dotenv
from fastapi import FastAPI, HTTPException
from openai import OpenAI
from openpyxl import Workbook, load_workbook
from pydantic import BaseModel

load_dotenv()

app = FastAPI(
    title="Lead Instant Response Bot",
    description="Recibe leads web, los califica con IA en segundos, responde al instante y guarda en Excel.",
    version="1.0.0",
)

EXCEL_PATH = Path(os.getenv("LEADS_EXCEL_PATH", "leads_calificados.xlsx"))


class LeadPayload(BaseModel):
    name: str
    email: str
    phone: str
    source: str = "web_form"
    message: str


def safe_excel_text(value: str) -> str:
    return "'" + value if value.lstrip().startswith(("=", "+", "-", "@")) else value


def qualify_lead(lead: LeadPayload) -> dict:
    """Califica el lead (ALTO / MEDIO / BAJO) y redacta respuesta inmediata usando OpenAI."""
    api_key = os.getenv("OPENAI_API_KEY")
    if not api_key:
        raise HTTPException(status_code=500, detail="OPENAI_API_KEY no está configurada")

    client = OpenAI(api_key=api_key)
    prompt = (
        f"Nombre: {lead.name}\nOrigen: {lead.source}\nMensaje: {lead.message}\n\n"
        "Responde SOLO en JSON válido con dos claves:\n"
        '- "score": "ALTO", "MEDIO" o "BAJO"\n'
        '- "instant_reply": mensaje corto y amable de respuesta inmediata.'
    )
    resp = client.chat.completions.create(
        model=os.getenv("OPENAI_MODEL", "gpt-4o-mini"),
        messages=[{"role": "user", "content": prompt}],
        response_format={"type": "json_object"},
        temperature=0.3,
    )
    content = resp.choices[0].message.content
    if not content:
        raise HTTPException(status_code=502, detail="Respuesta vacía de OpenAI")
    data = json.loads(content)
    score = str(data.get("score", "MEDIO")).upper()
    if score not in {"ALTO", "MEDIO", "BAJO"}:
        score = "MEDIO"
    return {"score": score, "instant_reply": str(data.get("instant_reply", ""))}


def save_lead_to_excel(lead: LeadPayload, score: str) -> str:
    """Guarda el lead calificado en un archivo Excel (.xlsx)."""
    if EXCEL_PATH.exists():
        wb = load_workbook(EXCEL_PATH)
        ws = wb.active
    else:
        wb = Workbook()
        ws = wb.active
        ws.title = "Leads"
        ws.append(["Fecha", "Nombre", "Email", "Teléfono", "Origen", "Score", "Mensaje"])

    ws.append([
        datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
        safe_excel_text(lead.name),
        safe_excel_text(lead.email),
        safe_excel_text(lead.phone),
        safe_excel_text(lead.source),
        score,
        safe_excel_text(lead.message),
    ])
    wb.save(EXCEL_PATH)
    return str(EXCEL_PATH)


def send_instant_email(to_email: str, subject: str, body: str) -> str:
    """Envía correo inmediato al lead vía SMTP."""
    smtp_user = os.getenv("SMTP_USER")
    smtp_pass = os.getenv("SMTP_PASS")
    if not smtp_user or not smtp_pass:
        raise HTTPException(status_code=500, detail="SMTP_USER y SMTP_PASS son obligatorios")

    msg = EmailMessage()
    msg["Subject"] = subject
    msg["From"] = smtp_user
    msg["To"] = to_email
    msg.set_content(body)

    with smtplib.SMTP(os.getenv("SMTP_HOST", "smtp.gmail.com"), int(os.getenv("SMTP_PORT", "587")), timeout=15) as server:
        server.starttls(context=ssl.create_default_context())
        server.login(smtp_user, smtp_pass)
        server.send_message(msg)
    return "sent"


@app.get("/health")
def health():
    return {"status": "ok", "excel_path": str(EXCEL_PATH)}


@app.post("/webhook/lead")
def process_lead(lead: LeadPayload):
    """Recibe un lead, lo califica, responde al instante y lo guarda en Excel."""
    qualification = qualify_lead(lead)
    excel_file = save_lead_to_excel(lead, qualification["score"])
    email_status = send_instant_email(
        to_email=lead.email,
        subject=f"Recibimos tu solicitud, {lead.name}",
        body=qualification["instant_reply"],
    )
    return {
        "status": "processed",
        "score": qualification["score"],
        "instant_reply": qualification["instant_reply"],
        "email_status": email_status,
        "saved_to": excel_file,
    }
