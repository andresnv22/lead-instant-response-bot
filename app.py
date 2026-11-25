import os
import smtplib
from datetime import datetime
from email.message import EmailMessage
from pathlib import Path

from dotenv import load_dotenv
from fastapi import FastAPI
from openpyxl import Workbook, load_workbook
from pydantic import BaseModel

load_dotenv()

app = FastAPI(
    title="Lead Instant Response Bot",
    description="Recibe leads web, los califica con IA en segundos, responde al instante y guarda en Excel/CRM.",
    version="1.0.0",
)

DEMO_MODE = os.getenv("DEMO_MODE", "true").lower() == "true"
EXCEL_PATH = Path(os.getenv("LEADS_EXCEL_PATH", "leads_calificados.xlsx"))


class LeadPayload(BaseModel):
    name: str
    email: str
    phone: str
    source: str = "web_form"
    message: str


def qualify_lead(lead: LeadPayload) -> dict:
    """Califica el lead (ALTO / MEDIO / BAJO) y redacta respuesta inmediata."""
    if DEMO_MODE or not os.getenv("OPENAI_API_KEY"):
        msg = lead.message.lower()
        if any(k in msg for k in ["cotizar", "precio", "urgente", "comprar", "presupuesto"]):
            score = "ALTO"
        elif len(msg) > 20:
            score = "MEDIO"
        else:
            score = "BAJO"

        reply = (
            f"¡Hola {lead.name}! Gracias por escribirnos desde {lead.source}. "
            "Recibimos tu solicitud y un asesor te contactará en menos de 5 minutos. "
            "¿Cuál es el mejor horario para llamarte?"
        )
        return {"score": score, "instant_reply": reply}

    from openai import OpenAI

    client = OpenAI(api_key=os.getenv("OPENAI_API_KEY"))
    prompt = (
        f"Nombre: {lead.name}\nOrigen: {lead.source}\nMensaje: {lead.message}\n\n"
        "1) Clasifica la intención de compra como ALTO, MEDIO o BAJO.\n"
        "2) Redacta un mensaje corto y amable de respuesta inmediata."
    )
    resp = client.chat.completions.create(
        model=os.getenv("OPENAI_MODEL", "gpt-4o-mini"),
        messages=[{"role": "user", "content": prompt}],
        temperature=0.3,
    )
    content = resp.choices[0].message.content or ""
    score = "ALTO" if "ALTO" in content.upper() else ("MEDIO" if "MEDIO" in content.upper() else "BAJO")
    return {"score": score, "instant_reply": content}


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
        lead.name,
        lead.email,
        lead.phone,
        lead.source,
        score,
        lead.message,
    ])
    wb.save(EXCEL_PATH)
    return str(EXCEL_PATH)


def send_instant_email(to_email: str, subject: str, body: str) -> str:
    """Envía correo inmediato al lead (o simula en DEMO_MODE)."""
    if DEMO_MODE or not os.getenv("SMTP_USER") or not os.getenv("SMTP_PASS"):
        return "simulated_in_demo_mode"

    msg = EmailMessage()
    msg["Subject"] = subject
    msg["From"] = os.getenv("SMTP_USER")
    msg["To"] = to_email
    msg.set_content(body)

    with smtplib.SMTP(os.getenv("SMTP_HOST", "smtp.gmail.com"), int(os.getenv("SMTP_PORT", "587"))) as server:
        server.starttls()
        server.login(os.getenv("SMTP_USER", ""), os.getenv("SMTP_PASS", ""))
        server.send_message(msg)
    return "sent"


@app.get("/health")
def health():
    return {"status": "ok", "demo_mode": DEMO_MODE, "excel_path": str(EXCEL_PATH)}


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
