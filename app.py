import hmac
import json
import logging
import os
import sqlite3
import smtplib
import ssl
import threading
from datetime import UTC, datetime
from email.message import EmailMessage
from pathlib import Path
from typing import Optional

from dotenv import load_dotenv
from fastapi import FastAPI, Header, HTTPException
from openai import OpenAI
from openpyxl import Workbook, load_workbook
from pydantic import BaseModel, EmailStr, Field

load_dotenv()

logger = logging.getLogger("lead_instant_response_bot")
logging.basicConfig(level=logging.INFO)

app = FastAPI(
    title="Lead Instant Response Bot",
    description="Recibe leads web, verifica autenticidad, evita duplicados en SQLite, califica con IA, responde por SMTP y exporta seguro a Excel.",
    version="1.0.0",
)

EXCEL_LOCK = threading.Lock()


def get_settings() -> dict:
    return {
        "openai_api_key": os.getenv("OPENAI_API_KEY", ""),
        "openai_model": os.getenv("OPENAI_MODEL", "gpt-4o-mini"),
        "webhook_secret": os.getenv("WEBHOOK_SECRET", ""),
        "smtp_host": os.getenv("SMTP_HOST", "smtp.gmail.com"),
        "smtp_port": int(os.getenv("SMTP_PORT", "587")),
        "smtp_user": os.getenv("SMTP_USER", ""),
        "smtp_pass": os.getenv("SMTP_PASS", ""),
        "database_path": os.getenv("DATABASE_PATH", "leads.db"),
        "excel_path": Path(os.getenv("LEADS_EXCEL_PATH", "leads_calificados.xlsx")),
    }


class LeadPayload(BaseModel):
    name: str = Field(min_length=1, max_length=120)
    email: EmailStr
    phone: str = Field(min_length=5, max_length=40)
    source: str = Field(default="web_form", max_length=80)
    message: str = Field(min_length=1, max_length=2000)


def init_db(db_path: str) -> None:
    with sqlite3.connect(db_path) as conn:
        conn.execute(
            """
            CREATE TABLE IF NOT EXISTS leads (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                email TEXT NOT NULL,
                source TEXT NOT NULL,
                name TEXT NOT NULL,
                phone TEXT NOT NULL,
                message TEXT NOT NULL,
                score TEXT NOT NULL,
                instant_reply TEXT NOT NULL,
                email_status TEXT NOT NULL,
                created_at TEXT NOT NULL,
                UNIQUE(email, source)
            )
            """
        )
        conn.commit()


def sanitize_excel_cell(value: str) -> str:
    """Evita Excel formula injection (=, +, -, @ al inicio)."""
    text = str(value).strip()
    if text and text[0] in ("=", "+", "-", "@", "\t", "\r"):
        return "'" + text
    return text


def qualify_lead(lead: LeadPayload, settings: dict) -> dict:
    """Califica el lead con OpenAI; si la API falla, registra fallback seguro para no perder el lead."""
    if not settings["openai_api_key"]:
        raise HTTPException(status_code=500, detail="OPENAI_API_KEY is not configured")

    try:
        client = OpenAI(api_key=settings["openai_api_key"], timeout=15.0)
        prompt = (
            f"Nombre: {lead.name}\nOrigen: {lead.source}\nMensaje: {lead.message}\n\n"
            "Responde SOLO en JSON válido con dos claves:\n"
            '- "score": "ALTO", "MEDIO" o "BAJO"\n'
            '- "instant_reply": mensaje corto y amable de respuesta inmediata en el mismo idioma del cliente.'
        )
        resp = client.chat.completions.create(
            model=settings["openai_model"],
            messages=[{"role": "user", "content": prompt}],
            response_format={"type": "json_object"},
            temperature=0.3,
        )
        data = json.loads(resp.choices[0].message.content or "{}")
        score = str(data.get("score", "MEDIO")).upper()
        if score not in {"ALTO", "MEDIO", "BAJO"}:
            score = "MEDIO"
        reply = str(data.get("instant_reply", "")).strip()
        if not reply:
            raise ValueError("Empty instant_reply")
        return {"score": score, "instant_reply": reply}
    except Exception as exc:
        logger.warning("OpenAI qualification fallback triggered: %s", exc)
        return {
            "score": "PENDIENTE",
            "instant_reply": (
                f"Hola {lead.name}, recibimos tu mensaje y un asesor te contactará muy pronto."
            ),
        }


def save_lead_to_excel(lead: LeadPayload, score: str, excel_path: Path) -> str:
    """Escribe de forma segura para hilos y sanitizada en Excel (.xlsx)."""
    with EXCEL_LOCK:
        excel_path.parent.mkdir(parents=True, exist_ok=True)
        if excel_path.exists():
            wb = load_workbook(excel_path)
            ws = wb.active
        else:
            wb = Workbook()
            ws = wb.active
            ws.title = "Leads"
            ws.append(["Fecha", "Nombre", "Email", "Teléfono", "Origen", "Score", "Mensaje"])

        ws.append([
            datetime.now(UTC).strftime("%Y-%m-%d %H:%M:%S"),
            sanitize_excel_cell(lead.name),
            sanitize_excel_cell(str(lead.email)),
            sanitize_excel_cell(lead.phone),
            sanitize_excel_cell(lead.source),
            score,
            sanitize_excel_cell(lead.message),
        ])
        tmp_path = excel_path.with_suffix(".tmp.xlsx")
        wb.save(tmp_path)
        tmp_path.replace(excel_path)
    return str(excel_path)


def send_instant_email(to_email: str, subject: str, body: str, settings: dict) -> str:
    """Envía correo inmediato al lead vía SMTP; nunca crashea el webhook si SMTP falla temporalmente."""
    if not settings["smtp_user"] or not settings["smtp_pass"]:
        raise HTTPException(status_code=500, detail="SMTP_USER and SMTP_PASS are required")

    msg = EmailMessage()
    msg["Subject"] = subject.replace("\r", "").replace("\n", "")
    msg["From"] = settings["smtp_user"]
    msg["To"] = to_email
    msg.set_content(body)

    try:
        with smtplib.SMTP(settings["smtp_host"], settings["smtp_port"], timeout=15) as server:
            server.starttls(context=ssl.create_default_context())
            server.login(settings["smtp_user"], settings["smtp_pass"])
            server.send_message(msg)
        return "sent"
    except Exception as exc:
        logger.error("SMTP send failed for %s: %s", to_email, exc)
        return "smtp_failed"


@app.get("/health")
def health():
    settings = get_settings()
    return {"status": "ok", "excel_path": str(settings["excel_path"])}


@app.post("/webhook/lead")
def process_lead(
    lead: LeadPayload,
    x_webhook_secret: Optional[str] = Header(None, alias="X-Webhook-Secret"),
):
    """Recibe un lead, valida secreto, deduplica en SQLite, responde al instante y guarda en Excel."""
    settings = get_settings()
    if settings["webhook_secret"]:
        if not x_webhook_secret or not hmac.compare_digest(x_webhook_secret, settings["webhook_secret"]):
            raise HTTPException(status_code=403, detail="Invalid webhook secret")

    init_db(settings["database_path"])
    normalized_email = str(lead.email).lower()

    with sqlite3.connect(settings["database_path"]) as conn:
        existing = conn.execute(
            "SELECT score, email_status FROM leads WHERE email = ? AND source = ?",
            (normalized_email, lead.source),
        ).fetchone()
        if existing:
            return {
                "status": "duplicate_ignored",
                "score": existing[0],
                "email_status": existing[1],
                "saved_to": str(settings["excel_path"]),
            }

    qualification = qualify_lead(lead, settings)
    email_status = send_instant_email(
        to_email=normalized_email,
        subject=f"Recibimos tu solicitud, {lead.name}",
        body=qualification["instant_reply"],
        settings=settings,
    )

    with sqlite3.connect(settings["database_path"]) as conn:
        conn.execute(
            """
            INSERT INTO leads (email, source, name, phone, message, score, instant_reply, email_status, created_at)
            VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)
            """,
            (
                normalized_email,
                lead.source,
                lead.name,
                lead.phone,
                lead.message,
                qualification["score"],
                qualification["instant_reply"],
                email_status,
                datetime.now(UTC).isoformat(),
            ),
        )
        conn.commit()

    excel_file = save_lead_to_excel(lead, qualification["score"], settings["excel_path"])
    return {
        "status": "processed",
        "score": qualification["score"],
        "instant_reply": qualification["instant_reply"],
        "email_status": email_status,
        "saved_to": excel_file,
    }
