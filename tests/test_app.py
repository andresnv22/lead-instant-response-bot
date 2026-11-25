from unittest.mock import MagicMock, patch
from openpyxl import load_workbook
from fastapi.testclient import TestClient

import app as lead_app

client = TestClient(lead_app.app)


def test_health():
    res = client.get("/health")
    assert res.status_code == 200
    assert res.json()["status"] == "ok"


def test_rejects_invalid_webhook_secret(monkeypatch):
    monkeypatch.setenv("WEBHOOK_SECRET", "top-secret")
    res = client.post(
        "/webhook/lead",
        json={
            "name": "Carlos",
            "email": "carlos@empresa.com",
            "phone": "+573001112233",
            "source": "Meta Ads",
            "message": "Quiero cotizar",
        },
        headers={"X-Webhook-Secret": "wrong"},
    )
    assert res.status_code == 403


def test_process_lead_deduplicates_and_sanitizes_excel(tmp_path, monkeypatch):
    monkeypatch.setenv("OPENAI_API_KEY", "sk-test")
    monkeypatch.setenv("WEBHOOK_SECRET", "top-secret")
    monkeypatch.setenv("SMTP_USER", "bot@example.com")
    monkeypatch.setenv("SMTP_PASS", "app-pass")
    monkeypatch.setenv("DATABASE_PATH", str(tmp_path / "leads.db"))
    excel_path = tmp_path / "leads.xlsx"
    monkeypatch.setenv("LEADS_EXCEL_PATH", str(excel_path))

    mock_resp = MagicMock()
    mock_resp.choices = [
        MagicMock(message=MagicMock(content='{"score": "ALTO", "instant_reply": "Hola!"}'))
    ]
    payload = {
        "name": "=CMD|'/C calc'!A0",
        "email": "carlos@empresa.com",
        "phone": "+573001112233",
        "source": "Meta Ads",
        "message": "Quiero cotizar urgente",
    }

    with patch("app.OpenAI") as mock_openai, patch("app.smtplib.SMTP") as mock_smtp:
        mock_openai.return_value.chat.completions.create.return_value = mock_resp
        r1 = client.post("/webhook/lead", json=payload, headers={"X-Webhook-Secret": "top-secret"})
        r2 = client.post("/webhook/lead", json=payload, headers={"X-Webhook-Secret": "top-secret"})

    assert r1.status_code == 200
    assert r1.json()["status"] == "processed"
    assert r1.json()["score"] == "ALTO"
    assert r2.status_code == 200
    assert r2.json()["status"] == "duplicate_ignored"
    mock_smtp.return_value.__enter__.return_value.send_message.assert_called_once()

    wb = load_workbook(excel_path)
    ws = wb.active
    assert ws.max_row == 2
    assert ws.cell(row=2, column=2).value.startswith("'")
