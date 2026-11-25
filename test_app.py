from unittest.mock import MagicMock, patch
from fastapi.testclient import TestClient

import app as lead_app

client = TestClient(lead_app.app)


def test_health():
    res = client.get("/health")
    assert res.status_code == 200
    assert res.json()["status"] == "ok"


def test_process_lead(tmp_path, monkeypatch):
    monkeypatch.setenv("OPENAI_API_KEY", "sk-test")
    monkeypatch.setenv("SMTP_USER", "bot@example.com")
    monkeypatch.setenv("SMTP_PASS", "app-pass")
    lead_app.EXCEL_PATH = tmp_path / "leads.xlsx"

    mock_resp = MagicMock()
    mock_resp.choices = [
        MagicMock(message=MagicMock(content='{"score": "ALTO", "instant_reply": "Hola Carlos!"}'))
    ]

    with patch("app.OpenAI") as mock_openai, patch("app.smtplib.SMTP") as mock_smtp:
        mock_openai.return_value.chat.completions.create.return_value = mock_resp
        res = client.post(
            "/webhook/lead",
            json={
                "name": "Carlos",
                "email": "carlos@empresa.com",
                "phone": "+573001112233",
                "source": "Meta Ads",
                "message": "Quiero cotizar urgente",
            },
        )

    assert res.status_code == 200
    data = res.json()
    assert data["score"] == "ALTO"
    assert data["email_status"] == "sent"
    assert (tmp_path / "leads.xlsx").exists()
    mock_smtp.return_value.__enter__.return_value.send_message.assert_called_once()
