from fastapi.testclient import TestClient

import app as module


def test_untrusted_lead_cannot_write_excel_formula():
    assert module.safe_excel_text("=HYPERLINK(\"https://example.com\")").startswith("'=")
    assert module.safe_excel_text("  +SUM(1,2)").startswith("'  +")
    assert module.safe_excel_text("Ana") == "Ana"


def test_production_missing_smtp_fails_visibly(tmp_path, monkeypatch):
    monkeypatch.setattr(module, "DEMO_MODE", False)
    monkeypatch.setattr(module, "EXCEL_PATH", tmp_path / "leads.xlsx")
    monkeypatch.delenv("SMTP_USER", raising=False)
    monkeypatch.delenv("SMTP_PASS", raising=False)
    client = TestClient(module.app, raise_server_exceptions=False)

    response = client.post(
        "/webhook/lead",
        json={"name": "Ana", "email": "ana@example.com", "phone": "123", "message": "Quiero cotizar"},
    )

    assert response.status_code == 500
    assert module.EXCEL_PATH.exists()


def test_demo_explicitly_reports_simulated_email(tmp_path, monkeypatch):
    monkeypatch.setattr(module, "DEMO_MODE", True)
    monkeypatch.setattr(module, "EXCEL_PATH", tmp_path / "leads.xlsx")
    client = TestClient(module.app)

    response = client.post(
        "/webhook/lead",
        json={"name": "Ana", "email": "ana@example.com", "phone": "123", "message": "Quiero cotizar"},
    )

    assert response.status_code == 200
    assert response.json()["email_status"] == "simulated_in_demo_mode"
