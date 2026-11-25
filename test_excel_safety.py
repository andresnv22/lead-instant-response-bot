from openpyxl import load_workbook

import app


def test_untrusted_lead_value_is_not_excel_formula(tmp_path, monkeypatch):
    monkeypatch.setattr(app, "EXCEL_PATH", tmp_path / "leads.xlsx")
    lead = app.LeadPayload(name="=HYPERLINK(\"https://example.com\")", email="a@example.com", phone="123", message="Hola")
    app.save_lead_to_excel(lead, "MEDIO")
    value = load_workbook(app.EXCEL_PATH).active.cell(row=2, column=2).value
    assert value.startswith("'=")
