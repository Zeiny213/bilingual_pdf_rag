import io
from types import SimpleNamespace
from unittest.mock import patch

import openpyxl
import pymupdf
from streamlit.testing.v1 import AppTest

import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT))
APP = str(ROOT / "app.py")


def fake_create(self, **kwargs):
    system = kwargs["messages"][0]["content"]
    assert kwargs["temperature"] == 0.0
    if "extract structured" in system:
        content = (
            '```json\n{"document_number": "X-1", "document_date": "2026-01-01", '
            '"vendor_or_second_party": "V", "client_or_first_party": "C", '
            '"total_amount": 5, "delay_or_late_penalty": null, "termination_notice": "30 days"}\n```'
        )
    else:
        content = "Answer [Al_Noor_Tech_Contract_AR.pdf, Page 1]"
    return SimpleNamespace(choices=[SimpleNamespace(message=SimpleNamespace(content=content))])


def fresh():
    at = AppTest.from_file(APP, default_timeout=180).run()
    assert not at.exception, at.exception
    return at


def load_sample(at):
    next(b for b in at.sidebar.button if "Sample" in b.label).click()
    at.run()
    assert not at.exception, at.exception
    return at


def test_initial_state():
    at = fresh()
    assert len(at.tabs) == 2


def test_sample_indexing():
    at = load_sample(fresh())
    assert len(at.session_state["indexed_docs"]) == 2
    assert at.session_state["total_chunks"] == 5
    pages = at.session_state["indexed_docs"]["Al_Noor_Tech_Contract_AR.pdf"]["pages"]
    assert [p["page_number"] for p in pages] == [1, 2]


def test_chat_and_sources():
    at = load_sample(fresh())
    at.sidebar.text_input[0].set_value("sk-test")
    with patch("openai.resources.chat.completions.Completions.create", fake_create):
        at.chat_input[0].set_value("ما هي مدة التسليم؟").run()
    assert not at.exception, at.exception
    hist = at.session_state["chat_history"]
    assert hist[-1]["role"] == "assistant" and "Page 1" in hist[-1]["content"]
    assert len(hist[-1]["sources"]) == 4
    assert any("Al_Noor" in s["filename"] for s in hist[-1]["sources"])


def test_chat_without_key_warns():
    at = load_sample(fresh())
    at.chat_input[0].set_value("hi").run()
    assert not at.exception
    assert at.warning


def test_extraction_and_excel():
    at = load_sample(fresh())
    at.sidebar.text_input[0].set_value("sk-test")
    with patch("openai.resources.chat.completions.Completions.create", fake_create):
        next(b for b in at.tabs[1].button if "Extract" in b.label).click()
        at.run()
    assert not at.exception, at.exception
    df = at.session_state["extracted_df"]
    assert len(df) == 2
    assert df.iloc[0]["delay_or_late_penalty"] == "N/A"
    assert df.iloc[0]["total_amount"] == "5"


def test_pure_functions_and_pdf():
    import app  # noqa: F401  (executes under bare mode; only used for helpers)
    from app import extract_pages, parse_json_response, sliding_window_chunk, to_excel_bytes

    chunks = sliding_window_chunk("a" * 2000)
    assert all(len(c) <= 900 for c in chunks) and len(chunks) == 3
    assert chunks[0][-200:] == chunks[1][:200]
    assert sliding_window_chunk("") == [] and sliding_window_chunk("short") == ["short"]

    rec = parse_json_response('Here:\n```JSON\n{"document_number": "A"}\n```')
    assert rec["document_number"] == "A" and rec["total_amount"] == "N/A"

    doc = pymupdf.open()
    doc.new_page().insert_text((72, 72), "Invoice INV-1 total 10")
    doc.new_page()  # blank page
    data = doc.tobytes()
    up = SimpleNamespace(name="t.pdf", getvalue=lambda: data)
    pages, errors = extract_pages([up, SimpleNamespace(name="bad.pdf", getvalue=lambda: b"junk")])
    assert pages == [{"filename": "t.pdf", "page_number": 1, "text": "Invoice INV-1 total 10"}]
    assert len(errors) == 1 and errors[0].startswith("bad.pdf")

    import pandas as pd

    wb = openpyxl.load_workbook(io.BytesIO(to_excel_bytes(pd.DataFrame([{"a": "x", "b": "y"}]))))
    assert wb.active["A2"].value == "x"
