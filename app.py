import html
import io
import json
import os
import re
import uuid
from typing import Any, Dict, List, Tuple

import chromadb
import pandas as pd
import pymupdf
import streamlit as st
from chromadb.utils.embedding_functions import SentenceTransformerEmbeddingFunction
from openai import OpenAI
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

st.set_page_config(page_title="Bilingual PDF RAG Analyzer", page_icon="📑", layout="wide")

st.markdown(
    """
    <style>
    .stChatMessage p, .stChatMessage li, .stMarkdown p, .stMarkdown li {
        unicode-bidi: plaintext;
        text-align: start;
        line-height: 1.7;
    }
    .source-card {
        background: rgba(128, 128, 128, 0.08);
        border-left: 4px solid #2563eb;
        border-radius: 6px;
        padding: 10px 14px;
        margin-bottom: 8px;
        unicode-bidi: plaintext;
        text-align: start;
        font-size: 0.92rem;
        line-height: 1.65;
    }
    .source-header { font-weight: 600; color: #2563eb; margin-bottom: 6px; font-size: 0.88rem; }
    </style>
    """,
    unsafe_allow_html=True,
)

CHUNK_SIZE = 900
CHUNK_OVERLAP = 200
TOP_K = 4
MAX_EXTRACTION_CHARS = 30000
EMBEDDING_MODEL = "paraphrase-multilingual-MiniLM-L12-v2"

MODEL_OPTIONS = [
    "openrouter/free",
    "google/gemma-4-31b-it:free",
    "google/gemini-2.5-flash",
    "openai/gpt-4o-mini",
]

EXTRACTION_KEYS = [
    "document_number",
    "document_date",
    "vendor_or_second_party",
    "client_or_first_party",
    "total_amount",
    "delay_or_late_penalty",
    "termination_notice",
]


def get_sample_pages() -> List[Dict[str, Any]]:
    ar_page_1 = (
        "عقد توريد وتطوير برمجيات مؤسسية\n"
        "رقم العقد: CNT-2026-084\n"
        "تاريخ تحرير العقد: 2026-03-15\n\n"
        "تم إبرام هذا العقد بين كل من:\n"
        "1. الطرف الأول (العميل): شركة الخليج للحلول الرقمية، ومقرها الرياض، المملكة العربية السعودية.\n"
        "2. الطرف الثاني (المورد): شركة النور للتقنية والبرمجيات (Al-Noor Tech)، ومقرها دبي، الإمارات العربية المتحدة.\n\n"
        "البند الأول - نطاق العمل:\n"
        "يلتزم الطرف الثاني بتوريد وتخصيص نظام تخطيط موارد المؤسسات (ERP) شاملاً وحدات الإدارة المالية "
        "وإدارة العقود والمشتريات ولوحات تقارير ثنائية اللغة لصالح الطرف الأول.\n\n"
        "البند الثاني - قيمة العقد وشروط الدفع:\n"
        "اتفق الطرفان على أن القيمة الإجمالية لهذا العقد تبلغ 18,500 دولار أمريكي ($18,500) "
        "شاملة كافة تكاليف التراخيص والتنفيذ والتدريب.\n"
        "- الدفعة الأولى: 40% عند توقيع العقد.\n"
        "- الدفعة الثانية: 40% عند التسليم الأولي واجتياز اختبارات القبول.\n"
        "- الدفعة النهائية: 20% عند التشغيل النهائي.\n\n"
        "البند الثالث - مدة التسليم:\n"
        "يلتزم الطرف الثاني بتسليم المشروع تسليماً نهائياً خلال 60 يوماً "
        "من تاريخ توقيع العقد واستلام الدفعة الأولى."
    )
    ar_page_2 = (
        "تابع عقد توريد وتطوير برمجيات مؤسسية - رقم العقد: CNT-2026-084\n\n"
        "البند الرابع - غرامة التأخير:\n"
        "في حال تأخر الطرف الثاني عن التسليم في الموعد المحدد (60 يوماً)، تُفرض عليه "
        "غرامة تأخير قدرها 250 دولاراً أمريكياً عن كل يوم تأخير ($250/day)، "
        "على ألا تتجاوز الغرامات في مجموعها 15% من إجمالي قيمة العقد.\n\n"
        "البند الخامس - الضمان والدعم الفني:\n"
        "يضمن الطرف الثاني خلو البرمجيات من العيوب لمدة 12 شهراً من تاريخ التسليم النهائي، "
        "مع دعم فني على مدار الساعة للحالات الحرجة.\n\n"
        "البند السادس - إنهاء العقد:\n"
        "يحق لأي من الطرفين إنهاء العقد عند إخلال الطرف الآخر بالتزاماته الجوهرية، "
        "بشرط توجيه إشعار كتابي مسبق مدته 30 يوماً دون معالجة المخالفة خلال فترة الإشعار.\n\n"
        "البند السابع - القانون الواجب التطبيق:\n"
        "يخضع هذا العقد للأنظمة المعمول بها، وحُرر من نسختين أصليتين باللغة العربية."
    )
    en_page_1 = (
        "CLOUDHOST ENTERPRISE SOLUTIONS INC.\n"
        "Enterprise Cloud Infrastructure & Managed Services Invoice\n"
        "----------------------------------------------------------------------\n"
        "Invoice ID: INV-99201\n"
        "Invoice Date: 2026-04-01\n"
        "Due Date: 2026-05-01 (Net 30)\n\n"
        "Vendor: CloudHost Enterprise Solutions Inc.\n"
        "Address: 450 Mission Street, Suite 800, San Francisco, CA 94105, USA\n\n"
        "Client (Billed To): Meridian Global Logistics Ltd.\n"
        "Address: Tower B, Financial District, King Fahd Road, Riyadh, KSA\n"
        "----------------------------------------------------------------------\n"
        "ITEMIZED SERVICES (April 1, 2026 - April 30, 2026):\n"
        "1. Dedicated GPU Inference Cluster (2x NVIDIA L40S Nodes) ...... $2,400.00\n"
        "2. Managed PostgreSQL & Vector Storage ......................... $1,100.00\n"
        "3. Enterprise WAF, DDoS Protection & 24/7 Gold SLA Support ..... $700.00\n"
        "----------------------------------------------------------------------\n"
        "Subtotal: $4,200\n"
        "VAT (15%): $630\n"
        "Total Amount Due: $4,830\n"
        "----------------------------------------------------------------------\n"
        "PAYMENT TERMS, LATE FEE & TERMINATION:\n"
        "- Payment Terms: Full payment of $4,830 is due within 30 calendar days of the invoice date.\n"
        "- Late Fee: Overdue balances accrue a late fee of 2.5%/month until settled in full.\n"
        "- Termination Notice: Recurring subscriptions require at least 15 days written notice "
        "before the next monthly billing cycle; otherwise the subscription renews automatically."
    )
    return [
        {"filename": "Al_Noor_Tech_Contract_AR.pdf", "page_number": 1, "text": ar_page_1},
        {"filename": "Al_Noor_Tech_Contract_AR.pdf", "page_number": 2, "text": ar_page_2},
        {"filename": "CloudHost_Invoice_EN.pdf", "page_number": 1, "text": en_page_1},
    ]


def extract_pages(uploaded_files: List[Any]) -> Tuple[List[Dict[str, Any]], List[str]]:
    """Extract text page-by-page. Returns (pages, error messages)."""
    pages: List[Dict[str, Any]] = []
    errors: List[str] = []
    for uploaded in uploaded_files:
        try:
            with pymupdf.open(stream=uploaded.getvalue(), filetype="pdf") as doc:
                found = False
                for idx, page in enumerate(doc):
                    text = "\n".join(l.strip() for l in page.get_text("text").splitlines()).strip()
                    text = re.sub(r"\n{3,}", "\n\n", text)
                    if text:
                        found = True
                        pages.append({"filename": uploaded.name, "page_number": idx + 1, "text": text})
                if not found:
                    errors.append(f"{uploaded.name}: no extractable text (scanned PDF?)")
        except Exception as exc:
            errors.append(f"{uploaded.name}: {exc}")
    return pages, errors


def sliding_window_chunk(text: str, chunk_size: int = CHUNK_SIZE, overlap: int = CHUNK_OVERLAP) -> List[str]:
    text = text.strip()
    if not text:
        return []
    if len(text) <= chunk_size:
        return [text]
    step = chunk_size - overlap
    chunks = []
    for start in range(0, len(text), step):
        chunk = text[start : start + chunk_size].strip()
        if chunk:
            chunks.append(chunk)
        if start + chunk_size >= len(text):
            break
    return chunks


@st.cache_resource(show_spinner="Loading embedding model...")
def get_chroma_resources() -> Tuple[Any, SentenceTransformerEmbeddingFunction]:
    embedding_fn = SentenceTransformerEmbeddingFunction(model_name=EMBEDDING_MODEL)
    return chromadb.EphemeralClient(), embedding_fn


def index_pages(pages: List[Dict[str, Any]]) -> None:
    """Replace the session's collection with a fresh index of the given pages."""
    client, embedding_fn = get_chroma_resources()

    old_name = st.session_state.get("collection_name")
    if old_name:
        try:
            client.delete_collection(name=old_name)
        except Exception:
            pass

    # Per-session collection name so concurrent sessions never overwrite each other.
    name = f"docs_{uuid.uuid4().hex}"
    collection = client.create_collection(
        name=name, embedding_function=embedding_fn, metadata={"hnsw:space": "cosine"}
    )

    ids, texts, metas = [], [], []
    docs: Dict[str, Dict[str, Any]] = {}
    for page in pages:
        fname, pnum = page["filename"], page["page_number"]
        doc = docs.setdefault(fname, {"pages": [], "chunk_count": 0})
        doc["pages"].append({"page_number": pnum, "text": page["text"]})
        for i, chunk in enumerate(sliding_window_chunk(page["text"]), start=1):
            ids.append(f"{len(ids)}")
            texts.append(chunk)
            metas.append({"filename": fname, "page_number": pnum, "chunk_index": i})
            doc["chunk_count"] += 1

    if ids:
        collection.add(ids=ids, documents=texts, metadatas=metas)

    for doc in docs.values():
        doc["pages"].sort(key=lambda p: p["page_number"])
        doc["full_text"] = "\n\n".join(f"[Page {p['page_number']}]\n{p['text']}" for p in doc["pages"])

    st.session_state.update(
        collection_name=name,
        indexed_docs=docs,
        total_chunks=len(ids),
        chat_history=[],
        extracted_df=None,
    )


def retrieve(query: str, k: int = TOP_K) -> List[Dict[str, Any]]:
    total = st.session_state["total_chunks"]
    if total == 0:
        return []
    client, embedding_fn = get_chroma_resources()
    collection = client.get_collection(
        name=st.session_state["collection_name"], embedding_function=embedding_fn
    )
    res = collection.query(
        query_texts=[query],
        n_results=min(k, total),
        include=["documents", "metadatas", "distances"],
    )
    return [
        {
            "filename": m["filename"],
            "page_number": m["page_number"],
            "chunk_index": m["chunk_index"],
            "distance": round(float(d), 4),
            "text": t,
        }
        for t, m, d in zip(res["documents"][0], res["metadatas"][0], res["distances"][0])
    ]


def get_client(api_key: str) -> OpenAI:
    return OpenAI(base_url="https://openrouter.ai/api/v1", api_key=api_key.strip())


def generate_answer(api_key: str, model: str, question: str, chunks: List[Dict[str, Any]]) -> str:
    context = "\n\n".join(
        f"--- Source {i} | {c['filename']} | Page {c['page_number']} ---\n{c['text']}"
        for i, c in enumerate(chunks, start=1)
    )
    system_prompt = (
        "You are a bilingual (Arabic & English) contract and invoice analyst. Rules:\n"
        "1. Reply in the SAME language as the user's question (Arabic or English).\n"
        "2. Answer ONLY from the retrieved context. Never use outside knowledge.\n"
        "3. After every claim, cite the source as [filename, Page X].\n"
        "4. If the answer is not in the context, say so explicitly and do not guess."
    )
    response = get_client(api_key).chat.completions.create(
        model=model,
        temperature=0.0,
        messages=[
            {"role": "system", "content": system_prompt},
            {"role": "user", "content": f"Context:\n{context}\n\nQuestion:\n{question}"},
        ],
    )
    return (response.choices[0].message.content or "").strip()


def parse_json_response(raw: str) -> Dict[str, str]:
    """Strip markdown fences / surrounding prose, parse JSON, normalize to the 7 keys."""
    cleaned = re.sub(r"```(?:json)?", "", raw, flags=re.IGNORECASE).strip()
    start, end = cleaned.find("{"), cleaned.rfind("}")
    if start == -1 or end <= start:
        raise ValueError("no JSON object in model response")
    parsed = json.loads(cleaned[start : end + 1])
    if not isinstance(parsed, dict):
        raise ValueError("JSON root is not an object")
    record = {}
    for key in EXTRACTION_KEYS:
        value = parsed.get(key)
        record[key] = "N/A" if value in (None, "") else str(value).strip()
    return record


def extract_record(api_key: str, model: str, filename: str, full_text: str) -> Dict[str, str]:
    system_prompt = (
        "You extract structured data from contracts and invoices (Arabic or English).\n"
        "Return ONLY a JSON object with exactly these keys: " + ", ".join(EXTRACTION_KEYS) + ".\n"
        "Write values in English where possible, keep amounts with currency, "
        'and use "N/A" for anything not stated. No markdown, no explanations.'
    )
    response = get_client(api_key).chat.completions.create(
        model=model,
        temperature=0.0,
        messages=[
            {"role": "system", "content": system_prompt},
            {
                "role": "user",
                "content": f"Filename: {filename}\n\n{full_text[:MAX_EXTRACTION_CHARS]}",
            },
        ],
    )
    return {"filename": filename, **parse_json_response(response.choices[0].message.content or "")}


def to_excel_bytes(df: pd.DataFrame) -> bytes:
    sheet = "Documents"
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name=sheet)
        ws = writer.sheets[sheet]
        side = Side(style="thin", color="D1D5DB")
        border = Border(left=side, right=side, top=side, bottom=side)
        for col_idx, col_name in enumerate(df.columns, start=1):
            header = ws.cell(row=1, column=col_idx)
            header.fill = PatternFill(start_color="1E3A8A", end_color="1E3A8A", fill_type="solid")
            header.font = Font(bold=True, color="FFFFFF")
            header.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
            header.border = border
            width = len(str(col_name))
            for row_idx in range(2, len(df) + 2):
                cell = ws.cell(row=row_idx, column=col_idx)
                cell.alignment = Alignment(vertical="top", wrap_text=True)
                cell.border = border
                width = max(width, len(str(cell.value or "")))
            ws.column_dimensions[get_column_letter(col_idx)].width = min(max(width + 4, 18), 48)
        ws.freeze_panes = "A2"
    return output.getvalue()


def render_sources(sources: List[Dict[str, Any]], expanded: bool = False) -> None:
    with st.expander(f"Sources ({len(sources)})", expanded=expanded):
        for i, s in enumerate(sources, start=1):
            body = html.escape(s["text"]).replace("\n", "<br>")
            st.markdown(
                f"<div class='source-card'><div class='source-header'>#{i} · "
                f"{html.escape(s['filename'])} · Page {s['page_number']} · distance {s['distance']}</div>"
                f"{body}</div>",
                unsafe_allow_html=True,
            )


for key, default in {
    "indexed_docs": {},
    "total_chunks": 0,
    "chat_history": [],
    "extracted_df": None,
    "collection_name": None,
}.items():
    st.session_state.setdefault(key, default)

get_chroma_resources()

with st.sidebar:
    st.header("Settings")
    api_key = st.text_input(
        "OpenRouter API Key", value=os.environ.get("OPENROUTER_API_KEY", ""), type="password"
    )
    model = st.selectbox("Model", MODEL_OPTIONS, index=0)

    st.divider()
    uploaded_pdfs = st.file_uploader("Upload PDFs", type=["pdf"], accept_multiple_files=True)
    if st.button("Index PDFs", type="primary", use_container_width=True):
        if not uploaded_pdfs:
            st.warning("Select at least one PDF.")
        else:
            with st.spinner("Indexing..."):
                pages, errors = extract_pages(uploaded_pdfs)
                if pages:
                    index_pages(pages)
                    st.success("PDFs indexed.")
                for err in errors:
                    st.error(err)

    if st.button("Load Built-in Bilingual Sample", use_container_width=True):
        with st.spinner("Indexing..."):
            index_pages(get_sample_pages())
        st.success("Sample indexed.")

    st.divider()
    docs = st.session_state["indexed_docs"]
    c1, c2 = st.columns(2)
    c1.metric("Documents", len(docs))
    c2.metric("Chunks", st.session_state["total_chunks"])
    for fname, info in docs.items():
        st.markdown(f"- `{fname}` ({len(info['pages'])} p. / {info['chunk_count']} chunks)")

st.title("📑 Bilingual PDF Contract & Invoice Analyzer")

tab_chat, tab_excel = st.tabs(["💬 Q&A Chat", "📊 Excel Exporter"])

with tab_chat:
    if not docs:
        st.info("Load the built-in sample or upload and index PDFs from the sidebar to begin.")
    else:
        if st.session_state["chat_history"] and st.button("Clear chat"):
            st.session_state["chat_history"] = []
            st.rerun()

    for msg in st.session_state["chat_history"]:
        with st.chat_message(msg["role"]):
            st.markdown(msg["content"])
            if msg.get("sources"):
                render_sources(msg["sources"])

    question = st.chat_input("Ask in Arabic or English...")
    if question:
        if not docs:
            st.warning("Index documents first.")
        elif not api_key.strip():
            st.warning("Enter your OpenRouter API key in the sidebar.")
        else:
            st.session_state["chat_history"].append({"role": "user", "content": question})
            with st.chat_message("user"):
                st.markdown(question)
            with st.chat_message("assistant"):
                chunks: List[Dict[str, Any]] = []
                try:
                    with st.spinner("Thinking..."):
                        chunks = retrieve(question)
                        answer = generate_answer(api_key, model, question, chunks)
                    if not answer:
                        answer = "The model returned an empty response. Try again or pick another model."
                except Exception as exc:
                    answer = f"**Error:** {exc}"
                st.markdown(answer)
                if chunks:
                    render_sources(chunks, expanded=True)
            st.session_state["chat_history"].append(
                {"role": "assistant", "content": answer, "sources": chunks}
            )

with tab_excel:
    if not docs:
        st.info("Load the built-in sample or upload and index PDFs from the sidebar to begin.")
    elif st.button("Extract structured data", type="primary"):
        if not api_key.strip():
            st.warning("Enter your OpenRouter API key in the sidebar.")
        else:
            records = []
            progress = st.progress(0.0)
            for i, (fname, info) in enumerate(docs.items(), start=1):
                try:
                    records.append(extract_record(api_key, model, fname, info["full_text"]))
                except Exception as exc:
                    st.error(f"{fname}: {exc}")
                    records.append({"filename": fname, **{k: "Error" for k in EXTRACTION_KEYS}})
                progress.progress(i / len(docs))
            progress.empty()
            st.session_state["extracted_df"] = pd.DataFrame(records)

    df = st.session_state["extracted_df"]
    if df is not None:
        st.dataframe(df, use_container_width=True, hide_index=True)
        st.download_button(
            "Download Excel (.xlsx)",
            data=to_excel_bytes(df),
            file_name="contracts_invoices.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            type="primary",
        )
