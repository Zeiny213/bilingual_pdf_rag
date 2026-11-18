# Bilingual (Arabic & English) PDF Contract & Invoice RAG Analyzer

Streamlit app for cited Q&A and Excel export over Arabic/English contracts and invoices. Pure Python: PyMuPDF, ChromaDB, sentence-transformers, OpenAI SDK (OpenRouter), pandas, openpyxl.

## Features
- Page-level PDF parsing with `filename` and 1-based `page_number` metadata.
- Sliding-window chunking (900 chars, 200 overlap).
- Multilingual retrieval (`paraphrase-multilingual-MiniLM-L12-v2`, in-memory ChromaDB, top 4).
- **Q&A tab:** answers in the question's language, grounded in retrieved text, with source snippets.
- **Excel tab:** extracts 7 fields per document to a table and a formatted `.xlsx`.
- Built-in sample: an Arabic contract (2 pages) and an English invoice (1 page).

## Setup
```powershell
python -m venv .venv
.\.venv\Scripts\Activate.ps1
pip install torch --index-url https://download.pytorch.org/whl/cpu
pip install -r requirements.txt
streamlit run app.py
```

Enter your OpenRouter API key in the sidebar (or set `OPENROUTER_API_KEY`).

## Tests
```powershell
pip install pytest
pytest tests
```
