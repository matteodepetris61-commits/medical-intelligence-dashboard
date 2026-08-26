# Medical Information & Pharmacovigilance Decision Intelligence Dashboard (v2)

Applicazione Streamlit offline e GDPR-compliant per l'analisi avanzata di ticket di Medical Information (MI), triage automatico di Farmacovigilanza (PV / PQC), normalizzazione geografica e raccomandazioni prescrittive Human-in-the-Loop.

## 🚀 Installazione Rapida

### 1. Prerequisiti
- Python 3.10, 3.11 o 3.12
- [Ollama](https://ollama.com/) (opzionale, per la Chat AI offline con `llama3`)

### 2. Installazione Librerie
```bash
pip install -r requirements.txt
```
oppure:
```bash
pip install streamlit pandas plotly requests openpyxl
```

### 3. Modello AI Locale (Ollama)
Per abilitare l'assistente conversazionale in locale senza inviare dati all'esterno:
```bash
ollama run llama3
```

### 4. Avvio dell'Applicazione
Posizionati nella cartella del progetto ed esegui:
```bash
streamlit run app.py
```
La dashboard si aprirà automaticamente nel tuo browser all'indirizzo `http://localhost:8501`.

---

## 📁 Struttura del Progetto
- `app.py`: Codice sorgente principale dell'applicazione Streamlit.
- `sample_medical_cases.csv`: Dataset di log reali pre-caricato per test immediati.
- `requirements.txt`: Elenco delle dipendenze Python.
- `README.md`: Guida all'installazione e all'uso.
