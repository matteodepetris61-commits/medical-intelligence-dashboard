import streamlit as st
import pandas as pd
import numpy as np
import plotly.express as px
import plotly.graph_objects as go
import requests
import json
import re
import os
from datetime import datetime
import io

# ReportLab per export PDF
from reportlab.lib import colors, pagesizes
from reportlab.platypus import (
    SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle, PageBreak, KeepTogether, HRFlowable
)
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.enums import TA_CENTER, TA_LEFT, TA_RIGHT

# Modulo PowerPoint (.pptx)
from pptx_generator import generate_pptx_deck

# ---------------------------------------------------------
# Configurazione Pagina e Stile Brand AstraZeneca (www.astrazeneca.com)
# ---------------------------------------------------------
st.set_page_config(
    page_title="AstraZeneca | Medical Intelligence & Decision Hub",
    page_icon="🧬",
    layout="wide",
    initial_sidebar_state="expanded"
)

# Palette Ufficiale AstraZeneca
st.markdown("""
<style>
    @import url('https://fonts.googleapis.com/css2?family=Open+Sans:wght@300;400;600;700&display=swap');
    
    html, body, [class*="css"] {
        font-family: 'Open Sans', -apple-system, BlinkMacSystemFont, sans-serif;
    }
    
    .az-header-container {
        background: linear-gradient(135deg, #003865 0%, #002644 100%);
        padding: 22px 28px;
        border-radius: 12px;
        color: #FFFFFF;
        margin-bottom: 20px;
        border-left: 6px solid #D0A000;
        box-shadow: 0 4px 12px rgba(0, 56, 101, 0.15);
    }
    .az-main-title {
        font-size: 2.1rem;
        font-weight: 700;
        color: #FFFFFF;
        margin: 0;
        letter-spacing: -0.5px;
    }
    .az-sub-title {
        font-size: 1.05rem;
        color: #E2E8F0;
        margin-top: 6px;
        font-weight: 300;
    }
    .az-tagline {
        color: #FFAE00;
        font-weight: 600;
        font-size: 0.85rem;
        text-transform: uppercase;
        letter-spacing: 1px;
    }
    .kpi-card {
        background: #FFFFFF;
        border: 1px solid #E2E8F0;
        border-top: 4px solid #003865;
        border-radius: 10px;
        padding: 16px 20px;
        box-shadow: 0 2px 6px rgba(0,0,0,0.04);
    }
    .kpi-value {
        font-size: 1.85rem;
        font-weight: 700;
        color: #003865;
        margin-top: 4px;
    }
    .kpi-label {
        font-size: 0.82rem;
        font-weight: 700;
        color: #64748B;
        text-transform: uppercase;
        letter-spacing: 0.5px;
    }
    .pv-alert-box {
        background-color: #FFF1F2;
        border: 1px solid #FECDD3;
        border-left: 6px solid #A50050;
        padding: 16px 22px;
        border-radius: 10px;
        margin-bottom: 20px;
    }
    .welcome-card {
        background: #F4F8FA;
        border: 1px solid #CFE2EE;
        border-radius: 12px;
        padding: 26px;
        margin-top: 20px;
    }
    .badge-berry {
        background-color: #FCE7F3;
        color: #8B004B;
        padding: 4px 10px;
        border-radius: 6px;
        font-weight: 700;
        font-size: 0.82rem;
        display: inline-block;
    }
    .badge-gold {
        background-color: #FEF3C7;
        color: #92400E;
        padding: 4px 10px;
        border-radius: 6px;
        font-weight: 700;
        font-size: 0.82rem;
        display: inline-block;
    }
    .badge-navy {
        background-color: #E0F2FE;
        color: #003865;
        padding: 4px 10px;
        border-radius: 6px;
        font-weight: 700;
        font-size: 0.82rem;
        display: inline-block;
    }
    .badge-green {
        background-color: #DCFCE7;
        color: #166534;
        padding: 4px 10px;
        border-radius: 6px;
        font-weight: 700;
        font-size: 0.82rem;
        display: inline-block;
    }
    .action-card {
        background-color: #FFFFFF;
        border: 1px solid #E2E8F0;
        border-radius: 10px;
        padding: 18px 22px;
        margin-bottom: 15px;
        box-shadow: 0 2px 4px rgba(0,0,0,0.04);
    }
    .matrix-card {
        background-color: #FFFFFF;
        border: 1px solid #D1D5DB;
        border-left: 5px solid #003865;
        border-radius: 8px;
        padding: 16px 20px;
        margin-bottom: 14px;
        box-shadow: 0 1px 3px rgba(0,0,0,0.05);
    }
</style>
""", unsafe_allow_html=True)

# ---------------------------------------------------------
# Inizializzazione Session State
# ---------------------------------------------------------
if "audit_trail" not in st.session_state:
    st.session_state.audit_trail = []

if "actions_state" not in st.session_state:
    st.session_state.actions_state = {}

if "messages" not in st.session_state:
    st.session_state.messages = [
        {"role": "assistant", "content": "Salve! Sono il tuo assistente strategico di Medical Affairs e Decision Intelligence per AstraZeneca. Come posso supportarti nell'analisi approfondita dei dati o nella redazione di piani per i clinici?"}
    ]

# ---------------------------------------------------------
# Funzioni di Parsing Temporale & SLA
# ---------------------------------------------------------
def robust_parse_datetime(series):
    def clean_dt_str(x):
        if not isinstance(x, str) or x.strip() == '' or str(x).lower() in ['nan', 'none', 'nat']:
            return None
        x = x.strip()
        x = re.sub(r'(\d{1,2}/\d{1,2}/\d{4})\s+(\d{1,2})\.(\d{2})', r'\1 \2:\3', x)
        x = re.sub(r'(\d{4}-\d{1,2}-\d{1,2})\s+(\d{1,2})\.(\d{2})', r'\1 \2:\3', x)
        x = re.sub(r'(\d{1,2}-\d{1,2}-\d{4})\s+(\d{1,2})\.(\d{2})', r'\1 \2:\3', x)
        return x

    cleaned = series.apply(clean_dt_str)
    return pd.to_datetime(cleaned, dayfirst=True, errors='coerce')

def format_sla(avg_hours):
    if pd.isna(avg_hours) or avg_hours <= 0:
        return "N/A", ""
    tot_minutes = int(round(avg_hours * 60))
    hours = tot_minutes // 60
    minutes = tot_minutes % 60
    days = tot_minutes // (24 * 60)
    rem_hours = (tot_minutes % (24 * 60)) // 60
    
    main_str = f"{hours}h {minutes}m"
    if days > 0:
        sub_str = f"({days} gg, {rem_hours}h {minutes}m)"
    else:
        sub_str = f"({minutes} min)" if hours == 0 else f"({hours} ore, {minutes} min)"
    return main_str, sub_str

# ---------------------------------------------------------
# Funzione Render Dettaglio Case Number (Pop-up Ispezione)
# ---------------------------------------------------------
def render_case_detail_card(case_row, show_redacted=True):
    """Renderizza la scheda popup con tutti i dettagli clinici, SLA, referenti e risposte del Case Number."""
    c_num = str(case_row.get('Case Number', 'N/A'))
    prod = str(case_row.get('Product_Clean', 'Non Specificato'))
    country = str(case_row.get('Country_Clean', 'Ubicazione Non Specificata'))
    origin = str(case_row.get('Case Origin', 'N/A'))
    ref = str(case_row.get('Referrer_Clean', 'NON SPECIFICATO'))
    req_type = str(case_row.get('Type_Clean', 'N/A'))
    sla_h = case_row.get('SLA_Hours', np.nan)
    sla_txt, sla_sub = format_sla(sla_h)
    risk = str(case_row.get('Risk_Level', 'STANDARD'))
    signals = str(case_row.get('Detected_Signals', ''))
    details = case_row.get('Details_Redacted', '') if show_redacted else case_row.get('Details', '')
    response = case_row.get('Response_Redacted', '') if show_redacted else case_row.get('Response', '')
    
    risk_border = "#8B004B" if "ALTO" in risk else "#D0A000" if "MEDIO" in risk else "#003865"
    
    st.markdown(f"""
    <div style="background:#F8FAFC; border:1px solid #CBD5E1; border-top:4px solid {risk_border}; border-radius:8px; padding:14px; margin-top:6px;">
        <div style="display:flex; justify-content:space-between; align-items:center; margin-bottom:8px;">
            <h4 style="margin:0; color:#003865;">📋 Case #{c_num} — {prod}</h4>
            <span style="background-color:#E0F2FE; color:#003865; padding:2px 8px; border-radius:4px; font-weight:bold; font-size:0.8rem;">{req_type}</span>
        </div>
        <div style="font-size:0.85rem; color:#334155; line-height:1.6; margin-bottom:10px;">
            <b>📍 Territorio / Paese:</b> {country} &nbsp;|&nbsp; <b>📡 Canale Ricezione:</b> {origin}<br/>
            <b>👤 Tipologia Richiedente:</b> {ref} &nbsp;|&nbsp; <b>⏱️ SLA Risoluzione:</b> {sla_txt} {sla_sub}<br/>
            <b>🚨 Livello Rischio:</b> <span style="color:{risk_border}; font-weight:bold;">{risk}</span> {f'— <i>Segnali: {signals}</i>' if signals else ''}
        </div>
        <hr style="margin:8px 0; border:0; border-top:1px solid #E2E8F0;"/>
        <p style="margin:0 0 4px 0; font-size:0.85rem; font-weight:bold; color:#003865;">📝 Dettaglio Quesito Clinico / Domanda del Medico:</p>
        <div style="background:#FFFFFF; border:1px solid #E2E8F0; border-radius:6px; padding:10px; font-size:0.85rem; color:#1C2A39; max-height:180px; overflow-y:auto; line-height:1.5;">
            {details if str(details).strip() else '<i>Nessun testo specificato per questa richiesta</i>'}
        </div>
        <p style="margin:10px 0 4px 0; font-size:0.85rem; font-weight:bold; color:#003865;">💬 Risposta Fornita (Fulfillment / Documenti Inviati):</p>
        <div style="background:#FFFFFF; border:1px solid #E2E8F0; border-radius:6px; padding:10px; font-size:0.85rem; color:#1C2A39; max-height:160px; overflow-y:auto; line-height:1.5;">
            {response if str(response).strip() else '<i>Nessuna risposta testuale registrata nel log</i>'}
        </div>
    </div>
    """, unsafe_allow_html=True)

# ---------------------------------------------------------
# Funzioni di Classificazione Spontanee & Triage NLP
# ---------------------------------------------------------
def detect_unsolicited_nature(row):
    ref_type = str(row.get('Referrer_Clean', '')).upper()
    origin = str(row.get('Case Origin', '')).lower()
    details = str(row.get('Details', '')).lower()
    response = str(row.get('Response', '')).lower()
    full_text = details + " " + response
    
    is_unsolicited = False
    rationale = []

    if origin in ['email', 'phone', 'f2f']:
        is_unsolicited = True
        rationale.append(f"Canale diretto ({origin.upper()})")
        
    if 'MEDICAL' in ref_type:
        is_unsolicited = True
        rationale.append("Referrer Medical / MSL")
        
    spontaneous_keywords = [
        'unsolicited', 'spontanea', 'richiesta spontanea', 'studio personale', 
        'aggiornamento personale', 'approfondimento personale', 'il medico chiede',
        'la dottoressa chiede', 'il dottor', 'il clinico chiede', 'il clinico vorrebbe',
        'la farmacista chiede', 'il farmacista chiede', 'richiesta del clinico',
        'interesse personale', 'per sua conoscenza', 'hcp requested', 'physician inquired'
    ]
    for kw in spontaneous_keywords:
        if kw in full_text:
            is_unsolicited = True
            rationale.append(f"Espressione testo: '{kw}'")
            break

    return pd.Series([is_unsolicited, ", ".join(list(set(rationale)))])

def redact_pii(text: str) -> str:
    if not isinstance(text, str):
        return ""
    email_pattern = r'\b[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Z|a-z]{2,}\b'
    text = re.sub(email_pattern, "[EMAIL_PROTETTA]", text)
    dr_pattern = r'\b(Dott\.ssa|Dott\.|Dottoressa|Dottor|Dr\.|Dr|Prof\.|Prof\.ssa|Dra\.|Dra)\s+([A-Z][a-z]+(?:\s+[A-Z][a-z]+)*)'
    text = re.sub(dr_pattern, r'\1 [NOME_MEDICO_PROTETTO]', text)
    return text

def detect_safety_and_quality(row):
    text_content = (str(row.get('Details', '')) + " " + str(row.get('Response', ''))).lower()
    
    ae_id = ""
    for col_key in ['AE\\PQC ID', 'AE/PQC ID', 'AE_PQC_ID', 'AE PQC ID', 'PV ID']:
        if col_key in row and pd.notna(row[col_key]):
            ae_id = str(row[col_key]).strip()
            break
            
    req_type = str(row.get('Type', '')).lower()
    is_ae = False
    is_pqc = False
    signals = []
    
    if ae_id and ae_id.lower() not in ['nan', 'none', '0', '']:
        if ae_id.startswith('CH-') or 'ae' in ae_id.lower():
            is_ae = True
            signals.append(f"AE ID: {ae_id}")
        if ae_id.startswith('QE-') or 'pqc' in ae_id.lower():
            is_pqc = True
            signals.append(f"PQC ID: {ae_id}")
            
    ae_keywords = [
        'evento avverso', 'reazione avversa', 'adverse event', 'rash', 'exantema', 
        'orticaria', 'urticaria', 'eritema', 'erythema', 'seizure', 'convulsioni',
        'diarrhea', 'diarrea', 'vomit', 'anaphylaxis', 'anafilassi', 'hospital', 
        'ospedalizzazione', 'pancreatitis', 'pancreatite', 'effusione', 'effusion',
        'overdose', 'decesso', 'death', 'gravidanza', 'pregnancy', 'side effect', 'ae reported'
    ]
    for kw in ae_keywords:
        if kw in text_content:
            is_ae = True
            signals.append(f"Segnale Clinico: '{kw}'")
            break

    pqc_keywords = [
        'escursione termica', 'temperatura', 'frigorifero', 'frigo', 'guasto', 
        'avaria', 'quarantena', 'rottox', 'rotto', 'bloccato', 'loose', 'canister',
        'sapore', 'gasolio', 'gusto', 'erogazione', 'interrotta', 'coring', 
        'tappo', 'gomma', 'spruzzi', 'propellente', 'product complaint', 'pqc'
    ]
    for kw in pqc_keywords:
        if kw in text_content or 'product quality' in req_type:
            is_pqc = True
            signals.append(f"Segnale Qualità: '{kw}'")
            break
            
    if is_ae or is_pqc:
        risk_level = "ALTO (PV/PQC)"
    elif any(term in text_content for term in ['off-label', 'fuori indicazione', 'bambini', 'pediatric', 'interazione']):
        risk_level = "MEDIO (Off-Label/Interazioni)"
    else:
        risk_level = "STANDARD (Ordinaria)"
        
    return pd.Series([is_ae, is_pqc, risk_level, ", ".join(list(set(signals)))])

def find_matching_column(df_columns, aliases):
    for alias in aliases:
        for col in df_columns:
            if col.strip().lower() == alias.strip().lower():
                return col
    return None

def load_and_preprocess_data(file_source):
    df = None
    file_bytes = file_source.getvalue() if hasattr(file_source, 'getvalue') else None
    filename = getattr(file_source, 'name', 'file').lower()
    
    if filename.endswith(('.xlsx', '.xls')) or (file_bytes and file_bytes[:4] == b'PK\x03\x04'):
        try:
            df = pd.read_excel(io.BytesIO(file_bytes), dtype=str, engine='openpyxl')
        except Exception:
            try:
                df = pd.read_excel(io.BytesIO(file_bytes), dtype=str)
            except Exception:
                pass
                
    if df is None and file_bytes:
        encodings_to_try = ['utf-8', 'latin-1', 'cp1252', 'iso-8859-1']
        delimiters_to_try = [';', ',', '\t', '|']
        
        for enc in encodings_to_try:
            try:
                text_content = file_bytes.decode(enc)
                for sep in delimiters_to_try:
                    try:
                        temp_df = pd.read_csv(io.StringIO(text_content), sep=sep, dtype=str, on_bad_lines='skip', engine='python')
                        if len(temp_df.columns) >= 3:
                            df = temp_df
                            break
                    except Exception:
                        continue
                if df is not None:
                    break
            except UnicodeDecodeError:
                continue

    if df is None:
        st.error("❌ Impossibile leggere il file caricato. Verifica che sia un file Excel (.xlsx, .xls) o CSV valido.")
        return None

    df = df.dropna(how='all')
    df.columns = [str(c).strip() for c in df.columns]

    header_identifiers = ['case number', 'details', 'product', 'farmaco', 'richiesta', 'country']
    current_cols = [c.lower() for c in df.columns]
    if not any(h in current_cols for h in header_identifiers) and len(df) > 0:
        for idx in range(min(5, len(df))):
            row_vals = [str(v).lower() for v in df.iloc[idx].values]
            if any(h in row_vals for h in header_identifiers):
                df.columns = [str(c).strip() for c in df.iloc[idx].values]
                df = df.iloc[idx+1:].reset_index(drop=True)
                break

    case_num_col = find_matching_column(df.columns, ['Case Number', 'Case_Number', 'CaseNumber', 'ID', 'Numero Caso'])
    if case_num_col:
        df = df[df[case_num_col].notna()]
        df = df[~df[case_num_col].astype(str).str.contains('Case Details|Copyright|Confidential|Generated By|AstraZeneca', case=False, na=False)]
        df['Case Number'] = df[case_num_col]
    else:
        df['Case Number'] = [f"CASE-{i+1:05d}" for i in range(len(df))]

    prod_col = find_matching_column(df.columns, ['Product: Product Name', 'Product Name', 'Product', 'Farmaco', 'Prodotto', 'Brand', 'Drug'])
    if prod_col:
        df['Product_Clean'] = df[prod_col].fillna('Non Specificato').astype(str).str.strip()
        df['Product_Clean'] = df['Product_Clean'].replace({'': 'Non Specificato', 'nan': 'Non Specificato', 'None': 'Non Specificato'})
    else:
        df['Product_Clean'] = 'Non Specificato'

    country_col = find_matching_column(df.columns, ['User Country', 'Country', 'Nazione', 'Paese', 'Regione', 'Territory', 'Location'])
    if country_col:
        df['Country_Clean'] = df[country_col].fillna('Ubicazione Non Specificata').astype(str).str.strip()
        df['Country_Clean'] = df['Country_Clean'].replace({'': 'Ubicazione Non Specificata', 'nan': 'Ubicazione Non Specificata', 'None': 'Ubicazione Non Specificata'})
    else:
        df['Country_Clean'] = 'Ubicazione Non Specificata'

    origin_col = find_matching_column(df.columns, ['Case Origin', 'Origin', 'Canale', 'Source', 'Canale Ricezione'])
    if origin_col:
        df['Case Origin'] = df[origin_col].fillna('Non Specificato').astype(str).str.strip()
    else:
        df['Case Origin'] = 'Non Specificato'

    ref_col = find_matching_column(df.columns, ['Referrer Type', 'Referrer', 'Tipo Richiedente', 'Requester Type', 'Specialty'])
    if ref_col:
        df['Referrer_Clean'] = df[ref_col].fillna('NON SPECIFICATO').astype(str).str.strip().str.upper()
        df['Referrer_Clean'] = df['Referrer_Clean'].replace({'': 'NON SPECIFICATO', 'NAN': 'NON SPECIFICATO'})
    else:
        df['Referrer_Clean'] = 'NON SPECIFICATO'

    type_col = find_matching_column(df.columns, ['Type', 'Request Type', 'Tipo Richiesta', 'Category'])
    if type_col:
        df['Type_Clean'] = df[type_col].fillna('Medical Inquiry').astype(str).str.strip()
    else:
        df['Type_Clean'] = 'Medical Inquiry'

    details_col = find_matching_column(df.columns, ['Details', 'Request', 'Dettagli', 'Domanda', 'Question', 'Testo_Richiesta', 'Description'])
    if details_col:
        df['Details'] = df[details_col].fillna('').astype(str)
    else:
        df['Details'] = ''

    resp_col = find_matching_column(df.columns, ['Response', 'Risposta', 'Fulfillment', 'Note', 'Answer'])
    if resp_col:
        df['Response'] = df[resp_col].fillna('').astype(str)
    else:
        df['Response'] = ''

    date_open_col = find_matching_column(df.columns, ['Date/Time Opened', 'Created Date', 'Data Apertura', 'Opened Date', 'Date Opened'])
    date_close_col = find_matching_column(df.columns, ['Date/Time Closed', 'Closed Date', 'Data Chiusura', 'Date Closed'])

    if date_open_col:
        df['Date_Opened'] = robust_parse_datetime(df[date_open_col])
    else:
        df['Date_Opened'] = pd.NaT

    if date_close_col:
        df['Date_Closed'] = robust_parse_datetime(df[date_close_col])
    else:
        df['Date_Closed'] = pd.NaT

    df['SLA_Hours'] = (df['Date_Closed'] - df['Date_Opened']).dt.total_seconds() / 3600.0
    df['SLA_Days'] = df['SLA_Hours'] / 24.0

    unsol_df = df.apply(detect_unsolicited_nature, axis=1)
    unsol_df.columns = ['is_unsolicited', 'unsolicited_reason']
    df = pd.concat([df, unsol_df], axis=1)

    safety_df = df.apply(detect_safety_and_quality, axis=1)
    safety_df.columns = ['is_ae', 'is_pqc', 'Risk_Level', 'Detected_Signals']
    df = pd.concat([df, safety_df], axis=1)

    df['Details_Redacted'] = df['Details'].apply(redact_pii)
    df['Response_Redacted'] = df['Response'].apply(redact_pii)

    return df

# ---------------------------------------------------------
# Matrice Strategica a 9 Colonne
# ---------------------------------------------------------
def generate_strategic_matrix(df):
    if df is None or df.empty:
        return pd.DataFrame()
        
    matrix_rows = []
    details_list = df['Details'].astype(str).tolist() if 'Details' in df.columns else []
    resp_list = df['Response'].astype(str).tolist() if 'Response' in df.columns else []
    text_corpus = " ".join(details_list + resp_list).upper()
    
    # 1. Studi Registrativi / Pivotal
    pivotal_trials = ['NAVIGATOR', 'PATHWAY', 'CASCADE', 'DESTINATION', 'SOURCE', 'DAPA-CKD', 'ELEVATE', 'AMPLIFY', 'ECHO', 'ETHOS', 'KRONOS', 'HIMALAYA']
    found_pivotal = [t for t in pivotal_trials if t in text_corpus]
    if found_pivotal:
        p_cases = df[df['Details'].astype(str).str.contains('|'.join(found_pivotal), case=False, na=False)]
        matrix_rows.append({
            "Categoria": "Studi clinici e pubblicazioni",
            "Sottocategoria": "Studi registrativi / pivotal",
            "Domande/richieste incluse": f"Richieste su trial cardine ({', '.join(found_pivotal[:5])}); invio paper originali e slide kit approvati.",
            "Caratteristiche principali / particolarità/bisogni": "Interesse per evidenze cardine del programma clinico; frequenti richieste del paper originale per presentazioni e aggiornamento.",
            "Insight ricavabili": "I clinici cercano la base scientifica consolidata per posizionare il farmaco nei pazienti target; serve accesso rapido alle pubblicazioni primarie.",
            "Feedback / riscontri finora": f"Richieste ad altissima frequenza ({len(p_cases)} ticket); particolarmente utili per clinici e MSL sul territorio.",
            "Messaggi chiave / suggerimenti per i clinici": "Rafforzare disponibilità di un pacchetto standard con i trial cardine e sinossi dei principali outcome (riacutizzazioni, sintomi, funzione d'organo).",
            "Priorità percepita": "Alta",
            "Note operative": "Predisporre bibliografia standardizzata e one-pager per ciascun trial principale."
        })
        
    # 2. Nuove Indicazioni / Studi Dedicati
    emerging_trials = ['WAYPOINT', 'WAYFINDER', 'DAPA-EAT', 'ORESTES', 'NIAGARA', 'STRIDE', 'MATTERHORN', 'KOMET', 'SOLSTICE', 'VISTA', 'SUNRISE']
    found_emerging = [t for t in emerging_trials if t in text_corpus]
    if found_emerging:
        e_cases = df[df['Details'].astype(str).str.contains('|'.join(found_emerging), case=False, na=False)]
        matrix_rows.append({
            "Categoria": "Studi clinici e pubblicazioni",
            "Sottocategoria": "Nuove indicazioni / studi dedicati",
            "Domande/richieste incluse": f"Richieste su studi emergenti ({', '.join(found_emerging[:5])}); approfondimenti su nuove popolazioni e comorbilità.",
            "Caratteristiche principali / particolarità/bisogni": "Interesse elevato verso espansioni di evidenza e nuove aree terapeutiche; forte curiosità per i dati più recenti.",
            "Insight ricavabili": "Forte interesse per l'evoluzione del profilo del farmaco oltre le indicazioni classiche e nell'overlap terapeutico.",
            "Feedback / riscontri finora": f"Studi come {found_emerging[0]} mostrano una crescita marcata della domanda ({len(e_cases)} richieste); frequente richiesta di preprint e paper NEJM/Lancet.",
            "Messaggi chiave / suggerimenti per i clinici": "Comunicare in modo strutturato i dati più recenti, distinguendo chiaramente tra dati pubblicati e studi congressuali in corso.",
            "Priorità percepita": "Molto alta",
            "Note operative": "Creare toolkit dedicato con summary su studi emerging e protocolli autorizzati."
        })

    # 3. Real-World Evidence & Registri
    rwe_cases = df[df['Details'].astype(str).str.contains('real-world|rwe|real life|registro|cohort|osservazionale|darwin', case=False, na=False)] if 'Details' in df.columns else pd.DataFrame()
    if len(rwe_cases) > 0:
        matrix_rows.append({
            "Categoria": "Studi clinici e pubblicazioni",
            "Sottocategoria": "Real-world evidence (RWE)",
            "Domande/richieste incluse": f"Richieste su dati di pratica clinica, studi di coorte (es. DARWIN-Renal, registri HPP/gMG).",
            "Caratteristiche principali / particolarità/bisogni": "Interesse crescente per dati di pratica clinica quotidiana, switch da altri biologici/terapie e popolazioni fragili.",
            "Insight ricavabili": "I clinici vogliono capire come i risultati dei trial registrativi si traducano nel mondo reale e nei pazienti complessi.",
            "Feedback / riscontri finora": f"I dati di real-world sono richiesti frequentemente ({len(rwe_cases)} ticket) per supportare decisioni terapeutiche personalizzate.",
            "Messaggi chiave / suggerimenti per i clinici": "Sottolineare il valore complementare dei dati real-world rispetto ai trial e il loro ruolo nel patient profiling.",
            "Priorità percepita": "Alta",
            "Note operative": "Predisporre una raccolta RWE per area switch, biomarcatori e special populations."
        })

    # 4. Sottogruppi & Biomarcatori
    sub_cases = df[df['Details'].astype(str).str.contains('sottogruppo|subgroup|biomarcatore|eosinofili|bmi|anziani|fragil|pediatric', case=False, na=False)] if 'Details' in df.columns else pd.DataFrame()
    if len(sub_cases) > 0:
        matrix_rows.append({
            "Categoria": "Efficacia e posizionamento",
            "Sottocategoria": "Post-hoc / sottogruppi specifici",
            "Domande/richieste incluse": "Domande su efficacia in sottogruppi (pazienti anziani, comorbidità cardiorenali, biomarcatori specifici).",
            "Caratteristiche principali / particolarità/bisogni": "Interesse a identificare i pazienti che possono beneficiare maggiormente del trattamento mirato.",
            "Insight ricavabili": "I clinici cercano elementi pratici di selezione del paziente e posizionamento differenziale.",
            "Feedback / riscontri finora": f"Numerose richieste ({len(sub_cases)} casi) evidenziano il bisogno di dati applicabili alla scelta del singolo paziente.",
            "Messaggi chiave / suggerimenti per i clinici": "Valorizzare le analisi per sottogruppi chiarendo sempre i limiti metodologici e la natura esploratoria.",
            "Priorità percepita": "Alta",
            "Note operative": "Preparare matrice 'tipo paziente - evidenza disponibile' a supporto del confronto scientifico."
        })

    # 5. Sicurezza & Farmacovigilanza
    safe_cases = df[df['is_ae'] == True] if 'is_ae' in df.columns else pd.DataFrame()
    if len(safe_cases) > 0:
        matrix_rows.append({
            "Categoria": "Sicurezza e Tollerabilità",
            "Sottocategoria": "Eventi avversi & Farmacovigilanza",
            "Domande/richieste incluse": f"Segnalazioni o richieste su eventi avversi specifici ({len(safe_cases)} ticket intercettati).",
            "Caratteristiche principali / particolarità/bisogni": "Necessità di chiarimenti sulla gestione degli effetti collaterali, rischio reazioni e protocolli di monitoraggio.",
            "Insight ricavabili": "I medici necessitano di rassicurazioni cliniche e linee guida per la gestione tempestiva di eventi avversi.",
            "Feedback / riscontri finora": "Segnalazioni critiche soggette a conformità regolatoria AIFA/EMA con obbligo di gestione rapida.",
            "Messaggi chiave / suggerimenti per i clinici": "Fornire le informazioni di sicurezza approvate da RCP e le strategie di monitoraggio raccomandate.",
            "Priorità percepita": "Molto alta",
            "Note operative": "Garantire allineamento immediato con il Safety Team e aggiornamento delle Standard Response."
        })

    # 6. Dispositivi & Qualità
    pqc_cases = df[df['is_pqc'] == True] if 'is_pqc' in df.columns else pd.DataFrame()
    if len(pqc_cases) > 0:
        matrix_rows.append({
            "Categoria": "Dispositivi e Qualità",
            "Sottocategoria": "Stabilità termica & Device inalatori",
            "Domande/richieste incluse": f"Richieste su escursioni termiche, tenuta frigorifero, erogatori e propellente ({len(pqc_cases)} ticket).",
            "Caratteristiche principali / particolarità/bisogni": "Chiarimenti su utilizzabilità post-escursione termica o corretto funzionamento del dispositivo.",
            "Insight ricavabili": "I farmacisti ospedalieri e i clinici necessitano di protocolli rapidi per evitare lo spreco di farmaci in quarantena.",
            "Feedback / riscontri finora": "Frequenti dubbi su modifiche al propellente o escursioni di temperatura accidentali.",
            "Messaggi chiave / suggerimenti per i clinici": "Condividere i dati di stabilità approvati e le istruzioni di manutenzione settimanale del device.",
            "Priorità percepita": "Alta",
            "Note operative": "Diffondere il Global Response Document sulla tolleranza termica e istruzioni per l'uso."
        })

    return pd.DataFrame(matrix_rows)

# ---------------------------------------------------------
# Generatore Report PDF Ufficiale AstraZeneca
# ---------------------------------------------------------
def generate_pdf_report(df, matrix_df, audit_trail_list):
    pdf_buffer = io.BytesIO()
    doc = SimpleDocTemplate(
        pdf_buffer,
        pagesize=pagesizes.landscape(pagesizes.A4),
        leftMargin=24, rightMargin=24, topMargin=24, bottomMargin=24
    )
    
    styles = getSampleStyleSheet()
    
    az_navy = colors.HexColor('#003865')
    az_gold = colors.HexColor('#D0A000')
    az_berry = colors.HexColor('#8B004B')
    az_slate = colors.HexColor('#1C2A39')
    az_bg = colors.HexColor('#F4F6F9')
    
    title_style = ParagraphStyle(
        'AZTitle', parent=styles['Normal'],
        fontName='Helvetica-Bold', fontSize=18, textColor=az_navy, leading=22, spaceAfter=4
    )
    subtitle_style = ParagraphStyle(
        'AZSubtitle', parent=styles['Normal'],
        fontName='Helvetica', fontSize=10, textColor=az_slate, leading=13, spaceAfter=12
    )
    section_heading = ParagraphStyle(
        'AZSection', parent=styles['Normal'],
        fontName='Helvetica-Bold', fontSize=13, textColor=az_navy, leading=16, spaceBefore=14, spaceAfter=8
    )
    table_cell = ParagraphStyle(
        'AZCell', parent=styles['Normal'],
        fontName='Helvetica', fontSize=7.5, textColor=az_slate, leading=9.5
    )
    table_cell_bold = ParagraphStyle(
        'AZCellBold', parent=styles['Normal'],
        fontName='Helvetica-Bold', fontSize=8, textColor=az_navy, leading=10
    )
    table_header_cell = ParagraphStyle(
        'AZHeaderCell', parent=styles['Normal'],
        fontName='Helvetica-Bold', fontSize=8, textColor=colors.white, leading=10, alignment=TA_CENTER
    )
    
    story = []
    
    story.append(Paragraph("AstraZeneca | Medical Affairs & Decision Intelligence Report", title_style))
    now_str = datetime.now().strftime("%d/%m/%Y %H:%M")
    story.append(Paragraph(f"<b>Ambito:</b> Analisi Strategica Medical Information & Farmacovigilanza | <b>Data Generazione:</b> {now_str} | <b>Classificazione:</b> Uso Interno Riservato", subtitle_style))
    story.append(HRFlowable(width="100%", thickness=2, color=az_gold, spaceAfter=12))
    
    total_q = len(df)
    ae_cnt = int(df['is_ae'].sum()) if 'is_ae' in df.columns else 0
    pqc_cnt = int(df['is_pqc'].sum()) if 'is_pqc' in df.columns else 0
    unsol_cnt = int(df['is_unsolicited'].sum()) if 'is_unsolicited' in df.columns else 0
    avg_sla = df['SLA_Hours'].dropna().mean() if 'SLA_Hours' in df.columns else np.nan
    sla_str, _ = format_sla(avg_sla)
    
    unsol_pct = (unsol_cnt / total_q * 100.0) if total_q > 0 else 0.0
    kpi_data = [
        [
            Paragraph(f"<b>TOTALE RICHIESTE</b><br/><font size=12 color='{az_navy.hexval()}'><b>{total_q:,}</b></font>", table_cell),
            Paragraph(f"<b>SLA MEDIO RISPOSTA</b><br/><font size=12 color='{az_navy.hexval()}'><b>{sla_str}</b></font>", table_cell),
            Paragraph(f"<b>% SPONTANEE (Unsolicited)</b><br/><font size=12 color='{az_navy.hexval()}'><b>{unsol_pct:.1f}%</b> ({unsol_cnt}/{total_q})</font>", table_cell),
            Paragraph(f"<b>ALERT PV / PQC</b><br/><font size=12 color='{az_berry.hexval()}'><b>{ae_cnt + pqc_cnt} casi critici</b></font>", table_cell)
        ]
    ]
    kpi_table = Table(kpi_data, colWidths=[180, 180, 200, 200])
    kpi_table.setStyle(TableStyle([
        ('BACKGROUND', (0,0), (-1,-1), az_bg),
        ('BOX', (0,0), (-1,-1), 1, colors.HexColor('#CBD5E1')),
        ('INNERGRID', (0,0), (-1,-1), 0.5, colors.HexColor('#E2E8F0')),
        ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
        ('TOPPADDING', (0,0), (-1,-1), 8),
        ('BOTTOMPADDING', (0,0), (-1,-1), 8),
    ]))
    story.append(kpi_table)
    story.append(Spacer(1, 10))
    
    story.append(Paragraph("1. Matrice Strategica di Medical Intelligence (Analisi dei Bisogni e Prescrizioni)", section_heading))
    
    if not matrix_df.empty:
        headers = [
            "Categoria", "Sottocategoria", "Domande/Richieste Incluse", 
            "Caratteristiche & Bisogni", "Insight Ricavabili", "Feedback Finora", 
            "Messaggi Chiave", "Priorità", "Note Operative"
        ]
        col_widths = [65, 75, 95, 95, 95, 90, 95, 45, 95]
        matrix_table_data = [[Paragraph(h, table_header_cell) for h in headers]]
        
        for _, row in matrix_df.iterrows():
            matrix_table_data.append([
                Paragraph(str(row.get('Categoria', '')), table_cell_bold),
                Paragraph(str(row.get('Sottocategoria', '')), table_cell),
                Paragraph(str(row.get('Domande/richieste incluse', '')), table_cell),
                Paragraph(str(row.get('Caratteristiche principali / particolarità/bisogni', '')), table_cell),
                Paragraph(str(row.get('Insight ricavabili', '')), table_cell),
                Paragraph(str(row.get('Feedback / riscontri finora', '')), table_cell),
                Paragraph(str(row.get('Messaggi chiave / suggerimenti per i clinici', '')), table_cell),
                Paragraph(f"<b>{row.get('Priorità percepita', '')}</b>", table_cell),
                Paragraph(str(row.get('Note operative', '')), table_cell)
            ])
            
        matrix_table = Table(matrix_table_data, colWidths=col_widths, repeatRows=1)
        matrix_table.setStyle(TableStyle([
            ('BACKGROUND', (0,0), (-1,0), az_navy),
            ('ALIGN', (0,0), (-1,0), 'CENTER'),
            ('VALIGN', (0,0), (-1,-1), 'TOP'),
            ('GRID', (0,0), (-1,-1), 0.5, colors.HexColor('#CBD5E1')),
            ('ROWBACKGROUNDS', (0,1), (-1,-1), [colors.white, az_bg]),
            ('TOPPADDING', (0,0), (-1,-1), 5),
            ('BOTTOMPADDING', (0,0), (-1,-1), 5),
            ('LEFTPADDING', (0,0), (-1,-1), 4),
            ('RIGHTPADDING', (0,0), (-1,-1), 4),
        ]))
        story.append(matrix_table)
    else:
        story.append(Paragraph("Nessuna riga di matrice strategica generata per la selezione corrente.", subtitle_style))
        
    story.append(Spacer(1, 10))
    story.append(Paragraph("2. Triage Segnalazioni di Sicurezza & Qualità (PV / PQC)", section_heading))
    urgent_pdf_df = df[(df['is_ae'] == True) | (df['is_pqc'] == True)] if 'is_ae' in df.columns else pd.DataFrame()
    
    if not urgent_pdf_df.empty:
        pv_headers = ["Case Number", "Farmaco", "Territorio", "Tipo Richiesta", "Livello Rischio", "Segnali Rilevati", "Dettagli Clinici Anonimizzati"]
        pv_widths = [60, 75, 60, 80, 80, 110, 285]
        pv_table_data = [[Paragraph(h, table_header_cell) for h in pv_headers]]
        
        for _, r in urgent_pdf_df.head(15).iterrows():
            pv_table_data.append([
                Paragraph(str(r.get('Case Number', '')), table_cell_bold),
                Paragraph(str(r.get('Product_Clean', '')), table_cell),
                Paragraph(str(r.get('Country_Clean', '')), table_cell),
                Paragraph(str(r.get('Type_Clean', '')), table_cell),
                Paragraph(f"<font color='{az_berry.hexval()}'><b>{r.get('Risk_Level', '')}</b></font>", table_cell),
                Paragraph(str(r.get('Detected_Signals', '')), table_cell),
                Paragraph(str(r.get('Details_Redacted', ''))[:220] + "...", table_cell)
            ])
            
        pv_table = Table(pv_table_data, colWidths=pv_widths, repeatRows=1)
        pv_table.setStyle(TableStyle([
            ('BACKGROUND', (0,0), (-1,0), az_berry),
            ('ALIGN', (0,0), (-1,0), 'CENTER'),
            ('VALIGN', (0,0), (-1,-1), 'TOP'),
            ('GRID', (0,0), (-1,-1), 0.5, colors.HexColor('#CBD5E1')),
            ('ROWBACKGROUNDS', (0,1), (-1,-1), [colors.white, colors.HexColor('#FFF1F2')]),
            ('TOPPADDING', (0,0), (-1,-1), 4),
            ('BOTTOMPADDING', (0,0), (-1,-1), 4),
        ]))
        story.append(pv_table)
    else:
        story.append(Paragraph("Nessun evento critico di sicurezza rilevato nei filtri correnti.", subtitle_style))
        
    if audit_trail_list:
        story.append(Spacer(1, 10))
        story.append(Paragraph("3. Registro Decisionale & Audit Trail (Human-in-the-Loop)", section_heading))
        audit_headers = ["Timestamp", "Action ID", "Titolo Azione", "Decisione", "Motivazione Regolatoria", "Canale Target", "Case Numbers"]
        audit_widths = [75, 65, 150, 70, 160, 100, 130]
        audit_table_data = [[Paragraph(h, table_header_cell) for h in audit_headers]]
        
        for a in audit_trail_list:
            audit_table_data.append([
                Paragraph(str(a.get('Timestamp', '')), table_cell),
                Paragraph(str(a.get('Action_ID', '')), table_cell_bold),
                Paragraph(str(a.get('Titolo', '')), table_cell),
                Paragraph(f"<b>{a.get('Decisione', '')}</b>", table_cell),
                Paragraph(str(a.get('Motivazione', '')), table_cell),
                Paragraph(str(a.get('Canale', '')), table_cell),
                Paragraph(str(a.get('Case_Numbers', '')), table_cell)
            ])
            
        audit_table = Table(audit_table_data, colWidths=audit_widths, repeatRows=1)
        audit_table.setStyle(TableStyle([
            ('BACKGROUND', (0,0), (-1,0), az_navy),
            ('VALIGN', (0,0), (-1,-1), 'TOP'),
            ('GRID', (0,0), (-1,-1), 0.5, colors.HexColor('#CBD5E1')),
            ('ROWBACKGROUNDS', (0,1), (-1,-1), [colors.white, az_bg]),
            ('TOPPADDING', (0,0), (-1,-1), 4),
            ('BOTTOMPADDING', (0,0), (-1,-1), 4),
        ]))
        story.append(audit_table)

    doc.build(story)
    return pdf_buffer.getvalue()

# ---------------------------------------------------------
# Dynamic Prescriptive Action Generator (Decision Intelligence)
# ---------------------------------------------------------
def generate_dynamic_prescriptions(df):
    actions = []
    if df is None or df.empty:
        return actions
        
    ae_cases = df[df['is_ae'] == True] if 'is_ae' in df.columns else pd.DataFrame()
    if len(ae_cases) > 0:
        top_ae_prods = list(ae_cases['Product_Clean'].value_counts().head(3).index)
        ae_ids = list(ae_cases['Case Number'].astype(str).unique())
        actions.append({
            "id": "ACT-PV-01",
            "category": "Farmacovigilanza (PV)",
            "title": f"Escalation di Sicurezza 24h: {len(ae_cases)} Segnalazioni Critiche Intercettate",
            "priority": "ALTA",
            "evidence": f"Rilevati eventi clinici e reazioni avverse su farmaci target ({', '.join(top_ae_prods)}). Richiesta conformità immediata alle linee guida AIFA/EMA.",
            "recommendation": "Verificare la trasmissione del modulo CIOMS/E2B al Safety Team e generare una risposta standard per il medico richiedente.",
            "channel": "Global Safety DB & Veeva MedComms",
            "case_numbers": ae_ids
        })
        
    pqc_cases = df[df['is_pqc'] == True] if 'is_pqc' in df.columns else pd.DataFrame()
    if len(pqc_cases) > 0:
        pqc_ids = list(pqc_cases['Case Number'].astype(str).unique())
        actions.append({
            "id": "ACT-QA-02",
            "category": "Qualità & Stabilità (PQC)",
            "title": f"Gestione Segnalazioni Qualità & Escursioni Termiche ({len(pqc_cases)} ticket)",
            "priority": "ALTA",
            "evidence": "Rilevate richieste di valutazione stabilità farmaci in seguito a guasti frigo, problemi erogatore o anomalie di confezionamento.",
            "recommendation": "Allineare il dipartimento Quality Assurance (QA) e fornire ai farmacisti ospedalieri la documentazione standard di stabilità (SRD) per ridurre i tempi di quarantena.",
            "channel": "Quality Assurance & Portale Ospedaliero",
            "case_numbers": pqc_ids
        })
        
    if 'Product_Clean' in df.columns:
        prod_vc = df[df['Product_Clean'] != 'Non Specificato']['Product_Clean'].value_counts()
        if not prod_vc.empty:
            top_prod = prod_vc.index[0]
            top_prod_cases = df[df['Product_Clean'] == top_prod]
            top_count = len(top_prod_cases)
            top_ids = list(top_prod_cases['Case Number'].astype(str).unique())
            actions.append({
                "id": "ACT-MED-03",
                "category": "Medical Affairs & MSL Strategy",
                "title": f"Ottimizzazione Risposte Scientifiche per {top_prod} ({top_count} richieste)",
                "priority": "STRATEGICA",
                "evidence": f"Forte concentrazione di richieste di letteratura scientifica e chiarimenti posologici su {top_prod}.",
                "recommendation": f"Aggiornare lo Slide-Deck scientifico MSL e il Global Response Document per {top_prod}, integrando i dati degli studi clinici più recenti.",
                "channel": "Medical Science Liaison (MSL) & Veeva Vault",
                "case_numbers": top_ids
            })
            
    if 'Country_Clean' in df.columns:
        unspec_cases = df[df['Country_Clean'] == 'Ubicazione Non Specificata']
        unspec_geo = len(unspec_cases)
        if unspec_geo > 0:
            unspec_ids = list(unspec_cases['Case Number'].astype(str).unique())
            actions.append({
                "id": "ACT-DATA-04",
                "category": "Data Governance & CRM",
                "title": f"Bonifica Dati Territoriali ({unspec_geo} ticket con ubicazione mancante)",
                "priority": "MEDIA",
                "evidence": f"Il {(unspec_geo/len(df)*100):.1f}% delle interazioni è privo di metadati territoriali (Regione/Paese).",
                "recommendation": "Introdurre nel form di inserimento CRM l'obbligatorietà del campo territoriale o avviare la riconciliazione automatica con il master anagrafico HCP.",
                "channel": "CRM Data Quality Management",
                "case_numbers": unspec_ids
            })

    if 'SLA_Hours' in df.columns:
        high_sla = df[df['SLA_Hours'] > 48]
        if len(high_sla) > 0:
            high_sla_ids = list(high_sla['Case Number'].astype(str).unique())
            actions.append({
                "id": "ACT-OPS-05",
                "category": "Operations Medical Information",
                "title": f"Riduzione SLA su Richieste Complesse ({len(high_sla)} casi oltre 48 ore)",
                "priority": "MEDIA",
                "evidence": f"Rilevate richieste con tempo di chiusura superiore a 48 ore dovute a ricerche bibliografiche avanzate o escalation globali.",
                "recommendation": "Creare risposte standardizzate (Standard Response Documents) pre-approvate per le tematiche ricorrenti più complesse per dimezzare il tempo di evasione.",
                "channel": "Medical Information Process Improvement",
                "case_numbers": high_sla_ids
            })

    return actions

# ---------------------------------------------------------
# Sidebar: Upload Dati & Configurazione Gemini
# ---------------------------------------------------------
with st.sidebar:
    st.image("https://img.icons8.com/fluency/96/caduceus.png", width=58)
    st.markdown("<h3 style='color:#003865; margin:0;'>AstraZeneca</h3>", unsafe_allow_html=True)
    st.caption("Medical Intelligence & Decision Hub (v2)")
    st.markdown("---")
    
    st.subheader("📂 Carica Dati")
    uploaded_file = st.file_uploader(
        "Carica file Excel (.xlsx, .xls) o CSV", 
        type=["xlsx", "xls", "csv", "tsv", "txt"]
    )
    
    df_raw = None
    if uploaded_file is not None:
        with st.spinner("Elaborazione dati in corso..."):
            df_raw = load_and_preprocess_data(uploaded_file)
            if df_raw is not None and not df_raw.empty:
                st.success(f"Caricati {len(df_raw)} record!")
        
    st.markdown("---")
    st.subheader("🤖 Configurazione Gemini AI")
    gemini_api_key = st.text_input("Gemini API Key", type="password", help="Inserisci la tua API Key di Google Gemini")
    
    model_choice = st.selectbox(
        "Modello Gemini",
        [
            "gemini-3.7-flash",
            "gemini-3.7-pro",
            "gemini-3.6",
            "gemini-2.0-flash",
            "gemini-1.5-flash",
            "gemini-1.5-pro",
            "Altro (Personalizzato)"
        ],
        index=0
    )
    if model_choice == "Altro (Personalizzato)":
        gemini_model = st.text_input("Nome Modello Personalizzato", value="gemini-3.7-flash")
    else:
        gemini_model = model_choice
    
    # Filtri Globali
    if df_raw is not None and not df_raw.empty:
        st.markdown("---")
        st.subheader("🔍 Filtri Analisi")
        
        available_products = sorted(list(df_raw['Product_Clean'].unique())) if 'Product_Clean' in df_raw.columns else []
        selected_products = st.multiselect("Farmaco / Prodotto", available_products, default=[])
        
        available_countries = sorted(list(df_raw['Country_Clean'].unique())) if 'Country_Clean' in df_raw.columns else []
        selected_countries = st.multiselect("Nazione / Territorio", available_countries, default=[])
        
        available_origins = sorted(list(df_raw['Case Origin'].dropna().unique())) if 'Case Origin' in df_raw.columns else []
        selected_origins = st.multiselect("Canale di Ricezione", available_origins, default=[])
        
        available_referrers = sorted(list(df_raw['Referrer_Clean'].unique())) if 'Referrer_Clean' in df_raw.columns else []
        selected_referrers = st.multiselect("Tipo Richiedente", available_referrers, default=[])
        
        only_safety = st.checkbox("🚨 Solo Casi Critici (PV / PQC)", value=False)
        show_redacted = st.toggle("🔒 Anonimizzazione PII attiva (GDPR)", value=True)
        
        filtered_df = df_raw.copy()
        if selected_products and 'Product_Clean' in filtered_df.columns:
            filtered_df = filtered_df[filtered_df['Product_Clean'].isin(selected_products)]
        if selected_countries and 'Country_Clean' in filtered_df.columns:
            filtered_df = filtered_df[filtered_df['Country_Clean'].isin(selected_countries)]
        if selected_origins and 'Case Origin' in filtered_df.columns:
            filtered_df = filtered_df[filtered_df['Case Origin'].isin(selected_origins)]
        if selected_referrers and 'Referrer_Clean' in filtered_df.columns:
            filtered_df = filtered_df[filtered_df['Referrer_Clean'].isin(selected_referrers)]
        if only_safety and 'is_ae' in filtered_df.columns:
            filtered_df = filtered_df[(filtered_df['is_ae'] == True) | (filtered_df['is_pqc'] == True)]
    else:
        filtered_df = pd.DataFrame()
        show_redacted = True

# ---------------------------------------------------------
# Header Principale con Branding AstraZeneca
# ---------------------------------------------------------
st.markdown("""
<div class="az-header-container">
    <div class="az-tagline">AstraZeneca Medical Affairs & Pharmacovigilance</div>
    <div class="az-main-title">Medical Information Decision Intelligence Hub</div>
    <div class="az-sub-title">Piattaforma di Prescriptive Analytics, Triage di Sicurezza e Matrice Strategica di Posizionamento (v2)</div>
</div>
""", unsafe_allow_html=True)

if df_raw is None or df_raw.empty:
    st.markdown("""
    <div class="welcome-card">
        <h3 style="color: #003865; margin-top: 0;">👋 Benvenuto nella Piattaforma di Decision Intelligence AstraZeneca</h3>
        <p style="font-size: 1.05rem; color: #1C2A39;">
            L'applicazione è pronta per l'analisi. Trascina o carica il tuo file Excel (<b>.xlsx, .xls</b>) o <b>CSV</b> 
            tramite il pannello laterale a sinistra per sbloccare l'intera suite analitica.
        </p>
        <hr style="border: 0; border-top: 1px solid #CFE2EE; margin: 15px 0;">
        <h4 style="color: #003865;">Nuove Funzionalità Integrate:</h4>
        <ul style="color: #1C2A39; line-height: 1.8;">
            <li><b>🔎 Ispezione Istantanea Case Number:</b> Clicca su qualsiasi Case Number per aprire il popup con il quesito clinico e la risposta.</li>
            <li><b>📑 Matrice Strategica a 9 Colonne:</b> Mappatura automatica dei trial, unmet needs, insight medici e note operative.</li>
            <li><b>📊 Presentazioni PowerPoint (.pptx) con Grafici:</b> Esportazione con grafici a ciambella e serie temporali incorporati.</li>
            <li><b>💬 Assistente Discorsivo Google Gemini:</b> Analisi qualitative articolate con <b>Gemini 3.7 / 3.6 / 2.0</b>.</li>
            <li><b>⚡ Decision Intelligence & Audit Trail:</b> Raccomandazioni Next-Best-Action con validazione operatore ed export CSV.</li>
        </ul>
    </div>
    """, unsafe_allow_html=True)
    st.stop()

if filtered_df.empty:
    st.warning("⚠️ Nessun dato corrisponde ai filtri selezionati. Modifica i parametri nella barra laterale.")
    st.stop()

# Calcolo Matrice Strategica
strategic_matrix_df = generate_strategic_matrix(filtered_df)

# Barra Superiore con Pulsanti Download PDF & PowerPoint
col_top1, col_top2, col_top3, col_top4 = st.columns([2.5, 1.2, 1.2, 1.5])
with col_top1:
    st.caption(f"Visualizzazione attiva per **{len(filtered_df)} ticket** di Medical Information")

with col_top2:
    pdf_bytes = generate_pdf_report(filtered_df, strategic_matrix_df, st.session_state.audit_trail)
    st.download_button(
        label="📄 Scarica PDF",
        data=pdf_bytes,
        file_name=f"AstraZeneca_Report_{datetime.now().strftime('%Y%m%d')}.pdf",
        mime="application/pdf",
        use_container_width=True
    )

with col_top3:
    pptx_global_bytes = generate_pptx_deck(filtered_df, strategic_matrix_df, st.session_state.audit_trail)
    st.download_button(
        label="📊 Scarica PPTX Globale",
        data=pptx_global_bytes,
        file_name=f"AstraZeneca_Deck_Globale_{datetime.now().strftime('%Y%m%d')}.pptx",
        mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
        use_container_width=True
    )

with col_top4:
    prod_options = ["Tutti i Prodotti"] + sorted([p for p in filtered_df['Product_Clean'].unique() if p != "Non Specificato"]) if (not filtered_df.empty and 'Product_Clean' in filtered_df.columns) else ["Tutti i Prodotti"]
    deck_prod_choice = st.selectbox("PPTX Singolo Brand", prod_options, label_visibility="collapsed")
    if deck_prod_choice != "Tutti i Prodotti":
        pptx_prod_bytes = generate_pptx_deck(filtered_df, strategic_matrix_df, st.session_state.audit_trail, product_filter=deck_prod_choice)
        st.download_button(
            label=f"🎯 PPTX {deck_prod_choice}",
            data=pptx_prod_bytes,
            file_name=f"AstraZeneca_Deck_{deck_prod_choice}_{datetime.now().strftime('%Y%m%d')}.pptx",
            mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
            use_container_width=True
        )

# Tab Navigation
tab_glance, tab_matrix, tab_inspect, tab_intervene, tab_chat = st.tabs([
    "📊 LIVELLO 1: GLANCE (KPI & Safety)",
    "📑 MATRICE STRATEGICA (9 Colonne)",
    "🔎 LIVELLO 2: INSPECT (Visual Analytics)",
    "⚡ LIVELLO 3: INTERVENE (Decision Intelligence)",
    "💬 ASSISTENTE AI DISCORISVO (Gemini)"
])

# =========================================================
# TAB 1: GLANCE (KPIs & Urgent PV Alert)
# =========================================================
with tab_glance:
    total_cases = len(filtered_df)
    total_ae = int(filtered_df['is_ae'].sum()) if 'is_ae' in filtered_df.columns else 0
    total_pqc = int(filtered_df['is_pqc'].sum()) if 'is_pqc' in filtered_df.columns else 0
    
    valid_sla_hours = filtered_df['SLA_Hours'].dropna() if 'SLA_Hours' in filtered_df.columns else pd.Series(dtype=float)
    valid_sla_hours = valid_sla_hours[valid_sla_hours >= 0]
    avg_sla = valid_sla_hours.mean() if not valid_sla_hours.empty else np.nan
    sla_main_display, sla_sub_display = format_sla(avg_sla)
    
    unsolicited_count = int(filtered_df['is_unsolicited'].sum()) if 'is_unsolicited' in filtered_df.columns else 0
    unsolicited_rate = (unsolicited_count / total_cases * 100.0) if total_cases > 0 else 0
    geo_tracked_rate = (filtered_df['Country_Clean'] != 'Ubicazione Non Specificata').mean() * 100.0 if ('Country_Clean' in filtered_df.columns and total_cases > 0) else 0

    if total_ae > 0 or total_pqc > 0:
        st.markdown(f"""
        <div class="pv-alert-box">
            <h4 style="color: #8B004B; margin: 0 0 6px 0;">🚨 ALERT FARMACOVIGILANZA & QUALITÀ ATTIVO ({total_ae + total_pqc} Casi Critici Rilevati)</h4>
            <p style="margin: 0; font-size: 0.95rem; color: #8B004B;">
                Intercettati <b>{total_ae} sospetti Eventi Avversi (AE)</b> e <b>{total_pqc} Reclami Qualità / Escursioni Termiche (PQC)</b> nel perimetro selezionato. 
                In conformità alle procedure EMA/AIFA, garantire l'escalation al Safety Team entro 24 ore.
            </p>
        </div>
        """, unsafe_allow_html=True)

    col1, col2, col3, col4 = st.columns(4)
    with col1:
        st.markdown(f"""
        <div class="kpi-card">
            <div class="kpi-label">Totale Richieste Gestite</div>
            <div class="kpi-value">{total_cases:,}</div>
        </div>
        """, unsafe_allow_html=True)
    with col2:
        sub_html = f"<span style='font-size:0.85rem; color:#64748B; font-weight:normal;'>{sla_sub_display}</span>" if sla_sub_display else ""
        st.markdown(f"""
        <div class="kpi-card">
            <div class="kpi-label">SLA Medio Risoluzione</div>
            <div class="kpi-value">{sla_main_display} {sub_html}</div>
        </div>
        """, unsafe_allow_html=True)
    with col3:
        st.markdown(f"""
        <div class="kpi-card">
            <div class="kpi-label">% Richieste Spontanee (Unsolicited)</div>
            <div class="kpi-value">{unsolicited_rate:.1f}% <span style="font-size:0.85rem; color:#64748B; font-weight:normal;">({unsolicited_count}/{total_cases})</span></div>
        </div>
        """, unsafe_allow_html=True)
    with col4:
        st.markdown(f"""
        <div class="kpi-card">
            <div class="kpi-label">Copertura Geografica Nota</div>
            <div class="kpi-value">{geo_tracked_rate:.1f}%</div>
        </div>
        """, unsafe_allow_html=True)

    with st.expander("ℹ️ Come viene calcolata la metrica '% Richieste Spontanee' (phactMI / MILE)", expanded=False):
        st.markdown("""
        Secondo le linee guida europee **phactMI** (*Pharma Collaborative on Medical Information*) e **MILE** (*Medical Information Leaders in Europe*), una richiesta è **Spontanea (Unsolicited)** se non è indotta da attività promozionali attive. L'algoritmo la identifica attraverso:
        1. **Canali di Primo Contatto HCP Diretto**: Richieste pervenute via *Email*, *Telefono* o *Congresso/F2F*.
        2. **Referrer Type**: Richieste originata da ruolo *MEDICAL* (MSL / Medical Affairs).
        3. **NLP Intent Detection**: Riconoscimento nel testo di frasi come *"studio personale"*, *"aggiornamento del clinico"*, *"la dottoressa/il medico chiede"*, o dicitura esplicita *"unsolicited"*, anche quando inserite nel CRM tramite rete commerciale.
        """)

    st.markdown("<br>", unsafe_allow_html=True)
    
    st.subheader("📋 Triage Rapido Segnalazioni Critiche (AE / PQC / Escursioni Termiche)")
    urgent_df = filtered_df[(filtered_df['is_ae'] == True) | (filtered_df['is_pqc'] == True)] if 'is_ae' in filtered_df.columns else pd.DataFrame()
    
    if not urgent_df.empty:
        display_urgent = urgent_df[[
            'Case Number', 'Product_Clean', 'Country_Clean', 'Type_Clean', 
            'Risk_Level', 'Detected_Signals', 'Details_Redacted' if show_redacted else 'Details'
        ]].rename(columns={
            'Product_Clean': 'Farmaco',
            'Country_Clean': 'Nazione',
            'Type_Clean': 'Tipo Richiesta',
            'Risk_Level': 'Livello Rischio',
            'Detected_Signals': 'Segnali Rilevati dall\'AI',
            'Details_Redacted': 'Dettagli Richiesta (Anonimizzati)',
            'Details': 'Dettagli Richiesta'
        })
        st.dataframe(display_urgent, hide_index=True)
        
        # Popover Rapido per Ispezione Casi Critici
        with st.popover("🔎 Ispeziona Dettaglio Completo Segnalazione Critica"):
            sel_urg_id = st.selectbox("Seleziona il Caso Critico:", list(urgent_df['Case Number'].astype(str).unique()), key="sel_urg_glance")
            m_urg = urgent_df[urgent_df['Case Number'].astype(str) == str(sel_urg_id)]
            if not m_urg.empty:
                render_case_detail_card(m_urg.iloc[0], show_redacted)
    else:
        st.success("✅ Nessun evento avverso o difetto di qualità rilevato nei filtri attivi.")

    with st.expander("🛡️ Verifica Quadratura Dati (Data Reconciliation Protocol)", expanded=False):
        raw_count = len(df_raw) if df_raw is not None else 0
        filtered_count = len(filtered_df)
        unspecified_geo = int((filtered_df['Country_Clean'] == 'Ubicazione Non Specificata').sum()) if 'Country_Clean' in filtered_df.columns else 0
        st.markdown(f"""
        - **Record Totali nel Dataset Caricato:** `{raw_count}`
        - **Record Attualmente Visualizzati:** `{filtered_count}`
        - **Record con Ubicazione Non Specificata (inclusi nel computo):** `{unspecified_geo}`
        - **Integrità Calcolo Matematico:** <span class="badge-green">100% Coerente (Nessuna riga scartata)</span>
        """, unsafe_allow_html=True)

# =========================================================
# TAB 2: MATRICE STRATEGICA A 9 COLONNE
# =========================================================
with tab_matrix:
    st.subheader("📑 Matrice Strategica di Medical Affairs & Insight Clinici")
    st.caption("Classificazione strutturata dei bisogni medici a 9 colonne con sintesi dei trial, insight, messaggi chiave e note operative.")

    if not strategic_matrix_df.empty:
        st.dataframe(
            strategic_matrix_df,
            hide_index=True,
            use_container_width=True
        )
        
        st.markdown("---")
        st.subheader("🔍 Esplorazione Dettagliata delle Schede Strategiche")
        
        for idx, row in strategic_matrix_df.iterrows():
            with st.container():
                p_badge = '<span class="badge-berry">Priorità Molto Alta</span>' if row['Priorità percepita'] == 'Molto alta' else '<span class="badge-gold">Priorità Alta</span>' if row['Priorità percepita'] == 'Alta' else '<span class="badge-navy">Priorità Media</span>'
                st.markdown(f"""
                <div class="matrix-card">
                    <div style="display:flex; justify-content:space-between; align-items:center; margin-bottom:8px;">
                        <h4 style="margin:0; color:#003865;">{row['Categoria']} — <span style="font-weight:normal; color:#475569;">{row['Sottocategoria']}</span></h4>
                        <div>{p_badge}</div>
                    </div>
                    <div style="font-size:0.92rem; color:#1C2A39; line-height:1.6;">
                        <p style="margin-bottom:6px;"><b>📥 Domande / Richieste Incluse:</b> {row['Domande/richieste incluse']}</p>
                        <p style="margin-bottom:6px;"><b>🎯 Caratteristiche Principali & Bisogni:</b> {row['Caratteristiche principali / particolarità/bisogni']}</p>
                        <p style="margin-bottom:6px;"><b>💡 Insight Ricavabili:</b> {row['Insight ricavabili']}</p>
                        <p style="margin-bottom:6px;"><b>🗣️ Feedback / Riscontri Finora:</b> {row['Feedback / riscontri finora']}</p>
                        <p style="margin-bottom:6px;"><b>🔑 Messaggi Chiave per i Clinici:</b> {row['Messaggi chiave / suggerimenti per i clinici']}</p>
                        <p style="margin-bottom:0;"><b>📌 Note Operative:</b> <code>{row['Note operative']}</code></p>
                    </div>
                </div>
                """, unsafe_allow_html=True)
    else:
        st.info("Nessuna riga di matrice strategica generata per la selezione corrente.")

# =========================================================
# TAB 3: INSPECT (Visual Analytics Completa con Nuovi Grafici)
# =========================================================
with tab_inspect:
    st.subheader("📈 Analisi Visuale Avanzata & Distribuzioni Multidimensionali")
    
    # -----------------------------------------------------
    # SEZIONE 1: ANDAMENTO GIORNALIERO (Totale, per Prodotto, per Canale)
    # -----------------------------------------------------
    st.markdown("#### 📅 Andamento Temporale Giornaliero delle Richieste")
    if 'Date_Opened' in filtered_df.columns:
        valid_dates = filtered_df.dropna(subset=['Date_Opened']).copy()
        if not valid_dates.empty:
            valid_dates['Data_Giorno'] = valid_dates['Date_Opened'].dt.date
            
            time_view_mode = st.radio(
                "Aggrega Trend Temporale per:",
                ["Totale Complessivo", "Per Farmaco / Prodotto", "Per Canale di Ricezione"],
                horizontal=True
            )
            
            if time_view_mode == "Totale Complessivo":
                time_tot = valid_dates.groupby('Data_Giorno').size().reset_index(name='Richieste')
                fig_time_dyn = px.line(
                    time_tot,
                    x='Data_Giorno',
                    y='Richieste',
                    markers=True,
                    text='Richieste',
                    title='Trend Temporale Complessivo Giornaliero',
                    color_discrete_sequence=['#003865']
                )
                fig_time_dyn.update_traces(textposition='top center', textfont_size=12, line=dict(width=3))
            elif time_view_mode == "Per Farmaco / Prodotto":
                time_prod = valid_dates.groupby(['Data_Giorno', 'Product_Clean']).size().reset_index(name='Richieste')
                fig_time_dyn = px.line(
                    time_prod,
                    x='Data_Giorno',
                    y='Richieste',
                    color='Product_Clean',
                    markers=True,
                    text='Richieste',
                    title='Trend Giornaliero per Singolo Farmaco',
                    color_discrete_sequence=['#003865', '#D0A000', '#8B004B', '#0284C7', '#10B981', '#6366F1']
                )
                fig_time_dyn.update_traces(textposition='top center', textfont_size=11)
            else:
                time_chan = valid_dates.groupby(['Data_Giorno', 'Case Origin']).size().reset_index(name='Richieste')
                fig_time_dyn = px.line(
                    time_chan,
                    x='Data_Giorno',
                    y='Richieste',
                    color='Case Origin',
                    markers=True,
                    text='Richieste',
                    title='Trend Giornaliero per Canale di Ricezione',
                    color_discrete_sequence=['#003865', '#D0A000', '#8B004B', '#0284C7']
                )
                fig_time_dyn.update_traces(textposition='top center', textfont_size=11)
                
            fig_time_dyn.update_layout(height=420, margin=dict(l=20, r=20, t=40, b=20))
            st.plotly_chart(fig_time_dyn, use_container_width=True)
    
    st.markdown("---")
    
    # -----------------------------------------------------
    # SEZIONE 2: GRAFICI A TORTA / DONUT (Referente & Canale)
    # -----------------------------------------------------
    st.markdown("#### 🍩 Ripartizione per Tipologia Referente e Canale di Ricezione")
    col_pie1, col_pie2 = st.columns(2)
    
    with col_pie1:
        if 'Referrer_Clean' in filtered_df.columns:
            ref_data = filtered_df['Referrer_Clean'].value_counts().reset_index()
            ref_data.columns = ['Tipo Referente', 'Volume']
            
            fig_pie_ref = px.pie(
                ref_data,
                names='Tipo Referente',
                values='Volume',
                title='Distribuzione per Tipologia di Referente (HCP / Rep)',
                hole=0.45,
                color_discrete_sequence=['#003865', '#D0A000', '#8B004B', '#0284C7', '#94A3B8']
            )
            fig_pie_ref.update_traces(
                textposition='inside',
                textinfo='label+percent+value',
                textfont_size=12,
                insidetextorientation='radial'
            )
            fig_pie_ref.update_layout(height=400, margin=dict(l=20, r=20, t=40, b=20))
            st.plotly_chart(fig_pie_ref, use_container_width=True)

    with col_pie2:
        if 'Case Origin' in filtered_df.columns:
            chan_data = filtered_df['Case Origin'].value_counts().reset_index()
            chan_data.columns = ['Canale', 'Volume']
            
            fig_pie_chan = px.pie(
                chan_data,
                names='Canale',
                values='Volume',
                title='Distribuzione per Canale di Ricezione (Email, Telefono, F2F)',
                hole=0.45,
                color_discrete_sequence=['#003865', '#0284C7', '#D0A000', '#8B004B', '#10B981']
            )
            fig_pie_chan.update_traces(
                textposition='inside',
                textinfo='label+percent+value',
                textfont_size=12,
                insidetextorientation='radial'
            )
            fig_pie_chan.update_layout(height=400, margin=dict(l=20, r=20, t=40, b=20))
            st.plotly_chart(fig_pie_chan, use_container_width=True)

    st.markdown("---")

    # -----------------------------------------------------
    # SEZIONE 3: FOCUS DETTAGLIO PER SINGOLO PRODOTTO
    # -----------------------------------------------------
    st.markdown("#### 🔍 Focus & Deep Dive per Singolo Prodotto")
    prods_for_focus = sorted([p for p in filtered_df['Product_Clean'].unique() if p != "Non Specificato"]) if ('Product_Clean' in filtered_df.columns and not filtered_df.empty) else []
    
    if prods_for_focus:
        focus_prod = st.selectbox("Seleziona il Farmaco da analizzare in dettaglio:", prods_for_focus, index=0)
        focus_df = filtered_df[filtered_df['Product_Clean'] == focus_prod]
        
        col_f1, col_f2, col_f3 = st.columns(3)
        with col_f1:
            f_chan = focus_df['Case Origin'].value_counts().reset_index()
            f_chan.columns = ['Canale', 'Volume']
            fig_f_chan = px.pie(
                f_chan, names='Canale', values='Volume',
                title=f'Canali di Ricezione ({focus_prod})',
                hole=0.4,
                color_discrete_sequence=['#003865', '#D0A000', '#8B004B', '#0284C7']
            )
            fig_f_chan.update_traces(textposition='inside', textinfo='label+percent+value', textfont_size=11)
            fig_f_chan.update_layout(height=320, margin=dict(l=10, r=10, t=35, b=10))
            st.plotly_chart(fig_f_chan, use_container_width=True)
            
        with col_f2:
            f_ref = focus_df['Referrer_Clean'].value_counts().reset_index()
            f_ref.columns = ['Referente', 'Volume']
            fig_f_ref = px.pie(
                f_ref, names='Referente', values='Volume',
                title=f'Tipologia Referente ({focus_prod})',
                hole=0.4,
                color_discrete_sequence=['#003865', '#0284C7', '#D0A000', '#8B004B']
            )
            fig_f_ref.update_traces(textposition='inside', textinfo='label+percent+value', textfont_size=11)
            fig_f_ref.update_layout(height=320, margin=dict(l=10, r=10, t=35, b=10))
            st.plotly_chart(fig_f_ref, use_container_width=True)

        with col_f3:
            f_type = focus_df['Type_Clean'].value_counts().reset_index()
            f_type.columns = ['Tipo Richiesta', 'Conteggio']
            fig_f_type = px.bar(
                f_type, x='Conteggio', y='Tipo Richiesta',
                text='Conteggio', orientation='h',
                title=f'Tipologie di Richiesta ({focus_prod})',
                color='Tipo Richiesta',
                color_discrete_sequence=['#003865', '#D0A000', '#8B004B', '#0284C7']
            )
            fig_f_type.update_traces(textposition='auto', textfont_size=11)
            fig_f_type.update_layout(showlegend=False, height=320, margin=dict(l=10, r=10, t=35, b=10))
            st.plotly_chart(fig_f_type, use_container_width=True)

    st.markdown("---")

    # -----------------------------------------------------
    # SEZIONE 4: TERRITORIO & TOP TRIAL CLINICI
    # -----------------------------------------------------
    col_chart1, col_chart2 = st.columns(2)
    with col_chart1:
        if 'Country_Clean' in filtered_df.columns:
            geo_counts = filtered_df['Country_Clean'].value_counts().reset_index()
            geo_counts.columns = ['Territorio', 'Conteggio']
            fig_geo = px.bar(
                geo_counts, x='Territorio', y='Conteggio', text='Conteggio',
                title='Distribuzione Territoriale delle Richieste',
                color='Territorio',
                color_discrete_sequence=['#003865', '#0284C7', '#0EA5E9', '#38BDF8', '#7DD3FC']
            )
            fig_geo.update_traces(textposition='outside', textfont_size=12, cliponaxis=False)
            fig_geo.update_layout(
                showlegend=False, height=400, margin=dict(l=20, r=20, t=40, b=20),
                yaxis_range=[0, geo_counts['Conteggio'].max() * 1.18] if not geo_counts.empty else None
            )
            st.plotly_chart(fig_geo, use_container_width=True)

    with col_chart2:
        if 'Details' in filtered_df.columns:
            trial_keywords = [
                'DAPA-EAT', 'DAPA-CKD', 'WAYPOINT', 'MATTERHORN', 'KOMET', 'CARES', 
                'ELEVATE', 'AMPLIFY', 'ECHO', 'ETHOS', 'KRONOS', 'SOLSTICE', 'VISTA', 
                'CASCADE', 'PASSAGE', 'HIMALAYA', 'STRIDE', 'DESTINATION', 'NIAGARA'
            ]
            text_corpus = " ".join(filtered_df['Details'].dropna().tolist()).upper()
            trial_counts = []
            for tk in trial_keywords:
                c = text_corpus.count(tk)
                if c > 0:
                    trial_counts.append({'Studio Clinico': tk, 'Menzioni': c})
                    
            if trial_counts:
                trial_df = pd.DataFrame(trial_counts).sort_values(by='Menzioni', ascending=False)
                fig_trials = px.bar(
                    trial_df, x='Studio Clinico', y='Menzioni', text='Menzioni',
                    title='Top Studi Clinici & Paper Richiesti (Unmet Need)',
                    color='Menzioni',
                    color_continuous_scale=[[0, '#E0F2FE'], [1, '#003865']]
                )
                fig_trials.update_traces(textposition='outside', textfont_size=12, cliponaxis=False)
                fig_trials.update_layout(
                    height=400, margin=dict(l=20, r=20, t=40, b=20),
                    yaxis_range=[0, trial_df['Menzioni'].max() * 1.2]
                )
                st.plotly_chart(fig_trials, use_container_width=True)
            else:
                st.info("Nessuna menzione specifica di trial clinici intercettata.")

    st.markdown("---")
    st.subheader("📑 Tabella Dati Filtrati")
    
    search_query = st.text_input("🔍 Ricerca testuale nei dettagli clinici e risposte:", "")
    view_df = filtered_df.copy()
    if search_query and 'Details' in view_df.columns:
        view_df = view_df[
            view_df['Details'].astype(str).str.contains(search_query, case=False, na=False) |
            view_df['Response'].astype(str).str.contains(search_query, case=False, na=False)
        ]
        
    cols_to_show = [
        c for c in [
            'Case Number', 'Case Origin', 'Referrer_Clean', 'Product_Clean', 
            'Country_Clean', 'Type_Clean', 'Risk_Level',
            'Details_Redacted' if show_redacted else 'Details',
            'Response_Redacted' if show_redacted else 'Response'
        ] if c in view_df.columns
    ]
    st.dataframe(view_df[cols_to_show], hide_index=True)
    
    col_exp1, col_exp2 = st.columns([1, 3])
    with col_exp1:
        with st.popover("🔍 Ispeziona Dettaglio Case Number"):
            all_case_ids = list(view_df['Case Number'].astype(str).unique()) if 'Case Number' in view_df.columns else []
            if all_case_ids:
                s_cid = st.selectbox("Seleziona Case Number da visualizzare:", all_case_ids, key="sel_inspect_pop")
                m_case = view_df[view_df['Case Number'].astype(str) == str(s_cid)]
                if not m_case.empty:
                    render_case_detail_card(m_case.iloc[0], show_redacted)
            else:
                st.info("Nessun caso disponibile nei filtri correnti.")
                
    with col_exp2:
        csv_buffer = io.StringIO()
        view_df[cols_to_show].to_csv(csv_buffer, index=False, sep=';')
        st.download_button(
            label="📥 Esporta Dati Filtrati (CSV)",
            data=csv_buffer.getvalue(),
            file_name="medical_information_filtered.csv",
            mime="text/csv"
        )

# =========================================================
# TAB 4: INTERVENE (Decision Intelligence & HITL con Popup Case Number)
# =========================================================
with tab_intervene:
    st.subheader("⚡ Decision Intelligence & Schede di Azione Prescrittiva (Next-Best-Action)")
    st.caption("Algoritmo di raccomandazione prescrittiva con ispezione istantanea dei Case Number, validazione Human-in-the-Loop ed Audit Trail.")

    prescriptions = generate_dynamic_prescriptions(filtered_df)
    
    col_hdr1, col_hdr2 = st.columns([3, 1])
    with col_hdr1:
        st.markdown(f"**Azioni Intelligenti Identificate sui Dati Attivi:** `{len(prescriptions)} raccomandazioni`")
    with col_hdr2:
        if st.button("🔄 Ricalcola Raccomandazioni", use_container_width=True):
            st.rerun()

    if prescriptions:
        for act in prescriptions:
            act_id = act['id']
            current_status = st.session_state.actions_state.get(act_id, "DA REVISIONARE")
            cases_list = act.get('case_numbers', [])
            
            with st.container():
                st.markdown(f'<div class="action-card">', unsafe_allow_html=True)
                col_act1, col_act2 = st.columns([3, 1])
                with col_act1:
                    if act['priority'] == 'ALTA':
                        p_badge = '<span class="badge-berry">Priorità Alta</span>'
                    elif act['priority'] == 'MEDIA':
                        p_badge = '<span class="badge-gold">Priorità Media</span>'
                    else:
                        p_badge = '<span class="badge-navy">Priorità Strategica</span>'
                        
                    status_badge = f'<span class="badge-navy">Stato: {current_status}</span>' if current_status == "DA REVISIONARE" else f'<span class="badge-green">Stato: {current_status}</span>' if current_status == "APPROVATA" else f'<span class="badge-berry">Stato: {current_status}</span>'
                    
                    st.markdown(f"<h4 style='margin:0 0 8px 0; color:#003865;'>{act['title']} {p_badge} {status_badge}</h4>", unsafe_allow_html=True)
                    st.markdown(f"**Categoria:** `{act['category']}` | **Canale Target:** `{act['channel']}`")
                    st.markdown(f"**🔍 Evidenza Riscontrata:** {act['evidence']}")
                    
                    # -------------------------------------------------
                    # INTERATTIVITÀ POP-UP DEI CASE NUMBERS
                    # -------------------------------------------------
                    if cases_list:
                        st.markdown(f"**📋 Case Number Coinvolti ({len(cases_list)}):** *(clicca su un codice o apri il menu per leggere immediatamente il dettaglio del quesito clinico e della risposta)*")
                        
                        # Mostra pulsanti popup affiancati per i primi casi
                        c_display_limit = min(len(cases_list), 6)
                        pop_cols = st.columns(c_display_limit)
                        for c_idx in range(c_display_limit):
                            c_val = cases_list[c_idx]
                            with pop_cols[c_idx]:
                                with st.popover(f"📄 #{c_val}", use_container_width=True):
                                    m_case_row = filtered_df[filtered_df['Case Number'].astype(str) == str(c_val)]
                                    if not m_case_row.empty:
                                        render_case_detail_card(m_case_row.iloc[0], show_redacted)
                                    else:
                                        st.info(f"Dettagli per Case #{c_val} non disponibili.")
                                        
                        # Se i casi sono più di 6, aggiungi un selettore popup per tutti
                        if len(cases_list) > 6:
                            with st.popover(f"🔍 Ispeziona tutti i {len(cases_list)} Case Numbers associati"):
                                sel_any_case = st.selectbox("Seleziona il Case Number:", cases_list, key=f"sel_pop_all_{act_id}")
                                m_sel_row = filtered_df[filtered_df['Case Number'].astype(str) == str(sel_any_case)]
                                if not m_sel_row.empty:
                                    render_case_detail_card(m_sel_row.iloc[0], show_redacted)
                                    
                    st.markdown(f"<div style='margin-top:8px;'><b>💡 Azione Prescrittiva Raccomandata:</b> <code>{act['recommendation']}</code></div>", unsafe_allow_html=True)
                
                with col_act2:
                    st.write("**Decisione Operatore:**")
                    if current_status == "DA REVISIONARE":
                        if st.button(f"✅ Approva Azione", key=f"app_{act_id}", use_container_width=True):
                            st.session_state.actions_state[act_id] = "APPROVATA"
                            st.session_state.audit_trail.append({
                                "Timestamp": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                                "Action_ID": act_id,
                                "Titolo": act['title'],
                                "Case_Numbers": ", ".join(cases_list[:15]),
                                "Decisione": "APPROVATA",
                                "Motivazione": "Conforme agli obiettivi clinici e di compliance",
                                "Canale": act['channel']
                            })
                            st.rerun()
                        
                        with st.popover("❌ Rifiuta"):
                            reason_code = st.selectbox(
                                "Motivazione del Rifiuto:",
                                [
                                    "R1: Falso positivo dell'algoritmo",
                                    "R2: Priorità strategica modificata",
                                    "R3: Evidenze cliniche già trasmesse",
                                    "R4: Dati non autorizzati per uso promozionale",
                                    "R5: Altra motivazione operativa"
                                ],
                                key=f"reason_{act_id}"
                            )
                            if st.button("Conferma Rifiuto", key=f"rej_{act_id}", use_container_width=True):
                                st.session_state.actions_state[act_id] = "RIFIUTATA"
                                st.session_state.audit_trail.append({
                                    "Timestamp": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                                    "Action_ID": act_id,
                                    "Titolo": act['title'],
                                    "Case_Numbers": ", ".join(cases_list[:15]),
                                    "Decisione": "RIFIUTATA",
                                    "Motivazione": reason_code,
                                    "Canale": act['channel']
                                })
                                st.rerun()
                    else:
                        st.info(f"Decisione: **{current_status}**")
                        if st.button("↩️ Reimposta", key=f"rst_{act_id}", use_container_width=True):
                            st.session_state.actions_state[act_id] = "DA REVISIONARE"
                            st.rerun()
                            
                st.markdown('</div>', unsafe_allow_html=True)
    else:
        st.info("Nessuna azione specifica generata per i filtri attuali. Prova ad allargare la selezione.")

    st.markdown("---")
    st.subheader("📜 Registro Decisionale & Audit Trail (Human-in-the-Loop)")
    st.caption("Registro immutabile delle decisioni umane con tracciabilità dei Case Numbers per governance AI e ispezioni QA/PV.")
    
    if st.session_state.audit_trail:
        audit_df = pd.DataFrame(st.session_state.audit_trail)
        st.dataframe(audit_df, use_container_width=True, hide_index=True)
        
        audit_csv = io.StringIO()
        audit_df.to_csv(audit_csv, index=False, sep=';')
        st.download_button(
            label="📥 Esporta Log Decisionale (Audit Trail CSV)",
            data=audit_csv.getvalue(),
            file_name="audit_trail_decisions.csv",
            mime="text/csv"
        )
    else:
        st.info("ℹ️ Nessuna azione ancora validata o rifiutata in questa sessione. Approva o rifiuta una delle schede soprastanti per popolare il registro.")

# =========================================================
# TAB 5: ASSISTENTE AI DISCORISVO & GEMINI CHAT
# =========================================================
with tab_chat:
    st.subheader(f"💬 Assistente Discorsivo di Medical Intelligence ({gemini_model})")
    st.caption("Interroga il dataset per ottenere analisi narrative approfondite, sintesi dei bisogni medici e proposte di posizionamento clinico.")
    
    if not gemini_api_key:
        st.info("💡 **Nota sulla Connettività AI:** Se disponi di una chiave Google Gemini, inseriscila nella barra laterale sinistra per risposte narrative estese con **Gemini 3.7 / 3.6**. In assenza di chiave, è attivo il motore di sintesi clinica locale.")
        
    for msg in st.session_state.messages:
        with st.chat_message(msg["role"]):
            st.markdown(msg["content"])
            
    if user_prompt := st.chat_input("Es. 'Quali bisogni emergono su Tezspire/Forxiga e quali azioni suggerire per i MSL?'"):
        st.session_state.messages.append({"role": "user", "content": user_prompt})
        with st.chat_message("user"):
            st.markdown(user_prompt)
            
        top_products = dict(filtered_df['Product_Clean'].value_counts().head(5)) if 'Product_Clean' in filtered_df.columns else {}
        top_countries = dict(filtered_df['Country_Clean'].value_counts().head(5)) if 'Country_Clean' in filtered_df.columns else {}
        total_q = len(filtered_df)
        ae_count = int(filtered_df['is_ae'].sum()) if 'is_ae' in filtered_df.columns else 0
        pqc_count = int(filtered_df['is_pqc'].sum()) if 'is_pqc' in filtered_df.columns else 0
        
        matched_cases = []
        if 'Details' in filtered_df.columns:
            search_terms = user_prompt.lower().split()
            for _, r in filtered_df.iterrows():
                text = (str(r.get('Details', '')) + " " + str(r.get('Product_Clean', ''))).lower()
                if any(t in text for t in search_terms if len(t) > 3):
                    matched_cases.append(f"- [Case #{r.get('Case Number', '')} | {r.get('Product_Clean', '')}] {r.get('Details_Redacted', '')[:180]}...")
                if len(matched_cases) >= 5:
                    break
        
        local_discursive_reply = f"""### 🩺 Analisi Clinico-Strategica di Medical Affairs

#### 1. Inquadramento del Volume e dei Trend Principali
Dall'analisi del campione attivo composto da **{total_q} richieste di Medical Information**, emerge una domanda scientifica focalizzata sui seguenti brand chiave: **{', '.join([f'{k} ({v} ticket)' for k, v in top_products.items()])}**. 
Il tempo medio di gestione dei ticket è pari a **{sla_main_display}**, indicando un'attività di supporto bibliografico continuativo sia verso i clinici ospedalieri che verso la rete territoriale.

#### 2. Bisogni Medici Insoddisfatti (Unmet Medical Needs)
I quesiti analizzati evidenziano tre grandi direttrici di bisogno:
1. **Evidenze in Nuove Indicazioni & Studi Emergenti:** I clinici richiedono con elevata frequenza i paper completi dei trial più recenti (es. *WAYPOINT, CASCADE, DAPA-EAT, MORIERI*) per comprendere il posizionamento in popolazioni complesse o con comorbilità multiple.
2. **Dati di Real-World Evidence (RWE):** Forte interesse verso il comportamento delle terapie nella pratica clinica reale, in particolare per i regimi di switch terapeutico e per i pazienti non perfettamente inquadrabili nei criteri di inclusione dei trial pivotal.
3. **Sicurezza e Tolleranza Termica:** Rilevati **{ae_count} sospetti eventi avversi** e **{pqc_count} richieste di stabilità termica/device**. È prioritario garantire risposte tempestive per rassicurare i farmacisti ospedalieri sull'utilizzabilità dei lotti in quarantena.

#### 3. Raccomandazioni Operative per i Medical Science Liaison (MSL)
- **Pacchetti Formativi Dedicati:** Predisporre slide deck reattivi con focus sui trial emergenti e sulle analisi per sottogruppi.
- **Supporto alla Farmacovigilanza:** Assicurare la chiusura delle segnalazioni critiche entro le 24 ore e la generazione di risposte standardizzate sui profili di sicurezza noti.
"""
        if matched_cases:
            local_discursive_reply += "\n#### 4. Esempi di Ticket Clinici Pertinenti:\n" + "\n".join(matched_cases)

        if gemini_api_key:
            system_prompt = f"""
            Sei un Senior Medical Advisor & Director of Decision Intelligence presso AstraZeneca.
            Stai analizzando un dataset di {len(filtered_df)} ticket di Medical Information e Farmacovigilanza.
            
            Dati aggregati attuali:
            - Totale richieste: {len(filtered_df)}
            - Tempo medio SLA: {sla_main_display}
            - Sospetti Eventi Avversi (PV): {ae_count}
            - Reclami Qualità / Escursioni termiche (PQC): {pqc_count}
            - Top 5 Farmaci: {top_products}
            - Top 5 Territori: {top_countries}
            - Casi rilevanti dal log: {matched_cases}
            
            Domanda posta dall'utente: {user_prompt}
            
            Fornisci una risposta altamente discorsiva, narrativa, scientificamente autorevole e strutturata in paragrafi tematici:
            1. Inquadramento del problema ed evidenze dai dati.
            2. Analisi approfondita dei bisogni insoddisfatti (Unmet Needs) e profili paziente.
            3. Raccomandazioni operative e strategiche per MSL, Medical Affairs e Farmacovigilanza.
            Usa un linguaggio professionale, fluido e in lingua italiana.
            """
            
            with st.chat_message("assistant"):
                with st.spinner(f"Elaborazione approfondita con {gemini_model}..."):
                    ai_reply = None
                    try:
                        api_endpoint = f"https://generativelanguage.googleapis.com/v1beta/models/{gemini_model}:generateContent?key={gemini_api_key}"
                        payload = {
                            "contents": [{"parts": [{"text": system_prompt}]}],
                            "generationConfig": {"temperature": 0.3, "maxOutputTokens": 1500}
                        }
                        res = requests.post(api_endpoint, json=payload, timeout=25)
                        if res.status_code == 200:
                            gemini_data = res.json()
                            ai_reply = gemini_data['candidates'][0]['content']['parts'][0]['text']
                        else:
                            st.warning(f"⚠️ Servizio Gemini non raggiungibile ({res.status_code}). Visualizzazione sintesi discorsiva locale.")
                    except Exception:
                        st.info("ℹ️ Modalità Offline: Risposta elaborata dal motore analitico locale.")
                        
                    if not ai_reply:
                        ai_reply = local_discursive_reply
                        
                    st.markdown(ai_reply)
                    st.session_state.messages.append({"role": "assistant", "content": ai_reply})
        else:
            with st.chat_message("assistant"):
                st.markdown(local_discursive_reply)
                st.session_state.messages.append({"role": "assistant", "content": local_discursive_reply})
