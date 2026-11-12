import streamlit as st
import pandas as pd
import os
import re
from io import BytesIO
from collections import defaultdict, Counter
from reportlab.lib.pagesizes import A4
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib import colors

st.set_page_config(page_title="B.E.N.I. - Van del Benessere", page_icon="🌿", layout="wide")

FILE_EXCEL = "BENI_Simulazione_Dataset_e_Matching.xlsx"

def trova_colonna(df, parole_chiave, default_idx=0):
    for col in df.columns:
        c_low = str(col).lower()
        if any(k in c_low for k in parole_chiave):
            return col
    return df.columns[default_idx] if len(df.columns) > default_idx else None

def normalizza_utenti(df):
    col_id = trova_colonna(df, ["id_utente", "id", "codice"], 0)
    col_cap = trova_colonna(df, ["cap", "zona"], 1)
    col_tag = trova_colonna(df, ["tag", "tema"], 2)
    col_fascia = trova_colonna(df, ["fascia", "orari", "turno"], 3)
    col_bisogni = trova_colonna(df, ["bisogn", "accessib", "requisit"], 4)

    res = pd.DataFrame()
    res["ID_Utente"] = df[col_id].astype(str)
    res["Zona_CAP"] = df[col_cap].astype(str).str.replace("CAP_", "", regex=False).str.strip()
    res["Tag_Tematici"] = df[col_tag].astype(str)
    res["Fascia_Oraria"] = df[col_fascia].astype(str)
    res["Bisogni_Pratici"] = df[col_bisogni].astype(str)
    return res

def normalizza_luoghi(df):
    col_id = trova_colonna(df, ["id_luogo", "id"], 0)
    col_nome = trova_colonna(df, ["nome", "luogo", "spazio"], 1)
    col_cap = trova_colonna(df, ["cap", "zona"], 2)
    col_esito = trova_colonna(df, ["esito", "stato", "verifica", "valid"], 3)
    col_sosta = trova_colonna(df, ["sosta", "van", "parchegg"], 4)

    res = pd.DataFrame()
    res["ID_Luogo"] = df[col_id].astype(str)
    res["Nome_Luogo"] = df[col_nome].astype(str)
    res["CAP"] = df[col_cap].astype(str).str.replace("CAP_", "", regex=False).str.strip()
    res["Esito_Verifica"] = df[col_esito].astype(str)
    res["Sosta_Van"] = df[col_sosta].astype(str) if col_sosta else "Verificata"
    # Concatena tutte le colonne testuali per il matching semantico
    res["Dettagli_Completi"] = df.apply(lambda r: " | ".join([str(v) for v in r.values if pd.notna(v)]), axis=1)
    return res

def carica_dati_iniziali():
    if os.path.exists(FILE_EXCEL):
        try:
            xl = pd.ExcelFile(FILE_EXCEL)
            s_utenti = [s for s in xl.sheet_names if "Richieste" in s or "Utenti" in s]
            s_luoghi = [s for s in xl.sheet_names if "Luoghi" in s or "OpenData" in s]
            df_u_raw = pd.read_excel(FILE_EXCEL, sheet_name=s_utenti[0] if s_utenti else 0)
            df_l_raw = pd.read_excel(FILE_EXCEL, sheet_name=s_luoghi[0] if s_luoghi else 1)
            return normalizza_utenti(df_u_raw), normalizza_luoghi(df_l_raw)
        except Exception:
            pass

    utenti_fallback = pd.DataFrame([
        {"ID_Utente": "USR_001", "Zona_CAP": "06034", "Tag_Tematici": "movimento, natura", "Fascia_Oraria": "Mattina", "Bisogni_Pratici": "ombra, terreno_pianeggiante"},
        {"ID_Utente": "USR_002", "Zona_CAP": "06034", "Tag_Tematici": "arte, socialita", "Fascia_Oraria": "Mattina", "Bisogni_Pratici": "sedute_disponibili, senza_barriere"},
        {"ID_Utente": "USR_003", "Zona_CAP": "06034", "Tag_Tematici": "movimento, natura, socialita", "Fascia_Oraria": "Mattina", "Bisogni_Pratici": "panchine_frequenti, terreno_pianeggiante"},
        {"ID_Utente": "USR_004", "Zona_CAP": "06034", "Tag_Tematici": "cultura, natura, socialita", "Fascia_Oraria": "Pomeriggio", "Bisogni_Pratici": "sedute_disponibili, servizi_igienici"},
        {"ID_Utente": "USR_005", "Zona_CAP": "06034", "Tag_Tematici": "movimento, natura", "Fascia_Oraria": "Mattina", "Bisogni_Pratici": "terreno_pianeggiante"},
        {"ID_Utente": "USR_006", "Zona_CAP": "06034", "Tag_Tematici": "arte, socialita", "Fascia_Oraria": "Mattina", "Bisogni_Pratici": "tavoli_accessibili, senza_barriere"},
        {"ID_Utente": "USR_007", "Zona_CAP": "06034", "Tag_Tematici": "movimento, natura", "Fascia_Oraria": "Mattina", "Bisogni_Pratici": "terreno_pianeggiante, senza_barriere"},
        {"ID_Utente": "USR_008", "Zona_CAP": "06034", "Tag_Tematici": "musica, socialita", "Fascia_Oraria": "Pomeriggio", "Bisogni_Pratici": "sedute_disponibili, ombra"},
        {"ID_Utente": "USR_009", "Zona_CAP": "06034", "Tag_Tematici": "natura, cultura", "Fascia_Oraria": "Mattina", "Bisogni_Pratici": "senza_barriere, panchine_frequenti"},
        {"ID_Utente": "USR_010", "Zona_CAP": "06034", "Tag_Tematici": "movimento, natura", "Fascia_Oraria": "Mattina", "Bisogni_Pratici": "sedute_disponibili, ombra"},
        {"ID_Utente": "USR_011", "Zona_CAP": "06034", "Tag_Tematici": "arte, natura", "Fascia_Oraria": "Mattina", "Bisogni_Pratici": "sedute_disponibili, servizi_igienici"},
        {"ID_Utente": "USR_012", "Zona_CAP": "06034", "Tag_Tematici": "socialita, cultura", "Fascia_Oraria": "Pomeriggio", "Bisogni_Pratici": "tavoli_accessibili, ombra"}
    ])

    luoghi_fallback = pd.DataFrame([
        {"ID_Luogo": "LOC_01", "Nome_Luogo": "Parco dei Canapè (Area Centrale)", "CAP": "06034", "Esito_Verifica": "VALIDATO", "Sosta_Van": "Sì (Via Nazario Sauro)", "Dettagli_Completi": "movimento, natura, arte, socialita, cultura | terreno_pianeggiante, senza_barriere, ombra, panchine_frequenti, sedute_disponibili, servizi_igienici"},
        {"ID_Luogo": "LOC_02", "Nome_Luogo": "Parco Fluviale Hoffmann", "CAP": "06034", "Esito_Verifica": "VALIDATO", "Sosta_Van": "Sì (Parcheggio adiacente)", "Dettagli_Completi": "movimento, natura, socialita | terreno_pianeggiante, senza_barriere, ombra, panchine_frequenti"},
        {"ID_Luogo": "LOC_03", "Nome_Luogo": "Giardino di Villa Candida", "CAP": "06034", "Esito_Verifica": "VALIDATO", "Sosta_Van": "Sì (Ingresso laterale)", "Dettagli_Completi": "arte, cultura, socialita, natura | senza_barriere, ombra, sedute_disponibili, tavoli_accessibili, servizi_igienici"},
        {"ID_Luogo": "LOC_08", "Nome_Luogo": "Area Verde Via Monte Pennino", "CAP": "06034", "Esito_Verifica": "SCARTATO", "Sosta_Van": "No", "Dettagli_Completi": "natura | gradini_ingresso, no_bagni"}
    ])
    return utenti_fallback, luoghi_fallback

def analizza_frase_vocale(frase, cap="06034"):
    testo = frase.lower()
    tags = []
    if any(k in testo for k in ["ginnastica", "cammin", "passi", "movimento", "stretching", "yoga", "respir", "schiena", "spalle", "piedi"]):
        tags.append("movimento")
    if any(k in testo for k in ["diping", "disegn", "argilla", "arte", "manual", "creativ", "acquerell", "lavorett"]):
        tags.append("arte")
    if any(k in testo for k in ["parco", "verde", "piante", "fiori", "giardino", "natura", "alberi", "aperto", "fuori"]):
        tags.append("natura")
    if any(k in testo for k in ["compagnia", "insieme", "chiacchier", "altri", "gruppo", "solo", "sola", "amici", "giochi"]):
        tags.append("socialita")
    if any(k in testo for k in ["storie", "poesie", "legger", "libri", "cultura", "memoria", "raccont"]):
        tags.append("cultura")
    if any(k in testo for k in ["cantar", "suonar", "musica", "canzoni"]):
        tags.append("musica")
    if not tags:
        tags.append("socialita")

    fascia = "Pomeriggio" if any(k in testo for k in ["pomeriggio", "sera", "dopo pranzo", "tardi"]) else "Mattina"

    bisogni = []
    if any(k in testo for k in ["ombra", "fresco", "sole", "caldo"]):
        bisogni.append("ombra")
    if any(k in testo for k in ["sedia", "sedie", "sedut", "stare in piedi", "ripos"]):
        bisogni.append("sedute_disponibili")
    if any(k in testo for k in ["panchin"]):
        bisogni.append("panchine_frequenti")
    if any(k in testo for k in ["scale", "gradin", "barriere", "deambulator", "carrozzin", "bastone"]):
        bisogni.append("senza_barriere")
    if any(k in testo for k in ["pianeggiant", "liscio", "sterrato", "camminare a lungo"]):
        bisogni.append("terreno_pianeggiante")
    if any(k in testo for k in ["bagn", "servizi", "toilette"]):
        bisogni.append("servizi_igienici")
    if any(k in testo for k in ["tavol", "appoggi"]):
        bisogni.append("tavoli_accessibili")
    if not bisogni:
        bisogni.append("senza_barriere")

    return {
        "Zona_CAP": cap,
        "Tag_Tematici": ", ".join(dict.fromkeys(tags)),
        "Fascia_Oraria": fascia,
        "Bisogni_Pratici": ", ".join(dict.fromkeys(bisogni))
    }

def calcola_cluster(df_u, df_l, soglia=3):
    df_l_validi = df_l[df_l["Esito_Verifica"].str.contains("VALIDATO", case=False, na=False)].copy()
    clusters = defaultdict(list)

    for _, row in df_u.iterrows():
        cap = str(row["Zona_CAP"]).strip()
        fascia = str(row["Fascia_Oraria"]).strip()
        tags = [t.strip().lower() for t in str(row["Tag_Tematici"]).split(",") if t.strip()]
        tema = tags[0] if tags else "socialita"
        clusters[(cap, fascia, tema)].append(row)

    risultati = []
    for (cap, fascia, tema), membri in clusters.items():
        n = len(membri)
        bisogni_cnt = Counter()
        utenti_ids = []
        for m in membri:
            utenti_ids.append(str(m["ID_Utente"]))
            for b in str(m["Bisogni_Pratici"]).split(","):
                b_clean = b.strip().lower()
                if b_clean and b_clean != "nan":
                    bisogni_cnt[b_clean] += 1

        miglior_luogo = None
        miglior_score = -1
        sosta_van = "Da verificare"
        dotazioni_van = []

        for _, l in df_l_validi.iterrows():
            if str(l["CAP"]).strip() != cap:
                continue
            testo_l = str(l["Dettagli_Completi"]).lower()
            score = 10 if tema in testo_l else 0
            mancanti = []
            for b, peso in bisogni_cnt.items():
                if b in testo_l:
                    score += 5 * peso
                else:
                    mancanti.append(b)
            if score > miglior_score:
                miglior_score = score
                miglior_luogo = l["Nome_Luogo"]
                sosta_van = l["Sosta_Van"]
                dotazioni_van = mancanti

        risultati.append({
            "Attivo": n >= soglia,
            "Tema": tema.upper(),
            "CAP": cap,
            "Fascia": fascia,
            "Partecipanti": n,
            "Utenti_Anonimi": ", ".join(utenti_ids),
            "Luogo_Consigliato": miglior_luogo or "Nessun luogo validato nel CAP",
            "Score": miglior_score if miglior_luogo else 0,
            "Sosta_Van": sosta_van,
            "Bisogni_Gruppo": dict(bisogni_cnt),
            "Dotazioni_Van": dotazioni_van
        })
    return sorted(risultati, key=lambda x: (x["Attivo"], x["Partecipanti"]), reverse=True)

def genera_pdf_bytes(cluster_attivi):
    buf = BytesIO()
    doc = SimpleDocTemplate(buf, pagesize=A4, rightMargin=36, leftMargin=36, topMargin=36, bottomMargin=36)
    styles = getSampleStyleSheet()
    title_style = ParagraphStyle('TitleBENI', parent=styles['Heading1'], fontSize=16, textColor=colors.HexColor("#1B5E20"), spaceAfter=10)
    sub_style = ParagraphStyle('SubBENI', parent=styles['Normal'], fontSize=9, textColor=colors.HexColor("#4E342E"), spaceAfter=14)
    h2_style = ParagraphStyle('H2BENI', parent=styles['Heading2'], fontSize=12, textColor=colors.HexColor("#2E7D32"), spaceBefore=10, spaceAfter=6)

    story = []
    story.append(Paragraph("B.E.N.I. - SCHEDA OPERATIVA VAN DEL BENESSERE", title_style))
    story.append(Paragraph("Partner Operativo: SPORTIVAMBIENTE ASD | Conformità GDPR: Solo dati pseudonimizzati (audio eliminato)", sub_style))

    if not cluster_attivi:
        story.append(Paragraph("Nessun cluster ha raggiunto la soglia minima di attivazione.", styles['Normal']))
    else:
        for idx, c in enumerate(cluster_attivi, 1):
            story.append(Paragraph(f"Tappa #{idx}: Attività {c['Tema']} ({c['Fascia']} - CAP {c['CAP']})", h2_style))
            dotazioni_str = ", ".join(c['Dotazioni_Van']) if c['Dotazioni_Van'] else "Nessuna dotazione extra richiesta (luogo già attrezzato)"
            data = [
                ["Luogo Selezionato (Open Data)", f"{c['Luogo_Consigliato']} (Score: {c['Score']} pt)"],
                ["Punto Sosta Van", c['Sosta_Van']],
                ["Partecipanti Anonimi", f"{c['Partecipanti']} utenti ({c['Utenti_Anonimi']})"],
                ["Dotazioni da caricare sul Van", dotazioni_str]
            ]
            t = Table(data, colWidths=[170, 350])
            t.setStyle(TableStyle([
                ('BACKGROUND', (0, 0), (0, -1), colors.HexColor("#E8F5E9")),
                ('TEXTCOLOR', (0, 0), (-1, -1), colors.black),
                ('FONTNAME', (0, 0), (0, -1), 'Helvetica-Bold'),
                ('FONTSIZE', (0, 0), (-1, -1), 9),
                ('GRID', (0, 0), (-1, -1), 0.5, colors.grey),
                ('VALIGN', (0, 0), (-1, -1), 'TOP'),
                ('PADDING', (0, 0), (-1, -1), 6),
            ]))
            story.append(t)
            story.append(Spacer(1, 10))

    doc.build(story)
    buf.seek(0)
    return buf.getvalue()

# --- INIZIALIZZAZIONE STATO ---
if "df_utenti" not in st.session_state or "df_luoghi" not in st.session_state:
    u_init, l_init = carica_dati_iniziali()
    st.session_state.df_utenti = u_init
    st.session_state.df_luoghi = l_init

# --- INTERFACCIA WEB ---
st.title("🌿 B.E.N.I. – Benessere Emozioni Natura Insieme")
st.caption("Prototipo operativo di matching territoriale con Open Data e **SPORTIVAMBIENTE ASD**")

with st.sidebar:
    st.header("⚙️ Parametri Simulazione")
    soglia_min = st.slider("Soglia minima partecipanti per tappa Van", min_value=2, max_value=8, value=3)
    cap_default = st.text_input("CAP di riferimento", value="06034")
    if st.button("🔄 Ripristina dataset iniziale"):
        u_init, l_init = carica_dati_iniziali()
        st.session_state.df_utenti = u_init
        st.session_state.df_luoghi = l_init
        st.rerun()

tab1, tab2, tab3 = st.tabs([
    "🎙️ 1. Assistente Vocale (Utenti)",
    "🗺️ 2. Open Data & Luoghi Validati",
    "🚐 3. Cruscotto Van & PDF Operativo"
])

with tab1:
    st.subheader("Raccolta desideri con Privacy by Design")
    st.info("🔒 **Minimizzazione dati attiva:** La frase inserita viene convertita istantaneamente in tag anonimi e subito cancellata dalla memoria.")

    with st.form("form_vocale", clear_on_submit=True):
        frase_input = st.text_area(
            "Scrivi la richiesta vocale simulata dell'utente:",
            placeholder="Es. Vorrei fare ginnastica dolce al parco la mattina all'ombra, ma ho bisogno di una sedia e niente scale..."
        )
        submitted = st.form_submit_button("Elabora richiesta e pseudonimizza")

    if submitted and frase_input.strip():
        estratti = analizza_frase_vocale(frase_input, cap=cap_default.strip() or "06034")
        del frase_input  # Eliminazione immediata dato grezzo
        nuovo_id = f"USR_{len(st.session_state.df_utenti) + 1:03d}"
        nuova_riga = pd.DataFrame([{
            "ID_Utente": nuovo_id,
            "Zona_CAP": estratti["Zona_CAP"],
            "Tag_Tematici": estratti["Tag_Tematici"],
            "Fascia_Oraria": estratti["Fascia_Oraria"],
            "Bisogni_Pratici": estratti["Bisogni_Pratici"]
        }])
        st.session_state.df_utenti = pd.concat([st.session_state.df_utenti, nuova_riga], ignore_index=True)
        st.success(f"✅ Richiesta elaborata come **{nuovo_id}**. Testo originale eliminato! Tag estratti: `{estratti['Tag_Tematici']}` | Fascia: `{estratti['Fascia_Oraria']}` | Bisogni: `{estratti['Bisogni_Pratici']}`")

    st.markdown("### Registro Richieste Anonimizzate")
    st.dataframe(st.session_state.df_utenti, use_container_width=True)

with tab2:
    st.subheader("Verifica e Bonifica Open Data Territoriali")
    solo_validati = st.checkbox("Mostra solo luoghi VALIDATI sul campo con SPORTIVAMBIENTE ASD", value=False)
    df_l_view = st.session_state.df_luoghi
    if solo_validati:
        df_l_view = df_l_view[df_l_view["Esito_Verifica"].str.contains("VALIDATO", case=False, na=False)]
    st.dataframe(df_l_view, use_container_width=True)

with tab3:
    st.subheader("Proposte Tappe Van del Benessere")
    risultati_cluster = calcola_cluster(st.session_state.df_utenti, st.session_state.df_luoghi, soglia=soglia_min)
    attivi = [r for r in risultati_cluster if r["Attivo"]]
    in_attesa = [r for r in risultati_cluster if not r["Attivo"]]

    col1, col2, col3 = st.columns(3)
    col1.metric("Utenti Anonimi Totali", len(st.session_state.df_utenti))
    luoghi_ok = len(st.session_state.df_luoghi[st.session_state.df_luoghi["Esito_Verifica"].str.contains("VALIDATO", case=False, na=False)])
    col2.metric("Luoghi Open Data Validati", luoghi_ok)
    col3.metric("Tappe Van Pronte", len(attivi))

    if attivi:
        pdf_bytes = genera_pdf_bytes(attivi)
        st.download_button(
            label="📄 Scarica Scheda Operativa PDF per SPORTIVAMBIENTE ASD",
            data=pdf_bytes,
            file_name="Scheda_Operativa_Van_BENI.pdf",
            mime="application/pdf"
        )

    st.markdown("### ✅ Tappe Attivate (Soglia Raggiunta)")
    for c in attivi:
        with st.expander(f"🚐 {c['Tema']} | {c['Fascia']} (CAP {c['CAP']}) — {c['Partecipanti']} partecipanti → {c['Luogo_Consigliato']}", expanded=True):
            st.write(f"**Codici Partecipanti:** `{c['Utenti_Anonimi']}`")
            st.write(f"**Sosta Van:** {c['Sosta_Van']} | **Punteggio Idoneità Luogo:** {c['Score']} pt")
            if c["Dotazioni_Van"]:
                st.warning(f"🎒 **Attrezzature extra da caricare sul Van:** {', '.join(c['Dotazioni_Van'])}")
            else:
                st.success("✨ Il luogo copre già tutti i bisogni pratici del gruppo.")

    if in_attesa:
        st.markdown("### ⏳ Gruppi in Formazione (Sotto Soglia)")
        for c in in_attesa:
            st.write(f"- **{c['Tema']}** ({c['Fascia']}, CAP {c['CAP']}): **{c['Partecipanti']}/{soglia_min}** partecipanti (`{c['Utenti_Anonimi']}`)")
