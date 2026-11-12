import streamlit as st
import pandas as pd
from collections import Counter, defaultdict
import datetime
import io
import os
from reportlab.lib.pagesizes import A4
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib import colors

# ==============================================================================
# CONFIGURAZIONE PAGINA STREAMLIT
# ==============================================================================
st.set_page_config(
    page_title="B.E.N.I. - Benessere Emozioni Natura Insieme",
    page_icon="🌿",
    layout="wide"
)

# ==============================================================================
# FUNZIONI DI CARICAMENTO E LOGICA DI DOMINIO
# ==============================================================================
FILE_EXCEL = "BENI_Simulazione_Dataset_e_Matching.xlsx"

def carica_dati_iniziali():
    if os.path.exists(FILE_EXCEL):
        try:
            df_utenti = pd.read_excel(FILE_EXCEL, sheet_name="1_Richieste_Utenti_Anonime")
            df_luoghi = pd.read_excel(FILE_EXCEL, sheet_name="2_Luoghi_OpenData_Validati")
            col_ammesse = ["ID_Utente", "Zona_CAP", "Tag_Tematici", "Fascia_Oraria", "Bisogni_Pratici"]
            col_presenti = [c for c in col_ammesse if c in df_utenti.columns]
            df_utenti = df_utenti[col_presenti].copy()
            return df_utenti, df_luoghi
        except Exception:
            pass

    # Dataset di fallback integrato (Area Foligno - CAP 06034)
    utenti_data = [
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
    ]
    luoghi_data = [
        {"ID_Luogo": "LOC_01", "Nome_Luogo": "Parco dei Canapè (Area Centrale)", "Zona_CAP": "06034", "Esito_Verifica": "VALIDATO", "Sosta_Van": "Sì (Via Nazario Sauro)", "Attributi_Verificati": "terreno_pianeggiante, senza_barriere, ombra, panchine_frequenti, sedute_disponibili, servizi_igienici", "Tag_Idonei": "movimento, natura, arte, socialita, cultura", "Note_Campo": "Vialetti asfaltati larghi, bagni accessibili vicino al chiosco."},
        {"ID_Luogo": "LOC_02", "Nome_Luogo": "Parco Fluviale Hoffmann", "Zona_CAP": "06034", "Esito_Verifica": "VALIDATO", "Sosta_Van": "Sì (Parcheggio adiacente)", "Attributi_Verificati": "terreno_pianeggiante, senza_barriere, ombra, panchine_frequenti", "Tag_Idonei": "movimento, natura, socialita", "Note_Campo": "Percorso pianeggiante lungo il fiume Topino, no servizi igienici fissi."},
        {"ID_Luogo": "LOC_03", "Nome_Luogo": "Giardino di Villa Candida", "Zona_CAP": "06034", "Esito_Verifica": "VALIDATO", "Sosta_Van": "Sì (Ingresso laterale)", "Attributi_Verificati": "senza_barriere, ombra, sedute_disponibili, tavoli_accessibili, servizi_igienici", "Tag_Idonei": "arte, cultura, socialita, natura", "Note_Campo": "Spazio raccolto e silenzioso, ideale per laboratori manuali e lettura."},
        {"ID_Luogo": "LOC_04", "Nome_Luogo": "Parco di Via XVI Giugno", "Zona_CAP": "06034", "Esito_Verifica": "VALIDATO", "Sosta_Van": "Sì (Area sosta frontale)", "Attributi_Verificati": "terreno_pianeggiante, senza_barriere, panchine_frequenti", "Tag_Idonei": "movimento, socialita", "Note_Campo": "Buona accessibilità, scarsa ombra nelle ore centrali."},
        {"ID_Luogo": "LOC_05", "Nome_Luogo": "Area Verde Via Monte Pennino", "Zona_CAP": "06034", "Esito_Verifica": "SCARTATO", "Sosta_Van": "No", "Attributi_Verificati": "gradini_ingresso, no_bagni", "Tag_Idonei": "natura", "Note_Campo": "Open Data obsoleti: presenti 3 gradini non segnalati all'ingresso."}
    ]
    return pd.DataFrame(utenti_data), pd.DataFrame(luoghi_data)

def analizza_frase_vocale(frase, cap="06034"):
    testo = frase.lower()
    tags = []
    if any(k in testo for k in ["ginnastica", "cammin", "passi", "stretching", "movimento", "yoga", "respir", "schiena", "spalle"]):
        tags.append("movimento")
    if any(k in testo for k in ["diping", "disegn", "argilla", "manual", "arte", "acquerello", "creare", "lavoretti"]):
        tags.append("arte")
    if any(k in testo for k in ["parco", "verde", "piante", "fiori", "alberi", "natura", "aperto", "giardino", "fresco"]):
        tags.append("natura")
    if any(k in testo for k in ["compagnia", "chiacchier", "altri", "insieme", "gruppo", "solo", "sola", "giochi"]):
        tags.append("socialita")
    if any(k in testo for k in ["leggere", "poesie", "storie", "memoria", "libri", "cultura"]):
        tags.append("cultura")
    if any(k in testo for k in ["suonare", "cantare", "musica", "canzoni"]):
        tags.append("musica")
    if not tags:
        tags.append("socialita")

    fascia = "Pomeriggio" if any(k in testo for k in ["pomeriggio", "sera", "dopo pranzo", "tardo"]) else "Mattina"

    bisogni = []
    if any(k in testo for k in ["ombra", "fresco", "sole", "caldo"]):
        bisogni.append("ombra")
    if any(k in testo for k in ["sedia", "sedie", "sedut", "stare in piedi", "ripos"]):
        bisogni.append("sedute_disponibili")
    if any(k in testo for k in ["panchin"]):
        bisogni.append("panchine_frequenti")
    if any(k in testo for k in ["pianeggiante", "liscio", "bastone", "deambulatore", "sterrato", "camminare a lungo"]):
        bisogni.append("terreno_pianeggiante")
    if any(k in testo for k in ["scale", "gradini", "barriere", "carrozzina", "accesso"]):
        bisogni.append("senza_barriere")
    if any(k in testo for k in ["bagn", "servizi", "toilette"]):
        bisogni.append("servizi_igienici")
    if any(k in testo for k in ["tavol", "appoggio"]):
        bisogni.append("tavoli_accessibili")
    if not bisogni:
        bisogni.append("senza_barriere")

    return {
        "Zona_CAP": str(cap).replace("CAP_", "").strip(),
        "Tag_Tematici": ", ".join(dict.fromkeys(tags)),
        "Fascia_Oraria": fascia,
        "Bisogni_Pratici": ", ".join(dict.fromkeys(bisogni))
    }

def calcola_matching(df_utenti, df_luoghi, soglia_minima=3):
    col_esito = [c for c in df_luoghi.columns if "Esito" in c or "Stato" in c]
    if col_esito:
        df_validi = df_luoghi[df_luoghi[col_esito[0]].astype(str).str.contains("VALIDATO", case=False)].copy()
    else:
        df_validi = df_luoghi.copy()

    col_cap_luoghi = [c for c in df_validi.columns if "CAP" in c]
    nome_col_cap = col_cap_luoghi[0] if col_cap_luoghi else "Zona_CAP"

    clusters = defaultdict(list)
    for _, row in df_utenti.iterrows():
        cap = str(row["Zona_CAP"]).replace("CAP_", "").strip()
        fascia = str(row["Fascia_Oraria"]).strip()
        tags = [t.strip().lower() for t in str(row["Tag_Tematici"]).split(",") if t.strip()]
        tema_principale = tags[0] if tags else "socialita"
        clusters[(cap, fascia, tema_principale)].append(row)

    risultati = []
    for (cap, fascia, tema), membri in clusters.items():
        num_utenti = len(membri)
        bisogni_gruppo = Counter()
        id_partecipanti = []
        for m in membri:
            id_partecipanti.append(str(m["ID_Utente"]))
            for b in str(m["Bisogni_Pratici"]).split(","):
                b_clean = b.strip().lower()
                if b_clean and b_clean != "nan":
                    bisogni_gruppo[b_clean] += 1

        attivo = num_utenti >= soglia_minima
        miglior_luogo = None

        if attivo:
            classifica = []
            for _, luogo in df_validi.iterrows():
                cap_luogo = str(luogo[nome_col_cap]).replace("CAP_", "").strip()
                if cap_luogo != cap:
                    continue
                testo_luogo = " ".join([str(val).lower() for val in luogo.values])
                score = 10 if tema in testo_luogo else 0
                ok, ko = [], []
                for bisogno, peso in bisogni_gruppo.items():
                    if bisogno in testo_luogo:
                        score += 5 * peso
                        ok.append(bisogno)
                    else:
                        ko.append(bisogno)
                nome_l = luogo.get("Nome_Luogo", str(luogo.iloc[1]))
                sosta_l = luogo.get("Sosta_Van", "Verificare in loco")
                classifica.append({
                    "score": score,
                    "nome": nome_l,
                    "sosta": sosta_l,
                    "ok": ok,
                    "ko": ko
                })
            classifica.sort(key=lambda x: x["score"], reverse=True)
            if classifica:
                miglior_luogo = classifica[0]

        risultati.append({
            "cap": cap,
            "fascia": fascia,
            "tema": tema,
            "num_utenti": num_utenti,
            "attivo": attivo,
            "id_partecipanti": id_partecipanti,
            "bisogni": dict(bisogni_gruppo),
            "luogo": miglior_luogo
        })

    risultati.sort(key=lambda x: (not x["attivo"], -x["num_utenti"]))
    return risultati

def genera_pdf_bytes(risultati, soglia_minima):
    buffer = io.BytesIO()
    doc = SimpleDocTemplate(buffer, pagesize=A4, rightMargin=36, leftMargin=36, topMargin=36, bottomMargin=36)
    styles = getSampleStyleSheet()
    title_style = ParagraphStyle('TitleBeni', parent=styles['Heading1'], fontSize=16, textColor=colors.HexColor("#1B5E20"), spaceAfter=6)
    sub_style = ParagraphStyle('SubBeni', parent=styles['Normal'], fontSize=10, textColor=colors.HexColor("#424242"), spaceAfter=14)
    h2_style = ParagraphStyle('H2Beni', parent=styles['Heading2'], fontSize=12, textColor=colors.HexColor("#2E7D32"), spaceBefore=12, spaceAfter=6)
    normal_style = ParagraphStyle('NormBeni', parent=styles['Normal'], fontSize=9.5, leading=13)

    story = []
    oggi = datetime.date.today().strftime("%d/%m/%Y")
    story.append(Paragraph("PROGETTO B.E.N.I. - SCHEDA OPERATIVA VAN DEL BENESSERE", title_style))
    story.append(Paragraph(f"Partner Operativo: SPORTIVAMBIENTE ASD | Data generazione: {oggi} | Soglia attivazione: {soglia_minima} utenti", sub_style))
    story.append(Paragraph("<b>Conformità Privacy by Design (GDPR):</b> Le richieste vocali originali sono state eliminate immediatamente dopo l'estrazione dei tag. Il presente documento contiene esclusivamente codici pseudonimizzati e requisiti logistici aggregati.", normal_style))
    story.append(Spacer(1, 12))

    tappe_attive = [r for r in risultati if r["attivo"] and r["luogo"]]
    story.append(Paragraph(f"1. Tappe Attivate dal Motore di Matching ({len(tappe_attive)} tappe)", h2_style))

    if not tappe_attive:
        story.append(Paragraph("Nessun cluster ha ancora raggiunto la soglia minima impostata.", normal_style))
    else:
        data_table = [["Tema / Fascia", "Partecipanti", "Luogo Selezionato (Open Data)", "Dotazioni da caricare sul Van"]]
        for r in tappe_attive:
            luogo = r["luogo"]
            ko_text = ", ".join(luogo["ko"]) if luogo["ko"] else "Nessuna extra (luogo completo)"
            col1 = Paragraph(f"<b>{r['tema'].upper()}</b><br/>{r['fascia']} (CAP {r['cap']})", normal_style)
            col2 = Paragraph(f"<b>{r['num_utenti']} utenti</b><br/>{', '.join(r['id_partecipanti'])}", normal_style)
            col3 = Paragraph(f"<b>{luogo['nome']}</b><br/>Score: {luogo['score']} pt<br/>Sosta: {luogo['sosta']}", normal_style)
            col4 = Paragraph(ko_text, normal_style)
            data_table.append([col1, col2, col3, col4])

        t = Table(data_table, colWidths=[105, 110, 175, 130])
        t.setStyle(TableStyle([
            ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor("#2E7D32")),
            ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),
            ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
            ('FONTSIZE', (0, 0), (-1, 0), 9),
            ('GRID', (0, 0), (-1, -1), 0.5, colors.HexColor("#BDBDBD")),
            ('VALIGN', (0, 0), (-1, -1), 'TOP'),
            ('ROWBACKGROUNDS', (0, 1), (-1, -1), [colors.white, colors.HexColor("#F1F8E9")]),
            ('PADDING', (0, 0), (-1, -1), 6),
        ]))
        story.append(t)

    story.append(Spacer(1, 14))
    story.append(Paragraph("2. Checklist Pre-Partenza Operatori SPORTIVAMBIENTE ASD", h2_style))
    checklist = [
        "[  ] Verifica carico attrezzature specifiche indicate nella colonna 'Dotazioni da caricare sul Van'.",
        "[  ] Controllo accessibilità ultimo metro (assenza ostacoli temporanei o cantieri nell'area sosta).",
        "[  ] Presenza kit primo soccorso, acqua e materiale informativo per gli attori locali coinvolti.",
        "[  ] Verifica bollettino meteo e attivazione spazio civico al coperto in caso di pioggia."
    ]
    for item in checklist:
        story.append(Paragraph(item, normal_style))
        story.append(Spacer(1, 4))

    doc.build(story)
    buffer.seek(0)
    return buffer.getvalue()

# ==============================================================================
# INIZIALIZZAZIONE STATO SESSIONE
# ==============================================================================
if "df_utenti" not in st.session_state or "df_luoghi" not in st.session_state:
    u_init, l_init = carica_dati_iniziali()
    st.session_state.df_utenti = u_init
    st.session_state.df_luoghi = l_init
    st.session_state.ultimo_log = None

# ==============================================================================
# INTERFACCIA UTENTE (HEADER E SIDEBAR)
# ==============================================================================
st.title("🌿 B.E.N.I. – Benessere Emozioni Natura Insieme")
st.markdown("Prototipo interattivo per la raccolta desideri con **Privacy by Design**, validazione **Open Data** e pianificazione tappe del **Van del Benessere** con **SPORTIVAMBIENTE ASD**.")

with st.sidebar:
    st.header("⚙️ Parametri Simulazione")
    soglia_minima = st.slider("Soglia minima partecipanti per attivare il Van", min_value=2, max_value=8, value=3)
    st.divider()
    st.subheader("📊 Stato Attuale")
    st.metric("Richieste anonime raccolte", len(st.session_state.df_utenti))
    luoghi_ok = len(st.session_state.df_luoghi[st.session_state.df_luoghi["Esito_Verifica"].astype(str).str.contains("VALIDATO", case=False)])
    st.metric("Luoghi Open Data validati", f"{luoghi_ok} / {len(st.session_state.df_luoghi)}")
    st.divider()
    if st.button("🔄 Ripristina Dataset Iniziale", use_container_width=True):
        u_init, l_init = carica_dati_iniziali()
        st.session_state.df_utenti = u_init
        st.session_state.df_luoghi = l_init
        st.session_state.ultimo_log = None
        st.rerun()

# ==============================================================================
# NAVIGAZIONE A SCHEDE (3 TAB)
# ==============================================================================
tab1, tab2, tab3 = st.tabs([
    "🎙️ 1. Assistente Vocale IA (Lato Utente)",
    "🗺️ 2. Verifica Open Data (Territorio)",
    "🚐 3. Cruscotto Van & PDF (SPORTIVAMBIENTE ASD)"
])

# --- TAB 1: ASSISTENTE VOCALE IA ---
with tab1:
    st.subheader("Simulazione Ascolto e Minimizzazione Dati (GDPR)")
    st.info("Inserisci una richiesta come se fosse pronunciata da una persona anziana o fragile. L'IA estrarrà solo i tag tematici e i bisogni pratici, eliminando subito il testo originale.")

    col_input, col_esempi = st.columns([2, 1])

    with col_input:
        with st.form("form_vocale", clear_on_submit=True):
            frase_utente = st.text_area(
                "Cosa desideri fare all'aperto?",
                placeholder="Es. Vorrei fare ginnastica dolce per la schiena al parco al mattino, ma ho bisogno di stare all'ombra e avere una sedia...",
                height=110
            )
            cap_scelto = st.selectbox("Zona / CAP di riferimento", ["06034 (Foligno)"], index=0)
            invia = st.form_submit_button("✨ Elabora Desiderio e Minimizza Dati", type="primary")

            if invia and frase_utente.strip():
                cap_clean = cap_scelto.split(" ")[0]
                metadati = analizza_frase_vocale(frase_utente, cap=cap_clean)
                nuovo_id = f"USR_{len(st.session_state.df_utenti) + 1:03d}"

                nuovo_record = {
                    "ID_Utente": nuovo_id,
                    "Zona_CAP": metadati["Zona_CAP"],
                    "Tag_Tematici": metadati["Tag_Tematici"],
                    "Fascia_Oraria": metadati["Fascia_Oraria"],
                    "Bisogni_Pratici": metadati["Bisogni_Pratici"]
                }

                # Eliminazione esplicita del testo grezzo (Privacy by Design)
                del frase_utente

                st.session_state.df_utenti = pd.concat(
                    [st.session_state.df_utenti, pd.DataFrame([nuovo_record])],
                    ignore_index=True
                )
                st.session_state.ultimo_log = nuovo_record
                st.rerun()

    with col_esempi:
        st.markdown("**💡 Frasi di esempio da copiare:**")
        st.caption("1. *Vorrei fare yoga sulla sedia all'ombra al mattino, senza fare scale.*")
        st.caption("2. *Mi piacerebbe leggere poesie in giardino nel pomeriggio vicino ai bagni.*")
        st.caption("3. *Vorrei dipingere all'aperto la mattina, ma mi serve un tavolo comodo.*")

    if st.session_state.ultimo_log:
        st.success(f"✅ Richiesta elaborata! Il testo originale è stato **cancellato**. Salvato solo il record pseudonimizzato **{st.session_state.ultimo_log['ID_Utente']}**:")
        st.json(st.session_state.ultimo_log)

    st.markdown("### Archivio Richieste Pseudonimizzate (Nessun dato vocale o clinico salvato)")
    st.dataframe(st.session_state.df_utenti, use_container_width=True, hide_index=True)

# --- TAB 2: VERIFICA OPEN DATA ---
with tab2:
    st.subheader("Bonifica e Validazione Open Data Territoriali")
    st.markdown("Confronto tra i dati estratti dai portali pubblici (OpenStreetMap / Open Data) e i sopralluoghi reali effettuati con **SPORTIVAMBIENTE ASD**.")

    filtro_stato = st.radio("Filtra luoghi per stato di verifica:", ["Tutti", "Solo VALIDATI", "Solo SCARTATI"], horizontal=True)
    df_l_view = st.session_state.df_luoghi.copy()
    if filtro_stato == "Solo VALIDATI":
        df_l_view = df_l_view[df_l_view["Esito_Verifica"].str.contains("VALIDATO", case=False)]
    elif filtro_stato == "Solo SCARTATI":
        df_l_view = df_l_view[df_l_view["Esito_Verifica"].str.contains("SCARTATO", case=False)]

    st.dataframe(df_l_view, use_container_width=True, hide_index=True)

# --- TAB 3: CRUSCOTTO OPERATIVO VAN ---
with tab3:
    st.subheader("Proposte Tappe Van del Benessere & Comunità Tematiche")
    risultati = calcola_matching(st.session_state.df_utenti, st.session_state.df_luoghi, soglia_minima=soglia_minima)

    tappe_pronte = [r for r in risultati if r["attivo"]]
    in_attesa = [r for r in risultati if not r["attivo"]]

    col_btn, _ = st.columns([1, 2])
    with col_btn:
        pdf_bytes = genera_pdf_bytes(risultati, soglia_minima)
        st.download_button(
            label="📄 Scarica Scheda Operativa PDF per SPORTIVAMBIENTE ASD",
            data=pdf_bytes,
            file_name="Scheda_Operativa_Van_BENI.pdf",
            mime="application/pdf",
            type="primary"
        )

    st.markdown(f"### 🟢 Tappe Pronte per l'Attivazione ({len(tappe_pronte)})")
    if not tappe_pronte:
        st.warning("Nessun gruppo ha ancora raggiunto la soglia minima. Prova ad abbassare la soglia nella barra laterale o ad aggiungere nuove richieste nella Scheda 1.")
    else:
        for r in tappe_pronte:
            with st.container(border=True):
                c1, c2, c3 = st.columns([1.2, 1.5, 1.3])
                with c1:
                    st.markdown(f"#### Tema: {r['tema'].upper()}")
                    st.write(f"**Fascia:** {r['fascia']} | **CAP:** {r['cap']}")
                    st.write(f"**Partecipanti ({r['num_utenti']}):** {', '.join(r['id_partecipanti'])}")
                with c2:
                    if r["luogo"]:
                        st.markdown(f"#### 📍 {r['luogo']['nome']}")
                        st.write(f"**Punteggio idoneità:** {r['luogo']['score']} pt")
                        st.write(f"**Sosta Van:** {r['luogo']['sosta']}")
                        st.write(f"**Requisiti coperti:** {', '.join(r['luogo']['ok']) if r['luogo']['ok'] else 'Base'}")
                with c3:
                    st.markdown("#### 🎒 Carico Van Richiesto")
                    if r["luogo"] and r["luogo"]["ko"]:
                        for item in r["luogo"]["ko"]:
                            st.warning(f"Portare: **{item}**")
                    else:
                        st.success("Nessuna dotazione extra necessaria.")

    if in_attesa:
        st.markdown(f"### 🟡 Gruppi in Formazione (Sotto soglia di {soglia_minima} utenti)")
        for r in in_attesa:
            st.write(f"- **{r['tema'].capitalize()}** ({r['fascia']}, CAP {r['cap']}): **{r['num_utenti']}** aderenti ({', '.join(r['id_partecipanti'])}) – mancano {soglia_minima - r['num_utenti']} adesioni.")
