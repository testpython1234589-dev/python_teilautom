import streamlit as st
import word_backend as wb

st.set_page_config(page_title="Word Vorlagen Generator", layout="wide")
st.title("Word Vorlagen Generator (Repo-Layout: Vorlagen im Root)")

with st.expander("📁 Vorlagen-Dateien im Repo", expanded=False):
    st.write(str(wb.VORLAGEN_DIR))
    st.write([p.name for p in wb.VORLAGEN_DIR.glob("*.docx")])

template_choice = st.selectbox(
    "Vorlage wählen",
    [
        ("Standard Schreiben", "standard"),
        ("130 Prozent", "130"),
        ("Totalschaden (konkret)", "ts_konkret"),
        ("Totalschaden konkret unter WBW", "ts_konkret_unter_wbw"),
        # optional später:
        # ("Schreiben Totalschaden", "schreibentotalschaden"),
        # ("Totalschaden (fiktiv)", "ts_fiktiv"),
    ],
    format_func=lambda x: x[0]
)[1]

st.subheader("Standarddaten (für alle Vorlagen)")
c1, c2, c3 = st.columns(3)
with c1:
    MANDANT_NACHNAME = st.text_input("Mandant Nachname")
    UNFALLE_STRASSE = st.text_input("Straße (Mandant)")
    UNFALL_DATUM = st.text_input("Unfalldatum")
with c2:
    MANDANT_VORNAME = st.text_input("Mandant Vorname")
    MANDANT_PLZ_ORT = st.text_input("PLZ / Ort")
    AKTENZEICHEN = st.text_input("Aktenzeichen")
with c3:
    FAHRZEUGTYP = st.text_input("Fahrzeugtyp")
    KENNZEICHEN = st.text_input("Kennzeichen")
    VORSTEUERBERECHTIGUNG = st.text_input("Vorsteuerberechtigt (JA/NEIN)")

SCHADENHERGANG = st.text_area("Schadenshergang", height=110)

data = {
    "MANDANT_NACHNAME": MANDANT_NACHNAME,
    "MANDANT_VORNAME": MANDANT_VORNAME,
    "UNFALLE_STRASSE": UNFALLE_STRASSE,
    "MANDANT_PLZ_ORT": MANDANT_PLZ_ORT,
    "UNFALL_DATUM": UNFALL_DATUM,
    "AKTENZEICHEN": AKTENZEICHEN,
    "FAHRZEUGTYP": FAHRZEUGTYP,
    "KENNZEICHEN": KENNZEICHEN,
    "VORSTEUERBERECHTIGUNG": VORSTEUERBERECHTIGUNG,
    "SCHADENHERGANG": SCHADENHERGANG,
}

st.subheader("Vorlagen-spezifische Angaben")

if template_choice == "standard":
    a, b, c, d = st.columns(4)
    with a:
        data["WERTMINDERUNG"] = st.text_input("Wertminderung", placeholder="1.001,99")
    with b:
        data["REPARATURKOSTEN"] = st.text_input("Reparaturkosten")
    with c:
        data["KOSTENPAUSCHALE"] = st.text_input("Kostenpauschale (optional)")
    with d:
        data["SACHVERST_KOSTEN"] = st.text_input("Sachverständigerkosten")

elif template_choice == "130":
    a, b, c = st.columns(3)
    with a:
        data["REPARATURKOSTEN"] = st.text_input("Reparaturkosten")
        data["MWST_BETRAG"] = st.text_input("Mehrwertsteuer")
    with b:
        data["NUTZUNGSAUSFALL"] = st.text_input("Nutzungsausfall")
        data["WIEDERBESCHAFFUNGSWERT"] = st.text_input("Wiederbeschaffungswert")
    with c:
        data["RESTWERT"] = st.text_input("Restwert")
        data["ZUSATZKOSTEN_BEZEICHNUNG"] = st.text_input("Zusatzkosten Bezeichnung (optional)")
        data["ZUSATZKOSTEN_BETRAG"] = st.text_input("Zusatzkosten Betrag (optional)")

elif template_choice == "ts_konkret":
    a, b, c = st.columns(3)
    with a:
        data["WIEDERBESCHAFFUNGSWERT"] = st.text_input("Wiederbeschaffungswert")
        data["WIEDERBESCHAFFUNGSAUFWAND"] = st.text_input("Wiederbeschaffungsaufwand")
    with b:
        data["RESTWERT"] = st.text_input("Restwert")
        data["ERSATZBESCHAFFUNG_MWST"] = st.text_input("Ersatzbeschaffungs-MWST")
    with c:
        data["ZUSATZKOSTEN_BEZEICHNUNG"] = st.text_input("Zusatzkosten Bezeichnung (optional)")
        data["ZUSATZKOSTEN_BETRAG"] = st.text_input("Zusatzkosten Betrag (optional)")

elif template_choice == "ts_konkret_unter_wbw":
    a, b, c = st.columns(3)
    with a:
        data["REPARATURKOSTEN"] = st.text_input("Reparaturkosten")
        data["MWST_BETRAG"] = st.text_input("Mehrwertsteuer")
    with b:
        data["WERTMINDERUNG"] = st.text_input("Wertminderung")
        data["NUTZUNGSAUSFALL"] = st.text_input("Nutzungsausfall")
    with c:
        data["ZUSATZKOSTEN_BEZEICHNUNG"] = st.text_input("Zusatzkosten Bezeichnung (optional)")
        data["ZUSATZKOSTEN_BETRAG"] = st.text_input("Zusatzkosten Betrag (optional)")

st.divider()

if st.button("✅ Word-Datei erzeugen", type="primary"):
    try:
        if template_choice == "standard":
            out_path = wb.vorlage_schreiben(data)
        elif template_choice == "130":
            out_path = wb.vorlage_130_prozent(data)
        elif template_choice == "ts_konkret":
            out_path = wb.vorlage_totalschaden_konkret(data)
        elif template_choice == "ts_konkret_unter_wbw":
            out_path = wb.vorlage_totalschaden_konkret_unter_wbw(data)
        else:
            st.error("Unbekannte Vorlage.")
            st.stop()

        st.success(f"Erstellt: {out_path.name}")

        with open(out_path, "rb") as f:
            st.download_button(
                "⬇️ Download .docx",
                data=f,
                file_name=out_path.name,
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            )

        st.caption(f"Gespeichert in: {wb.OUTPUT_DIR}")

    except Exception as e:
        st.error(f"Fehler: {e}")
