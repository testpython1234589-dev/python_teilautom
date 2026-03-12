import json
import streamlit as st
import word_backend as wb

st.set_page_config(page_title="Word Vorlagen Generator", layout="wide")
st.title("Word Vorlagen Generator (Streamlit)")

with st.expander("📁 .docx Vorlagen im Repo anzeigen", expanded=False):
    st.write(str(wb.VORLAGEN_DIR))
    st.write([p.name for p in wb.VORLAGEN_DIR.glob("*.docx")])

tab_form, tab_json = st.tabs(["🧾 Formular", "🤖 JSON/ChatGPT einfügen"])


# -----------------------------
# TAB 1: Formular
# -----------------------------
with tab_form:
    template_choice = st.selectbox(
        "Vorlage wählen",
        [
            ("Standard Schreiben", "standard"),
            ("130 Prozent", "130"),
            ("Totalschaden (konkret)", "ts_konkret"),
            ("Totalschaden konkret unter WBW", "konkret_unter_wbw"),
            ("Totalschaden (fiktiv)", "ts_fiktiv"),
            ("Schreiben Totalschaden", "schreibentotalschaden"),
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

    # NEU: Unfallort + Unfallstraße
    u1, u2 = st.columns(2)
    with u1:
        UNFALL_ORT = st.text_input("Unfallort (Stadt/Ort)")
    with u2:
        UNFALL_STRASSE = st.text_input("Unfallstraße (Straße/Hausnr.)")

    SCHADENHERGANG = st.text_area("Schadenshergang", height=110)
    SCHADENSNUMMER = st.text_input("Schadensnummer (optional)")

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
        "SCHADENSNUMMER": SCHADENSNUMMER,

        # NEU:
        "UNFALL_ORT": UNFALL_ORT,
        "UNFALL_STRASSE": UNFALL_STRASSE,
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

    elif template_choice == "konkret_unter_wbw":
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

    elif template_choice == "ts_fiktiv":
        a, b, c = st.columns(3)
        with a:
            data["WIEDERBESCHAFFUNGSWERT"] = st.text_input("Wiederbeschaffungswert")
            data["WIEDERBESCHAFFUNGSAUFWAND"] = st.text_input("Wiederbeschaffungsaufwand")
        with b:
            data["NUTZUNGSAUSFALL"] = st.text_input("Nutzungsausfall")
            data["RESTWERT"] = st.text_input("Restwert")
        with c:
            data["ZUSATZKOSTEN_BEZEICHNUNG"] = st.text_input("Zusatzkosten Bezeichnung (optional)")
            data["ZUSATZKOSTEN_BETRAG"] = st.text_input("Zusatzkosten Betrag (optional)")

    elif template_choice == "schreibentotalschaden":
        data["WIEDERBESCHAFFUNGSWERTAUFWAND"] = st.text_input(
            "Wiederbeschaffungswertaufwand",
            placeholder="z.B. 4.321,00"
        )

    st.divider()

    if st.button("✅ Word-Datei erzeugen (Formular)", type="primary"):
        try:
            if template_choice == "standard":
                out_path = wb.vorlage_schreiben(data)
            elif template_choice == "130":
                out_path = wb.vorlage_130_prozent(data)
            elif template_choice == "ts_konkret":
                out_path = wb.vorlage_totalschaden_konkret(data)
            elif template_choice == "konkret_unter_wbw":
                out_path = wb.vorlage_totalschaden_konkret_unter_wbw(data)
            elif template_choice == "ts_fiktiv":
                out_path = wb.vorlage_totalschaden_fiktiv(data)
            elif template_choice == "schreibentotalschaden":
                out_path = wb.vorlage_schreibentotalschaden(data)
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


# -----------------------------
# TAB 2: JSON/ChatGPT einfügen
# -----------------------------
with tab_json:
    st.subheader("JSON aus ChatGPT einfügen (ein Feld, alles automatisch)")

    st.caption(
        "Erwartet JSON mit einem Feld TEMPLATE, z.B. TEMPLATE='standard' oder '130' usw. "
        "Der Rest sind deine Variablenkeys."
    )

    example = {
        "TEMPLATE": "standard",
        "MANDANT_NACHNAME": "Huss",
        "MANDANT_VORNAME": "Roondf",
        "UNFALLE_STRASSE": "An der Magistrale 59",
        "MANDANT_PLZ_ORT": "0283 Hlef",
        "UNFALL_DATUM": "01.02.2023",
        "AKTENZEICHEN": "1342",
        "FAHRZEUGTYP": "VW Golf",
        "KENNZEICHEN": "HAL 2428F",
        "VORSTEUERBERECHTIGUNG": "JA",
        "UNFALL_ORT": "Halle (Saale)",
        "UNFALL_STRASSE": "Musterstraße 12",
        "SCHADENHERGANG": "Hergang...",
        "WERTMINDERUNG": "1000",
        "REPARATURKOSTEN": "1000",
        "KOSTENPAUSCHALE": "199",
        "SACHVERST_KOSTEN": "17",
        "SCHADENSNUMMER": ""
    }

    json_text = st.text_area(
        "JSON hier einfügen",
        height=260,
        value=json.dumps(example, ensure_ascii=False, indent=2)
    )

    colA, colB = st.columns(2)
    with colA:
        btn_validate = st.button("🔎 JSON prüfen")
    with colB:
        btn_generate = st.button("✅ Word erzeugen (JSON)", type="primary")

    parsed = None
    if btn_validate or btn_generate:
        try:
            parsed = json.loads(json_text)
            st.success("JSON ist gültig.")
            st.json(parsed)
        except Exception as e:
            st.error(f"JSON Fehler: {e}")
            parsed = None

    if btn_generate and parsed:
        try:
            tpl = (parsed.get("TEMPLATE") or "").strip()

            if tpl == "standard":
                out_path = wb.vorlage_schreiben(parsed)
            elif tpl == "130":
                out_path = wb.vorlage_130_prozent(parsed)
            elif tpl == "ts_konkret":
                out_path = wb.vorlage_totalschaden_konkret(parsed)
            elif tpl == "konkret_unter_wbw":
                out_path = wb.vorlage_totalschaden_konkret_unter_wbw(parsed)
            elif tpl == "ts_fiktiv":
                out_path = wb.vorlage_totalschaden_fiktiv(parsed)
            elif tpl == "schreibentotalschaden":
                out_path = wb.vorlage_schreibentotalschaden(parsed)
            else:
                st.error(
                    "Unbekanntes TEMPLATE. Erlaubt: standard, 130, ts_konkret, konkret_unter_wbw, ts_fiktiv, schreibentotalschaden"
                )
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
            st.error(f"Fehler beim Erzeugen: {e}")
