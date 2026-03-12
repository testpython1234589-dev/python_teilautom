import json
import streamlit as st
import word_backend as wb

st.set_page_config(page_title="Word Vorlagen Generator", layout="wide")
st.title("Word Vorlagen Generator (Streamlit)")

with st.expander("📁 .docx Vorlagen im Repo anzeigen", expanded=False):
    st.write(str(wb.VORLAGEN_DIR))
    st.write([p.name for p in wb.VORLAGEN_DIR.glob("*.docx")])

tab_form, tab_json = st.tabs(["🧾 Formular", "🤖 JSON/ChatGPT einfügen"])

PROMPTS = {
    "Standard Schreiben": """Gib NUR gültiges JSON zurück. Keine Erklärungen, kein Markdown, keine Codeblöcke.
Keys exakt so lassen. Unbekannte Werte als "".
Beträge als "1.234,56" oder "1234,56".
WICHTIG: VORSTEUERBERECHTIGUNG: bei JA -> "" (leer), bei NEIN -> "nicht".
WICHTIG: Versicherung NICHT eintragen (wird separat abgefragt).

{
  "TEMPLATE": "standard",
  "MANDANT_NACHNAME": "",
  "MANDANT_VORNAME": "",
  "UNFALLE_STRASSE": "",
  "MANDANT_PLZ_ORT": "",
  "UNFALL_DATUM": "",
  "UNFALL_ORT": "",
  "UNFALL_STRASSE": "",
  "AKTENZEICHEN": "",
  "FAHRZEUGTYP": "",
  "KENNZEICHEN": "",
  "VORSTEUERBERECHTIGUNG": "",
  "SCHADENHERGANG": "",
  "SCHADENSNUMMER": "",
  "WERTMINDERUNG": "",
  "REPARATURKOSTEN": "",
  "KOSTENPAUSCHALE": "",
  "SACHVERST_KOSTEN": ""
}
""",
    "130 Prozent": """Gib NUR gültiges JSON zurück. Keine Erklärungen, kein Markdown, keine Codeblöcke.
Keys exakt so lassen. Unbekannte Werte als "".
Beträge als "1.234,56" oder "1234,56".
WICHTIG: VORSTEUERBERECHTIGUNG: bei JA -> "" (leer), bei NEIN -> "nicht".
WICHTIG: Versicherung NICHT eintragen (wird separat abgefragt).

{
  "TEMPLATE": "130",
  "MANDANT_NACHNAME": "",
  "MANDANT_VORNAME": "",
  "UNFALLE_STRASSE": "",
  "MANDANT_PLZ_ORT": "",
  "UNFALL_DATUM": "",
  "UNFALL_ORT": "",
  "UNFALL_STRASSE": "",
  "AKTENZEICHEN": "",
  "FAHRZEUGTYP": "",
  "KENNZEICHEN": "",
  "VORSTEUERBERECHTIGUNG": "",
  "SCHADENHERGANG": "",
  "SCHADENSNUMMER": "",
  "REPARATURKOSTEN": "",
  "MWST_BETRAG": "",
  "NUTZUNGSAUSFALL": "",
  "WIEDERBESCHAFFUNGSWERT": "",
  "RESTWERT": "",
  "ZUSATZKOSTEN_BEZEICHNUNG": "",
  "ZUSATZKOSTEN_BETRAG": ""
}
""",
    "Totalschaden (konkret)": """Gib NUR gültiges JSON zurück. Keine Erklärungen, kein Markdown, keine Codeblöcke.
Keys exakt so lassen. Unbekannte Werte als "".
Beträge als "1.234,56" oder "1234,56".
WICHTIG: VORSTEUERBERECHTIGUNG: bei JA -> "" (leer), bei NEIN -> "nicht".
WICHTIG: Versicherung NICHT eintragen (wird separat abgefragt).

{
  "TEMPLATE": "ts_konkret",
  "MANDANT_NACHNAME": "",
  "MANDANT_VORNAME": "",
  "UNFALLE_STRASSE": "",
  "MANDANT_PLZ_ORT": "",
  "UNFALL_DATUM": "",
  "UNFALL_ORT": "",
  "UNFALL_STRASSE": "",
  "AKTENZEICHEN": "",
  "FAHRZEUGTYP": "",
  "KENNZEICHEN": "",
  "VORSTEUERBERECHTIGUNG": "",
  "SCHADENHERGANG": "",
  "SCHADENSNUMMER": "",
  "WIEDERBESCHAFFUNGSWERT": "",
  "WIEDERBESCHAFFUNGSAUFWAND": "",
  "RESTWERT": "",
  "ERSATZBESCHAFFUNG_MWST": "",
  "ZUSATZKOSTEN_BEZEICHNUNG": "",
  "ZUSATZKOSTEN_BETRAG": ""
}
""",
    "Konkret unter WBW": """Gib NUR gültiges JSON zurück. Keine Erklärungen, kein Markdown, keine Codeblöcke.
Keys exakt so lassen. Unbekannte Werte als "".
Beträge als "1.234,56" oder "1234,56".
WICHTIG: VORSTEUERBERECHTIGUNG: bei JA -> "" (leer), bei NEIN -> "nicht".
WICHTIG: Versicherung NICHT eintragen (wird separat abgefragt).

{
  "TEMPLATE": "konkret_unter_wbw",
  "MANDANT_NACHNAME": "",
  "MANDANT_VORNAME": "",
  "UNFALLE_STRASSE": "",
  "MANDANT_PLZ_ORT": "",
  "UNFALL_DATUM": "",
  "UNFALL_ORT": "",
  "UNFALL_STRASSE": "",
  "AKTENZEICHEN": "",
  "FAHRZEUGTYP": "",
  "KENNZEICHEN": "",
  "VORSTEUERBERECHTIGUNG": "",
  "SCHADENHERGANG": "",
  "SCHADENSNUMMER": "",
  "REPARATURKOSTEN": "",
  "MWST_BETRAG": "",
  "WERTMINDERUNG": "",
  "NUTZUNGSAUSFALL": "",
  "ZUSATZKOSTEN_BEZEICHNUNG": "",
  "ZUSATZKOSTEN_BETRAG": ""
}
""",
    "Totalschaden (fiktiv)": """Gib NUR gültiges JSON zurück. Keine Erklärungen, kein Markdown, keine Codeblöcke.
Keys exakt so lassen. Unbekannte Werte als "".
Beträge als "1.234,56" oder "1234,56".
WICHTIG: VORSTEUERBERECHTIGUNG: bei JA -> "" (leer), bei NEIN -> "nicht".
WICHTIG: Versicherung NICHT eintragen (wird separat abgefragt).

{
  "TEMPLATE": "ts_fiktiv",
  "MANDANT_NACHNAME": "",
  "MANDANT_VORNAME": "",
  "UNFALLE_STRASSE": "",
  "MANDANT_PLZ_ORT": "",
  "UNFALL_DATUM": "",
  "UNFALL_ORT": "",
  "UNFALL_STRASSE": "",
  "AKTENZEICHEN": "",
  "FAHRZEUGTYP": "",
  "KENNZEICHEN": "",
  "VORSTEUERBERECHTIGUNG": "",
  "SCHADENHERGANG": "",
  "SCHADENSNUMMER": "",
  "WIEDERBESCHAFFUNGSWERT": "",
  "WIEDERBESCHAFFUNGSAUFWAND": "",
  "NUTZUNGSAUSFALL": "",
  "RESTWERT": "",
  "ZUSATZKOSTEN_BEZEICHNUNG": "",
  "ZUSATZKOSTEN_BETRAG": ""
}
""",
    "Schreiben Totalschaden": """Gib NUR gültiges JSON zurück. Keine Erklärungen, kein Markdown, keine Codeblöcke.
Keys exakt so lassen. Unbekannte Werte als "".
Beträge als "1.234,56" oder "1234,56".
WICHTIG: VORSTEUERBERECHTIGUNG: bei JA -> "" (leer), bei NEIN -> "nicht".
WICHTIG: Versicherung NICHT eintragen (wird separat abgefragt).

{
  "TEMPLATE": "schreibentotalschaden",
  "MANDANT_NACHNAME": "",
  "MANDANT_VORNAME": "",
  "UNFALLE_STRASSE": "",
  "MANDANT_PLZ_ORT": "",
  "UNFALL_DATUM": "",
  "UNFALL_ORT": "",
  "UNFALL_STRASSE": "",
  "AKTENZEICHEN": "",
  "FAHRZEUGTYP": "",
  "KENNZEICHEN": "",
  "VORSTEUERBERECHTIGUNG": "",
  "SCHADENHERGANG": "",
  "SCHADENSNUMMER": "",
  "WIEDERBESCHAFFUNGSWERTAUFWAND": ""
}
"""
}

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

    u1, u2 = st.columns(2)
    with u1:
        UNFALL_ORT = st.text_input("Unfallort (Stadt/Ort)")
    with u2:
        UNFALL_STRASSE = st.text_input("Unfallstraße (Straße/Hausnr.)")

    v1, v2, v3 = st.columns(3)
    with v1:
        VERSICHERUNG = st.text_input("Versicherung")
    with v2:
        VER_STRASSE = st.text_input("Versicherung Straße")
    with v3:
        VER_ORT = st.text_input("Versicherung PLZ/Ort")

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
        "UNFALL_ORT": UNFALL_ORT,
        "UNFALL_STRASSE": UNFALL_STRASSE,
        "VERSICHERUNG": VERSICHERUNG,
        "VER_STRASSE": VER_STRASSE,
        "VER_ORT": VER_ORT,
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
        data["WIEDERBESCHAFFUNGSWERTAUFWAND"] = st.text_input("Wiederbeschaffungswertaufwand")

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
        except Exception as e:
            st.error(f"Fehler: {e}")

# -----------------------------
# TAB 2: JSON/ChatGPT einfügen (Versicherung separat!)
# -----------------------------
with tab_json:
    st.subheader("JSON aus ChatGPT einfügen (ein Feld) + Versicherung separat")

    st.markdown("### 1) Prompt auswählen & kopieren")
    prompt_name = st.selectbox("Prompt wählen", list(PROMPTS.keys()))
    st.code(PROMPTS[prompt_name], language="text")

    st.markdown("### 2) JSON von ChatGPT hier einfügen")
    json_text = st.text_area("JSON", height=260)

    st.markdown("### 3) Versicherung separat eingeben (nicht aus JSON)")
    v1, v2, v3 = st.columns(3)
    with v1:
        VERSICHERUNG_J = st.text_input("Versicherung", key="json_vers")
    with v2:
        VER_STRASSE_J = st.text_input("Versicherung Straße", key="json_ver_str")
    with v3:
        VER_ORT_J = st.text_input("Versicherung PLZ/Ort", key="json_ver_ort")

    colA, colB = st.columns(2)
    with colA:
        btn_validate = st.button("🔎 JSON prüfen", key="json_validate")
    with colB:
        btn_generate = st.button("✅ Word erzeugen (JSON)", type="primary", key="json_generate")

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
            parsed["VERSICHERUNG"] = VERSICHERUNG_J
            parsed["VER_STRASSE"] = VER_STRASSE_J
            parsed["VER_ORT"] = VER_ORT_J

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
                st.error("Unbekanntes TEMPLATE.")
                st.stop()

            st.success(f"Erstellt: {out_path.name}")
            with open(out_path, "rb") as f:
                st.download_button(
                    "⬇️ Download .docx",
                    data=f,
                    file_name=out_path.name,
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                )
        except Exception as e:
            st.error(f"Fehler beim Erzeugen: {e}")
