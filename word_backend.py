from __future__ import annotations

import json
from datetime import date, timedelta, datetime
from typing import Dict, Any, Set

import streamlit as st
import word_backend as wb


# -----------------------------
# Vorlagen
# -----------------------------
TEMPLATES = {
    "Standard Schreiben": ("vorlage_schreiben-1.docx", "Standard_schreiben"),
    "130 Prozent": ("vorlage_130_prozent-1.docx", "130_prozent"),
    "Totalschaden (konkret)": ("vorlage_totalschaden_konkret-1.docx", "totalschaden_konkret"),
    "Konkret unter WBW": ("vorlage_konkret_unter_wbw-1.docx", "konkret_unter_wbw"),
    "Totalschaden (fiktiv)": ("vorlage_totalschaden_fiktiv-1.docx", "totalschaden_fiktiv"),
    "Schreiben Totalschaden": ("vorlage_schreibentotalschaden-1.docx", "schreibentotalschaden"),
}


# -----------------------------
# Prompts (sichtbar + kopierbar via st.code)
# -----------------------------
PROMPTS = {
    "Standard Schreiben": """Gib NUR JSON zurück (keine Erklärungen).
Unbekannt -> "".
VORSTEUERBERECHTIGUNG: JA -> "" (leer), NEIN -> "nicht".

{
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
  "VERSICHERUNG": "",
  "VER_STRASSE": "",
  "VER_ORT": "",
  "WERTMINDERUNG": "",
  "REPARATURKOSTEN": "",
  "KOSTENPAUSCHALE": "",
  "SACHVERST_KOSTEN": ""
}
""",
    "130 Prozent": """Gib NUR JSON zurück (keine Erklärungen).
Unbekannt -> "".
VORSTEUERBERECHTIGUNG: JA -> "" (leer), NEIN -> "nicht".

{
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
  "VERSICHERUNG": "",
  "VER_STRASSE": "",
  "VER_ORT": "",
  "REPARATURKOSTEN": "",
  "MWST_BETRAG": "",
  "NUTZUNGSAUSFALL": "",
  "WIEDERBESCHAFFUNGSWERT": "",
  "RESTWERT": "",
  "ZUSATZKOSTEN_BEZEICHNUNG": "",
  "ZUSATZKOSTEN_BETRAG": ""
}
""",
    "Totalschaden (konkret)": """Gib NUR JSON zurück (keine Erklärungen).
Unbekannt -> "".
VORSTEUERBERECHTIGUNG: JA -> "" (leer), NEIN -> "nicht".

{
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
  "VERSICHERUNG": "",
  "VER_STRASSE": "",
  "VER_ORT": "",
  "WIEDERBESCHAFFUNGSWERT": "",
  "WIEDERBESCHAFFUNGSAUFWAND": "",
  "RESTWERT": "",
  "ERSATZBESCHAFFUNG_MWST": "",
  "ZUSATZKOSTEN_BEZEICHNUNG": "",
  "ZUSATZKOSTEN_BETRAG": ""
}
""",
    "Konkret unter WBW": """Gib NUR JSON zurück (keine Erklärungen).
Unbekannt -> "".
VORSTEUERBERECHTIGUNG: JA -> "" (leer), NEIN -> "nicht".

{
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
  "VERSICHERUNG": "",
  "VER_STRASSE": "",
  "VER_ORT": "",
  "REPARATURKOSTEN": "",
  "MWST_BETRAG": "",
  "WERTMINDERUNG": "",
  "NUTZUNGSAUSFALL": "",
  "ZUSATZKOSTEN_BEZEICHNUNG": "",
  "ZUSATZKOSTEN_BETRAG": ""
}
""",
    "Totalschaden (fiktiv)": """Gib NUR JSON zurück (keine Erklärungen).
Unbekannt -> "".
VORSTEUERBERECHTIGUNG: JA -> "" (leer), NEIN -> "nicht".

{
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
  "VERSICHERUNG": "",
  "VER_STRASSE": "",
  "VER_ORT": "",
  "WIEDERBESCHAFFUNGSWERT": "",
  "WIEDERBESCHAFFUNGSAUFWAND": "",
  "NUTZUNGSAUSFALL": "",
  "RESTWERT": "",
  "ZUSATZKOSTEN_BEZEICHNUNG": "",
  "ZUSATZKOSTEN_BETRAG": ""
}
""",
    "Schreiben Totalschaden": """Gib NUR JSON zurück (keine Erklärungen).
Unbekannt -> "".
VORSTEUERBERECHTIGUNG: JA -> "" (leer), NEIN -> "nicht".

{
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
  "VERSICHERUNG": "",
  "VER_STRASSE": "",
  "VER_ORT": "",
  "WIEDERBESCHAFFUNGSWERTAUFWAND": ""
}
""",
}


def normalize_vorsteuer(value: str) -> str:
    v = (value or "").strip().lower()
    if v in {"ja", "yes", "y", "true"}:
        return ""
    if v in {"nein", "no", "n", "false"}:
        return "nicht"
    return value


def default_values_for_keys(keys: Set[str]) -> Dict[str, str]:
    today = date.today().strftime("%d.%m.%Y")
    frist = (datetime.now() + timedelta(days=14)).strftime("%d.%m.%Y")
    out = {}
    if "HEUTDATUM" in keys:
        out["HEUTDATUM"] = today
    if "FIRST_DATUM" in keys:
        out["FIRST_DATUM"] = frist
    if "FRIST_DATUM" in keys:
        out["FRIST_DATUM"] = frist
    return out


def parse_json_text(text: str) -> Dict[str, Any]:
    text = (text or "").strip()
    if not text:
        return {}
    return json.loads(text)


def build_context(keys: Set[str], main_json: Dict[str, Any], insurance_json: Dict[str, Any]) -> Dict[str, Any]:
    ctx = {k: "" for k in keys}

    # main json
    for k in keys:
        if k in main_json:
            ctx[k] = main_json[k]

    # defaults
    for k, v in default_values_for_keys(keys).items():
        if not str(ctx.get(k, "")).strip():
            ctx[k] = v

    # vorsteuer rule
    if "VORSTEUERBERECHTIGUNG" in ctx:
        ctx["VORSTEUERBERECHTIGUNG"] = normalize_vorsteuer(str(ctx.get("VORSTEUERBERECHTIGUNG", "")))

    # insurance override (optional)
    for k in ["VERSICHERUNG", "VER_STRASSE", "VER_ORT"]:
        if k in insurance_json and str(insurance_json[k]).strip():
            ctx[k] = insurance_json[k]

    return ctx


# -----------------------------
# UI
# -----------------------------
st.set_page_config(page_title="JSON → Word", layout="wide")
st.title("Word-Vorlage aus JSON befüllen (JSON als Eingabefeld, ohne AI)")

with st.expander("📁 Vorlagen im Repo", expanded=False):
    st.write(str(wb.VORLAGEN_DIR))
    st.write([p.name for p in wb.VORLAGEN_DIR.glob("*.docx")])

template_label = st.selectbox("Vorlage wählen", list(TEMPLATES.keys()))
tpl_name, out_prefix = TEMPLATES[template_label]

st.subheader("1) Prompt-Auswahl (sichtbar & kopierbar)")
prompt_choice = st.selectbox("Prompt wählen", list(PROMPTS.keys()), index=list(PROMPTS.keys()).index(template_label))
st.code(PROMPTS[prompt_choice], language="json")

st.subheader("2) JSON Eingabe")
default_json = "{}"
json_text = st.text_area("Haupt-JSON (Paste hier rein)", height=260, value=default_json)

st.subheader("3) Versicherung (optional separat, überschreibt)")
insurance_text = st.text_area(
    "Versicherung-JSON (optional; überschreibt VERSICHERUNG/VER_STRASSE/VER_ORT)",
    height=140,
    value='{\n  "VERSICHERUNG": "",\n  "VER_STRASSE": "",\n  "VER_ORT": ""\n}'
)

show_debug = st.toggle("Debug: Kontext anzeigen", value=True)

st.divider()

if st.button("✅ Word erzeugen", type="primary"):
    try:
        main_json = parse_json_text(json_text)
        insurance_json = parse_json_text(insurance_text)

        keys = wb.get_template_vars(tpl_name)
        ctx = build_context(keys, main_json, insurance_json)

        out_path = wb.render_word(tpl_name, ctx, out_prefix)

        st.success(f"Word erstellt: {out_path.name}")
        with open(out_path, "rb") as f:
            st.download_button(
                "⬇️ Download .docx",
                data=f,
                file_name=out_path.name,
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            )

        if show_debug:
            with st.expander("🔎 Kontext (Debug)", expanded=True):
                st.json(ctx)

        st.caption(f"Gespeichert in: {wb.OUTPUT_DIR}")

    except json.JSONDecodeError as e:
        st.error(f"JSON Fehler: {e}")
    except Exception as e:
        st.error(f"Fehler: {e}")
