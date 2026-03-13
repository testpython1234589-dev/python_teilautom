from __future__ import annotations

import json
import os
import re
from datetime import date, timedelta, datetime
from typing import Dict, Any, Set

import streamlit as st
import fitz  # PyMuPDF
from openai import OpenAI

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
# Prompts (werden angezeigt & kopierbar via st.code)
# Versicherung soll aus JSON kommen -> Prompt weist darauf hin.
# -----------------------------
PROMPTS = {
    "Standard Schreiben": """Gib NUR gültiges JSON zurück. Keine Erklärungen, kein Markdown.
Unbekannt -> "".
Beträge im deutschen Format (z.B. 1.234,56).
VORSTEUERBERECHTIGUNG-Regel: wenn im Text "ja" -> "" (leer), wenn "nein" -> "nicht".
WICHTIG: Versicherungsdaten (VERSICHERUNG, VER_STRASSE, VER_ORT) NICHT ausfüllen (kommen aus separater JSON-Datei).

JSON:
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
  "SACHVERST_KOSTEN": "",
  "VERSICHERUNG": "",
  "VER_STRASSE": "",
  "VER_ORT": ""
}
""",
    "130 Prozent": """Gib NUR gültiges JSON zurück. Keine Erklärungen, kein Markdown.
Unbekannt -> "".
Beträge im deutschen Format (z.B. 1.234,56).
VORSTEUERBERECHTIGUNG-Regel: "ja" -> "" (leer), "nein" -> "nicht".
WICHTIG: Versicherungsdaten NICHT ausfüllen (kommen aus separater JSON-Datei).

JSON:
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
  "ZUSATZKOSTEN_BETRAG": "",
  "VERSICHERUNG": "",
  "VER_STRASSE": "",
  "VER_ORT": ""
}
""",
    "Totalschaden (konkret)": """Gib NUR gültiges JSON zurück. Keine Erklärungen, kein Markdown.
Unbekannt -> "".
Beträge im deutschen Format (z.B. 1.234,56).
VORSTEUERBERECHTIGUNG-Regel: "ja" -> "" (leer), "nein" -> "nicht".
WICHTIG: Versicherungsdaten NICHT ausfüllen (kommen aus separater JSON-Datei).

JSON:
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
  "ZUSATZKOSTEN_BETRAG": "",
  "VERSICHERUNG": "",
  "VER_STRASSE": "",
  "VER_ORT": ""
}
""",
    "Konkret unter WBW": """Gib NUR gültiges JSON zurück. Keine Erklärungen, kein Markdown.
Unbekannt -> "".
Beträge im deutschen Format (z.B. 1.234,56).
VORSTEUERBERECHTIGUNG-Regel: "ja" -> "" (leer), "nein" -> "nicht".
WICHTIG: Versicherungsdaten NICHT ausfüllen (kommen aus separater JSON-Datei).

JSON:
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
  "ZUSATZKOSTEN_BETRAG": "",
  "VERSICHERUNG": "",
  "VER_STRASSE": "",
  "VER_ORT": ""
}
""",
    "Totalschaden (fiktiv)": """Gib NUR gültiges JSON zurück. Keine Erklärungen, kein Markdown.
Unbekannt -> "".
Beträge im deutschen Format (z.B. 1.234,56).
VORSTEUERBERECHTIGUNG-Regel: "ja" -> "" (leer), "nein" -> "nicht".
WICHTIG: Versicherungsdaten NICHT ausfüllen (kommen aus separater JSON-Datei).

JSON:
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
  "ZUSATZKOSTEN_BETRAG": "",
  "VERSICHERUNG": "",
  "VER_STRASSE": "",
  "VER_ORT": ""
}
""",
    "Schreiben Totalschaden": """Gib NUR gültiges JSON zurück. Keine Erklärungen, kein Markdown.
Unbekannt -> "".
Beträge im deutschen Format (z.B. 1.234,56).
VORSTEUERBERECHTIGUNG-Regel: "ja" -> "" (leer), "nein" -> "nicht".
WICHTIG: Versicherungsdaten NICHT ausfüllen (kommen aus separater JSON-Datei).

JSON:
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
  "WIEDERBESCHAFFUNGSWERTAUFWAND": "",
  "VERSICHERUNG": "",
  "VER_STRASSE": "",
  "VER_ORT": ""
}
""",
}


# -----------------------------
# Helpers
# -----------------------------
def pdf_bytes_to_text(pdf_bytes: bytes) -> str:
    doc = fitz.open(stream=pdf_bytes, filetype="pdf")
    parts = []
    for i in range(doc.page_count):
        parts.append(doc.load_page(i).get_text("text"))
    return "\n".join(parts)


def normalize_vorsteuer(value: str) -> str:
    v = (value or "").strip().lower()
    if v in {"ja", "yes", "y", "true"}:
        return ""
    if v in {"nein", "no", "n", "false"}:
        return "nicht"
    return value


def standard_defaults(keys: Set[str]) -> Dict[str, str]:
    today = date.today().strftime("%d.%m.%Y")
    frist = (datetime.now() + timedelta(days=14)).strftime("%d.%m.%Y")
    out = {}
    if "HEUTDATUM" in keys:
        out["HEUTDATUM"] = today
    # Falls deine Vorlage FRIST_DATUM nutzt -> hier anpassen:
    if "FIRST_DATUM" in keys:
        out["FIRST_DATUM"] = frist
    return out


def build_json_schema(keys: Set[str]) -> str:
    keys_sorted = sorted(keys)
    lines = []
    for i, k in enumerate(keys_sorted):
        comma = "," if i < len(keys_sorted) - 1 else ""
        lines.append(f'  "{k}": ""{comma}')
    return "{\n" + "\n".join(lines) + "\n}"


def build_extraction_prompt(keys: Set[str], pdf_text: str) -> str:
    """
    Prompt, den wir intern für die KI verwenden (nicht die UI-Prompts).
    Wir sagen explizit: Versicherung aus JSON, daher leer lassen.
    """
    rules = (
        "Gib NUR JSON zurück. Keine Erklärungen.\n"
        "JSON muss ALLE Keys enthalten.\n"
        "Unbekannt -> \"\".\n"
        "VORSTEUERBERECHTIGUNG: ja -> \"\" (leer), nein -> \"nicht\".\n"
        "Beträge möglichst deutsches Format (z.B. 1.234,56).\n"
        "WICHTIG: Versicherungsdaten (VERSICHERUNG, VER_STRASSE, VER_ORT) NICHT ausfüllen -> \"\".\n"
    )
    schema = build_json_schema(keys)
    return f"{rules}\nJSON-SCHEMA:\n{schema}\n\nPDF_TEXT:\n<<<\n{pdf_text}\n>>>"


def openai_extract_json(keys: Set[str], pdf_text: str, model: str) -> Dict[str, Any]:
    api_key = os.getenv("OPENAI_API_KEY", "").strip()
    if not api_key:
        raise RuntimeError("OPENAI_API_KEY fehlt (Streamlit Secrets oder Umgebungsvariable).")

    client = OpenAI()
    prompt = build_extraction_prompt(keys, pdf_text)

    resp = client.responses.create(model=model, input=prompt)
    text = getattr(resp, "output_text", None)
    if not text:
        raise RuntimeError("OpenAI Antwort enthält keinen output_text.")

    # robustes JSON parsing
    try:
        data = json.loads(text)
    except json.JSONDecodeError:
        m = re.search(r"\{.*\}", text, re.DOTALL)
        if not m:
            raise
        data = json.loads(m.group(0))

    for k in keys:
        data.setdefault(k, "")

    if "VORSTEUERBERECHTIGUNG" in data:
        data["VORSTEUERBERECHTIGUNG"] = normalize_vorsteuer(str(data.get("VORSTEUERBERECHTIGUNG", "")))

    return data


def load_insurance_json(uploaded_json_file) -> Dict[str, Any]:
    """
    Erwartet JSON-Datei mit z.B.:
    {
      "VERSICHERUNG": "Allianz ...",
      "VER_STRASSE": "Musterstr. 1",
      "VER_ORT": "10115 Berlin"
    }
    """
    if uploaded_json_file is None:
        return {}

    raw = uploaded_json_file.read()
    try:
        data = json.loads(raw.decode("utf-8"))
    except Exception:
        # Fallback: manchmal kommt schon str
        data = json.loads(raw)

    # Nur die Versicherungskeys übernehmen
    out = {}
    for k in ["VERSICHERUNG", "VER_STRASSE", "VER_ORT"]:
        if k in data and str(data[k]).strip():
            out[k] = data[k]
    return out


# -----------------------------
# UI
# -----------------------------
st.set_page_config(page_title="PDF → Word (KI + Versicherungs-JSON)", layout="wide")
st.title("PDF-Gutachten → Word-Vorlage (KI extrahiert, Versicherung aus JSON)")

with st.expander("📁 Vorlagen im Repo", expanded=False):
    st.write(str(wb.VORLAGEN_DIR))
    st.write([p.name for p in wb.VORLAGEN_DIR.glob("*.docx")])

template_label = st.selectbox("Vorlage wählen", list(TEMPLATES.keys()))
tpl_name, out_prefix = TEMPLATES[template_label]

st.subheader("1) Dateien hochladen")
pdf_file = st.file_uploader("Gutachten als PDF hochladen", type=["pdf"])
insurance_json_file = st.file_uploader("Versicherung als JSON hochladen", type=["json"])

st.subheader("2) Prompt-Auswahl (sichtbar & kopierbar)")
prompt_choice = st.selectbox("Prompt wählen", list(PROMPTS.keys()), index=list(PROMPTS.keys()).index(template_label))
st.code(PROMPTS[prompt_choice], language="text")  # zeigt Prompt + Copy-Icon

st.subheader("3) KI Einstellungen")
model = st.selectbox("OpenAI Modell", ["gpt-4o-mini", "gpt-4.1-mini", "gpt-4o"], index=0)
show_debug = st.toggle("Debug: extrahierte Werte anzeigen", value=True)

st.divider()

disabled = (pdf_file is None)
if st.button("✅ PDF analysieren & Word erzeugen", type="primary", disabled=disabled):
    try:
        pdf_bytes = pdf_file.read()
        pdf_text = pdf_bytes_to_text(pdf_bytes)

        # Keys aus Template (Backend ohne KI)
        keys = wb.get_template_vars(tpl_name)

        # KI extrahiert alles (ohne Versicherung)
        extracted = openai_extract_json(keys, pdf_text, model=model)

        # Defaults für Datum/Frist
        extracted.update({k: v for k, v in standard_defaults(keys).items() if not str(extracted.get(k, "")).strip()})

        # Versicherung aus JSON überschreibt/ergänzt
        ins = load_insurance_json(insurance_json_file)
        for k, v in ins.items():
            extracted[k] = v  # bewusst überschreiben

        # Rendern
        out_path = wb.render_word(tpl_name, extracted, out_prefix)

        st.success(f"Word erstellt: {out_path.name}")
        with open(out_path, "rb") as f:
            st.download_button(
                "⬇️ Download .docx",
                data=f,
                file_name=out_path.name,
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            )

        if show_debug:
            with st.expander("🔎 Kontext (extrahierte Werte)", expanded=True):
                st.json(extracted)

        st.caption(f"Gespeichert in: {wb.OUTPUT_DIR}")

    except Exception as e:
        st.error(f"Fehler: {e}")
        st.info("Tipp: OPENAI_API_KEY in Streamlit Secrets setzen. Und prüfen, ob PDF echten Text enthält (kein Scan ohne OCR).")
