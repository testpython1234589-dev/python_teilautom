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
# Prompts (werden NICHT angezeigt, nur kopiert)
# Versicherung ist IM selben JSON enthalten.
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


# -----------------------------
# Helpers
# -----------------------------
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
    # Falls deine Vorlage FRIST_DATUM statt FIRST_DATUM nutzt:
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


def build_context(keys: Set[str], main_json: Dict[str, Any]) -> Dict[str, Any]:
    ctx = {k: "" for k in keys}

    # JSON -> Context (nur Keys, die die Vorlage kennt)
    for k in keys:
        if k in main_json:
            ctx[k] = main_json[k]

    # Defaults (Datum/Frist)
    for k, v in default_values_for_keys(keys).items():
        if not str(ctx.get(k, "")).strip():
            ctx[k] = v

    # Vorsteuer-Regel
    if "VORSTEUERBERECHTIGUNG" in ctx:
        ctx["VORSTEUERBERECHTIGUNG"] = normalize_vorsteuer(str(ctx.get("VORSTEUERBERECHTIGUNG", "")))

    return ctx


def copy_to_clipboard_button(text: str, button_label: str = "📋 Prompt kopieren"):
    # Prompt NICHT anzeigen, nur kopieren (JS Clipboard)
    safe = (
        text.replace("&", "&amp;")
            .replace("<", "&lt;")
            .replace(">", "&gt;")
    )
    st.components.v1.html(
        f"""
        <button id="copyBtn" style="
            padding:10px 14px;border-radius:10px;border:1px solid #333;
            background:#111;color:#fff;cursor:pointer;font-size:14px;">
            {button_label}
        </button>
        <textarea id="copyText" style="position:fixed;left:-10000px;top:-10000px;">{safe}</textarea>
        <script>
          const btn = document.getElementById('copyBtn');
          const txt = document.getElementById('copyText');
          btn.addEventListener('click', async () => {{
            try {{
              await navigator.clipboard.writeText(txt.value);
              btn.innerText = "✅ Kopiert!";
              setTimeout(() => btn.innerText = "{button_label}", 1200);
            }} catch(e) {{
              txt.select();
              document.execCommand('copy');
              btn.innerText = "✅ Kopiert!";
              setTimeout(() => btn.innerText = "{button_label}", 1200);
            }}
          }});
        </script>
        """,
        height=60,
    )


# -----------------------------
# UI
# -----------------------------
st.set_page_config(page_title="JSON → Word", layout="wide")
st.title("Word-Vorlage aus JSON befüllen (Prompt nur kopierbar, kein AI)")

with st.expander("📁 Vorlagen im Repo", expanded=False):
    st.write(str(wb.VORLAGEN_DIR))
    st.write([p.name for p in wb.VORLAGEN_DIR.glob("*.docx")])

template_label = st.selectbox("Vorlage wählen", list(TEMPLATES.keys()))
tpl_name, out_prefix = TEMPLATES[template_label]

st.subheader("1) Prompt kopieren (ohne Anzeige)")
prompt_choice = st.selectbox("Prompt wählen", list(PROMPTS.keys()), index=list(PROMPTS.keys()).index(template_label))
copy_to_clipboard_button(PROMPTS[prompt_choice], "📋 Prompt kopieren")

st.subheader("2) JSON Eingabe (enthält auch Versicherung)")
json_text = st.text_area("JSON hier einfügen", height=320, value="{}")

show_debug = st.toggle("Debug: Kontext anzeigen", value=True)

st.divider()

if st.button("✅ Word erzeugen", type="primary"):
    try:
        main_json = parse_json_text(json_text)
        keys = wb.get_template_vars(tpl_name)
        ctx = build_context(keys, main_json)

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
