from __future__ import annotations

import json
import re
from decimal import Decimal, InvalidOperation, ROUND_HALF_UP
from datetime import date, timedelta, datetime
from typing import Dict, Any, Set, Tuple

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
# Prompts
# -----------------------------
PROMPTS = {
    
    "Standard Schreiben": """Gib NUR JSON zurück (keine Erklärungen).
Unbekannt -> "". Mandanten Name mit Herr oder Frau siehe Gutachten (meist 1. Seite)
VORSTEUERBERECHTIGUNG: JA -> "" (leer), NEIN -> "nicht".

{
  "MANDANT_NACHNAME": "",
  "MANDANT_VORNAME": "",
  "MANDANT_STRASSE": "",
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
  "GUTACHTERKOSTEN": "",
  "SONSTIGE": ""
}
""",
    "130 Prozent": """Gib NUR JSON zurück (keine Erklärungen).
     Mandanten Name mit Herr oder Frau siehe Gutachten (meist 1. Seite)
Unbekannt -> "".
VORSTEUERBERECHTIGUNG: JA -> "" (leer), NEIN -> "nicht".

{
  "MANDANT_NACHNAME": "",
  "MANDANT_VORNAME": "",
  "MANDANT_STRASSE": "",
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
     Mandanten Name mit Herr oder Frau siehe Gutachten (meist 1. Seite)
Unbekannt -> "".
VORSTEUERBERECHTIGUNG: JA -> "" (leer), NEIN -> "nicht".

{
  "MANDANT_NACHNAME": "",
  "MANDANT_VORNAME": "",
  "MANDANT_STRASSE": "",
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
     Mandanten Name mit Herr oder Frau siehe Gutachten (meist 1. Seite)
Unbekannt -> "".
VORSTEUERBERECHTIGUNG: JA -> "" (leer), NEIN -> "nicht".

{
  "MANDANT_NACHNAME": "",
  "MANDANT_VORNAME": "",
  "MANDANT_STRASSE": "",
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
     Mandanten Name mit Herr oder Frau siehe Gutachten (meist 1. Seite)
Unbekannt -> "".
VORSTEUERBERECHTIGUNG: JA -> "" (leer), NEIN -> "nicht".

{
  "MANDANT_NACHNAME": "",
  "MANDANT_VORNAME": "",
  "MANDANT_STRASSE": "",
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
     Mandanten Name mit Herr oder Frau siehe Gutachten (meist 1. Seite)
Unbekannt -> "".
VORSTEUERBERECHTIGUNG: JA -> "" (leer), NEIN -> "nicht".

{
  "MANDANT_NACHNAME": "",
  "MANDANT_VORNAME": "",
  "MANDANT_STRASSE": "",
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
# Alias-Mapping
# Nur für wirklich gleichbedeutende / historisch gewachsene Felder
# -----------------------------
ALIASES: Dict[str, Tuple[str, ...]] = {
    "VERSICHERUNG": ("VRSICHERUNG",),
    "VRSICHERUNG": ("VERSICHERUNG",),

    "GUTACHTERKOSTEN": ("SACHVERST_KOSTEN",),
    "SACHVERST_KOSTEN": ("GUTACHTERKOSTEN",),

    "WIEDERBESCHAFFUNGSWERTAUFWAND": ("WIEDERBESCHAFFUNGSAUFWAND",),
    "WIEDERBESCHAFFUNGSAUFWAND": ("WIEDERBESCHAFFUNGSWERTAUFWAND",),
}


# Für automatische Kostensumme im Standard-Schreiben
SUM_CANDIDATES = (
    "WERTMINDERUNG",
    "REPARATURKOSTEN",
    "KOSTENPAUSCHALE",
    "GUTACHTERKOSTEN",
    "SACHVERST_KOSTEN",
    "SONSTIGE",
)


# -----------------------------
# Helpers
# -----------------------------
def normalize_vorsteuer(value: Any) -> str:
    v = str(value or "").strip().lower()
    if v in {"ja", "yes", "y", "true", "1"}:
        return ""
    if v in {"nein", "no", "n", "false", "0", "nicht"}:
        return "nicht"
    return str(value or "").strip()


def normalize_text(value: Any) -> str:
    if value is None:
        return ""
    return str(value).strip()


def parse_money(value: Any) -> Decimal | None:
    """
    Akzeptiert z. B.
    - 1234,56
    - 1.234,56
    - 1234.56
    - 1,234.56
    - 915,71 €
    """
    s = normalize_text(value)
    if not s:
        return None

    s = s.replace("€", "").replace("EUR", "").replace("eur", "").strip()
    s = re.sub(r"[^\d,.\-]", "", s)

    if not s:
        return None

    # Falls sowohl Punkt als auch Komma vorkommen:
    # Annahme: das letzte Trennzeichen ist das Dezimaltrennzeichen.
    if "," in s and "." in s:
        last_comma = s.rfind(",")
        last_dot = s.rfind(".")
        if last_comma > last_dot:
            # deutsch: 1.234,56
            s = s.replace(".", "").replace(",", ".")
        else:
            # englisch: 1,234.56
            s = s.replace(",", "")
    elif "," in s:
        # nur Komma -> deutsches Dezimaltrennzeichen
        s = s.replace(".", "").replace(",", ".")
    else:
        # nur Punkt oder gar nichts
        # mehrere Punkte = Tausenderpunkte
        if s.count(".") > 1:
            s = s.replace(".", "")

    try:
        return Decimal(s)
    except InvalidOperation:
        return None


def format_money(value: Decimal | None) -> str:
    """
    Ausgabe deutsch: 1.234,56
    """
    if value is None:
        return ""
    q = value.quantize(Decimal("0.01"), rounding=ROUND_HALF_UP)
    raw = f"{q:,.2f}"  # 1,234.56
    return raw.replace(",", "X").replace(".", ",").replace("X", ".")


def default_values_for_keys(keys: Set[str]) -> Dict[str, str]:
    today = date.today().strftime("%d.%m.%Y")
    frist = (datetime.now() + timedelta(days=14)).strftime("%d.%m.%Y")

    out: Dict[str, str] = {}
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

    # Toleranter Parser für kopierte AI-Antworten:
    # entfernt ```json ... ```
    text = re.sub(r"^\s*```(?:json)?\s*", "", text, flags=re.IGNORECASE)
    text = re.sub(r"\s*```\s*$", "", text)

    return json.loads(text)


def get_value_for_key(target_key: str, main_json: Dict[str, Any]) -> Any:
    """
    1) exakter Key
    2) Alias
    3) sonst leer
    """
    if target_key in main_json:
        return main_json[target_key]

    for alias in ALIASES.get(target_key, ()):
        if alias in main_json:
            return main_json[alias]

    return ""


def compute_kostensumme(main_json: Dict[str, Any], ctx: Dict[str, Any]) -> str:
    """
    Berechnet KOSTENSUMME_X aus Standard-Schadenpositionen.
    Doppelte Zählung von GUTACHTERKOSTEN/SACHVERST_KOSTEN wird vermieden.
    """
    total = Decimal("0.00")
    found_any = False
    already_counted = set()

    for key in SUM_CANDIDATES:
        value = ctx.get(key, "")
        parsed = parse_money(value)
        if parsed is None:
            continue

        canonical = "GUTACHTERKOSTEN" if key in {"GUTACHTERKOSTEN", "SACHVERST_KOSTEN"} else key
        if canonical in already_counted:
            continue

        already_counted.add(canonical)
        total += parsed
        found_any = True

    # Falls jemand KOSTENSUMME/KOSTENSUMME_X direkt liefert, hat das Vorrang
    direct_sum = (
        normalize_text(main_json.get("KOSTENSUMME_X", ""))
        or normalize_text(main_json.get("KOSTENSUMME", ""))
        or normalize_text(main_json.get("KOSTEN_SUMME", ""))
    )
    if direct_sum:
        parsed_direct = parse_money(direct_sum)
        if parsed_direct is not None:
            return format_money(parsed_direct)

    return format_money(total) if found_any else ""


def build_context(keys: Set[str], main_json: Dict[str, Any]) -> Dict[str, Any]:
    ctx = {k: "" for k in keys}

    # 1) Exakte Keys / Alias-Mapping
    for key in keys:
        ctx[key] = normalize_text(get_value_for_key(key, main_json))

    # 2) Default-Werte
    for key, value in default_values_for_keys(keys).items():
        if not normalize_text(ctx.get(key, "")):
            ctx[key] = value

    # 3) Vorsteuer-Regel
    if "VORSTEUERBERECHTIGUNG" in ctx:
        ctx["VORSTEUERBERECHTIGUNG"] = normalize_vorsteuer(ctx["VORSTEUERBERECHTIGUNG"])

    # 4) KOSTENSUMME_X berechnen, falls Vorlage das Feld nutzt
    if "KOSTENSUMME_X" in ctx and not normalize_text(ctx["KOSTENSUMME_X"]):
        ctx["KOSTENSUMME_X"] = compute_kostensumme(main_json, ctx)

    # 5) Optional auch KOSTENSUMME füllen, falls eine andere Vorlage das nutzt
    if "KOSTENSUMME" in ctx and not normalize_text(ctx["KOSTENSUMME"]):
        ctx["KOSTENSUMME"] = compute_kostensumme(main_json, ctx)

    return ctx


def analyze_context(keys: Set[str], main_json: Dict[str, Any], ctx: Dict[str, Any]) -> Dict[str, Any]:
    filled_keys = sorted([k for k, v in ctx.items() if normalize_text(v)])
    empty_keys = sorted([k for k, v in ctx.items() if not normalize_text(v)])
    json_unused_keys = sorted([k for k in main_json.keys() if k not in keys and k not in flatten_aliases()])
    return {
        "template_keys_count": len(keys),
        "filled_keys_count": len(filled_keys),
        "empty_keys_count": len(empty_keys),
        "filled_keys": filled_keys,
        "empty_keys": empty_keys,
        "json_unused_keys": json_unused_keys,
    }


def flatten_aliases() -> set[str]:
    out = set()
    for key, aliases in ALIASES.items():
        out.add(key)
        out.update(aliases)
    return out


def copy_to_clipboard_button(text: str, button_label: str = "📋 Prompt kopieren"):
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
st.title("Word-Vorlage aus JSON befüllen")

with st.expander("📁 Vorlagen im Repo", expanded=False):
    st.write(str(wb.VORLAGEN_DIR))
    st.write(wb.list_docx_templates())

template_label = st.selectbox("Vorlage wählen", list(TEMPLATES.keys()))
tpl_name, out_prefix = TEMPLATES[template_label]

st.subheader("1) Prompt kopieren (Aktenzeichen und Schadenshergang in Json Feld Umändern!)")
prompt_choice = st.selectbox(
    "Prompt wählen",
    list(PROMPTS.keys()),
    index=list(PROMPTS.keys()).index(template_label),
)
copy_to_clipboard_button(PROMPTS[prompt_choice], "📋 Prompt kopieren")

st.subheader("2) JSON Eingabe")
json_text = st.text_area("JSON hier einfügen", height=320, value="{}")

col1, col2, col3 = st.columns(3)
with col1:
    show_debug = st.toggle("Debug: Kontext anzeigen", value=True)
with col2:
    show_template_vars = st.toggle("Debug: Template-Variablen", value=False)
with col3:
    show_analysis = st.toggle("Debug: Feldanalyse", value=True)

st.divider()

if st.button("✅ Word erzeugen", type="primary"):
    try:
        main_json = parse_json_text(json_text)
        if not isinstance(main_json, dict):
            raise ValueError("Das eingefügte JSON muss ein Objekt sein, also mit { ... } beginnen.")

        keys = wb.get_template_vars(tpl_name)
        ctx = build_context(keys, main_json)
        analysis = analyze_context(keys, main_json, ctx)

        out_path = wb.render_word(tpl_name, ctx, out_prefix)

        st.success(f"Word erstellt: {out_path.name}")

        with open(out_path, "rb") as f:
            st.download_button(
                "⬇️ Download .docx",
                data=f,
                file_name=out_path.name,
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            )

        if show_template_vars:
            with st.expander("🧩 Variablen aus Vorlage", expanded=False):
                st.json(sorted(list(keys)))

        if show_debug:
            with st.expander("🔎 Kontext", expanded=True):
                st.json(ctx)

        if show_analysis:
            with st.expander("🧪 Feldanalyse", expanded=False):
                st.json(analysis)

        st.caption(f"Gespeichert in: {wb.OUTPUT_DIR}")

    except json.JSONDecodeError as e:
        st.error(f"JSON Fehler: {e}")
    except FileNotFoundError as e:
        st.error(str(e))
    except Exception as e:
        st.error(f"Fehler: {type(e).__name__}: {e}")




