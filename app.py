from pathlib import Path
from docxtpl import DocxTemplate
from datetime import date, timedelta, datetime
from flask import Flask, request, jsonify, send_from_directory

app = Flask(__name__)

BASE_DIR = Path(__file__).resolve().parent
VORLAGEN_DIR = BASE_DIR / "Vorlagen_word_file"
OUTPUT_DIR = BASE_DIR / "Output_wordvorlage"
OUTPUT_DIR.mkdir(exist_ok=True)

def safe_filename(s: str) -> str:
    return "".join(c for c in (s or "").strip() if c.isalnum() or c in ("-", "_"))

def euro_to_float(t: str) -> float:
    s = (t or "").strip()
    if not s:
        return 0.0
    s = s.replace("€", "").replace("EUR", "").strip()
    s = s.replace(" ", "")
    if "." in s and "," in s:
        s = s.replace(".", "").replace(",", ".")
    elif "," in s:
        s = s.replace(",", ".")
    return float(s)

heute = datetime.now()
first_datum_14_zukunft = datetime.now() + timedelta(days=14)

def standardabfrage(data: dict):
    datum = date.today().strftime("%d.%m.%Y")
    context_standard = {
        "MANDANT_NACHNAME": data.get("MANDANT_NACHNAME", ""),
        "MANDANT_VORNAME": data.get("MANDANT_VORNAME", ""),
        "MANDANT_PLZ_ORT": data.get("MANDANT_PLZ_ORT", ""),
        "AKTENZEICHEN": data.get("AKTENZEICHEN", ""),
        "HEUTDATUM": datum,
        "FIRST_DATUM": first_datum_14_zukunft.strftime("%d.%m.%Y"),
        "SCHADENHERGANG": data.get("SCHADENHERGANG", ""),
        "UNFALLE_STRASSE": data.get("UNFALLE_STRASSE", ""),
        "FAHRZEUGTYP": data.get("FAHRZEUGTYP", ""),
        "KENNZEICHEN": data.get("KENNZEICHEN", ""),
        "VORSTEUERBERECHTIGUNG": data.get("VORSTEUERBERECHTIGUNG", ""),
        "UNFALL_DATUM": data.get("UNFALL_DATUM", ""),
    }
    return context_standard

def save_word_bezeichg(tpl_name: str, context_standard: dict, out_prefix: str):
    tpl_path = VORLAGEN_DIR / tpl_name
    tpl = DocxTemplate(str(tpl_path))
    tpl.render(context_standard)

    nachname = safe_filename(context_standard.get("MANDANT_NACHNAME", "Unbekannt"))
    out_name = f"{out_prefix}_{nachname}_{datetime.now().strftime('%Y%m%d_%H%M')}.docx"
    out_path = OUTPUT_DIR / out_name
    tpl.save(str(out_path))
    return out_name

def vorlage_schreiben(data: dict):
    context_standard = standardabfrage(data)

    wert_minderung = data.get("WERTMINDERUNG", "")
    reparatur_kosten = data.get("REPARATURKOSTEN", "")
    kosten_pauschale = data.get("KOSTENPAUSCHALE", "")
    sachverst_kosten = data.get("SACHVERST_KOSTEN", "")

    kostensumme = (
        euro_to_float(wert_minderung)
        + euro_to_float(reparatur_kosten)
        + euro_to_float(kosten_pauschale)
        + euro_to_float(sachverst_kosten)
    )

    context_standard.update({
        "KOSTENSUMME_X": kostensumme,
        "WERTMINDERUNG": wert_minderung,
        "REPARATURKOSTEN": reparatur_kosten,
        "KOSTENPAUSCHALE": kosten_pauschale,
        "SACHVERST_KOSTEN": sachverst_kosten,
    })

    return save_word_bezeichg("vorlage_schreiben.docx", context_standard, "Standard_schreiben")

# --- HTTP Routes ---

@app.post("/generate")
def generate():
    data = request.get_json(force=True)
    template_key = data.get("template")

    if template_key == "standard":
        filename = vorlage_schreiben(data)
        return jsonify({"ok": True, "filename": filename, "download_url": f"/download/{filename}"})

    return jsonify({"ok": False, "error": "Unbekannte Vorlage"}), 400

@app.get("/download/<path:filename>")
def download(filename):
    return send_from_directory(OUTPUT_DIR, filename, as_attachment=True)

@app.get("/")
def root():
    # Frontend als statische Datei ausliefern (siehe unten)
    return send_from_directory(BASE_DIR, "index.html")

if __name__ == "__main__":
    app.run(host="127.0.0.1", port=5000, debug=True)