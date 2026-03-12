from pathlib import Path
from docxtpl import DocxTemplate
from datetime import date, timedelta, datetime

# entnimmt word vorlagen aus vorlagen folder
BASE_DIR = Path(__file__).resolve().parent
VORLAGEN_DIR = BASE_DIR / "Vorlagen_word_file"
OUTPUT_DIR = BASE_DIR / "Output_wordvorlage"
OUTPUT_DIR.mkdir(exist_ok=True)

# sorgt dafür das kein problem beim speichern word datei
def safe_filename(s: str) -> str:
    return "".join(c for c in (s or "").strip() if c.isalnum() or c in ("-", "_"))

# zahlen werden aus zahlen (eingabe) form in EURO umgewandelt
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

def euro_format(value: float) -> str:
    return f"{value:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")

# Datum für heute eingestellt
heute = datetime.now()
datumheut = heute.strftime("%d.%m.%Y")
first_datum_14_zukunft = datetime.now() + timedelta(days=14)

def standardabfrage(data: dict) -> dict:
    global first_datum_14_zukunft
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
        "UNFALL_DATUM": data.get("UNFALL_DATUM", "")
    }
    return context_standard

# speichert word vorlage mit nachnamen kennzeichen
def save_word_bezeichg(tpl_name: str, context_standard: dict, out_prefix: str) -> Path:
    tpl_path = VORLAGEN_DIR / tpl_name
    if not tpl_path.exists():
        raise FileNotFoundError(f"Vorlage nicht gefunden: {tpl_path}")

    tpl = DocxTemplate(str(tpl_path))
    tpl.render(context_standard)

    nachname = safe_filename(context_standard.get("MANDANT_NACHNAME", "Unbekannt"))
    out_name = f"{out_prefix}_{nachname}_{datetime.now().strftime('%Y%m%d_%H%M')}.docx"
    out_path = OUTPUT_DIR / out_name

    tpl.save(str(out_path))
    return out_path

# -------- Vorlagen-Funktionen (Namen beibehalten) --------

def vorlage_schreiben(data: dict) -> Path:
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
        "KOSTENSUMME_X": euro_format(kostensumme),
        "WERTMINDERUNG": wert_minderung,
        "REPARATURKOSTEN": reparatur_kosten,
        "KOSTENPAUSCHALE": kosten_pauschale,
        "SACHVERST_KOSTEN": sachverst_kosten,
    })

    return save_word_bezeichg("vorlage_schreiben.docx", context_standard, "Standard_schreiben")


def vorlage_totalschaden_konkret(data: dict) -> Path:
    context_standard = standardabfrage(data)

    wbw_txt = data.get("WIEDERBESCHAFFUNGSWERT", "")
    wba_txt = data.get("WIEDERBESCHAFFUNGSAUFWAND", "")
    rw_txt = data.get("RESTWERT", "")
    mwst_txt = data.get("ERSATZBESCHAFFUNG_MWST", "")

    zus_bez = data.get("ZUSATZKOSTEN_BEZEICHNUNG", "")
    zus_bet = data.get("ZUSATZKOSTEN_BETRAG", "")

    summe = (
        euro_to_float(wbw_txt)
        + euro_to_float(wba_txt)
        + euro_to_float(rw_txt)
        + euro_to_float(mwst_txt)
        + euro_to_float(zus_bet)
    )

    context_standard.update({
        "WIEDERBESCHAFFUNGSWERT": wbw_txt,
        "WIEDERBESCHAFFUNGSAUFWAND": wba_txt,
        "RESTWERT": rw_txt,
        "ERSATZBESCHAFFUNG_MWST": mwst_txt,
        "ZUSATZKOSTEN_BEZEICHNUNG": zus_bez,
        "ZUSATZKOSTEN_BETRAG": zus_bet,
        "KOSTENSUMME_X": euro_format(summe),
    })

    return save_word_bezeichg("vorlage_totalschaden_konkret-1.docx", context_standard, "totalschaden_konkret")


def vorlage_130_prozent(data: dict) -> Path:
    context_standard = standardabfrage(data)

    rep_txt = data.get("REPARATURKOSTEN", "")
    mwst_txt = data.get("MWST_BETRAG", "")
    na_txt = data.get("NUTZUNGSAUSFALL", "")
    wbw_txt = data.get("WIEDERBESCHAFFUNGSWERT", "")
    rw_txt = data.get("RESTWERT", "")

    zus_bez = data.get("ZUSATZKOSTEN_BEZEICHNUNG", "")
    zus_bet = data.get("ZUSATZKOSTEN_BETRAG", "")

    summe = (
        euro_to_float(rep_txt)
        + euro_to_float(mwst_txt)
        + euro_to_float(na_txt)
        + euro_to_float(wbw_txt)
        + euro_to_float(rw_txt)
        + euro_to_float(zus_bet)
    )

    context_standard.update({
        "REPARATURKOSTEN": rep_txt,
        "MWST_BETRAG": mwst_txt,
        "NUTZUNGSAUSFALL": na_txt,
        "WIEDERBESCHAFFUNGSWERT": wbw_txt,
        "RESTWERT": rw_txt,
        "ZUSATZKOSTEN_BEZEICHNUNG": zus_bez,
        "ZUSATZKOSTEN_BETRAG": zus_bet,
        "KOSTENSUMME_X": euro_format(summe),
    })

    return save_word_bezeichg("vorlage_130_prozent.docx", context_standard, "130_prozent")


def vorlage_totalschaden_konkret_unter_wbw(data: dict) -> Path:
    context_standard = standardabfrage(data)

    rep_txt = data.get("REPARATURKOSTEN", "")
    mwst_txt = data.get("MWST_BETRAG", "")
    wm_txt = data.get("WERTMINDERUNG", "")
    na_txt = data.get("NUTZUNGSAUSFALL", "")

    zus_bez = data.get("ZUSATZKOSTEN_BEZEICHNUNG", "")
    zus_bet = data.get("ZUSATZKOSTEN_BETRAG", "")

    summe = (
        euro_to_float(rep_txt)
        + euro_to_float(mwst_txt)
        + euro_to_float(wm_txt)
        + euro_to_float(na_txt)
        + euro_to_float(zus_bet)
    )

    context_standard.update({
        "REPARATURKOSTEN": rep_txt,
        "MWST_BETRAG": mwst_txt,
        "WERTMINDERUNG": wm_txt,
        "NUTZUNGSAUSFALL": na_txt,
        "ZUSATZKOSTEN_BEZEICHNUNG": zus_bez,
        "ZUSATZKOSTEN_BETRAG": zus_bet,
        "KOSTENSUMME_X": euro_format(summe),
    })

    return save_word_bezeichg("vorlage_totalschaden_konkret_unter_wbw-1.docx", context_standard, "ts_konkret_unter_wbw")
