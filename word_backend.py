from pathlib import Path
from docxtpl import DocxTemplate
from datetime import date, timedelta, datetime

# Repo-Root ist Vorlagen-Ordner (wie bei dir)
BASE_DIR = Path(__file__).resolve().parent
VORLAGEN_DIR = BASE_DIR

OUTPUT_DIR = BASE_DIR / "Output_wordvorlage"
OUTPUT_DIR.mkdir(exist_ok=True)


# -----------------------------
# Helpers
# -----------------------------
def safe_filename(s: str) -> str:
    return "".join(c for c in (s or "").strip() if c.isalnum() or c in ("-", "_"))


def euro_to_float(t: str) -> float:
    """
    Akzeptiert z.B.:
    '1000' | '1000,5' | '1000.50' | '1.000,50' | '€ 1 000,50' | ''
    Rückgabe: Euro als float ('' -> 0.0)
    """
    s = (t or "").strip()
    if not s:
        return 0.0

    s = s.replace("€", "").replace("EUR", "").strip()
    s = s.replace(" ", "")

    # Wenn sowohl '.' als auch ',' vorkommen -> '.' Tausender, ',' Dezimal
    if "." in s and "," in s:
        s = s.replace(".", "").replace(",", ".")
    elif "," in s:
        s = s.replace(",", ".")

    return float(s)


def euro_format(value: float) -> str:
    # 1234.5 -> '1.234,50'
    return f"{value:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")


# Datum/Frist (14 Tage)
first_datum_14_zukunft = datetime.now() + timedelta(days=14)


# -----------------------------
# Context builder (Standard)
# -----------------------------
def standardabfrage(data: dict) -> dict:
    datum = date.today().strftime("%d.%m.%Y")
    return {
        "MANDANT_NACHNAME": data.get("MANDANT_NACHNAME", ""),
        "MANDANT_VORNAME": data.get("MANDANT_VORNAME", ""),

        # Mandantenadresse (wie bisher)
        "UNFALLE_STRASSE": data.get("UNFALLE_STRASSE", ""),
        "MANDANT_PLZ_ORT": data.get("MANDANT_PLZ_ORT", ""),

        # Unfall-Daten (NEU ergänzt)
        "UNFALL_ORT": data.get("UNFALL_ORT", ""),
        "UNFALL_STRASSE": data.get("UNFALL_STRASSE", ""),

        "UNFALL_DATUM": data.get("UNFALL_DATUM", ""),
        "AKTENZEICHEN": data.get("AKTENZEICHEN", ""),  # besser String
        "FAHRZEUGTYP": data.get("FAHRZEUGTYP", ""),
        "KENNZEICHEN": data.get("KENNZEICHEN", ""),
        "VORSTEUERBERECHTIGUNG": data.get("VORSTEUERBERECHTIGUNG", ""),
        "SCHADENHERGANG": data.get("SCHADENHERGANG", ""),
        "SCHADENSNUMMER": data.get("SCHADENSNUMMER", ""),

        "HEUTDATUM": datum,

        # Achtung: In deinen Vorlagen evtl. FRIST_DATUM statt FIRST_DATUM.
        # Passe den Key an deine echten Platzhalter an, falls nötig.
        "FIRST_DATUM": first_datum_14_zukunft.strftime("%d.%m.%Y"),
    }


# -----------------------------
# Save renderer
# -----------------------------
def save_word_bezeichg(tpl_name: str, context_standard: dict, out_prefix: str) -> Path:
    tpl_path = VORLAGEN_DIR / tpl_name
    if not tpl_path.exists():
        raise FileNotFoundError(
            f"Vorlage nicht gefunden: {tpl_path}\n"
            f"Vorhandene .docx im Repo: {[p.name for p in VORLAGEN_DIR.glob('*.docx')]}"
        )

    tpl = DocxTemplate(str(tpl_path))
    tpl.render(context_standard)

    nachname = safe_filename(context_standard.get("MANDANT_NACHNAME", "Unbekannt"))
    out_name = f"{out_prefix}_{nachname}_{datetime.now().strftime('%Y%m%d_%H%M')}.docx"
    out_path = OUTPUT_DIR / out_name
    tpl.save(str(out_path))
    return out_path


# -----------------------------
# Vorlagen-Funktionen
# -----------------------------
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
        "WERTMINDERUNG": wert_minderung,
        "REPARATURKOSTEN": reparatur_kosten,
        "KOSTENPAUSCHALE": kosten_pauschale,
        "SACHVERST_KOSTEN": sachverst_kosten,
        "KOSTENSUMME_X": euro_format(kostensumme),
    })

    return save_word_bezeichg("vorlage_schreiben-1.docx", context_standard, "Standard_schreiben")


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

    return save_word_bezeichg("vorlage_130_prozent-1.docx", context_standard, "130_prozent")


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

    return save_word_bezeichg("vorlage_konkret_unter_wbw-1.docx", context_standard, "konkret_unter_wbw")


def vorlage_totalschaden_fiktiv(data: dict) -> Path:
    context_standard = standardabfrage(data)

    wbw_txt = data.get("WIEDERBESCHAFFUNGSWERT", "")
    wba_txt = data.get("WIEDERBESCHAFFUNGSAUFWAND", "")
    na_txt = data.get("NUTZUNGSAUSFALL", "")
    rw_txt = data.get("RESTWERT", "")
    zus_bez = data.get("ZUSATZKOSTEN_BEZEICHNUNG", "")
    zus_bet = data.get("ZUSATZKOSTEN_BETRAG", "")

    summe = (
        euro_to_float(wbw_txt)
        + euro_to_float(wba_txt)
        + euro_to_float(na_txt)
        + euro_to_float(rw_txt)
        + euro_to_float(zus_bet)
    )

    context_standard.update({
        "WIEDERBESCHAFFUNGSWERT": wbw_txt,
        "WIEDERBESCHAFFUNGSAUFWAND": wba_txt,
        "NUTZUNGSAUSFALL": na_txt,
        "RESTWERT": rw_txt,
        "ZUSATZKOSTEN_BEZEICHNUNG": zus_bez,
        "ZUSATZKOSTEN_BETRAG": zus_bet,
        "KOSTENSUMME_X": euro_format(summe),
    })

    return save_word_bezeichg("vorlage_totalschaden_fiktiv-1.docx", context_standard, "totalschaden_fiktiv")


def vorlage_schreibentotalschaden(data: dict) -> Path:
    context_standard = standardabfrage(data)

    # In deiner Vorlage dürfte es so ähnlich heißen; falls nicht: Key anpassen!
    wbau_txt = data.get("WIEDERBESCHAFFUNGSWERTAUFWAND", "")

    context_standard.update({
        "WIEDERBESCHAFFUNGSWERTAUFWAND": wbau_txt,
        "KOSTENSUMME_X": euro_format(euro_to_float(wbau_txt)),
    })

    return save_word_bezeichg("vorlage_schreibentotalschaden-1.docx", context_standard, "schreibentotalschaden")
