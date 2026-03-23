from __future__ import annotations

from pathlib import Path
from typing import Set, Dict, Any
from datetime import datetime

from docxtpl import DocxTemplate
from docx import Document  # 🔥 NEU

# Repo-Root ist Vorlagen-Ordner
BASE_DIR = Path(__file__).resolve().parent
VORLAGEN_DIR = BASE_DIR

OUTPUT_DIR = BASE_DIR / "Output_wordvorlage"
OUTPUT_DIR.mkdir(exist_ok=True)


def safe_filename(s: str) -> str:
    s = (s or "").strip()
    cleaned = "".join(c for c in s if c.isalnum() or c in ("-", "_"))
    return cleaned or "Unbekannt"


def list_docx_templates() -> list[str]:
    return sorted(p.name for p in VORLAGEN_DIR.glob("*.docx"))


def get_template_path(tpl_name: str) -> Path:
    tpl_path = VORLAGEN_DIR / tpl_name
    if not tpl_path.exists():
        raise FileNotFoundError(
            f"Vorlage nicht gefunden: {tpl_path}\n"
            f"Vorhandene .docx im Repo: {list_docx_templates()}"
        )
    return tpl_path


def get_template_vars(tpl_name: str) -> Set[str]:
    """Liest die in der Word-Vorlage verwendeten Variablen aus."""
    tpl_path = get_template_path(tpl_name)
    tpl = DocxTemplate(str(tpl_path))
    return set(tpl.get_undeclared_template_variables() or [])


# 🔥 NEU: NUR Tabellen bereinigen
def cleanup_empty_table_rows(doc: Document):
    """
    Löscht NUR Tabellenzeilen, wenn alle Zellen leer sind.
    Alles außerhalb der Tabelle bleibt unberührt.
    """
    for table in doc.tables:
        rows_to_delete = []

        for row in table.rows:
            texts = [cell.text.strip() for cell in row.cells]

            # wenn ALLE Zellen leer → löschen
            if all(t == "" for t in texts):
                rows_to_delete.append(row)

        for row in rows_to_delete:
            table._element.remove(row._element)


def render_word(tpl_name: str, context: Dict[str, Any], out_prefix: str) -> Path:
    """Rendert die Word-Vorlage mit context und speichert im Output-Ordner."""
    tpl_path = get_template_path(tpl_name)

    tpl = DocxTemplate(str(tpl_path))
    tpl.render(context)

    nachname = safe_filename(str(context.get("MANDANT_NACHNAME", "Unbekannt")))
    timestamp = datetime.now().strftime("%d-%m-%Y")
    out_name = f"{"01-AS_an_VR"}_{timestamp}.docx"
    out_path = OUTPUT_DIR / out_name

    tpl.save(str(out_path))

    # 🔥 NEU: nur Tabellen bereinigen
    doc = Document(str(out_path))
    cleanup_empty_table_rows(doc)
    doc.save(str(out_path))

    return out_path
