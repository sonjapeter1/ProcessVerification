import os
import json
import uuid
from pathlib import Path
from datetime import datetime
from xml.etree import ElementTree as ET

import bleach
from bleach.css_sanitizer import CSSSanitizer
from bs4 import BeautifulSoup, NavigableString, Tag
from defusedxml import ElementTree as SafeET

from flask import (
    Flask,
    render_template,
    request,
    redirect,
    url_for,
    session,
    send_file,
    abort
)

from werkzeug.utils import secure_filename

from docx import Document
from docx.shared import Inches, Pt, RGBColor


# --------------------------------------------------
# FLASK CONFIGURATION
# --------------------------------------------------

app = Flask(__name__)

app.secret_key = os.environ.get(
    "SECRET_KEY",
    "change-this-secret-key"
)

MAX_MB = 20

app.config["MAX_CONTENT_LENGTH"] = MAX_MB * 1024 * 1024

RESULT_DIR = Path(app.instance_path) / "results"
RESULT_DIR.mkdir(parents=True, exist_ok=True)

ALLOWED_EXTENSIONS = {".xml"}


# --------------------------------------------------
# HTML SANITIZATION
# --------------------------------------------------

ALLOWED_TAGS = {
    "a",
    "b",
    "strong",
    "i",
    "em",
    "u",
    "s",
    "p",
    "br",
    "div",
    "span",
    "ul",
    "ol",
    "li",
    "sub",
    "sup",
    "font"
}

ALLOWED_ATTRIBUTES = {
    "a": ["href", "title"],
    "span": ["style"],
    "div": ["style"],
    "p": ["style"],
    "font": ["size", "color", "face"]
}

CSS_SANITIZER = CSSSanitizer(
    allowed_css_properties=[
        "color",
        "background-color",
        "font-size",
        "font-weight",
        "font-style",
        "text-decoration",
        "font-family"
    ]
)


# --------------------------------------------------
# HELPER FUNCTIONS
# --------------------------------------------------

def allowed_file(filename):
    """
    Prüft, ob die hochgeladene Datei eine zulässige
    Dateiendung besitzt.
    """

    return Path(filename).suffix.lower() in ALLOWED_EXTENSIONS


def local_name(tag):
    """
    Entfernt:

    1. XML-Namespaces
    2. XML-Präfixe
    3. PAS-X-Java-Paketnamen

    Beispiel:

    com.werum.pasx.spec.components.speccore.global.CommonBFVO

    wird zu:

    CommonBFVO
    """

    name = str(tag).rsplit("}", 1)[-1]
    name = name.split(":")[-1]
    name = name.split(".")[-1]

    return name


def get_inner_content(element):
    """
    Liest den vollständigen inneren Inhalt eines
    instructionText-Elements.

    Unterstützt:

    - einfachen Text
    - HTML-Inhalt
    - verschachtelte XML-/HTML-Elemente
    """

    parts = []

    if element.text:
        parts.append(element.text)

    for child in list(element):
        serialized_child = ET.tostring(
            child,
            encoding="unicode",
            method="html"
        )

        parts.append(serialized_child)

    return "".join(parts).strip()


def get_direct_child_text(node, tag_name):
    """
    Liest den Text eines direkten Child-Elements.

    Es werden nur direkte Children durchsucht,
    nicht alle darunterliegenden Elemente.
    """

    expected_name = tag_name.lower()

    for child in list(node):
        child_name = local_name(child.tag).lower()

        if child_name == expected_name:
            return (child.text or "").strip()

    return None


def create_label(custom_id, description):
    """
    Erstellt eine lesbare Bezeichnung aus
    customId und description.

    Beispiele:

    BF1A020 - Datum Kuchen
    CBF1 - Common BF
    """

    custom_id = (custom_id or "").strip()
    description = (description or "").strip()

    if custom_id and description:
        return f"{custom_id} - {description}"

    if custom_id:
        return custom_id

    if description:
        return description

    return ""


# --------------------------------------------------
# XML EXTRACTION
# --------------------------------------------------

def extract_instruction_texts(xml_bytes):
    """
    Extrahiert alle instructionText-Elemente.

    Der Pfad besteht aus:

    CommonBFVO customId - description
    /
    SpecProp customId - description

    Beispiel:

    CBF1 - Common BF / BF1A020 - Datum Kuchen
    """

    root = SafeET.fromstring(xml_bytes)

    results = []

    def walk(
        node,
        current_spec_prop=None,
        current_common_bf=None
    ):
        """
        Durchläuft rekursiv den gesamten XML-Baum.

        current_spec_prop:
            Die aktuell nächstgelegene SpecProp.

        current_common_bf:
            Die aktuell nächstgelegene CommonBFVO.
        """

        node_name = local_name(node.tag)
        node_name_lower = node_name.lower()

        custom_id = get_direct_child_text(
            node,
            "customId"
        )

        description = get_direct_child_text(
            node,
            "description"
        )

        # ----------------------------------------------
        # CommonBFVO-Kontext aktualisieren
        # ----------------------------------------------

        if node_name_lower == "commonbfvo":
            current_common_bf = create_label(
                custom_id,
                description
            )

        # ----------------------------------------------
        # SpecProp-Kontext aktualisieren
        # ----------------------------------------------

        if node_name_lower.endswith("specpropvo"):
            current_spec_prop = create_label(
                custom_id,
                description
            )

        # ----------------------------------------------
        # Children durchlaufen
        # ----------------------------------------------

        for child in list(node):
            child_name = local_name(child.tag)
            child_name_lower = child_name.lower()

            if child_name_lower == "instructiontext":
                raw_html = get_inner_content(child)

                safe_html = bleach.clean(
                    raw_html,
                    tags=ALLOWED_TAGS,
                    attributes=ALLOWED_ATTRIBUTES,
                    css_sanitizer=CSS_SANITIZER,
                    protocols={
                        "http",
                        "https",
                        "mailto"
                    },
                    strip=True
                )

                text_content = BeautifulSoup(
                    raw_html,
                    "html.parser"
                ).get_text(
                    " ",
                    strip=True
                )

                path_parts = []

                if current_common_bf:
                    path_parts.append(
                        current_common_bf
                    )

                if current_spec_prop:
                    path_parts.append(
                        current_spec_prop
                    )

                path = " / ".join(path_parts)

                if not path:
                    path = "Unbekannter XML-Kontext"

                results.append({
                    "path": path,
                    "common_bf": current_common_bf or "",
                    "spec_prop": current_spec_prop or "",
                    "content_html": safe_html,
                    "content_text": text_content
                })

            walk(
                child,
                current_spec_prop,
                current_common_bf
            )

    walk(root)

    return results


# --------------------------------------------------
# HTML TO WORD FORMATTING
# --------------------------------------------------

def add_html_to_paragraph(paragraph, html):
    """
    Überträgt grundlegende HTML-Formatierungen
    in Word-Runs.

    Unterstützt unter anderem:

    - Fettschrift
    - Kursivschrift
    - Unterstreichung
    - Schriftfarbe
    - Schriftgröße
    - Zeilenumbrüche
    - Aufzählungen
    """

    soup = BeautifulSoup(
        html or "",
        "html.parser"
    )

    def normalize_font_size(value):
        """
        Konvertiert HTML-Font-Größen oder CSS-Größen
        in Word-Punktgrößen.
        """

        if value is None:
            return None

        value = str(value).strip().lower()

        html_font_sizes = {
            "1": 8,
            "2": 10,
            "3": 12,
            "4": 14,
            "5": 18,
            "6": 24,
            "7": 36
        }

        if value in html_font_sizes:
            return html_font_sizes[value]

        try:
            if value.endswith("pt"):
                return float(
                    value.removesuffix("pt").strip()
                )

            if value.endswith("px"):
                pixels = float(
                    value.removesuffix("px").strip()
                )

                return pixels * 0.75

            return float(value)

        except (ValueError, TypeError):
            return None

    def normalize_color(value):
        """
        Konvertiert eine Hex-Farbe in ein Format,
        das python-docx verwenden kann.
        """

        if not value:
            return None

        value = str(value).strip()
        value = value.replace("#", "").upper()

        if len(value) == 3:
            value = "".join(
                character * 2
                for character in value
            )

        if len(value) == 6:
            return value

        return None

    def add_node(node, state):
        """
        Verarbeitet einen BeautifulSoup-Knoten
        rekursiv.
        """

        if isinstance(node, NavigableString):
            value = str(node)

            if not value:
                return

            run = paragraph.add_run(value)

            run.bold = state["bold"]
            run.italic = state["italic"]
            run.underline = state["underline"]

            if state.get("color"):
                try:
                    run.font.color.rgb = RGBColor.from_string(
                        state["color"]
                    )
                except (ValueError, TypeError):
                    pass

            if state.get("size"):
                try:
                    run.font.size = Pt(
                        float(state["size"])
                    )
                except (ValueError, TypeError):
                    pass

            if state.get("font_name"):
                run.font.name = state["font_name"]

            return

        if not isinstance(node, Tag):
            return

        next_state = dict(state)

        name = node.name.lower()

        # ----------------------------------------------
        # Standard-HTML-Formatierungen
        # ----------------------------------------------

        if name in ("b", "strong"):
            next_state["bold"] = True

        elif name in ("i", "em"):
            next_state["italic"] = True

        elif name == "u":
            next_state["underline"] = True

        # ----------------------------------------------
        # Veraltetes HTML-font-Element unterstützen
        # ----------------------------------------------

        if name == "font":
            font_size = normalize_font_size(
                node.get("size")
            )

            font_color = normalize_color(
                node.get("color")
            )

            font_name = node.get("face")

            if font_size:
                next_state["size"] = font_size

            if font_color:
                next_state["color"] = font_color

            if font_name:
                next_state["font_name"] = font_name

        # ----------------------------------------------
        # Inline CSS auswerten
        # ----------------------------------------------

        style = node.get("style", "")

        for declaration in style.split(";"):
            if ":" not in declaration:
                continue

            key, value = [
                part.strip()
                for part in declaration.split(":", 1)
            ]

            key = key.lower()
            value_lower = value.lower()

            if key == "color":
                color = normalize_color(value)

                if color:
                    next_state["color"] = color

            elif key == "font-size":
                size = normalize_font_size(value_lower)

                if size:
                    next_state["size"] = size

            elif key == "font-weight":
                if value_lower in (
                    "bold",
                    "600",
                    "700",
                    "800",
                    "900"
                ):
                    next_state["bold"] = True

            elif key == "font-style":
                if value_lower == "italic":
                    next_state["italic"] = True

            elif key == "text-decoration":
                if "underline" in value_lower:
                    next_state["underline"] = True

            elif key == "font-family":
                font_name = value.strip(
                    "\"'"
                )

                if font_name:
                    next_state["font_name"] = font_name

        # ----------------------------------------------
        # Zeilenumbruch
        # ----------------------------------------------

        if name == "br":
            paragraph.add_run().add_break()
            return

        # ----------------------------------------------
        # Listenpunkt
        # ----------------------------------------------

        if name == "li":
            paragraph.add_run("• ")

        # ----------------------------------------------
        # Children rekursiv verarbeiten
        # ----------------------------------------------

        for child in node.children:
            add_node(
                child,
                next_state
            )

        # ----------------------------------------------
        # Blockelemente mit Zeilenumbruch abschließen
        # ----------------------------------------------

        if name in ("p", "div", "li"):
            paragraph.add_run().add_break()

    initial_state = {
        "bold": False,
        "italic": False,
        "underline": False,
        "color": None,
        "size": None,
        "font_name": None
    }

    for child in soup.contents:
        add_node(
            child,
            initial_state
        )


# --------------------------------------------------
# WORD DOCUMENT GENERATION
# --------------------------------------------------

def generate_word_document(rows, source_name):
    """
    Erstellt ein Word-Dokument mit:

    1. PAS-X-Pfad
    2. formatiertem InstructionText
    3. reinem InstructionText
    """

    doc = Document()

    section = doc.sections[0]

    section.top_margin = Inches(0.65)
    section.bottom_margin = Inches(0.65)
    section.left_margin = Inches(0.65)
    section.right_margin = Inches(0.65)

    # ----------------------------------------------
    # Dokumenttitel
    # ----------------------------------------------

    doc.add_heading(
        "InstructionText Extractor",
        level=0
    )

    # ----------------------------------------------
    # Metadaten
    # ----------------------------------------------

    doc.add_paragraph(
        f"Source file: {source_name}"
    )

    doc.add_paragraph(
        "Generated: "
        + datetime.now().strftime(
            "%d.%m.%Y %H:%M:%S"
        )
    )

    doc.add_paragraph(
        f"Instruction count: {len(rows)}"
    )

    # ----------------------------------------------
    # Tabelle
    # ----------------------------------------------

    table = doc.add_table(
        rows=1,
        cols=3
    )

    table.style = "Table Grid"

    header = table.rows[0].cells

    header[0].text = "PAS-X Path"
    header[1].text = "InstructionText"
    header[2].text = "Plain text InstructionText"

    for item in rows:
        cells = table.add_row().cells

        cells[0].text = item.get(
            "path",
            ""
        )

        formatted_paragraph = cells[1].paragraphs[0]

        add_html_to_paragraph(
            formatted_paragraph,
            item.get(
                "content_html",
                ""
            )
        )

        cells[2].text = item.get(
            "content_text",
            ""
        )

    # ----------------------------------------------
    # Dokument speichern
    # ----------------------------------------------

    output_path = (
        RESULT_DIR /
        f"{uuid.uuid4().hex}.docx"
    )

    doc.save(output_path)

    return output_path


# --------------------------------------------------
# INDEX / UPLOAD ROUTE
# --------------------------------------------------

@app.route("/", methods=["GET", "POST"])
def index():
    if request.method == "GET":
        return render_template(
            "index.html",
            error=None,
            max_mb=MAX_MB
        )

    uploaded_file = request.files.get(
        "xml_file"
    )

    # ----------------------------------------------
    # Upload prüfen
    # ----------------------------------------------

    if not uploaded_file or not uploaded_file.filename:
        return render_template(
            "index.html",
            error="Please select an XML file.",
            max_mb=MAX_MB
        ), 400

    filename = secure_filename(
        uploaded_file.filename
    )

    if not allowed_file(filename):
        return render_template(
            "index.html",
            error=(
                "Invalid file type. "
                "Please upload an XML file."
            ),
            max_mb=MAX_MB
        ), 400

    try:
        xml_bytes = uploaded_file.read()

        if not xml_bytes:
            raise ValueError(
                "The uploaded XML file is empty."
            )

        # ------------------------------------------
        # XML-Daten extrahieren
        # ------------------------------------------

        rows = extract_instruction_texts(
            xml_bytes
        )

        if not rows:
            raise ValueError(
                "No instructionText elements found."
            )

        # ------------------------------------------
        # Eindeutige Result-ID
        # ------------------------------------------

        result_id = uuid.uuid4().hex

        # ------------------------------------------
        # Ergebnisdaten speichern
        # ------------------------------------------

        payload = {
            "source_name": filename,
            "created_at": datetime.now().strftime(
                "%d.%m.%Y %H:%M:%S"
            ),
            "rows": rows
        }

        result_file = (
            RESULT_DIR /
            f"{result_id}.json"
        )

        result_file.write_text(
            json.dumps(
                payload,
                ensure_ascii=False,
                indent=2
            ),
            encoding="utf-8"
        )

        session["result_id"] = result_id

        return redirect(
            url_for("result")
        )

    except ET.ParseError as exc:
        return render_template(
            "index.html",
            error=f"XML parsing failed: {exc}",
            max_mb=MAX_MB
        ), 400

    except ValueError as exc:
        return render_template(
            "index.html",
            error=f"XML processing failed: {exc}",
            max_mb=MAX_MB
        ), 400

    except Exception:
        app.logger.exception(
            "Unexpected XML processing error"
        )

        return render_template(
            "index.html",
            error="An unexpected error occurred.",
            max_mb=MAX_MB
        ), 500


# --------------------------------------------------
# RESULT PAGE
# --------------------------------------------------

@app.route("/result")
def result():
    result_id = session.get(
        "result_id"
    )

    if not result_id:
        return redirect(
            url_for("index")
        )

    result_file = (
        RESULT_DIR /
        f"{result_id}.json"
    )

    if not result_file.is_file():
        session.pop(
            "result_id",
            None
        )

        return redirect(
            url_for("index")
        )

    data = json.loads(
        result_file.read_text(
            encoding="utf-8"
        )
    )

    return render_template(
        "result.html",
        source_name=data["source_name"],
        created_at=data["created_at"],
        rows=data["rows"],
        count=len(data["rows"])
    )


# --------------------------------------------------
# WORD DOWNLOAD
# --------------------------------------------------

@app.route("/download")
def download_word():
    result_id = session.get(
        "result_id"
    )

    if not result_id:
        return redirect(
            url_for("index")
        )

    result_file = (
        RESULT_DIR /
        f"{result_id}.json"
    )

    if not result_file.is_file():
        abort(404)

    data = json.loads(
        result_file.read_text(
            encoding="utf-8"
        )
    )

    docx_path = generate_word_document(
        data["rows"],
        data["source_name"]
    )

    download_name = (
        f"InstructionText_"
        f"{Path(data['source_name']).stem}.docx"
    )

    return send_file(
        docx_path,
        as_attachment=True,
        download_name=download_name,
        mimetype=(
            "application/vnd.openxmlformats-officedocument."
            "wordprocessingml.document"
        )
    )


# --------------------------------------------------
# FILE SIZE ERROR
# --------------------------------------------------

@app.errorhandler(413)
def file_too_large(error):
    return render_template(
        "index.html",
        error=(
            f"File too large. "
            f"Maximum size: {MAX_MB} MB."
        ),
        max_mb=MAX_MB
    ), 413


# --------------------------------------------------
# APPLICATION START
# --------------------------------------------------

if __name__ == "__main__":
    app.run(
        debug=True,
        host="0.0.0.0",
        port=5000
    )