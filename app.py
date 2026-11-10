import os
import json
import uuid
import subprocess
from copy import deepcopy
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


# ==================================================
# FLASK CONFIGURATION
# ==================================================

app = Flask(__name__)

app.secret_key = os.environ.get(
    "SECRET_KEY",
    "change-this-secret-key"
)

MAX_MB = 50

app.config["MAX_CONTENT_LENGTH"] = (
    MAX_MB * 1024 * 1024
)

RESULT_DIR = (
    Path(app.instance_path)
    / "results"
)

RESULT_DIR.mkdir(
    parents=True,
    exist_ok=True
)

ALLOWED_EXTENSIONS = {
    ".xml"
}


# ==================================================
# GITHUB COPILOT CONFIGURATION
# ==================================================
#AWS Platform 
#Bedrock Infrastructure (AWS) -> einfaches Modelle -> am Besten Deployen 
#Wie deploye ich Container -> BachemApp / container -> Modelle Bedrock freistellen / was für Modelle 
#Docker aufrufen Bedrock aufrufen -> kleines Modell
#Bedrock und Test von meiner Maschiene!
#Alessandra Passini DI (K7) 
COPILOT_EXE = (
    r"C:\Users\ets01\AppData\Local\Microsoft\WinGet\Packages"
    r"\GitHub.Copilot_Microsoft.Winget.Source_8wekyb3d8bbwe"
    r"\copilot.exe"
)

COPILOT_TIMEOUT_SECONDS = 120


# ==================================================
# HTML SANITIZATION
# ==================================================

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
    "a": [
        "href",
        "title",
        "class",
        "data-lexical-text"
    ],
    "b": [
        "class",
        "style",
        "data-lexical-text"
    ],
    "strong": [
        "class",
        "style",
        "data-lexical-text"
    ],
    "i": [
        "class",
        "style",
        "data-lexical-text"
    ],
    "em": [
        "class",
        "style",
        "data-lexical-text"
    ],
    "u": [
        "class",
        "style",
        "data-lexical-text"
    ],
    "s": [
        "class",
        "style",
        "data-lexical-text"
    ],
    "span": [
        "class",
        "style",
        "data-lexical-text"
    ],
    "div": [
        "class",
        "style",
        "data-lexical-text"
    ],
    "p": [
        "class",
        "style",
        "data-lexical-text"
    ],
    "ul": [
        "class",
        "style"
    ],
    "ol": [
        "class",
        "style"
    ],
    "li": [
        "class",
        "style"
    ],
    "sub": [
        "class",
        "style"
    ],
    "sup": [
        "class",
        "style"
    ],
    "font": [
        "size",
        "color",
        "face",
        "class",
        "style",
        "data-lexical-text"
    ]
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


# ==================================================
# GENERAL HELPER FUNCTIONS
# ==================================================

def allowed_file(filename):
    """
    Prüft, ob die Datei eine zulässige Dateiendung hat.
    """

    return (
        Path(filename).suffix.lower()
        in ALLOWED_EXTENSIONS
    )


def local_name(tag):
    """
    Entfernt XML-Namespaces, XML-Präfixe und
    PAS-X-Java-Paketnamen.
    """

    name = str(tag).rsplit(
        "}",
        1
    )[-1]

    name = name.split(":")[-1]
    name = name.split(".")[-1]

    return name


def get_inner_content(element):
    """
    Liest den vollständigen inneren Inhalt eines
    instructionText-Elements.
    """

    parts = []

    if element.text:
        parts.append(
            element.text
        )

    for child in list(element):
        serialized_child = ET.tostring(
            child,
            encoding="unicode",
            method="html"
        )

        parts.append(
            serialized_child
        )

    return "".join(parts).strip()


def get_direct_child_text(
    node,
    tag_name
):
    """
    Liest den Text eines direkten Child-Elements.
    """

    expected_name = tag_name.lower()

    for child in list(node):
        child_name = local_name(
            child.tag
        ).lower()

        if child_name == expected_name:
            return (
                child.text or ""
            ).strip()

    return None


def create_label(
    custom_id,
    description
):
    """
    Erstellt eine lesbare Bezeichnung aus
    customId und description.
    """

    custom_id = (
        custom_id or ""
    ).strip()

    description = (
        description or ""
    ).strip()

    if custom_id and description:
        return (
            f"{custom_id} - "
            f"{description}"
        )

    if custom_id:
        return custom_id

    if description:
        return description

    return ""


def create_context_text(row):
    """
    Erstellt den PAS-X-Kontext, der zusammen mit dem
    InstructionText an GitHub Copilot gesendet wird.
    """

    path = (
        row.get(
            "path",
            ""
        )
        or ""
    ).strip()

    common_bf = (
        row.get(
            "common_bf",
            ""
        )
        or ""
    ).strip()

    spec_prop = (
        row.get(
            "spec_prop",
            ""
        )
        or ""
    ).strip()

    context_parts = []

    if path:
        context_parts.append(
            f"PAS-X path: {path}"
        )

    if common_bf:
        context_parts.append(
            f"CommonBFVO: {common_bf}"
        )

    if spec_prop:
        context_parts.append(
            f"SpecProp: {spec_prop}"
        )

    if not context_parts:
        return (
            "No additional PAS-X context "
            "is available."
        )

    return "\n".join(
        context_parts
    )


# ==================================================
# HTML FORMATTING TRANSFER
# ==================================================

def sanitize_html(html):
    """
    Bereinigt HTML mit derselben Konfiguration,
    die auch während der XML-Extraktion verwendet wird.
    """

    return bleach.clean(
        html or "",
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


def contains_html_formatting(html):
    """
    Prüft, ob der Inhalt mindestens ein HTML-Element
    enthält.
    """

    soup = BeautifulSoup(
        html or "",
        "html.parser"
    )

    return soup.find() is not None


def get_meaningful_text_nodes(soup):
    """
    Gibt alle Textknoten zurück, die sichtbaren Text
    enthalten.
    """

    return [
        node
        for node in soup.find_all(string=True)
        if str(node).strip()
    ]


def get_single_top_level_tag(soup):
    """
    Gibt das einzige Top-Level-HTML-Element zurück,
    sofern keine zusätzlichen sichtbaren Textknoten
    daneben vorhanden sind.
    """

    top_level_tags = [
        node
        for node in soup.contents
        if isinstance(node, Tag)
    ]

    top_level_text_nodes = [
        node
        for node in soup.contents
        if (
            isinstance(node, NavigableString)
            and str(node).strip()
        )
    ]

    if (
        len(top_level_tags) == 1
        and not top_level_text_nodes
    ):
        return top_level_tags[0]

    return None


def create_formatting_wrapper(
    source_tag,
    corrected_text
):
    """
    Erstellt eine Kopie des ursprünglichen HTML-Tags,
    entfernt den alten Inhalt und fügt den korrigierten
    Text ein.

    Tagname und Attribute bleiben erhalten.
    """

    wrapper = BeautifulSoup(
        "",
        "html.parser"
    )

    copied_tag = wrapper.new_tag(
        source_tag.name
    )

    for attribute_name, attribute_value in (
        source_tag.attrs.items()
    ):
        copied_tag.attrs[
            attribute_name
        ] = deepcopy(
            attribute_value
        )

    copied_tag.string = corrected_text

    wrapper.append(
        copied_tag
    )

    return str(wrapper)


def apply_original_html_formatting(
    original_html,
    corrected_text
):
    """
    Überträgt die HTML-Formatierung des ursprünglichen
    InstructionText auf den korrigierten Text.

    Beispiele:

    <strong>Gebe Dtaum an</strong>

    wird zu:

    <strong>Gib das Datum an.</strong>

    Attribute wie class, style und data-lexical-text
    bleiben dabei erhalten.
    """

    original_html = (
        original_html or ""
    ).strip()

    corrected_text = (
        corrected_text or ""
    ).strip()

    if not corrected_text:
        return ""

    if not original_html:
        return corrected_text

    sanitized_original_html = sanitize_html(
        original_html
    )

    if not contains_html_formatting(
        sanitized_original_html
    ):
        return corrected_text

    soup = BeautifulSoup(
        sanitized_original_html,
        "html.parser"
    )

    text_nodes = get_meaningful_text_nodes(
        soup
    )

    # Fall 1:
    # Genau ein Textknoten ist vorhanden.
    # Dadurch kann der alte Text direkt ersetzt werden,
    # während die gesamte HTML-Struktur erhalten bleibt.
    if len(text_nodes) == 1:
        text_nodes[0].replace_with(
            corrected_text
        )

        return sanitize_html(
            str(soup)
        )

    # Fall 2:
    # Der vollständige Inhalt befindet sich in genau
    # einem Top-Level-Element.
    #
    # Das äußere Element und seine Attribute werden
    # erhalten. Der korrigierte Text wird in dieses
    # Element eingesetzt.
    single_top_level_tag = (
        get_single_top_level_tag(
            soup
        )
    )

    if single_top_level_tag is not None:
        formatted_html = create_formatting_wrapper(
            single_top_level_tag,
            corrected_text
        )

        return sanitize_html(
            formatted_html
        )

    # Fall 3:
    # Mehrere unabhängige Top-Level-Elemente sind
    # vorhanden.
    #
    # In diesem Fall wird geprüft, ob alle sichtbaren
    # Texte innerhalb desselben Formatierungselements
    # liegen.
    formatting_tags = {
        "a",
        "b",
        "strong",
        "i",
        "em",
        "u",
        "s",
        "span",
        "font",
        "sub",
        "sup"
    }

    common_formatting_tag = None

    for text_node in text_nodes:
        current_tag = text_node.parent

        while (
            current_tag is not None
            and isinstance(current_tag, Tag)
            and current_tag.name not in formatting_tags
        ):
            current_tag = current_tag.parent

        if (
            current_tag is None
            or not isinstance(current_tag, Tag)
        ):
            common_formatting_tag = None
            break

        if common_formatting_tag is None:
            common_formatting_tag = current_tag

        elif (
            current_tag.name
            != common_formatting_tag.name
            or current_tag.attrs
            != common_formatting_tag.attrs
        ):
            common_formatting_tag = None
            break

    if common_formatting_tag is not None:
        formatted_html = create_formatting_wrapper(
            common_formatting_tag,
            corrected_text
        )

        return sanitize_html(
            formatted_html
        )

    # Wenn keine sichere gemeinsame Formatierung
    # ermittelt werden kann, wird der korrigierte Text
    # ohne HTML zurückgegeben.
    return corrected_text


# ==================================================
# GITHUB COPILOT
# ==================================================

def check_grammar_with_copilot(
    text,
    context_text=""
):
    """
    Korrigiert Grammatik, Rechtschreibung und
    Zeichensetzung über GitHub Copilot CLI.

    context_text wird nur als Hintergrundinformation
    verwendet.
    """

    text = (
        text or ""
    ).strip()

    context_text = (
        context_text or ""
    ).strip()

    if not text:
        return ""

    if not Path(COPILOT_EXE).is_file():
        raise FileNotFoundError(
            "GitHub Copilot CLI was not found at: "
            f"{COPILOT_EXE}"
        )

    prompt = f"""
You are reviewing an instruction from a PAS-X manufacturing specification.

Background context:
{context_text or "No additional PAS-X context is available."}

Instruction to correct:
{text}

Task:
Correct the grammar, spelling, punctuation, and sentence structure of the
instruction.

Requirements:
- Preserve the original technical meaning.
- Preserve the original language of the instruction.
- Preserve PAS-X terminology.
- Preserve GMP terminology.
- Preserve equipment names.
- Preserve material names.
- Preserve parameter names.
- Preserve process terminology.
- Preserve identifiers, values, units, and abbreviations.
- Do not invent information.
- Do not remove technically relevant information.
- Do not include the background context in the result.
- Do not explain the changes.
- Do not use Markdown formatting.
- Do not use HTML formatting.
- Do not wrap the answer in quotation marks.
- Return only the corrected instruction.
""".strip()

    try:
        result = subprocess.run(
            [
                COPILOT_EXE,
                "-p",
                prompt,
                "-s"
            ],
            capture_output=True,
            text=True,
            encoding="utf-8",
            errors="replace",
            timeout=COPILOT_TIMEOUT_SECONDS,
            check=False,
            shell=False
        )

    except subprocess.TimeoutExpired as exc:
        raise RuntimeError(
            "GitHub Copilot did not respond within "
            f"{COPILOT_TIMEOUT_SECONDS} seconds."
        ) from exc

    except OSError as exc:
        raise RuntimeError(
            "GitHub Copilot CLI could not be started: "
            f"{exc}"
        ) from exc

    if result.returncode != 0:
        error_message = (
            result.stderr.strip()
            or result.stdout.strip()
            or "Unknown GitHub Copilot CLI error."
        )

        raise RuntimeError(
            error_message
        )

    corrected_text = (
        result.stdout or ""
    ).strip()

    if not corrected_text:
        raise RuntimeError(
            "GitHub Copilot returned an empty response."
        )

    return corrected_text


def correct_rows_with_copilot(rows):
    """
    Sendet jeden InstructionText zusammen mit seinem
    PAS-X-Kontext an GitHub Copilot.

    Die ursprüngliche HTML-Formatierung wird danach
    auf den korrigierten Text übertragen.

    Falls Copilot für eine Zeile fehlschlägt, wird
    der vollständige ursprüngliche HTML-Inhalt verwendet.
    """

    for number, row in enumerate(
        rows,
        start=1
    ):
        original_text = (
            row.get(
                "content_text",
                ""
            )
            or ""
        ).strip()

        original_html = (
            row.get(
                "content_html",
                ""
            )
            or ""
        ).strip()

        context_text = create_context_text(
            row
        )

        row["number"] = number
        row["context_text"] = context_text
        row["corrected_content_html"] = ""
        row["corrected_content_text"] = ""
        row["copilot_error"] = ""

        if not original_text:
            continue

        try:
            corrected_text = (
                check_grammar_with_copilot(
                    text=original_text,
                    context_text=context_text
                )
            )

            corrected_html = (
                apply_original_html_formatting(
                    original_html=original_html,
                    corrected_text=corrected_text
                )
            )

            row["corrected_content_text"] = (
                corrected_text
            )

            row["corrected_content_html"] = (
                corrected_html
            )

        except Exception as exc:
            app.logger.exception(
                "GitHub Copilot correction failed "
                "for instruction %s",
                number
            )

            row["corrected_content_text"] = (
                original_text
            )

            row["corrected_content_html"] = (
                original_html
                or original_text
            )

            row["copilot_error"] = str(
                exc
            )

    return rows


# ==================================================
# XML EXTRACTION
# ==================================================

def extract_instruction_texts(xml_bytes):
    """
    Extrahiert alle instructionText-Elemente.

    Der Pfad besteht aus:

    CommonBFVO customId - description
    /
    SpecProp customId - description
    """

    root = SafeET.fromstring(
        xml_bytes
    )

    results = []

    def walk(
        node,
        current_spec_prop=None,
        current_common_bf=None
    ):
        node_name = local_name(
            node.tag
        )

        node_name_lower = (
            node_name.lower()
        )

        custom_id = get_direct_child_text(
            node,
            "customId"
        )

        description = get_direct_child_text(
            node,
            "description"
        )

        # CommonBFVO-Kontext aktualisieren
        if node_name_lower == "commonbfvo":
            current_common_bf = create_label(
                custom_id,
                description
            )

        # SpecProp-Kontext aktualisieren
        if node_name_lower.endswith(
            "specpropvo"
        ):
            current_spec_prop = create_label(
                custom_id,
                description
            )

        for child in list(node):
            child_name = local_name(
                child.tag
            )

            child_name_lower = (
                child_name.lower()
            )

            if child_name_lower == "instructiontext":
                raw_html = get_inner_content(
                    child
                )

                safe_html = sanitize_html(
                    raw_html
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

                path = " / ".join(
                    path_parts
                )

                if not path:
                    path = (
                        "Unbekannter XML-Kontext"
                    )

                results.append(
                    {
                        "path": path,
                        "common_bf": (
                            current_common_bf
                            or ""
                        ),
                        "spec_prop": (
                            current_spec_prop
                            or ""
                        ),
                        "content_html": safe_html,
                        "content_text": text_content
                    }
                )

            walk(
                child,
                current_spec_prop,
                current_common_bf
            )

    walk(root)

    return results


# ==================================================
# HTML TO WORD FORMATTING
# ==================================================

def add_html_to_paragraph(
    paragraph,
    html
):
    """
    Überträgt grundlegende HTML-Formatierungen
    in Word-Runs.
    """

    soup = BeautifulSoup(
        html or "",
        "html.parser"
    )

    def normalize_font_size(value):
        if value is None:
            return None

        value = str(
            value
        ).strip().lower()

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
                    value.removesuffix(
                        "pt"
                    ).strip()
                )

            if value.endswith("px"):
                pixels = float(
                    value.removesuffix(
                        "px"
                    ).strip()
                )

                return pixels * 0.75

            return float(value)

        except (
            ValueError,
            TypeError
        ):
            return None

    def normalize_color(value):
        if not value:
            return None

        value = str(
            value
        ).strip()

        value = value.replace(
            "#",
            ""
        ).upper()

        if len(value) == 3:
            value = "".join(
                character * 2
                for character in value
            )

        if len(value) == 6:
            return value

        return None

    def add_node(
        node,
        state
    ):
        if isinstance(
            node,
            NavigableString
        ):
            value = str(node)

            if not value:
                return

            run = paragraph.add_run(
                value
            )

            run.bold = state["bold"]
            run.italic = state["italic"]
            run.underline = state["underline"]
            run.font.strike = state["strike"]

            if state.get("color"):
                try:
                    run.font.color.rgb = (
                        RGBColor.from_string(
                            state["color"]
                        )
                    )
                except (
                    ValueError,
                    TypeError
                ):
                    pass

            if state.get("size"):
                try:
                    run.font.size = Pt(
                        float(
                            state["size"]
                        )
                    )
                except (
                    ValueError,
                    TypeError
                ):
                    pass

            if state.get("font_name"):
                run.font.name = state[
                    "font_name"
                ]

            if state.get("subscript"):
                run.font.subscript = True

            if state.get("superscript"):
                run.font.superscript = True

            return

        if not isinstance(
            node,
            Tag
        ):
            return

        next_state = dict(
            state
        )

        name = node.name.lower()

        if name in (
            "b",
            "strong"
        ):
            next_state["bold"] = True

        elif name in (
            "i",
            "em"
        ):
            next_state["italic"] = True

        elif name == "u":
            next_state["underline"] = True

        elif name == "s":
            next_state["strike"] = True

        elif name == "sub":
            next_state["subscript"] = True
            next_state["superscript"] = False

        elif name == "sup":
            next_state["superscript"] = True
            next_state["subscript"] = False

        if name == "font":
            font_size = normalize_font_size(
                node.get("size")
            )

            font_color = normalize_color(
                node.get("color")
            )

            font_name = node.get(
                "face"
            )

            if font_size:
                next_state["size"] = font_size

            if font_color:
                next_state["color"] = font_color

            if font_name:
                next_state["font_name"] = font_name

        style = node.get(
            "style",
            ""
        )

        for declaration in style.split(";"):
            if ":" not in declaration:
                continue

            key, value = [
                part.strip()
                for part in declaration.split(
                    ":",
                    1
                )
            ]

            key = key.lower()
            value_lower = value.lower()

            if key == "color":
                color = normalize_color(
                    value
                )

                if color:
                    next_state["color"] = color

            elif key == "font-size":
                size = normalize_font_size(
                    value_lower
                )

                if size:
                    next_state["size"] = size

            elif key == "font-weight":
                try:
                    numeric_weight = int(
                        value_lower
                    )
                except ValueError:
                    numeric_weight = None

                if (
                    value_lower == "bold"
                    or (
                        numeric_weight is not None
                        and numeric_weight >= 600
                    )
                ):
                    next_state["bold"] = True

            elif key == "font-style":
                if value_lower == "italic":
                    next_state["italic"] = True

            elif key == "text-decoration":
                if "underline" in value_lower:
                    next_state["underline"] = True

                if "line-through" in value_lower:
                    next_state["strike"] = True

            elif key == "font-family":
                font_name = value.strip(
                    "\"'"
                )

                if font_name:
                    next_state["font_name"] = font_name

        if name == "br":
            paragraph.add_run().add_break()
            return

        if name == "li":
            paragraph.add_run(
                "• "
            )

        for child in node.children:
            add_node(
                child,
                next_state
            )

        if name in (
            "p",
            "div",
            "li"
        ):
            paragraph.add_run().add_break()

    initial_state = {
        "bold": False,
        "italic": False,
        "underline": False,
        "strike": False,
        "subscript": False,
        "superscript": False,
        "color": None,
        "size": None,
        "font_name": None
    }

    for child in soup.contents:
        add_node(
            child,
            initial_state
        )


# ==================================================
# WORD DOCUMENT GENERATION
# ==================================================

def generate_word_document(
    rows,
    source_name
):
    """
    Erstellt ein Word-Dokument mit:

    1. PAS-X-Pfad
    2. formatiertem InstructionText
    3. reinem InstructionText
    4. Copilot-korrigiertem und formatiertem
       InstructionText
    """

    doc = Document()

    section = doc.sections[0]

    section.top_margin = Inches(0.65)
    section.bottom_margin = Inches(0.65)
    section.left_margin = Inches(0.65)
    section.right_margin = Inches(0.65)

    doc.add_heading(
        "InstructionText Extractor",
        level=0
    )

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

    table = doc.add_table(
        rows=1,
        cols=4
    )

    table.style = "Table Grid"

    header = table.rows[0].cells

    header[0].text = "PAS-X Path"

    header[1].text = (
        "Formatted InstructionText"
    )

    header[2].text = (
        "Plain text InstructionText"
    )

    header[3].text = (
        "Copilot corrected InstructionText"
    )

    for item in rows:
        cells = table.add_row().cells

        cells[0].text = item.get(
            "path",
            ""
        )

        formatted_paragraph = (
            cells[1].paragraphs[0]
        )

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

        corrected_paragraph = (
            cells[3].paragraphs[0]
        )

        add_html_to_paragraph(
            corrected_paragraph,
            item.get(
                "corrected_content_html",
                ""
            )
        )

        copilot_error = item.get(
            "copilot_error",
            ""
        )

        if copilot_error:
            error_paragraph = (
                cells[3].add_paragraph()
            )

            error_run = error_paragraph.add_run(
                f"Copilot error: {copilot_error}"
            )

            error_run.italic = True

            error_run.font.color.rgb = RGBColor(
                192,
                0,
                0
            )

    output_path = (
        RESULT_DIR
        / f"{uuid.uuid4().hex}.docx"
    )

    doc.save(
        output_path
    )

    return output_path


# ==================================================
# INDEX / UPLOAD ROUTE
# ==================================================

@app.route(
    "/",
    methods=[
        "GET",
        "POST"
    ]
)
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

    if (
        not uploaded_file
        or not uploaded_file.filename
    ):
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

        rows = extract_instruction_texts(
            xml_bytes
        )

        if not rows:
            raise ValueError(
                "No instructionText elements found."
            )

        rows = correct_rows_with_copilot(
            rows
        )

        result_id = uuid.uuid4().hex

        payload = {
            "source_name": filename,
            "created_at": datetime.now().strftime(
                "%d.%m.%Y %H:%M:%S"
            ),
            "rows": rows
        }

        result_file = (
            RESULT_DIR
            / f"{result_id}.json"
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
            error=(
                f"XML parsing failed: {exc}"
            ),
            max_mb=MAX_MB
        ), 400

    except ValueError as exc:
        return render_template(
            "index.html",
            error=(
                f"XML processing failed: {exc}"
            ),
            max_mb=MAX_MB
        ), 400

    except Exception as exc:
        app.logger.exception(
            "Unexpected XML processing error"
        )

        return render_template(
            "index.html",
            error=(
                "An unexpected error occurred: "
                f"{exc}"
            ),
            max_mb=MAX_MB
        ), 500


# ==================================================
# RESULT PAGE
# ==================================================

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
        RESULT_DIR
        / f"{result_id}.json"
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


# ==================================================
# WORD DOWNLOAD
# ==================================================

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
        RESULT_DIR
        / f"{result_id}.json"
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
        "InstructionText_"
        f"{Path(data['source_name']).stem}"
        ".docx"
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


# ==================================================
# FILE SIZE ERROR
# ==================================================

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


# ==================================================
# APPLICATION START
# ==================================================

if __name__ == "__main__":
    app.run(
        debug=True,
        host="0.0.0.0",
        port=5000
    )