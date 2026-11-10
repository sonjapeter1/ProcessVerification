# InstructionText Extractor with GitHub Copilot

A Flask web application for extracting `instructionText` elements from PAS-X XML files, displaying their process context, correcting the instruction text with GitHub Copilot, and exporting the results to a formatted Word document.

## Features

- Upload PAS-X XML files through a web interface
- Extract all `instructionText` elements
- Identify the associated PAS-X context
- Extract `CommonBFVO` and `SpecPropVO` information
- Generate readable PAS-X paths using `customId` and `description`
- Preserve HTML formatting from the original instruction text
- Correct grammar, spelling, punctuation, and sentence structure using GitHub Copilot CLI
- Preserve the original language and technical meaning
- Preserve PAS-X and GMP terminology, equipment names, material names, parameters, identifiers, values, units, and abbreviations
- Reapply the original HTML formatting after Copilot correction
- Display original and corrected instructions in the web application
- Export the results to a formatted Word document

## Workflow

The application processes an XML file in the following sequence:

1. Upload PAS-X XML file
2. Parse XML securely using `defusedxml`
3. Locate all `instructionText` elements
4. Determine the corresponding `CommonBFVO` and `SpecPropVO`
5. Build the PAS-X context/path
6. Extract the original instruction text and HTML formatting
7. Send the instruction and PAS-X context to GitHub Copilot
8. Correct grammar and spelling while preserving technical meaning
9. Reapply the original HTML formatting
10. Display the results
11. Export the results to Word

## PAS-X Context

For every `instructionText`, the application attempts to determine its surrounding PAS-X structure.

The displayed path is built from:

- `CommonBFVO`
- `SpecPropVO`

using their:

- `customId`
- `description`

Example:

```text
CommonBF-001 - Preparation / SP-010 - Check equipment
```

The same context is supplied to GitHub Copilot as background information when reviewing the instruction.

## GitHub Copilot Integration

The application uses the GitHub Copilot CLI to review every extracted instruction.

Copilot is instructed to correct:

- Grammar
- Spelling
- Punctuation
- Sentence structure

while preserving:

- Original language
- Technical meaning
- PAS-X terminology
- GMP terminology
- Equipment names
- Material names
- Parameter names
- Process terminology
- Identifiers
- Values
- Units
- Abbreviations

Copilot is explicitly instructed not to invent information or include explanations in its response.

If Copilot processing fails for an instruction, the application retains the original content and records the Copilot error.

## HTML Formatting

The application preserves supported formatting from the original `instructionText`.

Supported formatting includes:

- Bold
- Italic
- Underline
- Strikethrough
- Subscript
- Superscript
- Font size
- Font color
- Font family
- Basic paragraphs and lists

For example:

```html
<strong>Gebe Dtaum an</strong>
```

may be corrected by Copilot to:

```html
<strong>Gib das Datum an.</strong>
```

The text changes, while the original formatting is retained.

HTML content is sanitized using `bleach` before further processing.

## Word Export

The generated Word document contains four columns:

| Column | Content |
|---|---|
| PAS-X Path | Associated CommonBFVO and SpecProp context |
| Formatted InstructionText | Original instruction with HTML formatting |
| Plain text InstructionText | Original instruction without HTML formatting |
| Copilot corrected InstructionText | Corrected instruction with formatting reapplied |

The document also contains:

- Source XML filename
- Generation timestamp
- Number of extracted instructions

## Installation

Install the required Python dependencies:

```bash
pip install -r requirements.txt
```

The application uses packages including:

```text
Flask
bleach
beautifulsoup4
defusedxml
python-docx
Werkzeug
```

GitHub Copilot CLI must also be installed and accessible on the machine running the application.

## Configuration

### File size

The default maximum XML upload size is:

```text
50 MB
```

This can be changed using `MAX_MB` in `app.py`.

### Allowed files

Currently only XML files are accepted:

```text
.xml
```

### GitHub Copilot CLI

The path to the Copilot executable is configured through:

```python
COPILOT_EXE
```

The current application expects the GitHub Copilot CLI executable to be available at the configured location.

The Copilot request timeout is:

```text
120 seconds
```

and can be changed using:

```python
COPILOT_TIMEOUT_SECONDS
```

### Flask secret key

For production use, configure the `SECRET_KEY` environment variable instead of relying on the development fallback.

## Running the Application

Start the Flask application with:

```bash
python app.py
```

The application runs on:

```text
http://localhost:5000
```

The Flask server is configured with:

```text
Host: 0.0.0.0
Port: 5000
```

## Usage

1. Start the application.
2. Open the application in your browser.
3. Upload a PAS-X XML file.
4. The XML is parsed and all `instructionText` elements are extracted.
5. Each instruction is sent to GitHub Copilot together with its PAS-X context.
6. Review the original and corrected instructions on the result page.
7. Download the results as a Word document.

## Result Storage

Processing results are temporarily stored as JSON files under the Flask instance directory:

```text
instance/results/
```

Generated Word documents are also created in the result directory before being returned to the user.

## Project Structure

A typical project structure is:

```text
.
├── app.py
├── requirements.txt
├── README.md
├── templates/
│   ├── index.html
│   └── result.html
└── instance/
    └── results/
```

## Security

The application includes several security-related controls:

- XML parsing with `defusedxml`
- HTML sanitization with `bleach`
- Restricted XML file uploads
- Secure filenames using Werkzeug
- Maximum upload size of 50 MB
- Sanitized CSS properties
- Restricted HTML tags and attributes

For production deployment, the Flask development server and default secret key should not be used.

## Requirements

- Python 3.8+
- GitHub Copilot CLI
- Flask
- bleach
- beautifulsoup4
- defusedxml
- python-docx
- Werkzeug

## Purpose

The tool supports review and quality improvement of PAS-X manufacturing instructions by combining deterministic XML extraction and context identification with GitHub Copilot-based language correction.

The XML structure, PAS-X context, and formatting are processed locally by the application. GitHub Copilot is used specifically for reviewing and correcting the extracted instruction text.

## License

This tool is provided as-is for internal use.