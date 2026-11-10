# InstructionText Extractor

A Flask web application that extracts all `<instructionText>` tags from XML files and generates a formatted Word document with their paths and content.

## Features

- Extracts all `<instructionText>` tags from XML files
- Shows the full XML path to each instruction
- Preserves HTML formatting (bold, colors, font sizes, italics, underline)
- Generates a professional Word document with proper formatting
- Drag & drop file upload interface
- Bootstrap 5 responsive design

## Installation

1. Install Python dependencies:
```bash
pip install -r requirements.txt
```

2. Run the application:
```bash
python app.py
```

3. Open your browser and navigate to:
```
http://localhost:5000
```

## Usage

1. **Upload XML File**: Drag and drop or click to select your XML file
2. **Generate Document**: Click "Generate Word Document" button
3. **Download**: The Word document will automatically download

## Sample XML Format

The tool looks for `<instructionText>` tags in your XML. The content can include:
- Plain text
- HTML formatted content with entities like `&lt;`, `&gt;`, `&amp;`, `&auml;`
- HTML styling with tags like `<strong>`, `<span>`, `<font>`, `<em>`, `<u>`

Example:
```xml
<instructionText>&lt;html&gt;&lt;font size=5&gt;&lt;strong&gt;Note:&lt;/strong&gt; Important instruction here.&lt;/html&gt;</instructionText>
```

## Output Format

The generated Word document includes:

1. **Header Section**:
   - Document title
   - Source filename
   - Generation timestamp
   - Total instruction count

2. **For Each Instruction**:
   - Instruction number
   - Full XML path
   - Formatted content with preserved HTML styling

## Configuration

- **Max File Size**: 50 MB (configurable in `app.py`)
- **Allowed Formats**: XML only
- **Port**: 5000 (configurable in `app.py`)

## Technical Details

### HTML Formatting Support

The application supports the following HTML elements:
- `<strong>`, `<b>` → Bold text
- `<em>`, `<i>` → Italic text
- `<u>` → Underlined text
- `<span style="color: #rrggbb;">` → Colored text
- `<font size="1-7">` → Different font sizes

### Path Generation

The tool generates XPath-like paths showing the exact location of each `<instructionText>` tag in the XML structure, including indices for repeated elements.

## Requirements

- Python 3.8+
- Flask 3.0.0
- python-docx 1.1.0
- beautifulsoup4 4.12.2
- lxml 5.1.0

## Troubleshooting

**Issue**: "No <instructionText> tags found"
- **Solution**: Verify your XML file contains `<instructionText>` tags

**Issue**: "XML parsing error"
- **Solution**: Ensure your XML file is well-formed and valid

**Issue**: "File too large"
- **Solution**: Reduce file size or increase `MAX_CONTENT_LENGTH` in `app.py`

## License

This tool is provided as-is for internal use.