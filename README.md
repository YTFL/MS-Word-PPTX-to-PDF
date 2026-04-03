# MS Word/PPTX to PDF Converter

A small Python script that converts Microsoft Word and PowerPoint files in the current folder to PDF.

## What It Converts

- `.doc`
- `.docx`
- `.pptx`

The script scans the folder it is run from and creates a `.pdf` file for each supported input file.

## Requirements

- Windows
- Microsoft Word and Microsoft PowerPoint installed (desktop Office apps)
- Python 3.8+
- Python package: `comtypes`

## Setup

1. Clone or download this repository.
2. Open a terminal in the project folder.
3. Install dependency:

```bash
pip install comtypes
```

## Usage

1. Put your `.doc`, `.docx`, and/or `.pptx` files in the same folder as `convert_to_pdf.py`.
2. Run:

```bash
python convert_to_pdf.py
```

3. The script will generate PDFs with matching file names:

- `Report.docx` -> `Report.pdf`
- `Slides.pptx` -> `Slides.pdf`

## Behavior

- Uses Microsoft Office via COM automation.
- Converts files in the **current working directory**.
- Skips conversion if a target PDF already exists.
- Continues processing other files even if one file fails.
- Closes Word/PowerPoint automatically when finished.

## Notes

- If no supported files are found, the script prints a message and exits.
- Macro-enabled files (`.docm`, `.pptm`) are not included by default.
- Password-protected or corrupted files may fail to convert.

## Troubleshooting

- If you see COM-related errors, ensure Office apps are installed and activated.
- Make sure no blocking dialogs are open in Word or PowerPoint.
- Try running the terminal as your regular desktop user (same user profile that can open Office normally).

## License

Check [LICENSE](LICENSE)
