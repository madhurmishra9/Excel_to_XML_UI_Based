# Excel ↔ XML Converter — GUI

A Tkinter desktop front-end for the [`excel_xml`](../Excel_to_xml) backend. It
exposes every backend operation in one window:

| Tab            | Backend call                          | What it does |
|----------------|---------------------------------------|--------------|
| Excel → XML    | `excel_to_xml`                        | One `.xml` file per sheet |
| XML → Excel    | `xml_to_excel`                        | Combine a folder of XML into one workbook |
| Compare        | `check_xml_data`                      | Report cells that differ between Excel and XML |
| Fetch          | `remote_to_excel` / `remote_to_xml`   | Pull an RSS/XML/HTML/cloud source into Excel or XML |
| Search         | `file_search`                         | Find a file by name across all drives (Windows) |

Differences (Compare) and matches (Search) are printed by the backend; the UI
captures that output and shows it in the **Output** pane. Long operations run on
a background thread so the window stays responsive.

## Layout

This UI lives **next to** the backend repository:

```
Downloads/
├── Excel_to_xml/            <- backend (the `excel_xml` package)
└── Excel_to_XML_UI_Based/   <- this UI
```

`UI.py` adds the sibling `../Excel_to_xml` folder to `sys.path` automatically.
If your backend lives elsewhere, point `EXCEL_XML_BACKEND` at the folder that
contains the `excel_xml` package.

## Running

The UI imports the backend, whose features need the backend's dependencies
(openpyxl, xlrd, beautifulsoup4, lxml, requests, feedparser). The simplest way
is to use the backend's virtual environment:

```bat
:: Windows — from this folder
..\Excel_to_xml\.venv\Scripts\python.exe UI.py
```

Or install the requirements into whatever Python you use:

```bash
pip install -r ../Excel_to_xml/requirements.txt
python UI.py
```

`EXCEL_XML_BACKEND` can override the backend location:

```bat
set EXCEL_XML_BACKEND=C:\path\to\Excel_to_xml
python UI.py
```

## Legacy files

`UI_V1.py`, `convert_excel_to_xml.py`, `compare_excel_to_xml.py` and
`transpose.py` are the original standalone scripts. They are superseded by the
`excel_xml` backend and are kept only for reference — `UI.py` no longer uses
them.
