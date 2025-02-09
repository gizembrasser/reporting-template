# 📝 Dynamic DOCX Form Filler

This is a simple Flask web application that allows users to upload a DOCX template file, dynamically fill in placeholders with custom data, and download the completed document.

## 🚀 Features

- **Upload a DOCX template:** The program scans the <ins>first column</ins> of each table in your DOCX file for placeholders like 'Name', 'Location', etc.
- **Fill out the form:** After uploading the template, a form is generated with corresponding input fields:
    - Placeholders containing the text 'date/time' will have date pickers.
    - All other placeholders will have multi-line text areas.
- **Download the filled document:** Once the form is submitted, the <ins>second column</ins> of the DOCX file is filled with the provided data, and will be ready for download.


## ⚙️ Installation

1. Clone this repository and navigate to the directory:

```bash
git clone https://github.com/your-repo/docx-form-filler.git
cd docx-form-filler
```

2. Set up a Python environment:
```bash
python -m venv venv
source venv/bin/activate  # On Windows: venv\Scripts\activate
```

3. Install the dependencies:
```bash
pip install -r requirements.txt
```

## 💡 Usage

1. Start the Flask server:
```bash
python app.py
```

2. Open your browser and go to:
```bash
http://127.0.0.1:5000/
```

3. Follow the steps:
    - Upload your DOCX template.
    - Provide a name for the output file (e.g., `filled_report.docx`).
    - Fill out the dynamic form.
    - Generate and download the completed document.

> Make sure your template file follows the same table format as in the provided `example_template_docx` in this repository.

## 🗂️ Project Structure

```
.
├── app.py
├── utils.py
├── uploads/                # Folder for uploaded template files
├── docs/                   # Folder for generated filled documents
├── templates/
│   ├── setup.html          # Upload template & set output file name
│   └── index.html          # Dynamic form for placeholder inputs
├── static/
│   └── opmaak.css          # Basic CSS styling for the forms
└── README.md
```


