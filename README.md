# PDF → Word — GitHub + Render

Flask/PyMuPDF/python-docx manuscript converter. Output uses Times New Roman 12 pt and single spacing.

## GitHub
```bash
git init
git add .
git commit -m "Initial PDF to Word converter"
git branch -M main
git remote add origin https://github.com/YOUR_USERNAME/pdf-to-word.git
git push -u origin main
```

## Render
Create a Web Service connected to the GitHub repo.
Build: `pip install -r requirements.txt`
Start: `gunicorn app:app`

`render.yaml` is included for Blueprint deployment.

The app expects selectable-text PDFs. Scanned PDFs need OCR.
