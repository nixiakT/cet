# CET4 Word Checker

Upload a `.txt`, `.docx`, or `.pdf` file (or paste text) and choose CET4 or CET6 to list out-of-level words. The UI can export highlighted HTML, PDF, and DOCX.

## Run locally

```bash
pip install -r requirements.txt
python cet4_web.py
```

Open `http://localhost:5000`.

## Deploy (Render)

- Create a new Web Service.
- Runtime: Python
- Build Command: `pip install -r requirements.txt`
- Start Command: `gunicorn cet4_web:app`
