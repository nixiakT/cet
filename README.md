# CET4 Word Checker

Upload a `.txt` file or paste English text to list words not found in the CET4 vocabulary list.

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
