# ICMPC Inventory Management System - Full System

Default login:

Username: admin
Password: admin123

Run locally:

```bash
python3 -m venv .venv
source .venv/bin/activate
python -m pip install -r requirements.txt
python app.py
```

Render settings:

Build Command:
```bash
pip install -r requirements.txt
```

Start Command:
```bash
gunicorn app:app
```

Logo:
Place your ICMPC logo at:
static/images/logo.png
