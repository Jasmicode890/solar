# Solar Power Prediction System

A Flask application for solar power prediction using weather inputs. It includes
user and admin dashboards, prediction history and Excel/PDF exports, live
weather data from Open-Meteo, and optional email and WhatsApp notifications.

## Requirements

- Python 3.11 or later
- The trained models in `model/` (included in this repository)

The model files use scikit-learn serialization. Install the pinned dependencies
from `requirements.txt` to avoid compatibility problems.

## Run locally

From the project directory, create and activate a virtual environment:

```powershell
py -m venv .venv
.\.venv\Scripts\Activate.ps1
python -m pip install -r requirements.txt
Copy-Item .env.example .env
```

Edit `.env` and set a unique `SECRET_KEY`, `ADMIN_USERNAME`, and
`ADMIN_PASSWORD` (at least 12 characters). The first admin account is created
when the app starts. Email and WhatsApp settings are optional; leave them blank
to disable those notifications.

Start the development server:

```powershell
python app.py
```

Open <http://127.0.0.1:5000/login>. Register a standard user at
`/register`, or sign in with the admin account configured in `.env`.

The built-in Flask server is for local development only. Do not expose it
directly to the internet or enable `FLASK_DEBUG` in a public deployment.

## Project contents

- `app.py` — Flask application and SQLite database initialization
- `model/` — trained model artifacts
- `templates/`, `static/` — web interface
- `train_*.py` — model training scripts

Datasets, the local SQLite database, project reports, and presentations are
intentionally excluded from this public repository. The training scripts expect
their input CSV files under `dataset/`; provide those files locally before
training. The app creates `database/solar.db` automatically.

## Configuration

See `.env.example` for supported settings. Do not commit `.env` or paste
credentials into source files. For production, use a securely generated secret
key, HTTPS, and a production WSGI server.
