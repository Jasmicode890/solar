from flask import (
    Flask, render_template, request, jsonify, send_file,
    redirect, url_for, session, flash
)
import sqlite3
import psycopg
import secrets
from functools import wraps
import joblib
import pandas as pd
import requests
import os
import re
import smtplib
from email.message import EmailMessage
from dotenv import load_dotenv
from io import BytesIO
from math import ceil
from datetime import timedelta
from werkzeug.security import generate_password_hash, check_password_hash
from psycopg.rows import tuple_row

from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill

from reportlab.lib import colors
from reportlab.lib.pagesizes import landscape, A4
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer
from reportlab.lib.styles import getSampleStyleSheet

load_dotenv()

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
DATABASE_PATH = os.path.join(BASE_DIR, "database", "solar.db")
DATABASE_URL = os.getenv("DATABASE_URL", "").strip()
MODEL_PATH = os.path.join(BASE_DIR, "model", "solar_five_input_model.pkl")

app = Flask(__name__)
secret_key = os.getenv("SECRET_KEY")
if os.getenv("VERCEL") == "1" and not secret_key:
    raise RuntimeError("SECRET_KEY must be configured when running on Vercel.")
app.config["SECRET_KEY"] = secret_key or secrets.token_hex(32)
app.config["SESSION_COOKIE_HTTPONLY"] = True
app.config["SESSION_COOKIE_SAMESITE"] = "Lax"
app.config["SESSION_COOKIE_SECURE"] = os.getenv(
    "SESSION_COOKIE_SECURE", "false"
).lower() == "true"
app.permanent_session_lifetime = timedelta(hours=8)

if not os.getenv("SECRET_KEY"):
    app.logger.warning(
        "SECRET_KEY is not set in .env. Sessions will reset when Flask restarts."
    )


# ------------------------- EMAIL NOTIFICATION -------------------------
def send_email_notification(
    predicted_power, temperature, humidity,
    solar_irradiance, wind_speed, cloud_cover
):
    try:
        sender_email = os.getenv("EMAIL_ADDRESS")
        app_password = os.getenv("EMAIL_APP_PASSWORD")
        receiver_email = os.getenv("EMAIL_RECIPIENT")

        if not sender_email or not app_password or not receiver_email:
            app.logger.warning("Email settings are missing in .env")
            return False

        message = EmailMessage()
        message["Subject"] = "Solar Power Prediction Report"
        message["From"] = sender_email
        message["To"] = receiver_email
        message.set_content(f"""
Hello!

A new solar power prediction has been generated.

Prediction Details:
-------------------
Predicted Solar Power: {predicted_power} kW

Temperature: {temperature} °C
Humidity: {humidity} %
Solar Irradiance: {solar_irradiance} W/m²
Wind Speed: {wind_speed} m/s
Cloud Cover: {cloud_cover} %

Thank you,
Solar Power Prediction System
""")

        with smtplib.SMTP("smtp.gmail.com", 587, timeout=20) as server:
            server.starttls()
            server.login(sender_email, app_password)
            server.send_message(message)

        app.logger.info("Email notification sent successfully.")
        return True

    except Exception:
        app.logger.exception("Email notification failed.")
        return False


# ------------------------- WHATSAPP NOTIFICATION -------------------------
def send_whatsapp_notification(
    predicted_power, temperature, humidity,
    solar_irradiance, wind_speed, cloud_cover
):
    """
    Sends a WhatsApp template message using Meta WhatsApp Cloud API.

    Required approved template:
      solar_prediction_alert
    It must contain six BODY placeholders: {{1}} through {{6}}.
    """

    access_token = os.getenv("WHATSAPP_ACCESS_TOKEN")
    phone_number_id = os.getenv("WHATSAPP_PHONE_NUMBER_ID")
    recipient = os.getenv("WHATSAPP_RECIPIENT")
    api_version = os.getenv("WHATSAPP_API_VERSION", "v25.0")
    template_name = os.getenv(
        "WHATSAPP_TEMPLATE_NAME", "solar_prediction_alert"
    )
    language_code = os.getenv("WHATSAPP_LANGUAGE_CODE", "en_US")

    if not all([
        access_token, phone_number_id, recipient,
        api_version, template_name, language_code
    ]):
        app.logger.warning(
            "WhatsApp configuration is incomplete. Check WHATSAPP_* in .env."
        )
        return False

    # Meta expects the destination number with country code, without + or spaces.
    recipient = re.sub(r"\D", "", recipient)
    if not recipient:
        app.logger.error("WHATSAPP_RECIPIENT is invalid.")
        return False

    api_version = api_version.strip().strip("/")
    if not api_version.startswith("v"):
        api_version = "v" + api_version

    url = (
        f"https://graph.facebook.com/{api_version}/"
        f"{phone_number_id}/messages"
    )

    # These values must match {{1}} to {{6}} in the approved template.
    parameters = [
        {"type": "text", "text": f"{predicted_power:.2f}"},
        {"type": "text", "text": str(temperature)},
        {"type": "text", "text": str(humidity)},
        {"type": "text", "text": str(solar_irradiance)},
        {"type": "text", "text": str(wind_speed)},
        {"type": "text", "text": str(cloud_cover)}
    ]

    payload = {
        "messaging_product": "whatsapp",
        "recipient_type": "individual",
        "to": recipient,
        "type": "template",
        "template": {
            "name": template_name,
            "language": {"code": language_code},
            "components": [{
                "type": "body",
                "parameters": parameters
            }]
        }
    }

    headers = {
        "Authorization": f"Bearer {access_token}",
        "Content-Type": "application/json"
    }

    try:
        response = requests.post(
            url, headers=headers, json=payload, timeout=20
        )
        response.raise_for_status()
        result = response.json()

        messages = result.get("messages", [])
        if messages:
            app.logger.info(
                "WhatsApp notification accepted by Meta. Message ID: %s",
                messages[0].get("id", "not returned")
            )
            return True

        app.logger.error("WhatsApp API response has no message ID: %s", result)
        return False

    except requests.exceptions.HTTPError:
        # Do not log the access token. Meta's response body contains the API error.
        app.logger.error(
            "WhatsApp API HTTP error: %s",
            response.text[:2000] if "response" in locals() else "No response"
        )
        return False
    except requests.exceptions.RequestException as error:
        app.logger.error("WhatsApp request failed: %s", error)
        return False
    except ValueError as error:
        app.logger.error("WhatsApp API returned invalid JSON: %s", error)
        return False


# ------------------------- LOAD MODEL -------------------------
model = joblib.load(MODEL_PATH)


# ------------------------- DATABASE -------------------------
class PostgresCursor:
    def __init__(self, cursor):
        self.cursor = cursor

    def execute(self, query, parameters=()):
        self.cursor.execute(query.replace("?", "%s"), parameters)
        return self

    def fetchone(self):
        return self.cursor.fetchone()

    def fetchall(self):
        return self.cursor.fetchall()


class PostgresConnection:
    def __init__(self, connection):
        self.connection = connection

    def cursor(self):
        return PostgresCursor(self.connection.cursor())

    def execute(self, query, parameters=()):
        return PostgresCursor(self.connection.execute(
            query.replace("?", "%s"), parameters
        ))

    def commit(self):
        self.connection.commit()

    def close(self):
        self.connection.close()


def get_db_connection():
    if DATABASE_URL:
        return PostgresConnection(psycopg.connect(
            DATABASE_URL,
            connect_timeout=20,
            row_factory=tuple_row
        ))

    if os.getenv("VERCEL") == "1":
        raise RuntimeError(
            "DATABASE_URL must be configured with a persistent PostgreSQL "
            "database when running on Vercel."
        )

    os.makedirs(os.path.dirname(DATABASE_PATH), exist_ok=True)
    # Keep the default tuple rows used by the original HTML/Chart.js templates.
    conn = sqlite3.connect(DATABASE_PATH, timeout=20)
    return conn


def login_required(view_function):
    @wraps(view_function)
    def wrapped_view(*args, **kwargs):
        if not session.get("user_id"):
            flash("Please log in to continue.", "warning")
            return redirect(url_for("login"))
        return view_function(*args, **kwargs)
    return wrapped_view


def admin_required(view_function):
    @wraps(view_function)
    def wrapped_view(*args, **kwargs):
        if not session.get("user_id"):
            flash("Please log in to continue.", "warning")
            return redirect(url_for("login"))
        if session.get("role") != "admin":
            flash("You do not have permission to access the Admin dashboard.", "danger")
            return redirect(url_for("home"))
        return view_function(*args, **kwargs)
    return wrapped_view


def user_required(view_function):
    """Allow only authenticated standard users into the User workspace."""
    @wraps(view_function)
    def wrapped_view(*args, **kwargs):
        if not session.get("user_id"):
            flash("Please log in to continue.", "warning")
            return redirect(url_for("login"))
        if session.get("role") != "user":
            flash("This page is for User accounts.", "warning")
            return redirect(url_for("admin_dashboard"))
        return view_function(*args, **kwargs)
    return wrapped_view


@app.context_processor
def inject_session_user():
    return {
        "session_username": session.get("username"),
        "session_role": session.get("role")
    }


def init_db():
    """Create tables and safely migrate existing prediction records."""
    conn = get_db_connection()
    cursor = conn.cursor()
    id_definition = (
        "SERIAL PRIMARY KEY"
        if DATABASE_URL
        else "INTEGER PRIMARY KEY AUTOINCREMENT"
    )

    cursor.execute(f"""
        CREATE TABLE IF NOT EXISTS users (
            id {id_definition},
            username TEXT NOT NULL UNIQUE,
            password_hash TEXT NOT NULL,
            role TEXT NOT NULL DEFAULT 'user' CHECK(role IN ('admin', 'user')),
            created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
        )
    """)

    cursor.execute(f"""
        CREATE TABLE IF NOT EXISTS predictions (
            id {id_definition},
            temperature REAL,
            humidity REAL,
            solar_irradiance REAL,
            wind_speed REAL,
            cloud_cover REAL,
            predicted_power REAL,
            timestamp TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
            user_id INTEGER,
            FOREIGN KEY (user_id) REFERENCES users(id)
        )
    """)

    # Existing databases won't have user_id. Add it without deleting old records.
    if DATABASE_URL:
        cursor.execute("""
            SELECT column_name
            FROM information_schema.columns
            WHERE table_schema = current_schema()
              AND table_name = 'predictions'
        """)
        prediction_columns = {row[0] for row in cursor.fetchall()}
    else:
        cursor.execute("PRAGMA table_info(predictions)")
        prediction_columns = {row[1] for row in cursor.fetchall()}
    if "user_id" not in prediction_columns:
        cursor.execute("ALTER TABLE predictions ADD COLUMN user_id INTEGER")

    # Create the first administrator from .env. Never hard-code a default password.
    admin_username = os.getenv("ADMIN_USERNAME", "").strip()
    admin_password = os.getenv("ADMIN_PASSWORD", "")
    if admin_username and len(admin_password) >= 12:
        cursor.execute(
            "SELECT id, role FROM users WHERE username = ?",
            (admin_username,)
        )
        existing_admin = cursor.fetchone()
        if existing_admin is None:
            cursor.execute(
                "INSERT INTO users (username, password_hash, role) VALUES (?, ?, 'admin')",
                (admin_username, generate_password_hash(admin_password))
            )
        elif existing_admin[1] != "admin":
            cursor.execute(
                "UPDATE users SET role = 'admin', password_hash = ? WHERE username = ?",
                (generate_password_hash(admin_password), admin_username)
            )
    else:
        app.logger.warning(
            "Set ADMIN_USERNAME and ADMIN_PASSWORD (at least 12 characters) in .env to create the first admin account."
        )

    conn.commit()
    conn.close()


# ------------------------- AUTHENTICATION -------------------------
@app.route("/login", methods=["GET", "POST"])
def login():
    if session.get("user_id"):
        return redirect(url_for("home"))

    if request.method == "POST":
        username = request.form.get("username", "").strip()
        password = request.form.get("password", "")

        conn = get_db_connection()
        user = conn.execute(
            "SELECT id, username, password_hash, role FROM users WHERE username = ?",
            (username,)
        ).fetchone()
        conn.close()

        if user and check_password_hash(user[2], password):
            session.clear()
            session.permanent = True
            session["user_id"] = user[0]
            session["username"] = user[1]
            session["role"] = user[3]
            flash("Login successful.", "success")
            return redirect(url_for("home"))

        flash("Invalid username or password.", "danger")

    return render_template("login.html")


@app.route("/register", methods=["GET", "POST"])
def register():
    if session.get("user_id"):
        return redirect(url_for("home"))

    if request.method == "POST":
        username = request.form.get("username", "").strip()
        password = request.form.get("password", "")
        confirm_password = request.form.get("confirm_password", "")

        if len(username) < 3 or len(username) > 30 or not username.replace("_", "").isalnum():
            flash("Username must be 3–30 characters and contain only letters, numbers or underscores.", "danger")
        elif len(password) < 8:
            flash("Password must contain at least 8 characters.", "danger")
        elif password != confirm_password:
            flash("Passwords do not match.", "danger")
        else:
            conn = get_db_connection()
            try:
                conn.execute(
                    "INSERT INTO users (username, password_hash, role) VALUES (?, ?, 'user')",
                    (username, generate_password_hash(password))
                )
                conn.commit()
                flash("Account created. Please log in.", "success")
                return redirect(url_for("login"))
            except (sqlite3.IntegrityError, psycopg.IntegrityError):
                flash("That username is already registered.", "danger")
            finally:
                conn.close()

    return render_template("register.html")


@app.route("/logout", methods=["GET", "POST"])
def logout():
    session.clear()
    flash("You have been logged out.", "info")
    return redirect(url_for("login"))


# ------------------------- ROLE-BASED DASHBOARDS -------------------------
@app.route("/")
def home():
    if not session.get("user_id"):
        return redirect(url_for("login"))
    if session.get("role") == "admin":
        return redirect(url_for("admin_dashboard"))
    return redirect(url_for("user_dashboard"))


@app.route("/user/dashboard")
@user_required
def user_dashboard():
    user_id = session["user_id"]
    conn = get_db_connection()
    cursor = conn.cursor()

    cursor.execute("""
        SELECT predicted_power, timestamp
        FROM predictions
        WHERE user_id = ?
        ORDER BY id ASC
    """, (user_id,))
    records = cursor.fetchall()

    cursor.execute("""
        SELECT predicted_power, temperature, solar_irradiance,
               humidity, cloud_cover
        FROM predictions
        WHERE user_id = ?
        ORDER BY id DESC
        LIMIT 1
    """, (user_id,))
    latest = cursor.fetchone()

    cursor.execute("SELECT COUNT(*) FROM predictions WHERE user_id = ?", (user_id,))
    total_predictions = cursor.fetchone()[0]
    conn.close()

    if latest is not None:
        latest_prediction = latest[0]
        latest_temperature = latest[1]
        latest_irradiance = latest[2]
        latest_humidity = latest[3]
        latest_cloud_cover = latest[4]
    else:
        latest_prediction = None
        latest_temperature = None
        latest_irradiance = None
        latest_humidity = None
        latest_cloud_cover = None

    return render_template(
        "index.html",
        records=records,
        latest_prediction=latest_prediction,
        latest_temperature=latest_temperature,
        latest_irradiance=latest_irradiance,
        latest_humidity=latest_humidity,
        latest_cloud_cover=latest_cloud_cover,
        total_predictions=total_predictions
    )


@app.route("/admin/dashboard")
@admin_required
def admin_dashboard():
    conn = get_db_connection()
    cursor = conn.cursor()
    cursor.execute("SELECT COUNT(*) FROM users")
    total_users = cursor.fetchone()[0]
    cursor.execute("SELECT COUNT(*) FROM predictions")
    total_predictions = cursor.fetchone()[0]
    cursor.execute("SELECT AVG(predicted_power) FROM predictions")
    average_row = cursor.fetchone()
    cursor.execute("""
        SELECT predicted_power, temperature, solar_irradiance, timestamp
        FROM predictions ORDER BY id DESC LIMIT 1
    """)
    latest = cursor.fetchone()
    cursor.execute("""
        SELECT p.temperature, p.humidity, p.solar_irradiance,
               p.wind_speed, p.cloud_cover, p.predicted_power,
               p.timestamp, COALESCE(u.username, 'Legacy record') AS username
        FROM predictions p
        LEFT JOIN users u ON u.id = p.user_id
        ORDER BY p.id DESC LIMIT 10
    """)
    recent_records = cursor.fetchall()
    cursor.execute("""
        SELECT u.username, u.role, u.created_at,
               COUNT(p.id) AS prediction_count,
               MAX(p.timestamp) AS last_prediction
        FROM users u
        LEFT JOIN predictions p ON p.user_id = u.id
        GROUP BY u.id
        ORDER BY u.created_at DESC
        LIMIT 12
    """)
    users_list = cursor.fetchall()

    # Admin graph: include EVERY stored prediction, including legacy rows
    # whose user_id is NULL. This is read-only and does not alter ownership.
    cursor.execute("SELECT predicted_power, timestamp FROM predictions ORDER BY id ASC")
    chart_records = cursor.fetchall()

    cursor.execute("SELECT COUNT(*) FROM predictions WHERE user_id IS NULL")
    unassigned_records = cursor.fetchone()[0]

    conn.close()

    return render_template(
        "admin_dashboard.html",
        total_users=total_users,
        total_predictions=total_predictions,
        average_prediction=(average_row[0] if average_row and average_row[0] is not None else 0),
        latest=latest,
        recent_records=recent_records,
        users_list=users_list,
        chart_records=chart_records,
        unassigned_records=unassigned_records
    )


# ------------------------- LIVE WEATHER -------------------------
@app.route("/api/weather")
@login_required
def get_current_weather():
    """Fetch current weather and the matching hourly solar-radiation value.

    Open-Meteo exposes temperature, humidity, cloud cover and wind speed through
    its `current` parameter. `shortwave_radiation` is an hourly variable, so it
    must be requested through `hourly` and matched to the current local hour.
    """
    url = "https://api.open-meteo.com/v1/forecast"
    params = {
        "latitude": 12.9716,
        "longitude": 77.5946,
        "current": (
            "temperature_2m,"
            "relative_humidity_2m,"
            "cloud_cover,"
            "wind_speed_10m"
        ),
        "hourly": "shortwave_radiation",
        "forecast_days": 1,
        "timezone": "Asia/Kolkata"
    }

    try:
        response = requests.get(url, params=params, timeout=20)
        response.raise_for_status()
        payload = response.json()
        current = payload["current"]
        hourly = payload["hourly"]

        # Current time and hourly timestamps use the requested Asia/Kolkata zone.
        current_hour = current["time"][:13]  # e.g. 2026-11-21T11
        hourly_times = hourly["time"]
        radiation_values = hourly["shortwave_radiation"]
        radiation_index = next(
            (i for i, timestamp in enumerate(hourly_times)
             if timestamp[:13] == current_hour),
            None
        )
        if radiation_index is None:
            raise ValueError(
                f"No hourly solar-radiation value found for {current_hour}"
            )

        # Open-Meteo returns wind speed in km/h by default. The prediction form
        # expects m/s, so the frontend should populate wind_speed_mps.
        wind_speed_kmh = float(current["wind_speed_10m"])
        wind_speed_mps = wind_speed_kmh / 3.6

        return jsonify({
            "success": True,
            "location": "Bengaluru",
            "timestamp": current["time"],
            "temperature": current["temperature_2m"],
            "humidity": current["relative_humidity_2m"],
            "solar_irradiance": radiation_values[radiation_index],
            "wind_speed_kmh": wind_speed_kmh,
            "wind_speed_mps": round(wind_speed_mps, 2),
            "cloud_cover": current["cloud_cover"]
        })

    except (requests.RequestException, KeyError, ValueError, TypeError, IndexError) as error:
        app.logger.exception("Weather API error: %s", error)
        return jsonify({
            "success": False,
            "message": "Unable to fetch weather data. Please try again."
        }), 502


# ------------------------- PREDICTION -------------------------
@app.route("/predict", methods=["GET", "POST"])
@login_required
def predict():
    prediction = None

    if request.method == "POST":
        try:
            temperature = float(request.form["temperature"])
            humidity = float(request.form["humidity"])
            solar_irradiance = float(request.form["solar_irradiance"])
            wind_speed = float(request.form["wind_speed"])
            cloud_cover = float(request.form["cloud_cover"])

            if not 0 <= humidity <= 100:
                raise ValueError("Humidity must be between 0 and 100.")
            if not 0 <= cloud_cover <= 100:
                raise ValueError("Cloud cover must be between 0 and 100.")
            if solar_irradiance < 0 or wind_speed < 0:
                raise ValueError(
                    "Solar irradiance and wind speed cannot be negative."
                )

            # The form uses m/s, while the model's training data uses km/h.
            wind_speed_kmh = wind_speed * 3.6

            input_data = pd.DataFrame({
                "Temperature": [temperature],
                "Humidity": [humidity],
                "Solar_Irradiance": [solar_irradiance],
                "Wind_Speed": [wind_speed_kmh],
                "Cloud_Cover": [cloud_cover]
            })

            result = model.predict(input_data)
            prediction = round(float(result[0]), 2)

            # Save and commit the prediction before sending notifications.
            conn = get_db_connection()
            try:
                cursor = conn.cursor()
                cursor.execute("""
                    INSERT INTO predictions (
                        temperature, humidity, solar_irradiance,
                        wind_speed, cloud_cover, predicted_power, user_id
                    )
                    VALUES (?, ?, ?, ?, ?, ?, ?)
                """, (
                    temperature, humidity, solar_irradiance,
                    wind_speed, cloud_cover, prediction, session["user_id"]
                ))
                conn.commit()
            finally:
                conn.close()

            # Notification failures are logged and do not cancel the prediction.
            send_email_notification(
                prediction, temperature, humidity,
                solar_irradiance, wind_speed, cloud_cover
            )

            send_whatsapp_notification(
                prediction, temperature, humidity,
                solar_irradiance, wind_speed, cloud_cover
            )

        except ValueError as error:
            return f"Invalid input: {error}", 400

    return render_template("predict.html", prediction=prediction)


# ------------------------- HISTORY -------------------------
@app.route("/history")
@login_required
def history():
    search = request.args.get("search", "").strip()
    page = request.args.get("page", 1, type=int) or 1
    page = max(1, page)
    per_page = 20

    search_clause = """
        (CAST(temperature AS TEXT) LIKE ?
         OR CAST(humidity AS TEXT) LIKE ?
         OR CAST(solar_irradiance AS TEXT) LIKE ?
         OR CAST(wind_speed AS TEXT) LIKE ?
         OR CAST(cloud_cover AS TEXT) LIKE ?
         OR CAST(predicted_power AS TEXT) LIKE ?
         OR CAST(timestamp AS TEXT) LIKE ?)
    """
    term = f"%{search}%"
    conn = get_db_connection()
    cursor = conn.cursor()

    if session.get("role") == "admin":
        base_query = "FROM predictions WHERE " + search_clause
        params = [term] * 7
        cursor.execute("SELECT COUNT(*) FROM predictions")
    else:
        base_query = "FROM predictions WHERE user_id = ? AND " + search_clause
        params = [session["user_id"]] + [term] * 7
        cursor.execute("SELECT COUNT(*) FROM predictions WHERE user_id = ?", (session["user_id"],))
    total_predictions = cursor.fetchone()[0]

    cursor.execute("SELECT COUNT(*) " + base_query, params)
    total_filtered = cursor.fetchone()[0]
    total_pages = max(1, ceil(total_filtered / per_page))
    page = min(page, total_pages)
    offset = (page - 1) * per_page

    cursor.execute("""
        SELECT temperature, humidity, solar_irradiance,
               wind_speed, cloud_cover, predicted_power, timestamp
    """ + base_query + """
        ORDER BY id DESC LIMIT ? OFFSET ?
    """, params + [per_page, offset])
    records = cursor.fetchall()

    if session.get("role") == "admin":
        cursor.execute("SELECT AVG(predicted_power) FROM predictions")
        average_row = cursor.fetchone()
        cursor.execute("SELECT predicted_power FROM predictions ORDER BY id DESC LIMIT 1")
    else:
        cursor.execute("SELECT AVG(predicted_power) FROM predictions WHERE user_id = ?", (session["user_id"],))
        average_row = cursor.fetchone()
        cursor.execute("SELECT predicted_power FROM predictions WHERE user_id = ? ORDER BY id DESC LIMIT 1", (session["user_id"],))
    latest_row = cursor.fetchone()
    conn.close()

    average = average_row[0] if average_row and average_row[0] is not None else 0
    latest = latest_row[0] if latest_row else 0

    return render_template(
        "history.html",
        records=records,
        total=total_predictions,
        average=average,
        latest=latest,
        search=search,
        page=page,
        total_pages=total_pages,
        total_filtered=total_filtered
    )


# ------------------------- EXPORT HELPER -------------------------
def get_all_history():
    conn = get_db_connection()
    cursor = conn.cursor()
    if session.get("role") == "admin":
        cursor.execute("""
            SELECT temperature, humidity, solar_irradiance,
                   wind_speed, cloud_cover, predicted_power, timestamp
            FROM predictions ORDER BY id DESC
        """)
    else:
        cursor.execute("""
            SELECT temperature, humidity, solar_irradiance,
                   wind_speed, cloud_cover, predicted_power, timestamp
            FROM predictions WHERE user_id = ? ORDER BY id DESC
        """, (session["user_id"],))
    records = cursor.fetchall()
    conn.close()
    return records


# ------------------------- EXPORT TO EXCEL -------------------------
@app.route("/history/export/excel")
@login_required
def export_history_excel():
    records = get_all_history()

    workbook = Workbook()
    sheet = workbook.active
    sheet.title = "Prediction History"

    headers = [
        "Temperature (°C)", "Humidity (%)", "Solar Irradiance (W/m²)",
        "Wind Speed (m/s)", "Cloud Cover (%)", "Predicted Power (kW)",
        "Date & Time"
    ]
    sheet.append(headers)

    for cell in sheet[1]:
        cell.font = Font(bold=True, color="FFFFFF")
        cell.fill = PatternFill(fill_type="solid", fgColor="176B50")

    for record in records:
        sheet.append(list(record))

    sheet.freeze_panes = "A2"
    sheet.auto_filter.ref = sheet.dimensions

    for column in sheet.columns:
        max_length = max(len(str(cell.value or "")) for cell in column)
        sheet.column_dimensions[column[0].column_letter].width = min(
            max_length + 3, 30
        )

    output = BytesIO()
    workbook.save(output)
    output.seek(0)

    return send_file(
        output,
        as_attachment=True,
        download_name="solar_prediction_history.xlsx",
        mimetype=(
            "application/vnd.openxmlformats-officedocument."
            "spreadsheetml.sheet"
        )
    )


# ------------------------- EXPORT TO PDF -------------------------
@app.route("/history/export/pdf")
@login_required
def export_history_pdf():
    records = get_all_history()
    output = BytesIO()

    document = SimpleDocTemplate(
        output,
        pagesize=landscape(A4),
        rightMargin=25,
        leftMargin=25,
        topMargin=25,
        bottomMargin=25
    )

    styles = getSampleStyleSheet()
    content = [
        Paragraph("Solar Power Prediction History", styles["Title"]),
        Spacer(1, 12)
    ]

    data = [[
        "Temperature", "Humidity", "Irradiance",
        "Wind Speed", "Cloud Cover", "Power", "Date & Time"
    ]]

    for record in records:
        data.append([
            str(value) if value is not None else ""
            for value in record
        ])

    if not records:
        content.append(Paragraph(
            "No prediction records are available.",
            styles["Normal"]
        ))
    else:
        table = Table(data, repeatRows=1)
        table.setStyle(TableStyle([
            ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#176B50")),
            ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
            ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
            ("FONTSIZE", (0, 0), (-1, -1), 7),
            ("GRID", (0, 0), (-1, -1), 0.4, colors.grey),
            ("ROWBACKGROUNDS", (0, 1), (-1, -1),
             [colors.white, colors.HexColor("#EAF5EF")]),
            ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
            ("LEFTPADDING", (0, 0), (-1, -1), 5),
            ("RIGHTPADDING", (0, 0), (-1, -1), 5),
        ]))
        content.append(table)

    document.build(content)
    output.seek(0)

    return send_file(
        output,
        as_attachment=True,
        download_name="solar_prediction_history.pdf",
        mimetype="application/pdf"
    )


# Create the database table when the app starts.
init_db()


if __name__ == "__main__":
    app.run(debug=os.getenv("FLASK_DEBUG", "false").lower() == "true")
