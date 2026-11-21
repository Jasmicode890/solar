from flask import Flask, render_template, request, jsonify, send_file
import sqlite3
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

from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill

from reportlab.lib import colors
from reportlab.lib.pagesizes import landscape, A4
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer
from reportlab.lib.styles import getSampleStyleSheet

load_dotenv()
app = Flask(__name__)


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
model = joblib.load("model/solar_five_input_model.pkl")


# ------------------------- DATABASE -------------------------
def get_db_connection():
    return sqlite3.connect("database/solar.db")


def init_db():
    conn = get_db_connection()
    cursor = conn.cursor()
    cursor.execute("""
        CREATE TABLE IF NOT EXISTS predictions (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            temperature REAL,
            humidity REAL,
            solar_irradiance REAL,
            wind_speed REAL,
            cloud_cover REAL,
            predicted_power REAL,
            timestamp DATETIME DEFAULT CURRENT_TIMESTAMP
        )
    """)
    conn.commit()
    conn.close()


# ------------------------- DASHBOARD -------------------------
@app.route("/")
def home():
    conn = get_db_connection()
    cursor = conn.cursor()

    cursor.execute("""
        SELECT predicted_power, timestamp
        FROM predictions
        ORDER BY id ASC
    """)
    records = cursor.fetchall()

    cursor.execute("""
        SELECT predicted_power, temperature, solar_irradiance,
               humidity, cloud_cover
        FROM predictions
        ORDER BY id DESC
        LIMIT 1
    """)
    latest = cursor.fetchone()

    cursor.execute("SELECT COUNT(*) FROM predictions")
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


# ------------------------- LIVE WEATHER -------------------------
@app.route("/api/weather")
def get_current_weather():
    url = "https://api.open-meteo.com/v1/forecast"
    params = {
        "latitude": 12.9716,
        "longitude": 77.5946,
        "current": (
            "temperature_2m,"
            "relative_humidity_2m,"
            "cloud_cover,"
            "wind_speed_10m,"
            "shortwave_radiation"
        ),
        "timezone": "Asia/Kolkata"
    }

    try:
        response = requests.get(url, params=params, timeout=15)
        response.raise_for_status()
        current = response.json()["current"]

        return jsonify({
            "success": True,
            "location": "Bengaluru",
            "timestamp": current["time"],
            "temperature": current["temperature_2m"],
            "humidity": current["relative_humidity_2m"],
            "solar_irradiance": current["shortwave_radiation"],
            "wind_speed_kmh": current["wind_speed_10m"],
            "cloud_cover": current["cloud_cover"]
        })

    except (requests.RequestException, KeyError, ValueError) as error:
        app.logger.error("Weather API error: %s", error)
        return jsonify({
            "success": False,
            "message": "Unable to fetch weather data. Please try again."
        }), 502


# ------------------------- PREDICTION -------------------------
@app.route("/predict", methods=["GET", "POST"])
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
                        wind_speed, cloud_cover, predicted_power
                    )
                    VALUES (?, ?, ?, ?, ?, ?)
                """, (
                    temperature, humidity, solar_irradiance,
                    wind_speed, cloud_cover, prediction
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
def history():
    search = request.args.get("search", "").strip()
    page = request.args.get("page", 1, type=int) or 1
    page = max(1, page)
    per_page = 20

    conn = get_db_connection()
    cursor = conn.cursor()

    base_query = """
        FROM predictions
        WHERE CAST(temperature AS TEXT) LIKE ?
           OR CAST(humidity AS TEXT) LIKE ?
           OR CAST(solar_irradiance AS TEXT) LIKE ?
           OR CAST(wind_speed AS TEXT) LIKE ?
           OR CAST(cloud_cover AS TEXT) LIKE ?
           OR CAST(predicted_power AS TEXT) LIKE ?
           OR CAST(timestamp AS TEXT) LIKE ?
    """
    term = f"%{search}%"
    params = [term] * 7

    cursor.execute("SELECT COUNT(*) " + base_query, params)
    total_filtered = cursor.fetchone()[0]

    total_pages = max(1, ceil(total_filtered / per_page))
    page = min(page, total_pages)
    offset = (page - 1) * per_page

    cursor.execute("""
        SELECT temperature, humidity, solar_irradiance,
               wind_speed, cloud_cover, predicted_power, timestamp
    """ + base_query + """
        ORDER BY id DESC
        LIMIT ? OFFSET ?
    """, params + [per_page, offset])
    records = cursor.fetchall()

    cursor.execute("SELECT COUNT(*) FROM predictions")
    total_predictions = cursor.fetchone()[0]

    cursor.execute("""
        SELECT AVG(predicted_power), MAX(id)
        FROM predictions
    """)
    average_row = cursor.fetchone()

    cursor.execute("""
        SELECT predicted_power
        FROM predictions
        ORDER BY id DESC LIMIT 1
    """)
    latest_row = cursor.fetchone()
    conn.close()

    average = (
        average_row[0]
        if average_row and average_row[0] is not None
        else 0
    )
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
    cursor.execute("""
        SELECT temperature, humidity, solar_irradiance,
               wind_speed, cloud_cover, predicted_power, timestamp
        FROM predictions
        ORDER BY id DESC
    """)
    records = cursor.fetchall()
    conn.close()
    return records


# ------------------------- EXPORT TO EXCEL -------------------------
@app.route("/history/export/excel")
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
    app.run(debug=True)
