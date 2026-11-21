from flask import Flask, render_template, request
import sqlite3
import joblib
import pandas as pd

app = Flask(__name__)

# Load the trained ML model
model = joblib.load("model/solar_model.pkl")


# Initialize the database
def init_db():
    conn = sqlite3.connect("database/solar.db")

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


# Home page
@app.route("/")
def home():
    conn = sqlite3.connect("database/solar.db")
    cursor = conn.cursor()

    cursor.execute("""
        SELECT predicted_power, timestamp
        FROM predictions
        ORDER BY id ASC
    """)

    records = cursor.fetchall()
    conn.close()

    return render_template(
        "index.html",
        records=records
    )

# Prediction page
@app.route("/predict", methods=["GET", "POST"])
def predict():

    prediction = None

    if request.method == "POST":

        # Get values entered by the user
        temperature = float(request.form["temperature"])
        humidity = float(request.form["humidity"])
        solar_irradiance = float(request.form["solar_irradiance"])
        wind_speed = float(request.form["wind_speed"])
        cloud_cover = float(request.form["cloud_cover"])

        # Create input data for the ML model
        input_data = pd.DataFrame({
            "Temperature": [temperature],
            "Humidity": [humidity],
            "Solar_Irradiance": [solar_irradiance],
            "Wind_Speed": [wind_speed],
            "Cloud_Cover": [cloud_cover]
        })

        # Make prediction
        result = model.predict(input_data)

        prediction = round(result[0], 2)

        # Save prediction to the database
        conn = sqlite3.connect("database/solar.db")

        cursor = conn.cursor()

        cursor.execute("""
            INSERT INTO predictions
            (
                temperature,
                humidity,
                solar_irradiance,
                wind_speed,
                cloud_cover,
                predicted_power
            )
            VALUES (?, ?, ?, ?, ?, ?)
        """, (
            temperature,
            humidity,
            solar_irradiance,
            wind_speed,
            cloud_cover,
            prediction
        ))

        conn.commit()
        conn.close()

    return render_template(
        "predict.html",
        prediction=prediction
    )


# History page
# History page
@app.route("/history")
def history():

    conn = sqlite3.connect("database/solar.db")

    cursor = conn.cursor()

    cursor.execute("""
        SELECT
            temperature,
            humidity,
            solar_irradiance,
            wind_speed,
            cloud_cover,
            predicted_power,
            timestamp
        FROM predictions
        ORDER BY id DESC
    """)

    records = cursor.fetchall()

    conn.close()     

    return render_template(
        "history.html",
        records=records
    )

# Create database when application starts
init_db()


# Start the Flask application
if __name__ == "__main__":
    app.run(debug=True)