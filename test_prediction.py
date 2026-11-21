import joblib
import pandas as pd

# Load the trained model
model = joblib.load("model/solar_model.pkl")

# New input data
new_data = pd.DataFrame({
    "Temperature": [31],
    "Humidity": [50],
    "Solar_Irradiance": [800],
    "Wind_Speed": [3.0],
    "Cloud_Cover": [15]
})

# Make prediction
prediction = model.predict(new_data)

print("Solar Power Prediction")
print("----------------------")
print("Temperature:", new_data["Temperature"].iloc[0], "°C")
print("Humidity:", new_data["Humidity"].iloc[0], "%")
print("Solar Irradiance:", new_data["Solar_Irradiance"].iloc[0])
print("Wind Speed:", new_data["Wind_Speed"].iloc[0])
print("Cloud Cover:", new_data["Cloud_Cover"].iloc[0], "%")
print("Predicted Power:", round(prediction[0], 2), "kW")