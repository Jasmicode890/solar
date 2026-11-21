
import requests

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
    data = response.json()

    current = data["current"]

    print("\n--- LIVE WEATHER API TEST ---")
    print("Location: Bengaluru")
    print("Timestamp:", current["time"])
    print("Temperature:", current["temperature_2m"], "°C")
    print("Humidity:", current["relative_humidity_2m"], "%")
    print("Cloud cover:", current["cloud_cover"], "%")
    print("Wind speed:", current["wind_speed_10m"], "km/h")
    print(
        "Solar irradiance:",
        current["shortwave_radiation"],
        "W/m²"
    )

    print("\nAPI connection successful!")

except Exception as error:
    print("API test failed:", error)
