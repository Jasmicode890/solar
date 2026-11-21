import pandas as pd
from sklearn.model_selection import train_test_split
from sklearn.ensemble import RandomForestRegressor
from sklearn.metrics import mean_absolute_error, r2_score
import joblib

# Load the dataset
data = pd.read_csv("dataset/solar_data.csv")

# Input features
X = data[
    [
        "Temperature",
        "Humidity",
        "Solar_Irradiance",
        "Wind_Speed",
        "Cloud_Cover"
    ]
]

# Target value
y = data["Power_Generated"]

# Split the data into training and testing sets
X_train, X_test, y_train, y_test = train_test_split(
    X, y, test_size=0.2, random_state=42
)

# Create the Random Forest model
model = RandomForestRegressor(
    n_estimators=100,
    random_state=42
)

# Train the model
model.fit(X_train, y_train)

# Make predictions
predictions = model.predict(X_test)

# Evaluate the model
mae = mean_absolute_error(y_test, predictions)
r2 = r2_score(y_test, predictions)

print("Model Training Completed!")
print("Mean Absolute Error:", round(mae, 2))
print("R2 Score:", round(r2, 2))

# Save the trained model
joblib.dump(model, "model/solar_model.pkl")

print("Model saved successfully!")