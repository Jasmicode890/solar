
import pandas as pd
import joblib
from sklearn.ensemble import RandomForestRegressor
from sklearn.metrics import mean_absolute_error, r2_score

data = pd.read_csv("dataset/solar_plant1_prepared.csv")

features = [
    "Temperature",
    "Module_Temperature",
    "Solar_Irradiance"
]

X = data[features]
y = data["Power_Generated"]

split_index = int(len(data) * 0.8)

X_train = X.iloc[:split_index]
X_test = X.iloc[split_index:]
y_train = y.iloc[:split_index]
y_test = y.iloc[split_index:]

model = RandomForestRegressor(
    n_estimators=100,
    random_state=42,
    n_jobs=-1
)

model.fit(X_train, y_train)

predictions = model.predict(X_test)

mae = mean_absolute_error(y_test, predictions)
r2 = r2_score(y_test, predictions)

print("Real Dataset Model Training Completed!")
print("Training records:", len(X_train))
print("Testing records:", len(X_test))
print("MAE:", round(mae, 2))
print("R2 Score:", round(r2, 4))

import os
os.makedirs("model", exist_ok=True)

joblib.dump(model, "model/solar_real_model.pkl")
print("New model saved successfully!")
