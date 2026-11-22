import joblib
import pandas as pd
from pathlib import Path

MODEL_PATH = (
    Path(__file__).resolve().parents[1]
    / "models"
    / "performance_model.pkl"
)

model = joblib.load(MODEL_PATH)


def predict_performance(features):

    data = pd.DataFrame([features])

    prediction = model.predict(data)[0]

    probabilities = model.predict_proba(data)[0]

    confidence = round(float(max(probabilities)) * 100, 2)

    return {
        "prediction": prediction,
        "confidence": confidence
    }