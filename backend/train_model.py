
import numpy as np
import pandas as pd
import joblib
import os

from sklearn.ensemble import RandomForestClassifier
from sklearn.model_selection import train_test_split
from sklearn.metrics import accuracy_score, classification_report

# Reproducible synthetic student dataset
np.random.seed(42)
n = 1200

data = pd.DataFrame({
    "attendance": np.random.uniform(30, 100, n),
    "study_hours": np.random.uniform(0, 10, n),
    "previous_marks": np.random.uniform(20, 100, n),
    "assignment_score": np.random.uniform(10, 100, n),
    "sleep_hours": np.random.uniform(3, 10, n)
})

# Generate target labels using academic performance criteria
score = (
    0.30 * data["previous_marks"]
    + 0.25 * data["attendance"]
    + 0.25 * data["assignment_score"]
    + 0.15 * (data["study_hours"] * 10)
    + 0.05 * (data["sleep_hours"] * 10)
)

data["performance"] = np.where(
    score >= 75, "Excellent",
    np.where(score >= 50, "Average", "At Risk")
)

X = data.drop("performance", axis=1)
y = data["performance"]

X_train, X_test, y_train, y_test = train_test_split(
    X, y, test_size=0.2, random_state=42, stratify=y
)

# Train Random Forest
model = RandomForestClassifier(
    n_estimators=150,
    max_depth=10,
    random_state=42
)

model.fit(X_train, y_train)

# Evaluate model
predictions = model.predict(X_test)
accuracy = accuracy_score(y_test, predictions)

print("Model Training Completed!")
print("Accuracy:", round(accuracy * 100, 2), "%")
print("\nClassification Report:")
print(classification_report(y_test, predictions))

# Save trained model
os.makedirs("models", exist_ok=True)
joblib.dump(model, "models/performance_model.pkl")

print("\nModel saved successfully!")
print("Location: models/performance_model.pkl")
