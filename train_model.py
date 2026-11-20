import pandas as pd
from sklearn.model_selection import train_test_split
from sklearn.ensemble import RandomForestClassifier
import joblib

print("Reading CSV file...")

# 1. Read the CSV file
df = pd.read_csv('fitness_data.csv')

# 2. Split the data into features (X) and labels (y)
X = df.drop('exercise_name', axis=1)
y = df['exercise_name']

# 3. Split the dataset into training (80%) and testing (20%)
X_train, X_test, y_train, y_test = train_test_split(
    X, y, test_size=0.2, random_state=42
)

print("Training the AI model...")

# 4. Create and train the Random Forest model
model = RandomForestClassifier(n_estimators=100, random_state=42)
model.fit(X_train, y_train)

# 5. Evaluate the model accuracy
accuracy = model.score(X_test, y_test)
print(f"\nSuccess! Model Accuracy: {accuracy * 100:.2f}%")

# 6. Save the trained model
joblib.dump(model, 'fitness_model.pkl')
print("Model saved as 'fitness_model.pkl'")