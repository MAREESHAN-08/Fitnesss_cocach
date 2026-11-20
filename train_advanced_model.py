import pandas as pd
import numpy as np
import joblib
import json
import matplotlib.pyplot as plt
import seaborn as sns

from sklearn.model_selection import train_test_split, StratifiedKFold, cross_val_score, GridSearchCV
from sklearn.preprocessing import StandardScaler
from sklearn.ensemble import RandomForestClassifier, GradientBoostingClassifier
from sklearn.svm import SVC
from sklearn.neural_network import MLPClassifier
from sklearn.metrics import accuracy_score, precision_recall_fscore_support, classification_report, confusion_matrix

from feature_engineering import process_csv_dataframe

print("="*60)
print("DATA SCIENCE MACHINE LEARNING PIPELINE TRAINING")
print("="*60)

# 1. Load Dataset
print("\n[1/6] Loading raw dataset 'fitness_data.csv'...")
raw_df = pd.read_csv('fitness_data.csv')
print(f"      Raw dataset shape: {raw_df.shape}")

# 2. Feature Engineering
print("\n[2/6] Running Biomechanical Feature Engineering...")
df = process_csv_dataframe(raw_df)
print(f"      Enriched dataset shape: {df.shape}")

X = df.drop('exercise_name', axis=1)
y = df['exercise_name']
feature_names = list(X.columns)

# 3. Train / Test Split
print("\n[3/6] Splitting dataset (80% Train, 20% Test) with Stratification...")
X_train, X_test, y_train, y_test = train_test_split(
    X, y, test_size=0.20, random_state=42, stratify=y
)

# Feature Scaling
scaler = StandardScaler()
X_train_scaled = scaler.fit_transform(X_train)
X_test_scaled = scaler.transform(X_test)

# 4. Multi-Model Benchmarking with 5-Fold Cross Validation
print("\n[4/6] Benchmarking Candidate ML Algorithms (5-Fold Stratified CV)...")
candidate_models = {
    'Random Forest': RandomForestClassifier(n_estimators=100, random_state=42),
    'Gradient Boosting': GradientBoostingClassifier(n_estimators=100, random_state=42),
    'Support Vector Machine (SVC)': SVC(probability=True, kernel='rbf', random_state=42),
    'Multi-Layer Perceptron (Neural Net)': MLPClassifier(hidden_layer_sizes=(128, 64), max_iter=500, random_state=42)
}

benchmark_results = {}
skf = StratifiedKFold(n_splits=5, shuffle=True, random_state=42)

for name, clf in candidate_models.items():
    # Fit model
    clf.fit(X_train_scaled, y_train)
    y_pred = clf.predict(X_test_scaled)
    acc = accuracy_score(y_test, y_pred)
    
    # Cross Validation
    cv_scores = cross_val_score(clf, X_train_scaled, y_train, cv=skf, scoring='accuracy')
    
    prec, rec, f1, _ = precision_recall_fscore_support(y_test, y_pred, average='weighted')
    
    benchmark_results[name] = {
        'Test Accuracy': round(acc * 100, 2),
        'CV Mean Accuracy': round(cv_scores.mean() * 100, 2),
        'CV Std': round(cv_scores.std() * 100, 2),
        'Weighted Precision': round(prec * 100, 2),
        'Weighted Recall': round(rec * 100, 2),
        'Weighted F1-Score': round(f1 * 100, 2)
    }
    
    print(f"   -> {name}: Test Acc = {acc*100:.2f}%, CV Acc = {cv_scores.mean()*100:.2f}% (+/- {cv_scores.std()*100:.2f}%)")

# 5. Hyperparameter Tuning on Top Candidate (Random Forest)
print("\n[5/6] Performing Hyperparameter Tuning with GridSearchCV on Random Forest...")
param_grid = {
    'n_estimators': [100, 150],
    'max_depth': [None, 15, 25],
    'min_samples_split': [2, 5],
    'criterion': ['gini', 'entropy']
}

grid_search = GridSearchCV(
    RandomForestClassifier(random_state=42),
    param_grid,
    cv=skf,
    scoring='accuracy',
    n_jobs=-1
)
grid_search.fit(X_train_scaled, y_train)

best_model = grid_search.best_estimator_
print(f"      Best Hyperparameters: {grid_search.best_params_}")

# Evaluate the tuned estimator on the untouched test split before full-data fit.
holdout_pred = best_model.predict(X_test_scaled)
holdout_accuracy = accuracy_score(y_test, holdout_pred)
report_text = classification_report(y_test, holdout_pred, zero_division=0)
report_dict = classification_report(y_test, holdout_pred, output_dict=True, zero_division=0)
print(f"\n[+] Tuned held-out test accuracy: {holdout_accuracy * 100:.2f}%")
print("\nHeld-out classification report:\n", report_text)

# Fit the deployable estimator on all available data after evaluation.
print("\n[+] Retraining Best Model on Complete Dataset for Maximum Accuracy & Generalization...")
scaler_full = StandardScaler()
X_full_scaled = scaler_full.fit_transform(X)
best_model.fit(X_full_scaled, y)

# 6. Generate & Save Visual Artifacts
print("\n[6/6] Generating Visual Artifacts & Saving Pipeline...")

# A. Confusion Matrix Plot
plt.figure(figsize=(8, 6))
cm = confusion_matrix(y_test, holdout_pred, labels=best_model.classes_)
sns.heatmap(cm, annot=True, fmt='d', cmap='Blues', xticklabels=best_model.classes_, yticklabels=best_model.classes_)
plt.title("Held-Out Test Confusion Matrix - Fitness Pose Classification", fontsize=14, pad=12)
plt.xlabel("Predicted Exercise Class", fontsize=12)
plt.ylabel("True Exercise Class", fontsize=12)
plt.tight_layout()
plt.savefig("confusion_matrix.png", dpi=300)
plt.close()
print("      Saved 'confusion_matrix.png'")

# B. Feature Importance Plot
if hasattr(best_model, 'feature_importances_'):
    plt.figure(figsize=(10, 6))
    importances = best_model.feature_importances_
    indices = np.argsort(importances)[::-1][:15] # Top 15 features
    top_features = [feature_names[i] for i in indices]
    top_importances = importances[indices]
    
    sns.barplot(x=top_importances, y=top_features, hue=top_features, palette="viridis", legend=False)
    plt.title("Top 15 Most Important Features for Exercise Classification", fontsize=14, pad=12)
    plt.xlabel("Relative Feature Importance Score", fontsize=12)
    plt.ylabel("Engineered Feature", fontsize=12)
    plt.tight_layout()
    plt.savefig("feature_importance.png", dpi=300)
    plt.close()
    print("      Saved 'feature_importance.png'")

# C. Save Model Metrics Summary to JSON
metrics_summary = {
    'benchmark_results': benchmark_results,
    'best_hyperparameters': grid_search.best_params_,
    'final_accuracy': round(holdout_accuracy * 100, 2),
    'evaluation_type': 'held-out test split',
    'classes': list(best_model.classes_),
    'classification_report': report_dict
}

with open("model_metrics.json", "w") as f:
    json.dump(metrics_summary, f, indent=4)
print("      Saved 'model_metrics.json'")

# D. Save Full Pipeline Object
pipeline_data = {
    'scaler': scaler_full,
    'model': best_model,
    'feature_names': feature_names,
    'classes': best_model.classes_,
    'metrics': metrics_summary
}
joblib.dump(pipeline_data, 'fitness_pipeline_v2.pkl')
print("      Saved model pipeline as 'fitness_pipeline_v2.pkl'")

print("\n"+"="*60)
print("MACHINE LEARNING PIPELINE TRAINING COMPLETED SUCCESSFULLY!")
print("="*60)
