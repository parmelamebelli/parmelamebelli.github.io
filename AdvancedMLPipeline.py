# Import necessary libraries
from sklearn.datasets import load_iris
from sklearn.model_selection import train_test_split, GridSearchCV
from sklearn.ensemble import RandomForestClassifier
from sklearn.preprocessing import StandardScaler, PolynomialFeatures
from sklearn.decomposition import PCA
from sklearn.pipeline import Pipeline, FeatureUnion
from sklearn.base import BaseEstimator, TransformerMixin
from sklearn.metrics import make_scorer, accuracy_score, classification_report, confusion_matrix
import numpy as np

# Custom transformer for adding advanced custom features
class AdvancedCustomFeatures(BaseEstimator, TransformerMixin):
    def __init__(self, add_poly_features=True):
        self.add_poly_features = add_poly_features
    
    def fit(self, X, y=None):
        return self
    
    def transform(self, X):
        # Adding square root features
        sqrt_features = np.sqrt(X)
        
        if self.add_poly_features:
            # Adding polynomial features
            poly_transformer = PolynomialFeatures(degree=2, include_bias=False)
            poly_features = poly_transformer.fit_transform(X)
        else:
            poly_features = X

        # Compute statistical features
        mean_features = np.mean(X, axis=1, keepdims=True)
        std_features = np.std(X, axis=1, keepdims=True)
        
        # Combine all features
        features_combined = np.hstack((X, sqrt_features, poly_features, mean_features, std_features))
        return features_combined

# Load the Iris dataset
iris = load_iris()
X = iris.data
y = iris.target

# Split the dataset
X_train, X_test, y_train, y_test = train_test_split(X, y, test_size=0.3, random_state=42)

# Pipeline with feature union
pipeline = Pipeline([
    ('features', FeatureUnion([
        ('advanced_custom_features', AdvancedCustomFeatures(add_poly_features=True)),
        ('std_scaler', StandardScaler())
    ])),
    ('pca', PCA(n_components=3)),
    ('clf', RandomForestClassifier(random_state=42))
])

# Parameters for GridSearch
parameters = {
    'clf__n_estimators': [100, 150],
    'clf__max_depth': [None, 10, 15],
    'pca__n_components': [2, 3, 4]
}

# Custom scoring function
def custom_accuracy(y_true, y_pred):
    return accuracy_score(y_true, y_pred)

custom_scorer = make_scorer(custom_accuracy)

# GridSearchCV for hyperparameter tuning
grid_search = GridSearchCV(pipeline, parameters, cv=5, scoring=custom_scorer, n_jobs=-1)

# Fit the model
grid_search.fit(X_train, y_train)

# Print the best parameters
print("Best parameters found:", grid_search.best_params_)

# Predictions on the test set
predictions = grid_search.predict(X_test)

# Model evaluation
print(classification_report(y_test, predictions))
print(confusion_matrix(y_test, predictions))
