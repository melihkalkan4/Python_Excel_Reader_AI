import pandas as pd
from sklearn.model_selection import train_test_split
from sklearn.ensemble import RandomForestRegressor
from sklearn.metrics import mean_squared_error

def load_and_preprocess(file_path):
    df = pd.read_excel(file_path)
    columns_of_interest = ['Unnamed: 0', 'Unnamed: 1', 'Unnamed: 2', 'Unnamed: 3', 'Unnamed: 4']
    df = df[columns_of_interest]
    df.fillna(method='ffill', inplace=True)
    return df

def extract_features(df):
    df['feature_sum'] = df.sum(axis=1)
    df['feature_mean'] = df.mean(axis=1)
    return df

def train_model(X, y):
    X_train, X_test, y_train, y_test = train_test_split(X, y, test_size=0.2, random_state=42)
    model = RandomForestRegressor(n_estimators=100, random_state=42)
    model.fit(X_train, y_train)
    predictions = model.predict(X_test)
    mse = mean_squared_error(y_test, predictions)
    print(f"Test MSE: {mse}")
    return model, X_test, y_test, predictions

def generate_comments(predictions):
    comments = []
    threshold = predictions.mean()  # Ortalama bir eşik değeri
    for pred in predictions:
        if pred > threshold:
            comments.append("Bu parametreler pozitif bir etkiye işaret ediyor.")
        else:
            comments.append("Bu parametreler olumsuz bir etkiye işaret ediyor.")
    return comments

def update_model_with_feedback(model, X_train, y_train):
    model.fit(X_train, y_train, warm_start=True)
    return model

# Dosya yolu ve veri yükleme
file_path = 'path_to_your_file.xlsx'
data = load_and_preprocess(file_path)

# Özellik çıkarımı
data = extract_features(data)

# Model eğitimi ve test
X = data[['feature_sum', 'feature_mean']]
y = data['outcome_variable']  # Bu sütun veri setinde mevcut olmalı
model, X_test, y_test, predictions = train_model(X, y)

# Yorumlar üretme
comments = generate_comments(predictions)
for comment in comments:
    print(comment)

# Modeli kullanıcı geri bildirimleriyle güncelleme
# X_train ve y_train yeni verilerle güncellenebilir
# model = update_model_with_feedback(model, X_train, y_train)
