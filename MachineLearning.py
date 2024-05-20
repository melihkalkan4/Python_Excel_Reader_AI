import pandas as pd
import numpy as np
from sklearn.preprocessing import MinMaxScaler
from keras.models import Sequential
from keras.layers import Dense, LSTM
import pmdarima as pm
import matplotlib.pyplot as plt

# Veri okuma
file_path = r'C:\Users\melih.kalkan\Desktop\Sirius Brick Analizi (18. Hafta)-Copy-(1).xlsx'
turkey_data = pd.read_excel(file_path, sheet_name='Türkiye Analiz')

# Veri hazırlığı
time_series_data = turkey_data.melt(id_vars=['Grup', 'Bölge', 'UTT', 'Brick Name', 'Market', 'Marka'], 
                                    var_name='Date', 
                                    value_name='Value')
time_series_data['Date'] = pd.to_datetime(time_series_data['Date'])
time_series_data = time_series_data[['Date', 'Value']].dropna()
data = time_series_data.set_index('Date').resample('MS').sum()  # 'M' yerine 'MS' (month start) kullanıldı

# Veriyi kontrol etme
print("Data Shape:", data.shape)
print("Data Head:\n", data.head())

# ARIMA modelinin çalışması için veriyi uygun formata dönüştürme
data = data.asfreq('MS')  # Verinin frekansını belirle

# ARIMA modeli
if len(data) >= 24:  # En az 2 yıllık veri gerekliliği
    arima_model = pm.auto_arima(data, seasonal=True, m=12, stepwise=True, suppress_warnings=True)
    arima_forecast = arima_model.predict(n_periods=3)
else:
    arima_forecast = None
    print("ARIMA modeli için yeterli veri yok. En az 24 aylık veri gerekiyor.")

# LSTM modeli
scaler = MinMaxScaler(feature_range=(0, 1))
scaled_data = scaler.fit_transform(data)
train_size = int(len(scaled_data) * 0.8)
train, test = scaled_data[:train_size], scaled_data[train_size:]

def create_dataset(dataset, look_back=1):
    X, Y = [], []
    for i in range(len(dataset) - look_back - 1):
        a = dataset[i:(i + look_back), 0]
        X.append(a)
        Y.append(dataset[i + look_back, 0])
    return np.array(X), np.array(Y)

look_back = 12
trainX, trainY = create_dataset(train, look_back)
testX, testY = create_dataset(test, look_back)

# Verilerin doğru boyutta olup olmadığını kontrol etme
print(f"trainX shape: {trainX.shape}")
print(f"trainY shape: {trainY.shape}")
print(f"testX shape: {testX.shape}")
print(f"testY shape: {testY.shape}")

# Eğer trainX veya testX boşsa (veri noktası yetersizse) hata vermeden durması için kontrol ekleyelim
if trainX.shape[0] == 0 or testX.shape[0] == 0:
    raise ValueError("Yeterli veri noktası yok. LSTM modeli eğitimi için daha fazla veriye ihtiyaç var.")

trainX = np.reshape(trainX, (trainX.shape[0], trainX.shape[1], 1))
testX = np.reshape(testX, (testX.shape[0], testX.shape[1], 1))

model = Sequential()
model.add(LSTM(50, return_sequences=True, input_shape=(look_back, 1)))
model.add(LSTM(50))
model.add(Dense(1))
model.compile(loss='mean_squared_error', optimizer='adam')
model.fit(trainX, trainY, epochs=100, batch_size=1, verbose=2)

trainPredict = model.predict(trainX)
testPredict = model.predict(testX)
trainPredict = scaler.inverse_transform(trainPredict)
trainY = scaler.inverse_transform([trainY])
testPredict = scaler.inverse_transform(testPredict)
testY = scaler.inverse_transform([testY])

last_train = scaled_data[-look_back:]
next_3_months = []
input_seq = last_train.reshape((1, look_back, 1))
for _ in range(3):
    prediction = model.predict(input_seq)
    next_3_months.append(prediction[0, 0])
    input_seq = np.append(input_seq[:, 1:, :], [[prediction]], axis=1)

lstm_forecast = scaler.inverse_transform(np.array(next_3_months).reshape(-1, 1)).flatten()

# Tahminlerin ortalamasını al
if arima_forecast is not None:
    combined_forecast = (arima_forecast + lstm_forecast) / 2
else:
    combined_forecast = lstm_forecast

# Tahminleri görselleştirme
plt.figure(figsize=(12, 6))
plt.plot(data.index, data, label='Original Data')
plt.plot(pd.date_range(start=data.index[-1], periods=4, freq='MS')[1:], combined_forecast, label='Combined Forecast')
plt.legend()
plt.title('Combined ARIMA and LSTM Model Forecast')
plt.show()
