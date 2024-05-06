import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
import tensorflow as tf
from tensorflow.keras.models import Sequential
from tensorflow.keras.layers import Dense

# Excel dosyasının yolu kullanıcıdan alınacak veya chat prompt ile istenecek
excel_path = input("Lütfen Excel dosyasının yolunu girin: ")

# Excel dosyasını oku
try:
    df = pd.read_excel(excel_path)
except FileNotFoundError:
    print("Belirtilen dosya bulunamadı. Lütfen geçerli bir yol girin.")
    exit()

# Bağımsız ve bağımlı değişkenlerin ayrılması
X = df.drop('target_column_name', axis=1)  # Bağımsız değişkenler
y = df['target_column_name']  # Bağımlı değişken

# Yapay Sinir Ağı Modelinin Oluşturulması
model = Sequential([
    Dense(64, activation='relu', input_shape=(X.shape[1],)),
    Dense(32, activation='relu'),
    Dense(1, activation='sigmoid')
])

model.compile(optimizer='adam',
              loss='binary_crossentropy',
              metrics=['accuracy'])

# Modelin Eğitilmesi
history = model.fit(X, y, epochs=20, batch_size=32, validation_split=0.2)

# Sonuçların Görselleştirilmesi
plt.plot(history.history['accuracy'], label='accuracy')
plt.plot(history.history['val_accuracy'], label='val_accuracy')
plt.xlabel('Epoch')
plt.ylabel('Accuracy')
plt.legend(loc='lower right')
plt.savefig('accuracy_plot.png')  # Grafik dosyasını kaydet
plt.show()

# Modelin kaydedilmesi
model.save('model.h5')  # Modeli kaydet
print("Model başarıyla kaydedildi.")
