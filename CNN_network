import tensorflow as tf
from tensorflow.keras.models import Sequential
from tensorflow.keras.layers import Conv2D, MaxPooling2D, Flatten, Dense, Dropout

# Modelinizi tanımlayın
model = Sequential([
    Conv2D(32, (3, 3), activation='relu', input_shape=(64, 64, 3)),
    MaxPooling2D(2, 2),
    Conv2D(64, (3, 3), activation='relu'),
    MaxPooling2D(2, 2),
    Conv2D(128, (3, 3), activation='relu'),
    MaxPooling2D(2, 2),
    Flatten(),
    Dense(512, activation='relu'),
    Dropout(0.5),
    Dense(1, activation='sigmoid')  # 1 çıkış nöronu, var/yok için
])

# Modeli derleyin
model.compile(optimizer='adam', loss='binary_crossentropy', metrics=['accuracy'])

# Modelinizi eğitmeye hazırsınız, bu adımı gerçekleştirmek için önce veri setinizi hazırlamanız gerekmekte
# Eğitim verilerinizi ve etiketlerinizi kullanarak modeli eğitin
# model.fit(x_train, y_train, epochs=10, batch_size=32, validation_data=(x_val, y_val))
