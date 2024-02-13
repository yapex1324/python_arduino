import serial
import time
import numpy as np
import pandas as pd
from scipy.fft import fft, ifft
import matplotlib.pyplot as plt
import pickle
from sklearn.pipeline import Pipeline

# Enable interactive mode
plt.ion()

# Membuat 6 subplot
fig, axs = plt.subplots(2, 3, figsize=(12, 6))

# Membuat rentang nilai untuk nilai y-tick
y_ticks_range = np.linspace(100000, 900000, 9)

while True:
    ser = serial.Serial("COM3", 115200)
    arduinoData_string = ser.readline().decode(errors='replace')

    print(arduinoData_string)

    # Memisahkan data menjadi list
    data_list = arduinoData_string.strip().split(";")

    # Membagi list menjadi kelompok-kelompok berukuran 6
    grouped_data = [data_list[i:i+6] for i in range(0, len(data_list), 6)]

    if len(grouped_data) == 51:
        grouped_data = [[int(item) for item in inner_list if item] for inner_list in grouped_data]
        print(grouped_data)

        file_path = r'C:\Users\moh_y\project_all\02_ekg\data_df.xlsx'

        data = grouped_data
        
        df = pd.DataFrame(data)

        df.to_excel(file_path, index=False)

        ####################    tampilkan gambar yang sudah diparsing
        for i in range(2):
            for j in range(3):
                subplot_data = [grouped_data[k][i * 3 + j] for k in range(len(grouped_data)) if i * 3 + j < len(grouped_data[k])]
                axs[i, j].clear()  # Clear the previous plot
                axs[i, j].plot(subplot_data)
                axs[i, j].set_yticks(y_ticks_range)  # Atur nilai y-tick

        # Menampilkan plot
        plt.tight_layout()
        plt.draw()
        plt.pause(0.1)  # Add a small pause to allow the plot to update

        ##############################  fft cek
        df = pd.read_excel(r'C:\Users\moh_y\project_all\02_ekg\data_df.xlsx')

        # Mengambil data dari setiap kolom
        data_columns = [df[col].tolist() for col in df.columns]

        # Menerapkan transformasi Fourier dan filtering pada setiap kolom
        filtered_data_columns = []

        for i, col_data in enumerate(data_columns, start=1):
            # Mengaplikasikan transformasi Fourier
            fft_result_col = fft(col_data)

            # Menghapus komponen frekuensi tinggi dengan zeroing pada frekuensi tertentu
            cutoff_frequency = 0.1  # Frekuensi cutoff yang diinginkan
            fft_result_filtered_col = fft_result_col.copy()
            num_elements = len(fft_result_col)
            cutoff_index = int(cutoff_frequency * num_elements)
            fft_result_filtered_col[cutoff_index:num_elements-cutoff_index] = 0

            # Mengaplikasikan inverse transform untuk mendapatkan data hasil filtering
            filtered_data_col = ifft(fft_result_filtered_col).real

            # Menambahkan data hasil filtering ke dalam list
            filtered_data_columns.append(filtered_data_col)

            # Menambahkan data hasil filtering ke dalam DataFrame
            df[f'Filtered Col {i}'] = filtered_data_col

        # Menyimpan hasil perubahan ke dalam file Excel baru
        output_file_path = r'C:\Users\moh_y\project_all\02_ekg\filtered_data.xlsx'
        df.to_excel(output_file_path, index=False)

        ################################ predict nilai svm
        #   load model coba
        filename = r'C:\Users\moh_y\project_all\02_ekg\model_svm.sav'
        loaded_model = pickle.load(open(filename, 'rb'))

        data_test_arduino = pd.read_excel(r"C:\Users\moh_y\project_all\02_ekg\filtered_data.xlsx")
        data_test_arduino = data_test_arduino[['Filtered Col 1','Filtered Col 2','Filtered Col 3','Filtered Col 4'
                                            , 'Filtered Col 5', 'Filtered Col 6']]

        hasil = loaded_model.predict(data_test_arduino)

        # Menghitung banyaknya nilai 0 dan 1
        count_0 = np.sum(np.array(hasil) == 0)
        count_1 = np.sum(np.array(hasil) == 1)

        # Menampilkan hasil
        if count_0 >= count_1:
            print("grafik normal")
        if count_0 < count_1:
            print("grafik tidak normal")
        

    time.sleep(1)

    ser.close()
