import serial
import time
import openpyxl

while True:
    ser = serial.Serial("COM4", 115200)
    arduinoData_string = ser.readline().decode(errors='replace')

    print(arduinoData_string)

    # Memisahkan data menjadi list
    data_list = arduinoData_string.strip().split(";")

    # Membagi list menjadi kelompok-kelompok berukuran 6
    grouped_data = [data_list[i:i+6] for i in range(0, len(data_list), 6)]

    if len(grouped_data) == 101:
        grouped_data = [[int(item) for item in inner_list if item] for inner_list in grouped_data]
        print(grouped_data)

        data = grouped_data

        # Nama file Excel
        file_path = 'data.xlsx'

        # Membuka file Excel
        try:
            wb = openpyxl.load_workbook(file_path)
        except FileNotFoundError:
            # Jika file tidak ditemukan, buat workbook baru
            wb = openpyxl.Workbook()

        # Mendapatkan sheet aktif (default)
        sheet = wb.active

        # Menambahkan data ke sheet
        for row_data in data:
            sheet.append(row_data)

        # Menyimpan perubahan ke file Excel
        wb.save(file_path)

    time.sleep(1)

    ser.close()
