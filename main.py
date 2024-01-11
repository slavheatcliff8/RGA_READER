import tkinter as tk
from tkinter import filedialog
import matplotlib.pyplot as plt
from exel_editor import raw_input
averages_probe = []

class ProbeDataExtractor:
    def __init__(self):
        self.file_path = ""
        self.probes = []

    def choose_file(self):
        root = tk.Tk()
        root.withdraw()
        self.file_path = filedialog.askopenfilename(title="Выберите файл с данными")

    def extract_data(self):
        with open(self.file_path, 'r') as file:
            lines = file.readlines()

        start_index = 0
        for i, line in enumerate(lines):
            if any(char.isdigit() or char == '.' or char == '-' or char == ',' for char in line):
                start_index = i
                break

        for line in lines[start_index:]:
            data = [float(value.replace(',', '.')) if float(value.replace(',', '.')) >= 0 else float(value.replace(',', '.')) for value in line.split()]
            probe_number = int(data[0])
            probe_values = data[1:]
            probe = {'Номер пробы': probe_number, 'Значения пробы': probe_values}
            if 1401 <= probe['Номер пробы'] <= 1500: 
                self.probes.append(probe)
        return (self.probes)


# Пример использования класса
if __name__ == "__main__":
    data_extractor = ProbeDataExtractor()

    # Выбор файла
    data_extractor.choose_file()

    # Извлекаем данные из файла
    probes_data = data_extractor.extract_data()
    raw_input(probes_data,"data")

    data_extractor.choose_file()
    probes_reference = data_extractor.extract_data()
    raw_input(probes_reference,"reference")


