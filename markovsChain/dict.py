import numpy as np
import pandas as pd
import openpyxl
from collections import defaultdict

def build_transition_matrix(text):
    # Русский алфавит
    alphabet = 'абвгдеёжзийклмнопрстуфхцчшщъыьэюя'
    text = ''.join([c for c in text.lower() if c in alphabet])

    # Инициализируем словарь для хранения частот переходов
    transition_counts = defaultdict(lambda: defaultdict(int))
    total_counts = defaultdict(int)

    # Подсчитываем частоты переходов
    prev_char = None
    for char in text:
        if prev_char is not None:
            transition_counts[prev_char][char] += 1
            total_counts[prev_char] += 1
        prev_char = char

    # Создаем список всех уникальных букв
    unique_chars = sorted(set(text))

    # Инициализируем матрицу переходных вероятностей
    size = len(unique_chars)
    matrix = np.zeros((size, size))

    # Заполняем матрицу вероятностей
    for i, char1 in enumerate(unique_chars):
        for j, char2 in enumerate(unique_chars):
            if total_counts[char1] > 0:
                matrix[i][j] = transition_counts[char1][char2] / total_counts[char1]

    # Преобразуем в DataFrame
    df = pd.DataFrame(matrix, index=unique_chars, columns=unique_chars)
    return df

# Пример использования
file = open('Example.txt', 'r')
text = file.read()
transition_df = build_transition_matrix(text)

# Запись в Excel файл
transition_df.to_excel("transition_matrix.xlsx")

print("Матрица переходных вероятностей сохранена в файл transition_matrix.xlsx")
