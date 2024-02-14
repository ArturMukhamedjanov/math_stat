import math
import matplotlib.pyplot as plt
import openpyxl

def input_data(file_name):
    data = []
    maximum = 0
    minimum = 10000
    file = open(file_name)
    lines = file.readlines()
    for line in lines:
        temp = list(map(float, line.split(" ")))
        maximum = max(maximum, max(temp))
        minimum = min(minimum, min(temp))
        data.append(temp)
    return data, minimum, maximum

def count_elements(left, right, data, M):
    ans = 0
    for ar in data:
        for item in ar:
            if(item >= left and item < right):
                ans+=1
                continue
            if right >= M and item <=right and item >= left:
                ans+=1
                continue
    return ans

def make_intervals(data, n, M, h, a, summury_count, result_matrix):
    temp = []
    temp.append(a)
    temp.append(round(a + h,2))
    temp.append(round(a + h / 2, 2))
    count = count_elements(a, round(a + h,2), data, M)
    temp.append(count)
    temp.append(count/n)
    summury_count += count/n
    temp.append(round(summury_count, 2))
    temp.append(round(count/ n / h, 2))
    result_matrix.append(temp)
    if(a + h >= M):
        return result_matrix
    else:
        return make_intervals(data, n, M, h, round(a+h,2), summury_count, result_matrix)
    
def generate_gistogramm(result_matrix):
    bin_edges = [row[:2] for row in result_matrix]
    bin_values = [row[4] for row in result_matrix]
    bin_labels = [row[2] for row in result_matrix]  # Названия столбцов
    # Создаем гистограмму
    plt.bar(range(len(bin_edges)), bin_values, width=1, edgecolor='black')
    # Устанавливаем метки оси x и подписи столбцов
    plt.xticks(range(len(bin_edges)), bin_labels)
    # Добавляем подписи к осям
    plt.xlabel('Столбцы')
    plt.ylabel('Значение столбца')
    # Добавляем заголовок
    plt.title('Гистограмма на основе массива данных')
    plt.savefig('gistogramm.png')

def generate_grafic(data):
    plt.clf()
    x_values = [row[2] for row in data]
    y_values = [row[5] for row in data]
    # Создаем график
    plt.plot(x_values, y_values, 'b-')  # График с точками
    # Добавляем пунктирные линии от линий координат к каждой точке
    for i in range(len(data)):
        plt.plot([data[i][2], data[i][2]], [0, data[i][5]], 'k--', linewidth=0.5)
        plt.plot([0, data[i][2]], [data[i][5], data[i][5]], 'k--', linewidth=0.5)
    # Устанавливаем названия осей
    plt.xlabel('x')
    plt.ylabel('Fn(x)')
    plt.xticks(x_values)
    plt.yticks(y_values)
    plt.xlim(min(x_values), max(x_values))
    # Добавляем заголовок
    plt.title('График точек с пунктирными линиями к координатным осям')
    # Отображаем график
    plt.grid(True)
    plt.savefig('grafic.png')

def create_excel_table(data):
    wb = openpyxl.Workbook()
    ws = wb.active
    # Добавляем заголовки для столбцов
    headers = ['Lower Bound', 'Upper Bound', 'Middle', 'Frequency', 'Value 1', 'Value 2', 'Value 3']
    ws.append(headers)
    # Добавляем данные
    for row in data:
        ws.append(row)
    # Сохраняем файл
    wb.save("matrix.xlsx")

def count_math_estimate(data, n):
    temp = 0
    for ar in data:
        temp += ar[2] * ar[3]
    return temp / n 

def count_math_dispersion(data, n, math_estimate):
    temp = 0
    for ar in data:
        temp += ar[2] * ar[2] * ar[3]
    return temp / n - math_estimate * math_estimate

data, m, M = input_data("input_data.txt")
n = len(data) * 10
k = 1 + 3.32 * math.log10(n)
h = round((M - m) / k,2)
a = m
print(m, M, k, h)
result_matrix = make_intervals(data, n, M, h, a, 0, [])
generate_gistogramm(result_matrix)
generate_grafic(result_matrix)
create_excel_table(result_matrix)
math_estimate = count_math_estimate(result_matrix, n)
dispersion = round(count_math_dispersion(result_matrix, n, math_estimate),2)
point_dispersion = round(n / (n - 1) * dispersion,2)