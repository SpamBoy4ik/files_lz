import pandas as pd
import matplotlib.pyplot as plt
import csv

df = pd.read_parquet('titanic.parquet')
df.to_csv('titanic.csv')

third_class = 0 # сколько было пассажиров 3-го класса
third_class_survived = 0 # сколько выжило пассажиров 3-го класса
second_class = 0
second_class_survived = 0
first_class = 0
first_class_survived = 0

with open('titanic.csv', 'r') as stat:
    data = csv.reader(stat)

    for column in data:
        match column[3]:
            case '3':
                third_class += 1
                if column[2] == '1':
                    third_class_survived += 1
            case '2':
                second_class += 1
                if column[2] == '1':
                    second_class_survived += 1
            case '1':
                first_class += 1
                if column[2] == '1':
                    first_class_survived += 1

labels = ['Первый класс', 'Второй класс', 'Третий класс']
survived = [first_class_survived, second_class_survived, third_class_survived]
died = [first_class - first_class_survived, second_class - second_class_survived, third_class - third_class_survived]


'''Создание гистограмммы'''

plt.style.use('fivethirtyeight')
plt.title('Выживаемости пассажиров Титаника')
plt.bar(labels, died, label = 'Умерли')
plt.bar(labels, survived, bottom=died, label = 'Выжили')
plt.legend(loc='upper center')
plt.show()
