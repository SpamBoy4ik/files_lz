import docx
from docx import Document
import pandas as pd
import matplotlib.pyplot as plt


doc = docx.Document("lion.docx")
symbols = [',', '.', '!', '?', ':', ';', '"', '*', '(', ')', '[', ']', '«', '»', '_', '—', '-', '0', '1', '2', '3', '4', '5', '6', '7', '8', '9']
input_data = []
for paragraph in doc.paragraphs:
    input_data.append(paragraph.text)

# делаем нижний регистр для всех слов
for i in range(len(input_data)):
    input_data[i] = input_data[i].lower()
###

# разбиваем предложения на слова
text = []
for i in range(len(input_data)):
    text.extend(input_data[i].split())
###
 
# разбиваем слова на буквы
letters = []
for word in text:
    for letter in range(len(word)):
        letters.extend(word[letter])
###

# убераем лишние символы из СЛОВ
dirty_word = False
for word in text:
    clear_word = word
    for symbol in symbols:
        if symbol in word:
            dirty_word = True
            clear_word = clear_word.replace(symbol, '')
    if dirty_word:
        dirty_word = False
        text.insert(text.index(word), clear_word)
        text.remove(word)
###

# убераем лишние символы из списка БУКВ
clear_letters = [] # список букв без symbols
skip = False
for letter in letters:
    for symbol in symbols:
        if letter == symbol:
            skip = True
            break
    if skip:
        skip = False
        pass
    else:
        clear_letters.extend(letter)
###

# делаем список из встречаемых СЛОВ
key_words = text
key_words = set(key_words)
key_words = list(key_words)
###

# делаем список из встречаемых БУКВ
key_letters = set(clear_letters)
key_letters = list(key_letters)
###

# считаем повторения СЛОВ
words_repetition = dict.fromkeys(key_words, 0)
for word in text:
    words_repetition[word] += 1
###

words_repetition.pop('')

# считаем повторения БУКВ
letters_repetition = dict.fromkeys(key_letters, 0)
for letter in clear_letters:
    letters_repetition[letter] += 1
###


'''Создание таблицы Excel'''

words_amount = [] # повторения каждого слова
words_amount_sum = 0 # кол-во всех слов
for key in words_repetition:
    words_amount.append(words_repetition[key])
    words_amount_sum += words_repetition[key]

words_rate = [] # повторения каждого слова в %
for i in range(len(words_amount)):
    words_rate.append(words_amount[i] / words_amount_sum * 100 )

df = pd.DataFrame(list(words_repetition.items()), columns=['Слово', 'Частота встречи, раз'])
df['Частота встречи, %'] = words_rate

df.to_excel('result.xlsx', index = False)


'''Создание гистограмммы'''

plt.bar(letters_repetition.keys(), letters_repetition.values())
plt.xlabel("Буква")
plt.ylabel("Встречаемость, раз")
plt.show()
