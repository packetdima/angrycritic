import openpyxl as Workbook
import os
import pandas as pd
import numpy as np
from pymorphy3 import MorphAnalyzer
import string
from config import *
import sys
import time
from tqdm import tqdm


def main():
    start_time = time.time()

    if os.path.exists('data.xlsx') and os.path.exists('config.py'):
        print("Файл обнаружен. Начинаем обработку...")

        df = pd.read_excel('data.xlsx', index_col=None, header=0)
        new_df = df.assign(Lemma='', Type='', Type_Keywords='', Mood='', Mood_Keywords='', Index='', Test = '')
        arr = new_df.to_numpy()

        new_arr = []

        for item in tqdm(arr):
            time.sleep(0.001)
            prep = prepare(str(item[0]).lower())
            lemtext = prep[0]
            char = prep[1]
            type_keywords = prep[2]

            result = mood_define(lemtext, char)
            i = result[0]
            mood = result[1]
            mood_keywords = result[2]

            item[2] = lemtext
            item[3] = char
            item[4] = type_keywords
            item[5] = mood
            item[6] = mood_keywords
            item[7] = i
            new_arr.append(item)


        resultdf = pd.DataFrame(new_arr)
        resultdf.columns= ['Text', 'Answer', 'Lemma', 'Type', 'Type_Keywords', 'Mood', 'Mood_Keywords', 'Index']
        resultdf.to_excel('data_done.xlsx')

        end_time = time.time()  # время окончания выполнения
        execution_time = end_time - start_time  # вычисляем время выполнения

        print(f"Время выполнения программы: {execution_time} секунд")

        print("Обработка выполнена! Проверьте файл data_done.xlsx")
    else:
        print("Файл с данными (data.xlsx) или config.py не найден. Пожалуйста проверьте их наличие в папке с программой!")

def mood_define(lemtext, char):

    index = 0
    sum_words = ''
    words = []
    new_sum_words = []

    if char != 'undef' and char != 'skip':
        for key, value in keywords[char].items():
            if key in lemtext:
                words.append(key)
                if sum_words == '':
                    sum_words = key
                else:
                    sum_words += ', ' + key
            else:
                continue
    elif char == 'undef' or char == 'skip':
        index = 0

    if len(words) == 1:
        for word in words:
            if keywords[char][word] == 'positive':
                new_sum_words.append(word)
                index += 1
            elif keywords[char][word] == 'negative':
                new_sum_words.append(word)
                index -= 1
            elif keywords[char][word][:4] == 'near':
                for item in near[char]:
                    if lemtext[lemtext.index(word) - 1] == item or \
                            lemtext[lemtext.index(word) - 2] == item:
                        if keywords[char][word][5:] == 'negative':
                            new_sum_words.append(item + ' + ' + word)
                            index -= 1
                        elif keywords[char][word][5:] == 'positive':
                            new_sum_words.append(item + ' + ' + word)
                            index += 1
                        else:
                            index += 0
                        break
                    else:
                        continue
            elif keywords[char][word] == 'presence':
                for key, value in presence[char].items():
                    if key in lemtext:
                        if value == 'negative':
                            new_sum_words.append(word + ' + ' + key)
                            index -= 1
                        else:
                            new_sum_words.append(word + ' + ' + key)
                            index += 1
                    continue
            elif keywords[char][word] == 'together':
                for item in near[char]:
                    if lemtext[lemtext.index(word) - 1] == item or \
                            lemtext[lemtext.index(word) - 2] == item:
                        for key, value in presence[char].items():
                            if key in lemtext:
                                if value == 'negative':
                                    new_sum_words.append(item + ' + ' + word + ' + ' + key)
                                    index -= 1
                                else:
                                    new_sum_words.append(item + ' + ' + word + ' + ' + key)
                                    index += 1
                            continue
            else:
                index += 0

    elif len(words) > 1:
        for word in words:
            if keywords[char][word] == 'positive':
                new_sum_words.append(word)
                index += 1
            elif keywords[char][word] == 'negative':
                new_sum_words.append(word)
                index -= 1
            elif keywords[char][word][:4] == 'near':
                for item in near[char]:
                    if lemtext[lemtext.index(word) - 1] == item or \
                            lemtext[lemtext.index(word) - 2] == item:
                        if keywords[char][word][5:] == 'negative':
                            new_sum_words.append(item + ' + ' + word)
                            index -= 1
                            break
                        elif keywords[char][word][5:] == 'positive':
                            new_sum_words.append(item + ' + ' + word)
                            index += 1
                            break
                        else:
                            index += 0
                            break
                    else:
                        continue
            elif keywords[char][word] == 'presence':
                for key, value in presence[char].items():
                    if key in lemtext:
                        if value == 'negative':
                            new_sum_words.append(word + ' + ' + key)
                            index -= 1
                        else:
                            new_sum_words.append(word + ' + ' + key)
                            index += 1
                    continue
            elif keywords[char][word] == 'together':
                for item in near[char]:
                    if lemtext[lemtext.index(word) - 1] == item or \
                            lemtext[lemtext.index(word) - 2] == item:
                        for key, value in presence[char].items():
                            if key in lemtext:
                                if value == 'negative':
                                    new_sum_words.append(item + ' + ' + word + ' + ' + key)
                                    index -= 1
                                else:
                                    new_sum_words.append(item + ' + ' + word + ' + ' + key)
                                    index += 1
                            continue
            else:
                index += 0

    if index > 0:
        mood = 'positive'
    elif index < 0:
        mood = 'negative'
    else:
        mood = 'undef'
    return index, mood, new_sum_words


def prepare(text):
    morph = MorphAnalyzer()
    tokens = []
    skip_status = 0
    char = ''
    keyword = []
    arr = []
    new_text=''

    spec_chars = string.punctuation + '\xa0«»\t—…'
    text = str(text).replace('-', ' ')
    text = str(text).replace('/', ' ')
    text = str(text).replace('.', ' ')
    text = str(text).replace(',', ' ')

    text = "".join([ch for ch in text if ch not in spec_chars])

    for word in skip:
        if word in str(text):
            skip_status = 1
            char = 'skip'
            keyword.append(word)
            break
        else:
            skip_status = 0
            continue
    if skip_status == 0:
        for token in text.split():
            token = token.strip()
            token = morph.normal_forms(token)[0]
            new_text += ' ' + token
            tokens.append(token)

        for k, v in char_keywords.items():
            if k in tokens:
                if v == 'presence':
                    for key, value in char_presence.items():
                        if key in tokens:
                            arr.append(value)
                            keyword.append(k)
                            keyword.append(key)
                        else:
                            continue
                else:
                    arr.append(v)
                    keyword.append(k)

        if len(arr) > 0:
            char = {i: arr.count(i) for i in arr};
            char = list(char.keys())[0]
        else:
            char = 'undef'

        if char != 'undef' and char != 'skip':
            for key, value in replace['general'].items():
                if key in new_text:
                    new_text = new_text.replace(key,value)
                else:
                    continue
            for key, value in replace[char].items():
                if key in new_text:
                    new_text = new_text.replace(key,value)
                else:
                    continue

    return tokens, char, keyword


if __name__ == '__main__':
    main()