import openpyxl as Workbook
import os
from pymorphy2 import MorphAnalyzer
import string
from config import *
import sys

def main(param):
    if os.path.exists('data.xlsx') and os.path.exists('config.py'):
        print("Файл обнаружен. Начинаем обработку...")
        wb = Workbook.load_workbook('data.xlsx')
        ws = wb.active
        column = ws['A']
        i = 0

        for cell in column:
            if cell.value is not None:
                if i == 0:
                    i += 1
                    if param == 'debug':
                        ws['B' + str(i)].value = 'Lemma'
                    ws['C' + str(i)].value = 'Type'
                    ws['D' + str(i)].value = 'Mood'
                    ws['E' + str(i)].value = 'Keywords'
                    continue
                i += 1
                print(i)
                prep = prepare(str(cell.value).lower())
                lemtext = prep[0]
                char = prep[1]

                if param == 'debug':
                    if prep[1] == 'skip':
                        ws['B' + str(i)].value = str(lemtext)
                        ws['C' + str(i)].value = 'skiped'
                        ws['D' + str(i)].value = ''
                        ws['E' + str(i)].value = prep[2]
                        continue
                    elif prep[1] == 'undef':
                        ws['B' + str(i)].value = str(lemtext)
                        ws['C' + str(i)].value = 'undef'
                        ws['D' + str(i)].value = ''
                        ws['E' + str(i)].value = ''
                        continue
                    else:
                        if len(lemtext) > 3:
                            result = mood_define(lemtext, char)
                            if result[0] > 0:
                                ws['B' + str(i)].value = str(lemtext)
                                ws['c' + str(i)].value = result[1]
                                ws['d' + str(i)].value = 'positive'
                                ws['e' + str(i)].value = result[2]
                                ws['f' + str(i)].value = result[0]
                            elif result[0] < 0:
                                ws['B' + str(i)].value = str(lemtext)
                                ws['c' + str(i)].value = result[1]
                                ws['d' + str(i)].value = 'negative'
                                ws['e' + str(i)].value = result[2]
                                ws['f' + str(i)].value = result[0]
                            else:
                                ws['B' + str(i)].value = str(lemtext)
                                ws['c' + str(i)].value = result[1]
                                ws['d' + str(i)].value = 'undef'
                                ws['e' + str(i)].value = result[2]
                                ws['f' + str(i)].value = result[0]
                else:
                    if prep[1] == 'skip':
                        ws['B' + str(i)].value = str(lemtext)
                        ws['C' + str(i)].value = 'skiped'
                        ws['D' + str(i)].value = ''
                        ws['E' + str(i)].value = prep[2]
                        continue
                    elif prep[1] == 'undef':
                        ws['B' + str(i)].value = str(lemtext)
                        ws['C' + str(i)].value = 'undef'
                        ws['D' + str(i)].value = ''
                        ws['E' + str(i)].value = ''
                        continue
                    else:
                        if len(lemtext) > 3:
                            result = mood_define(lemtext, char)
                            if result[0] > 0:
                                ws['B' + str(i)].value = str(lemtext)
                                ws['c' + str(i)].value = result[1]
                                ws['d' + str(i)].value = 'positive'
                                ws['e' + str(i)].value = result[2]
                                ws['f' + str(i)].value = result[0]
                            elif result[0] < 0:
                                ws['B' + str(i)].value = str(lemtext)
                                ws['c' + str(i)].value = result[1]
                                ws['d' + str(i)].value = 'negative'
                                ws['e' + str(i)].value = result[2]
                                ws['f' + str(i)].value = result[0]
                            else:
                                ws['B' + str(i)].value = str(lemtext)
                                ws['c' + str(i)].value = result[1]
                                ws['d' + str(i)].value = 'undef'
                                ws['e' + str(i)].value = result[2]
                                ws['f' + str(i)].value = result[0]

                wb.save(r'data_done.xlsx')

        print("Обработка выполнена! Проверьте файл data_done.xlsx")
    else:
        print("Файл с данными (data.xlsx) или config.py не найден. Пожалуйста проверьте их наличие в папке с программой!")

def mood_define(lemtext, char):

    index = 0
    sum_words = ''
    words = []

    for key, value in keywords[char].items():
        if key in lemtext:
            words.append(key)
            if sum_words == '':
                sum_words = key
            else:
                sum_words += ', ' + key
        else:
            continue

    if len(words) == 1:
        for word in words:
            if keywords[char][word] == 'positive':
                index += 1
            elif keywords[char][word] == 'negative':
                index -= 1
            elif keywords[char][word][:4] == 'near':
                for item in near[char]:
                    if lemtext[lemtext.index(word) - 1] == item or \
                            lemtext[lemtext.index(word) - 2] == item:
                        if keywords[char][word][5:] == 'negative':
                            index -= 1
                        elif keywords[char][word][5:] == 'positive':
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
                            index -= 1
                        else:
                            index += 1
                    continue
            elif keywords[char][word] == 'together':
                for item in near[char]:
                    if lemtext[lemtext.index(word) - 1] == item or \
                            lemtext[lemtext.index(word) - 2] == item:
                        for key, value in presence[char].items():
                            if key in lemtext:
                                if value == 'negative':
                                    index -= 1
                                else:
                                    index += 1
                            continue
            else:
                index += 0

    elif len(words) > 1:
        for word in words:
            if keywords[char][word] == 'positive':
                index += 1
            elif keywords[char][word] == 'negative':
                index -= 1
            elif keywords[char][word][:4] == 'near':
                for item in near[char]:
                    if lemtext[lemtext.index(word) - 1] == item or \
                            lemtext[lemtext.index(word) - 2] == item:
                        if keywords[char][word][5:] == 'negative':
                            index -= 1
                        elif keywords[char][word][5:] == 'positive':
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
                            index -= 1
                        else:
                            index += 1
                    continue
            elif keywords[char][word] == 'together':
                for item in near[char]:
                    if lemtext[lemtext.index(word) - 1] == item or \
                            lemtext[lemtext.index(word) - 2] == item:
                        for key, value in presence[char].items():
                            if key in lemtext:
                                if value == 'negative':
                                    index -= 1
                                else:

                                    index += 1
                            continue
            else:
                index += 0

    return index, char, sum_words


def prepare(text):
    morph = MorphAnalyzer()
    tokens = []
    skip_status = 0
    char = ''
    keyword = ''
    arr = []
    new_text=''

    spec_chars = string.punctuation + '\xa0«»\t—…'
    text = text.replace('-', ' ')
    text = text.replace('/', ' ')
    text = text.replace('.', ' ')
    text = text.replace(',', ' ')

    text = "".join([ch for ch in text if ch not in spec_chars])

    for word in skip:
        if word in text:
            skip_status = 1
            char = 'skip'
            keyword = word
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
                        else:
                            continue
                else:
                    arr.append(v)

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
    for param in sys.argv:
        print (param)
    main(param[1])