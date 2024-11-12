import yaml
from yaml.loader import SafeLoader
import os
import pandas as pd
from pymorphy3 import MorphAnalyzer
import string
import time
from tqdm import tqdm
from nltk.corpus import stopwords
from nltk.probability import FreqDist
import jarowinkler


def main():
    global skip
    global types_keywords
    global replace
    global rustopwords
    rustopwords = stopwords.words('russian')
    start_time = time.time()
    active_types = []

    if os.path.exists('1.xlsx'):
        print("Файл обнаружен. Начинаем обработку...")

        df = pd.read_excel('1.xlsx', index_col=None, header=0)
        new_df = df.assign(Lemma='', Type='', Type_Keywords='', Mood='', Mood_Keywords='', Index='')
        arr = new_df.to_numpy()

        with open('./data/skip.yaml', encoding='utf8') as f:
            skip = yaml.load(f, Loader=SafeLoader)

        with open('./data/types.yaml', encoding='utf8') as f:
            types_keywords = yaml.load(f, Loader=SafeLoader)

        # with open('./data/replace.yaml', encoding='utf8') as f:
        #     replace = yaml.load(f, Loader=SafeLoader)

        for item in types_keywords:
            # if item != 'undef' and item != 'presence_kw':
            active_types.append(item)
            # else:
            #     continue

        prep = prepare(arr)
        prepdf = pd.DataFrame(prep[0])
        countdf = pd.DataFrame(prep[1])
        prepdf.columns= ['Text', 'Lemma', 'Type', 'Type_Keywords', 'Mood', 'Mood_Keywords', 'Index']
        countdf.columns= ['Слово', 'Частота']
        filtred = prepdf.query("Type in ('skip', 'undef')").to_numpy()
        df1 = pd.DataFrame(filtred)

        frames = []
        data = {}
        writer = pd.ExcelWriter('data_done.xlsx', engine='xlsxwriter')
        for item in active_types:
            val = str(item).lower()
            df2 = prepdf.query("Type == @val")
            arr2 = df2.to_numpy()
            new_arr = mood_define(arr2, item)
            mood_df = pd.DataFrame(new_arr[0])
            frames.append(mood_df)
            if len(new_arr[1]) > 0:
                data[item] = new_arr[1]

        frames.append(df1)
        result = pd.concat(frames)

        result.columns= ['Text', 'Lemma', 'Type', 'Type_Keywords', 'Mood', 'Mood_Keywords', 'Index']
        result.to_excel(writer, sheet_name='Total', index=False)
        countdf.to_excel(writer, sheet_name='Частота по документу', index=False)

        for k, v in data.items():
            r = pd.DataFrame(v)
            r.columns = ['Слово', 'Частота']
            r.to_excel(writer, sheet_name=k, index=False)

        writer.close()

        end_time = time.time()
        execution_time = end_time - start_time

        print(f"Время выполнения программы: {execution_time} секунд")
        print("Обработка выполнена! Проверьте файл data_done.xlsx")
    else:
        print("Файл с данными (data.xlsx) не найден. Пожалуйста проверьте их наличие в папке с программой!")


def prepare(arr):
    print('Начинаем предварительную обработку...')
    array = arr
    morph = MorphAnalyzer()
    textcount = []

    for item in tqdm(array):
        time.sleep(0.001)
        tokens = []
        keyword = []
        type_arr = []
        new_text = ''
        char = ''
        skip_status = 0

        spec_chars = string.punctuation + '\xa0«»\t—…'
        text = str(item[0]).lower().replace('-', ' ')
        text = str(item[0]).lower().replace('/', ' ')
        text = str(item[0]).lower().replace('.', ' ')
        text = str(item[0]).lower().replace(',', ' ')
        text = str(item[0]).lower().replace('¶', ' ')
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
                if token and token not in rustopwords:
                    if not token.isdigit():
                        textcount.append(token)
                new_text += ' ' + token
                tokens.append(token)

            for k, v in types_keywords.items():
                for i in v:
                    y = 0
                    for word in tokens:
                        if word is not None and len(word) > 2:
                            if i.isalpha() == True:  # Тональность определяется по одному слову
                                simular = jarowinkler.jarowinkler_similarity(i, word)
                                if simular >= 0.9:
                                    type_arr.append(k)
                                    keyword.append(i)
                            else:
                                kwrd = i.split('_')
                                simular = jarowinkler.jarowinkler_similarity(word, kwrd[1])
                                if simular >= 0.9:
                                    if len(kwrd[0]) > 0 and len(kwrd[2]) == 0:  # Только near
                                        result = check_near(tokens, y, kwrd[0])
                                        if result[0] == True:
                                            type_arr.append(k)
                                            keyword.append(
                                                str(result[1]).replace('"', '').replace('[', '').replace(']',
                                                                                                         '').replace(
                                                    '\\', '').replace("'", "") + ' *' + kwrd[1] + '*')

                                    elif len(kwrd[0]) == 0 and len(kwrd[2]) > 0:  # Только presence
                                        res = check_presence(tokens, kwrd[2])
                                        if res[0] == True:

                                            type_arr.append(k)
                                            keyword.append(
                                                '*' + kwrd[1] + '*' + str(res[1]).replace('"', '').replace('[', '').replace(
                                                    ']', '').replace('\\', '').replace("'", ""))
                                    elif len(kwrd[0]) > 0 and len(kwrd[2]) > 0:  # Near & Presence
                                        result = check_near(tokens, y, kwrd[0])
                                        if result[0] == True:
                                            res = check_presence(tokens, kwrd[2])
                                            if res[0] == True:
                                                type_arr.append(k)
                                                keyword.append(
                                                    str(result[1]).replace('"', '').replace('[', '').replace(']',
                                                                                                             '').replace(
                                                        '\\', '').replace("'", "") + ' *' + kwrd[1] + '*' + str(
                                                        res[1]).replace('"', '').replace('[', '').replace(']',
                                                                                                          '').replace(
                                                        '\\', '').replace("'", ""))
                        y += 1

            if len(type_arr) > 0:
                char = {i: type_arr.count(i) for i in type_arr};
                char = list(char.keys())[0]
            else:
                char = 'undef'

        item[1] = tokens
        item[2] = char
        item[3] = keyword

    fdist = FreqDist(textcount)
    print('Предварительная обработка завершена!')

    return arr, fdist.most_common()


def mood_define(arr, char):

    print('Начинаем процесс определения тональности - ' + char)

    neg = []
    pos = []
    textcount = []

    path = './data/mood_keywords_' + char + '.yaml'

    with open(path, encoding='utf8') as f:
        keywords = yaml.safe_load(f)

    for key, value in keywords.items():         # собираем ключевые слова из yaml в переменные
        if key == 'negative':
            neg = value
        elif key == 'positive':
            pos = value

    for item in tqdm(arr):
        index = 0
        lemtext = item[1]
        sum_words = []
        y = 0

        for i in lemtext:
            if len(i) > 2:
                if i and i not in rustopwords:
                    if not i.isdigit():
                        textcount.append(i)
                for wrd in neg:
                    if wrd is not None:
                        if wrd.isalpha() == True:                                       # Тональность определяется по одному слову
                            simular = jarowinkler.jarowinkler_similarity(i, wrd)
                            if simular >= 0.95:
                                sum_words.append('*' + wrd + '*')
                                index -= 1
                        else:
                            kwrd = wrd.split('_')
                            simular = jarowinkler.jarowinkler_similarity(i, kwrd[1])
                            if simular >= 0.95:
                                if len(kwrd[0]) > 0 and len(kwrd[2]) == 0:                  # Только near
                                    result = check_near(lemtext, y, kwrd[0])
                                    if result[0]== True:
                                        sum_words.append(str(result[1]).replace('"','').replace('[','').replace(']','').replace('\\','').replace("'","") + ' *' + kwrd[1] + '*')
                                        index -= 1
                                elif len(kwrd[0]) == 0 and len(kwrd[2]) > 0:               # Только presence
                                    res = check_presence(lemtext, kwrd[2])
                                    if res[0]== True:
                                        sum_words.append('*' + kwrd[1] + '*' + str(res[1]).replace('"','').replace('[','').replace(']','').replace('\\','').replace("'",""))
                                        index -= 1
                                elif len(kwrd[0]) > 0 and len(kwrd[2]) > 0:                # Near & Presence
                                    result = check_near(lemtext, y, kwrd[0])
                                    if result[0]== True:
                                        res = check_presence(lemtext, kwrd[2])
                                        if res[0] == True:
                                            sum_words.append(str(result[1]).replace('"','').replace('[','').replace(']','').replace('\\','').replace("'","") + ' *' + kwrd[1] + '*' + str(res[1]).replace('"','').replace('[','').replace(']','').replace('\\','').replace("'",""))
                                            index -= 1
                                        # else:
                                        #     sum_words.append(str(result).replace('"','').replace('[','').replace(']','').replace('\\','') + ' *' + i + '*' + str(res[1]).replace('"','').replace('[','').replace(']','').replace('\\',''))

                for wrd in pos:
                    if wrd is not None:
                        if wrd.isalpha() == True:  # Тональность определяется по одному слову
                            simular = jarowinkler.jarowinkler_similarity(i, wrd)
                            if simular >= 0.95:
                                sum_words.append('*' + wrd + '*')
                                index += 1
                                continue
                        else:
                            kwrd = wrd.split('_')
                            simular = jarowinkler.jarowinkler_similarity(i, kwrd[1])
                            if simular >= 0.95:
                                if len(kwrd[0]) > 0 and len(kwrd[2]) == 0:  # Только near
                                    result = check_near(lemtext, y, kwrd[0])
                                    if result[0]== True:
                                        sum_words.append(str(result[1]).replace('"','').replace('[','').replace(']','').replace('\\','').replace("'","") + ' *' + kwrd[1] + '*')
                                        index += 1
                                elif len(kwrd[0]) == 0 and len(kwrd[2]) > 0:  # Только presence
                                    res = check_presence(lemtext, kwrd[2])
                                    if res[0]== True:
                                        sum_words.append('*' + kwrd[1] + '*' + str(res[1]).replace('"','').replace('[','').replace(']','').replace('\\','').replace("'",""))
                                        index += 1
                                elif len(kwrd[0]) > 0 and len(kwrd[2]) > 0:  # Near & Presence
                                    result = check_near(lemtext, y, kwrd[0])
                                    if result[0]== True:
                                        res = check_presence(lemtext, kwrd[2])
                                        if res[0] == True:
                                            sum_words.append(str(result[1]).replace('"','').replace('[','').replace(']','').replace('\\','').replace("'","") + ' *' + kwrd[1] + '*' + str(res[1]).replace('"','').replace('[','').replace(']','').replace('\\','').replace("'",""))
                                            index += 1
                                        # else:
                                        #     sum_words.append(str(result).replace('"','').replace('[','').replace(']','').replace('\\','') + ' *' + i + '*' + str(
                                        #         res[1]).replace('"','').replace('[','').replace(']','').replace('\\',''))

            y += 1
        if index > 0:
            mood = 'positive'
        elif index < 0:
            mood = 'negative'
        else:
            mood = 'undef'

        item[4] = mood
        item[5] = sum_words
        item[6] = index

    fdist = FreqDist(textcount)

    return arr, fdist.most_common()


def check_near (lemtext, i, keywords):
    text = []
    status = False

    if keywords.find('+') >= 0 and keywords.find('-') >= 0:
        if keywords.find('+') == 0:
            plus = keywords[1:keywords.find('-')].split(',')
            minus = keywords[keywords.find('-') + 1:].split(',')
        else:
            plus = keywords[keywords.find('-') + 1:].split(',')
            minus = keywords[1:keywords.find('-')].split(',')

        if i == 1:
            if lemtext[i - 1] in plus:
                text.append('+' + lemtext[i - 1] + ' -' + str(minus))
                status = True
        elif i > 2:
            if lemtext[i - 1] in plus:
                if lemtext[i - 2] not in minus:
                    text.append('+' + lemtext[i - 1] + ' -' + str(minus))
                    status = True
            elif lemtext[i - 2] in plus:
                if lemtext[i - 1] not in minus:
                    text.append('+' + lemtext[i - 2] + ' -' + str(minus))
                    status = True

    elif keywords.find('+') >= 0 and keywords.find('-') < 0:
        plus = keywords[1:].split(',')
        if i == 1:
            if lemtext[i - 1] in plus:
                text.append('+' + lemtext[i - 1])
                status = True
        elif i >= 2:
            if lemtext[i - 1] in plus:
                text.append('+' + lemtext[i - 1])
                status = True
            elif lemtext[i - 2] in plus:
                text.append('+' + lemtext[i - 2])
                status = True

    elif keywords.find('+') < 0 and keywords.find('-') >= 0:
        minus = keywords[1:].split(',')
        if i == 1:
            if lemtext[i - 1] not in minus:
                text.append('+' + lemtext[i - 1])
                status = True
        elif i >= 2:
            if lemtext[i - 1] not in minus and lemtext[i - 2] not in minus:
                text.append('-' + str(minus))
                status = True

    return status, text

def check_presence (lemtext, keywords):
    text = []
    status = False

    if keywords.find('+') >= 0 and keywords.find('-') >= 0:
        if keywords.find('+') == 0:
            plus = keywords[1:keywords.find('-')].split(',')
            minus = keywords[keywords.find('-') + 1:].split(',')
        else:
            plus = keywords[keywords.find('-') + 1:].split(',')
            minus = keywords[1:keywords.find('-')].split(',')

        res = []
        # res_min = []
        for pl in plus:
            for lem in lemtext:
                simular = jarowinkler.jarowinkler_similarity(pl, lem)
                if simular >= 0.95:
                    for min in minus:
                        for l in lemtext:
                            simular = jarowinkler.jarowinkler_similarity(min, l)
                            if simular >= 0.95:
                                # res_min.append('-' + min)
                                break
                        else:
                            res.append(' +' + pl + ' -' + str(minus))
                            continue
                        break

        if len(set(res)) > 1:
            unique = list(set(res))
            status = True
            text.append(unique)
        elif len(set(res)) == 1:
            status = True
            text.append(res[0])
        else:
            # if len(set(res_min)) > 1:
            #     unique = list(set(res_min))
            #     text.append(unique)
            # elif len(set(res_min)) == 1:
            #     text.append(res_min)
            # else:
            text = []

    elif keywords.find('+') >= 0 and keywords.find('-') < 0:
        plus = keywords[1:].split(',')

        for pl in plus:
            for lem in lemtext:
                simular = jarowinkler.jarowinkler_similarity(pl, lem)
                if simular >= 0.95:
                    text.append(' +' + pl)
                    status = True

    elif keywords.find('+') < 0 and keywords.find('-') >= 0:
        minus = keywords[1:].split(',')
        res = []
        for min in minus:
            for lem in lemtext:
                simular = jarowinkler.jarowinkler_similarity(min, lem)
                if simular >= 0.95:
                    break
            else:
                res.append(' -' + str(minus))
                continue
            break

        if len(set(res)) > 1:
            unique = list(set(res))
            status = True
            text.append(unique)
        elif len(set(res)) == 1:
            status = True
            text.append(res[0])
        else:
            text = []

    return status, text


if __name__ == '__main__':
    main()