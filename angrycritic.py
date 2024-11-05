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


def main():
    global skip
    global types_keywords
    global replace
    global rustopwords
    rustopwords = stopwords.words('russian')
    start_time = time.time()
    active_types = []

    if os.path.exists('data.xlsx'):
        print("Файл обнаружен. Начинаем обработку...")

        df = pd.read_excel('data.xlsx', index_col=None, header=0, usecols='A')
        new_df = df.assign(Lemma='', Type='', Type_Keywords='', Mood='', Mood_Keywords='', Index='')
        arr = new_df.to_numpy()

        with open('./data/skip.yaml', encoding='utf8') as f:
            skip = yaml.load(f, Loader=SafeLoader)

        with open('./data/types.yaml', encoding='utf8') as f:
            types_keywords = yaml.load(f, Loader=SafeLoader)

        with open('./data/replace.yaml', encoding='utf8') as f:
            replace = yaml.load(f, Loader=SafeLoader)

        for item in types_keywords:
            if item != 'undef' and item != 'presence_kw':
                active_types.append(item)
            else:
                continue

        prep = prepare(arr)
        prepdf = pd.DataFrame(prep[0])
        countdf = pd.DataFrame(prep[1]) #.from_dict(counter, orient='index')
        prepdf.columns= ['Text', 'Lemma', 'Type', 'Type_Keywords', 'Mood', 'Mood_Keywords', 'Index']
        countdf.columns= ['Слово','Частота']
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
        countdf.to_excel(writer, sheet_name='Частотность по документу', index=False)

        for k, v in data.items():
            r = pd.DataFrame(v)
            r.columns = ['Слово','Частота']
            r.to_excel(writer, sheet_name=k, index=False)

        writer.close()

        end_time = time.time()
        execution_time = end_time - start_time

        print(f"Время выполнения программы: {execution_time} секунд")
        print("Обработка выполнена! Проверьте файл data_done.xlsx")
    else:
        print("Файл с данными (data.xlsx) не найден. Пожалуйста проверьте их наличие в папке с программой!")


def prepare(arr):
    print('Начинаем предварительную обработку.')
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
                if k != 'presence_kw':
                    for i in v:
                        if i in tokens:
                            if k == 'undef':
                                for key, value in types_keywords['presence_kw'].items():
                                    for i in value:
                                        if i in tokens:
                                            type_arr.append(key)
                                            keyword.append(i)
                                    else:
                                        continue
                            else:
                                type_arr.append(k)
                                keyword.append(i)

            if len(type_arr) > 0:
                char = {i: type_arr.count(i) for i in type_arr};
                char = list(char.keys())[0]
            else:
                char = 'undef'

            if char != 'undef' and char != 'skip':
                try:
                    for key, value in replace['general'].items():
                        if key in new_text:
                            if value is None:
                                new_text = new_text.replace(key,'')
                            else:
                                new_text = new_text.replace(key, value)
                        else:
                            continue
                except KeyError:
                    pass

                try:
                    for key, value in replace[char].items():
                        if key in new_text:
                            if value is None:
                                new_text = new_text.replace(key,'')
                            else:
                                new_text = new_text.replace(key, value)
                        else:
                            continue
                except KeyError:
                    pass

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
    near_pos = []
    near_neg = []
    near_kw = []
    presence_pos = []
    presence_neg = []
    presence_kw = []
    textcount = []

    path = './data/mood_keywords_' + char + '.yaml'

    with open(path, encoding='utf8') as f:
        keywords = yaml.safe_load(f)

    for key, value in keywords.items():         # собираем ключевые слова из yaml в переменные
        if key == 'negative':
            neg = value
        elif key == 'positive':
            pos = value
        elif key == 'near':
            for item in value:
                for k, v in item.items():
                    if v != 'None':
                        if k == 'negative':
                            near_neg = v
                        elif k == 'positive':
                            near_pos = v
                        elif k == 'near_kw':
                            near_kw = v
        elif key == 'presence':
            for item in value:
                for k, v in item.items():
                    if v != 'None':
                        if k == 'negative':
                            presence_neg = v
                        elif k == 'positive':
                            presence_pos = v
        elif key == 'presence_kw':
            presence_kw = value

    for item in tqdm(arr):
        index = 0
        lemtext = item[1]
        sum_words = []

        for i in lemtext:
            if i and i not in rustopwords:
                if not i.isdigit():
                    textcount.append(i)
            if i in neg:
                sum_words.append(i)
                index -= 1
            elif i in pos:
                sum_words.append(i)
                index += 1
            elif i in near_neg:
                if lemtext[lemtext.index(i) - 1] in near_kw:
                    sum_words.append(lemtext[lemtext.index(i) - 1] + ' + ' + i)
                    index -= 1
                elif lemtext[lemtext.index(i) - 2] in near_kw:
                    sum_words.append(lemtext[lemtext.index(i) - 2] + ' + ' + i)
                    index -= 1
            elif i in near_pos:
                if lemtext[lemtext.index(i) - 1] in near_kw:
                    sum_words.append(lemtext[lemtext.index(i) - 1] + ' + ' + i)
                    index += 1
                elif lemtext[lemtext.index(i) - 2] in near_kw:
                    sum_words.append(lemtext[lemtext.index(i) - 2] + ' + ' + i)
                    index += 1
            elif i in presence_kw:
                if lemtext[lemtext.index(i) - 1] in near_kw:
                    for word in presence_neg:
                        if word in lemtext:
                            sum_words.append(lemtext[lemtext.index(i) - 1] + ' + ' + i + ' + ' + word)
                            index -= 1
                    for word in presence_pos:
                        if word in lemtext:
                            sum_words.append(lemtext[lemtext.index(i) - 1] + ' + ' + i + ' + ' + word)
                            index += 1
                elif lemtext[lemtext.index(i) - 2] in near_kw:
                    for word in presence_neg:
                        if word in lemtext:
                            sum_words.append(lemtext[lemtext.index(i) - 2] + ' + ' + i + ' + ' + word)
                            index -= 1
                    for word in presence_pos:
                        if word in lemtext:
                            sum_words.append(lemtext[lemtext.index(i) - 2] + ' + ' + i + ' + ' + word)
                            index += 1

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

if __name__ == '__main__':
    main()
