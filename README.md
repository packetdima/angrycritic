# Инструкция по использованию программного продукта AngryCritic


AngryCritic - инструмент для анализа данных и определения их тональности, на основании заданных параметров и ключевых слов.
## Компоненты программы
- **angrycritic.py** – исполняемый файл;
- **config.py** – конфигурационный файл, для настройки работы кода под ваш контекст данных.

## Список пакетов, необходимых для работы программы
Перед запуском убедитесь, что все указанные ниже пакеты установлены
- **openpyxl**;
- **os**;
- **pymorph3**;
- **string**;
- **sys**;
- **time**;
- **tqdm**.

## Запуск программы
Перед запуском программы, Excel-файл с данными необходимо переименовать в «**data.xlsx**» и поместить в папку с программой. Данные для анализа должны находиться в столбце «А». По окончанию обработки в данной папке будет создан файл «**data_done.xlsx**», который будет содержать результаты обработки. Исходя из заданных параметров в config.py, в обработанном файле будут добавлены следующие данные:
- Леммы слов в данных (только в режиме debug);
- Категория, к которой относятся данные;
- Ключевые слова, на основании которых была определена категория;
- Тональность;
- Список ключевых слов, на основании которых определена тональность;
- Значение index, на основании которого определяется тональность данных (Больше 0 – позитив, меньше 0 – негатив).
Если в анализируемом тексте встречается несколько ключевых слов, то на основании значения их тональности вычисляется значение index (positive +1, negative -1, undef = 0).

## Скорость обработки
На ПК с процессором Intel(R) Core(TM) i5-11400, 8Gb RAM - 1000 строк/мин

## Как настроить программу под свои данные
Откройте файл «config.py», он разделен на следующие сегменты:
1. Ключевые слова и фразы, позволяющие сразу исключить спам, массовые рассылки и прочие шаблонные сообщения.
	+ **skip** – массив ключевых слов и словосочетаний, при обнаружении которых обработка данных производиться не будет. 

2. Ключевые слова для замены. Позволяет облегчить определения тональности данных
	+ **replace** – на стадии предобработке текста имеется возможность произвести замену ключевых слов или фраз и внести эти данные в словарь keywords, для однозначного и верного определения категории данных. Кроме того, с помощью замены можно исключить, данные не имеющие значения, но часто встречающиеся и влияющие на результат (Например, фраза “Отправлено с iPhone”).
```
{
  'general': {
  			'отправить с ipad':'', 'отправить с iphone':''
		  },
  'fishing': {
  			'не согласен': 'не_согласен', 'мочь подтвердить': 					'могу_подтвердить'
    		  }
}
```
- ‘general’, ‘fishing’ – категории данных, к которым будут применены правила замены;
- ‘general’ – общие правила замены для всех категорий;
- ‘отправить с ipad’ – слово или фраза, которую необходимо заменить;
- ‘не_согласен’ – слово или фраза, на что требуется заменить. Заменяйте пробелы на _ для получения единых словосочетаний, которые в последствии можно будет добавить в ключевые слова для определения тональности.

3. Ключевые слова для определения категории данных содержит два словаря:
	+ **char_keywords** – данные словарь содержит ключевое слов и категорию данных, к которой по этому слову можно отнести данные. Добавлять следует в формате ‘ключ’:’категория’. Если однозначно определить категорию по ключевому слову не представляется возможным, то задается категория ‘presence’ (‘письмо’:’presence’). То есть, будет проведен дополнительный анализ данных, по ключевым словам, из словаря char_presence и на основании их наличия будет определяться категория данных; 
	+ **char_presence** – словарь содержит дополнительные ключи для обеспечения более точного определения категории данных.
Как это работает:
	> Например, ваши данные содержат слово «письмо», по которому определить к какой категории их отнести достаточно трудно. Поэтому в char_keywords мы его записываем со значением ‘presence’ ('письмо':'presence'). Теперь, если в ваших данных содержится слово ‘письмо’ обработчик будет искать наличие других ключевых слов из словаря char_presence и на основании их наличия будет определять категорию данных. Если ему это не удастся, данные будут помечены как нераспознанные - ‘undef’.

4. Основные ключевые слова для анализа тональности данных
	+ **keywords** – словарь, в который необходимо добавлять ключевые слова для определения тональности конкретных категорий данных.
```
{
   'fishing': {
      	'проверить':'positive', 'ошибиться':'negative', 'пропустить':'near_negative', 'открывать':'near_positive', 'пересылать':'together'
      	}
}
```
- ‘fishing’ – категория, к которой относятся указанные далее ключи;
- ‘проверить’ – ключевые слова, на основании которых определяется тональность;
- ‘positive’ – тональность сообщения при наличии данного ключевого слова. 

Тональность может принимать следующие значения:
| Наименование     | Описание                                                      |
|------------------|---------------------------------------------------------------|
| positive	  | при наличии данного слова в тексте тональность положительная |
| negative        | при наличии данного слова в тексте тональность отрицательная |
| near_positive  | при наличии данного ключа в тексте и сопутствующих слов из словаря near перед ключевым в пределах 2 позиций - тональность положительная |
| near_negative   | при наличии данного ключа в тексте и сопутствующих слов из словаря near перед ключевым в пределах 2 позиций - тональность отрицательная |
| together	   | при наличии данного ключа в тексте, сопутствующих слов из словаря near перед ключевым в пределах 2 позиций, а также при наличии вспомогательных слов из словаря presence - тональность будет определяться исходя из значения ключей presence |

5. Вспомогательные ключевые слова, которые располагаются в пределах двух позиций перед основным ключевым словом
	+ **near** – в словаре содержится данные о категории и список сопутствующих слов, при наличии или отсутствии которых принимается решении о тональности данных
	> Например, возьмем слово «пропустить», для которого в keywords задано значение 'пропустить':'near_negative'. Если в данных будет фраза «не пропустить», то указанное выражение будет считаться негативным.
	+ **presence** – в словаре содержатся данные о категории и вспомогательные ключи, нахождение которых с основными ключами и вспомогательными ключами near определяют тональность данных


