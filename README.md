# Анализ остатков товаров на складе

## Краткое описание
Данная задача выполняется в рамках реальной оптимизации процесса для упрощения аналитической работы. Данные, представленные в Exel-файлах, содержат информацию за прошедший месяц.

Необходимо:
- Привести данные в читаемый вид для дальнейшей обработки;
- Выбрать данные по определенному складу;
- Выделить когорты по товарам и посчитать необходимые метрики;
- Загрузить итоговые таблицы с аналитикой на отдельные Exel-листы;
- По каждой когорте подготовить список товаров для дальнейшей работы.

Весь процесс выполнялся ранее вручную через Exel, но, в связи с тем, что данная работа выполняется ежемесячно и занимает слишком много времени, то было принято решение создать код для автоматической обработки.

## Данный код еще находится в процессе доработки. 
Уже с текущими изменениями можно сэкономить 2 часа ручной работы.

Исходные данные можно посмотреть в файле "Остатки.xlsx". Итоговые таблицы и данные в фале "Аналитика.xlsx".

Стэк:

- Python
- JupiterHub
- Exel

## Библиотеки для работы с Exel

```
#pip install pandas
#pip install openpyxl

import pandas as pd
import numpy as np
import os
import openpyxl
from openpyxl import load_workbook
from openpyxl import Workbook
```

## Форматирование данных и получение основных метрик

На первом этапе преобразуем файл Exel для дальнейшего чтения и редактирования. Убираем ненужные строки и слобцы, а также исключаем строки с подзаголовками, чтобы убрать задвоение цифр.
Для работы выбираем только один склад. Параллельно считаем нужные показатели и выводим все в таблицу.

```
itog = pd.DataFrame({'Склад': ['Угданский - масло', 'Угданский - запасные части', 'Итого', 'Продажи за ГОД в ценах с/с', 'Коэфф. оборачиваемости товара'], 
                    'Шт.': [Quanity_oil, Quanity_prod, Quanity_oil+Quanity_prod, Sales_oil+Sales_prod, '0'], 
                    'Сумма.': [Sum_oil, Sum_prod, Sum_oil+Sum_prod, Sales_sum_oil+Sales_sum_prod, (Sales_sum_oil+Sales_sum_prod)/((Sum_begin_oil+Sum_begin_prod+Sum_prod+Sum_oil)/2)]})
itog = itog.set_index('Склад').round(2)

itog
```
![image](https://user-images.githubusercontent.com/100629361/230099732-06dce5d6-a8a3-4e2b-85be-92a2efccc856.png)

Эти данные загружаем в новую таблицу.
```
itog.to_excel ('Анализ.xlsx')
```

## Анализ групп

Второй этап подразумевает разбиение товаров на группы для дальнейшей детальной работы.

Как пример - считаем показатели по первым 3-м группам.
```
Count_group_1 = Ugdan.query('Приход_Кол == 0 and Расход_кол == 0')['Name'].count()
Sum_group_1 = Ugdan.query('Приход_Кол == 0 and Расход_кол == 0')['Кон_ост_сумм'].sum()

Count_group_1_5 = Ugdan.query('Приход_Кол == 0 and Расход_кол == 0 and (Кон_ост_сумм/Кон_ост_кол) > 5000')['Name'].count()
Sum_group_1_5 = Ugdan.query('Приход_Кол == 0 and Расход_кол == 0 and (Кон_ост_сумм/Кон_ост_кол) > 5000')['Кон_ост_сумм'].sum()

Count_group_1_10 = Ugdan.query('Приход_Кол == 0 and Расход_кол == 0 and (Кон_ост_сумм/Кон_ост_кол) > 10000')['Name'].count()
Sum_group_1_10 = Ugdan.query('Приход_Кол == 0 and Расход_кол == 0 and (Кон_ост_сумм/Кон_ост_кол) > 10000')['Кон_ост_сумм'].sum()
```
Информация по группам:

![image](https://user-images.githubusercontent.com/100629361/230100594-7077a15d-0271-467f-ac62-cb97b77609b7.png)

В результате все новые данные также добавляем в файл "Аналитика" из первого этапа.

```
with pd.ExcelWriter('Анализ.xlsx', mode='a', if_sheet_exists= 'replace') as writer:  
    groups.to_excel(writer, sheet_name='Общая информация',index = True)
    Group_1.to_excel(writer, sheet_name='Группа 1',index = False)
    Group_1_5.to_excel(writer, sheet_name='Группа 1 (>5)',index = False)
    Group_1_10.to_excel(writer, sheet_name='Группа 1 (>10)',index = False)
    Group_2.to_excel(writer, sheet_name='Группа 2',index = False)
    Group_3.to_excel(writer, sheet_name='Группа 3',index = False)
    Group_4.to_excel(writer, sheet_name='Группа 4',index = False)
    Group_5.to_excel(writer, sheet_name='Группа 5',index = False)
```
