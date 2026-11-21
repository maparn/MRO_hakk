# MRO_hakk

## Windowsi nupp

Topeltklõps `MRO_Test.bat` failil. Nupp kontrollib kolm sisendfaili, eemaldab vanad
genereeritud raportid, käivitab automaattestid, loob ühe põhiraporti ja avab
`excel_files/output` kausta. Iga käivitus asendab sama faili
`STAT_UOM_vordlusraport.xlsx`; uusi nummerdatud koopiaid ei looda.

Uue testi jaoks kopeeri uued sama nimega failid `excel_files/input` kausta. Windows
küsib, kas olemasolevad failid asendada. Vali asendamine ja käivita `MRO_Test.bat`.

Sisendkaustas peab olema täpselt kolm faili: `STAT_INITIAL - Copy.xlsx`,
`AMOS_uus_süsteem_info.xlsx` ja `AMOS_UOM.xlsx`.

`MRO_Puhasta.bat` eemaldab pärast kinnitust kõik `config.json` failis nimetatud input-
ja output-Excelid. Kasuta seda ainult enne uute sisendfailide laadimist.

Failinimed ja STAT-i veerunumbrid asuvad `config.json` failis. Hinna valiku,
ajaloolise hinnalisa ning mõõtühikute aliased saab muuta `business_rules` plokis.
Praegune kinnitatud loogika kasutab EUR-is olevat AMOS `avg_price` väärtust,
teisendab koguse UOM `rate / exchange_base` järgi ja käsitleb `LT` ühikut `L`
aliasena. Kui `avg_price` on null, kasutatakse sama Task Cardi ajaloolist `AVG`
väärtust 20% lisaga. `Service Summary` lehe viimane rida `SUM_AVG` on lõplik EUR
summa üle kõigi teenuste.

Подготовка справочника AMOS и сверка расходных строк STAT для Magnetic MRO. Основной результат для разработчика Excel — лист `STAT decisions`: одна запись на содержательную строку листа `INITIAL`, включая строки с проблемными ключами. Стоимость и разрешение на перевод единиц определяет расчётный процесс команды.

## Запуск в Ubuntu

Из корня проекта, в окружении Python 3.12:

```bash
cd ~/MRO_hakk
source .venv/bin/activate
python scripts/amos_prepare.py
python scripts/sorter.py --amos "excel_files/input/AMOS_uus_süsteem_info.xlsx"
python scripts/compare_stat_uom_full.py
```

Для новой копии проекта сначала создайте окружение командами `python3 -m venv .venv` и `.venv/bin/python -m pip install -r requirements.txt`. Корневая команда `python amos_prepare.py` также запускает подготовку. Вариант `python scripts/sorter.py` выполняет сверку только по STAT и UOM. Параметры каждого скрипта доступны через `--help`.

## Файлы

Положите `STAT_INITIAL - Copy.xlsx`, `AMOS_UOM.xlsx` и `AMOS_uus_süsteem_info.xlsx` в `excel_files/input/`. По умолчанию скрипты также поддерживают прежнее расположение исходников в корне. Результаты создаются в `excel_files/output/`.

| Книга результата | Назначение |
| --- | --- |
| `amos_prepared.xlsx` | `Prepared` содержит одну строку на запись AMOS, обе цены, их исходные значения и состояния. `UOM_Map` хранит правила, исходные факторы, описания и доступные служебные поля. `Issues` и `Summary` показывают причины проверки и контрольные числа. |
| `STAT_UOM_vordlusraport.xlsx` | `STAT decisions` содержит данные для Excel. `Vastavused` связывает эту строку со всеми правилами-кандидатами по номеру строки. В книге есть отдельные листы для проблем UOM, конфликтов и повторов. |
| `STAT_UOM_vordlusraport_friend.xlsx` | Подробный отчёт по всем парам STAT × UOM. Сравнивает единицу STAT с внутренней единицей правила. Уникальные артикулы, строки STAT и пары имеют отдельные счётчики. |

Рабочую книгу с формулами храните отдельно от генерируемых результатов. При повторном запуске скрипт заменяет свой отчёт целиком, включая ручные значения в нём. Запись идёт через отдельный временный файл и атомарную замену. Произвольная существующая книга, результат другого скрипта, входной файл, выходной симлинк или существующая книга с формулами ячеек защищены от замены. Для первоначального обновления старые отчёты распознаются по штатному имени и составу листов. Входные книги открываются в режиме чтения. Для входных ячеек с формулами нужен актуальный результат пересчёта, сохранённый Excel.

## Контракт для Excel

- Источник количества — `STAT decisions`. Ключ — `stat_row` в пределах конкретной исходной книги. Для предоставленного STAT используются F для артикула, I для единицы, H для количества и C для Task Card. Номера колонок можно изменить параметрами командной строки. `reference` остаётся пустым до явного указания `--reference-col N`.
- Поля `partno`, `stat_unit`, `stat_quantity`, `task_card` и `reference` относятся к исходной расходной строке. Одинаковый артикул в двух расходных строках даёт две записи. Лист `Vastavused` может содержать несколько строк на один расход. Его строки предназначены для поиска вариантов и не подлежат суммированию.
- При `--amos` поиск по артикулу игнорирует регистр и крайние пробелы. Для единственной записи `amos_lookup_status=UNIQUE` выводятся `amos_source_row`, `amos_unit`, `avg_price`, `purch_price` и состояния обеих цен. При нескольких совпадениях используется `MULTIPLE`; произвольная запись не выбирается. `NO_MATCH` и `NOT_REQUESTED` различают отсутствие записи и запуск без большого AMOS.
- `part_rule_rows` — число правил для артикула, `unit_match_rows` — число правил с совпадением единицы с любой стороны, `clean_unit_match_rows` — число таких записей без выявленных дефектов. `unit_decision` объясняет результат проверки, а `conversion_approved` в этой версии всегда равно `FALSE`. Подробности доступны в `match_side` и `rule_issues` на листе кандидатов. Несколько внешних единиц одного артикула допустимы. Точным повтором считается совпадение артикула, двух единиц и обоих числовых факторов.
- В `Prepared` поле `uom_status` отражает текстовое совпадение. Для интерпретации используйте также `uom_conversion_status` и оба счётчика кандидатов. Например, совпадение внутренней единицы при нулевом коэффициенте даёт `REVIEW_INVALID_RULE`. Для цен состояния `MISSING`, `ZERO`, `NEGATIVE`, `INVALID` и `POSITIVE` различаются. Сырые цены и ключи остаются доступны для проверки.
- Выбор цены, валюта, единица цены, псевдонимы единиц, направленные переводы, ограничения цены, историческая надбавка и правила агрегирования требуют утверждения владельцем расчёта. До этого итоговая стоимость из кандидатов не вычисляется.

## Проверка

```bash
.venv/bin/python -B -m unittest discover -s tests -v
```

Тесты создают собственные входные и выходные книги во временной копии проекта. Контрольные показатели поставленной выгрузки — 66 535 записей AMOS, 1 417 правил UOM, 42 373 записи с двумя нулевыми ценами. У `S15/90` одна строка STAT, четыре правила и единица `LT`, которой среди этих правил нет. Эти числа характеризуют именно текущие файлы.
