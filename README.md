# MRO_hakk

## Veebirakendus

Brauseriversioon käivitub failist `streamlit_app.py`. Kasutaja laadib veebilehel
kolm Exceli faili, vajutab arvutamise nuppu, kontrollib teenuste kokkuvõtet ning
laadib alla ühe faili `STAT_UOM_vordlusraport.xlsx`. Üleslaaditud failid töödeldakse
eraldi ajutises kaustas ja kustutatakse kohe pärast raporti loomist.

Kohalik testkäivitus:

```bash
python -m pip install -r requirements.txt
streamlit run streamlit_app.py
```

Avaldamiseks ühenda GitHubi repositoorium Streamlit Community Cloudiga, vali haru
`main` ja rakenduse failiks `streamlit_app.py`. Eraldi domeeni ei ole vaja;
Streamlit annab rakendusele püsiva `streamlit.app` aadressi.

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
ajaloolise hinnalisa, kinnitamise limiidi ning mõõtühikute aliased saab muuta
`business_rules` plokis. `approval_limit_eur` võib olla 500, 1000 või muu positiivne
summa. Limiit märgib rea kontrollimiseks, kuid ei vähenda arvutatud hinda.

Kinnitatud arvutusloogika kasutab EUR-is olevat AMOS `avg_price` väärtust ja
teisendab koguse ainult täpse, suunaga UOM-reegli järgi:
`teisendatud kogus = STAT kogus × rate / exchange_base`. `LT` käsitletakse `L`
aliasena. Kui AMOS `avg_price` on tühi või null või vajalik UOM-reegel puudub,
kasutatakse sama STAT-i materjalirea ajaloolist EUR-ühikuhinda (veerus K) 20%
juurdehindlusega. Task Cardi vana `AVG` on ainult võrdlusnäitaja ning seda ei
kasutata materjali ühikuhinnana. Kui tootekood puudub AMOS-ist, vasteid on mitu,
UOM-reegel on vastuoluline või vajalik ajalooline ühikuhind puudub, jäetakse rida
käsitsi kontrolli; seda ei muudeta nulliks ega arvata kinnitatud kogusummasse.

Arvutusjärjekord on: materjalirea hind → materjalide summa `Reference` kaupa →
`Reference` summade keskmine `Task Card` kaupa → `SUM_AVG`. EUR/USD kurss laaditakse
automaatselt Euroopa Keskpanga päevakursist; võrguvea korral kasutatakse viimast
edukalt salvestatud ECB kurssi. Kurss tähendab USD-d ühe EUR-i kohta, seega
`USD = EUR × EUR/USD`. `approval_limit_eur` on ühikuhinna kontrolllävi, mitte
hinnalagi: süsteem ei vähenda selle tõttu hinda.

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
| `STAT_UOM_vordlusraport.xlsx` | Основной итог. `Executive Summary` показывает контрольные числа и `SUM_AVG`; `Material Calculations` — расчёт каждой строки; `Reference Summary` — суммы по Reference; `Task Card Summary` — средние по Task Card и сравнение со старым AVG; `Exceptions` — fallback и строки для ручной проверки. Технические листы сохраняют диагностические связи и проблемы UOM. |
| `STAT_UOM_vordlusraport_friend.xlsx` | Подробный отчёт по всем парам STAT × UOM. Сравнивает единицу STAT с внутренней единицей правила. Уникальные артикулы, строки STAT и пары имеют отдельные счётчики. |

Рабочую книгу с формулами храните отдельно от генерируемых результатов. При повторном запуске скрипт заменяет свой отчёт целиком, включая ручные значения в нём. Запись идёт через отдельный временный файл и атомарную замену. Произвольная существующая книга, результат другого скрипта, входной файл, выходной симлинк или существующая книга с формулами ячеек защищены от замены. Для первоначального обновления старые отчёты распознаются по штатному имени и составу листов. Входные книги открываются в режиме чтения. Для входных ячеек с формулами нужен актуальный результат пересчёта, сохранённый Excel.

## Контракт для Excel

- Источник расчёта — каждая материальная строка листа `INITIAL`. Ключ — номер исходной строки. По умолчанию используются A для Reference, C для Task Card, D для NRC, E для Factor, F для артикула, G для описания, H для количества, I для единицы и K для исторической EUR-цены за единицу. Номера колонок можно изменить в `config.json` или параметрами командной строки.
- Поля `partno`, `stat_unit`, `stat_quantity`, `task_card` и `reference` относятся к исходной расходной строке. Одинаковый артикул в двух расходных строках даёт две записи. Лист `Vastavused` может содержать несколько строк на один расход. Его строки предназначены для поиска вариантов и не подлежат суммированию.
- При `--amos` поиск по артикулу игнорирует регистр и крайние пробелы. Для единственной записи `amos_lookup_status=UNIQUE` выводятся `amos_source_row`, `amos_unit`, `avg_price`, `purch_price` и состояния обеих цен. При нескольких совпадениях используется `MULTIPLE`; произвольная запись не выбирается. `NO_MATCH` и `NOT_REQUESTED` различают отсутствие записи и запуск без большого AMOS.
- `part_rule_rows` — число правил для артикула, `unit_match_rows` — число правил с совпадением единицы с любой стороны, `clean_unit_match_rows` — число таких записей без выявленных дефектов. `unit_decision` объясняет результат проверки, а `conversion_approved` в этой версии всегда равно `FALSE`. Подробности доступны в `match_side` и `rule_issues` на листе кандидатов. Несколько внешних единиц одного артикула допустимы. Точным повтором считается совпадение артикула, двух единиц и обоих числовых факторов.
- В `Prepared` поле `uom_status` отражает текстовое совпадение. Для интерпретации используйте также `uom_conversion_status` и оба счётчика кандидатов. Например, совпадение внутренней единицы при нулевом коэффициенте даёт `REVIEW_INVALID_RULE`. Для цен состояния `MISSING`, `ZERO`, `NEGATIVE`, `INVALID` и `POSITIVE` различаются. Сырые цены и ключи остаются доступны для проверки.
- AMOS `avg_price` применяется только при единственном точном совпадении артикула и совместимых единицах. При нулевой/пустой цене или отсутствии направленного UOM-правила применяется историческая цена именно этой строки STAT с настраиваемой надбавкой. Отсутствующий в AMOS артикул автоматически не оценивается.
- `Reference Summary` получает сумму только когда рассчитаны все строки Reference. `Task Card Summary` усредняет только полные Reference и явно помечает Task Card как `INCOMPLETE`, если хотя бы один Reference не завершён. Поэтому неполный результат нельзя принять за окончательную цену.

## Проверка

```bash
.venv/bin/python -B -m unittest discover -s tests -v
```

Тесты создают собственные входные и выходные книги во временной копии проекта. Контрольные показатели поставленной выгрузки — 66 535 записей AMOS, 1 417 правил UOM, 42 373 записи с двумя нулевыми ценами. У `S15/90` одна строка STAT, четыре правила и единица `LT`, которой среди этих правил нет. Эти числа характеризуют именно текущие файлы.
