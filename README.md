# MRO_hakk

## Veebirakendus

Brauseriversioon käivitatakse failist `streamlit_app.py`. Kasutaja laadib veebilehel üles kolm Exceli faili, vajutab arvutamise nuppu, kontrollib teenuste kokkuvõtet ning laadib alla faili `STAT_UOM_vordlusraport.xlsx`. Üleslaaditud faile töödeldakse eraldi ajutises kaustas ja need kustutatakse kohe pärast raporti koostamist.

Kohalik testkäivitus:

```bash
python -m pip install -r requirements.txt
streamlit run streamlit_app.py
```

Rakenduse avaldamiseks ühenda GitHubi repositoorium Streamlit Community Cloudiga, vali haru `main` ja rakenduse käivitusfailiks `streamlit_app.py`. Eraldi domeeni pole vaja, sest Streamlit annab rakendusele püsiva `streamlit.app` aadressi.

Veebilehe mõõtühikute konverter on eraldiseisev käsitsi kasutatav abivahend. See ei asenda detailinumbripõhiseid AMOS UOM-reegleid ega osale hinnaarvutuses.

## Windowsi nupp

Topeltklõpsa failil `MRO_Test.bat`. See kontrollib kolme sisendfaili olemasolu, eemaldab vanad genereeritud raportid, käivitab automaattestid, koostab ühe põhiraporti ja avab kausta `excel_files/output`. Iga käivitus asendab sama faili `STAT_UOM_vordlusraport.xlsx`; uusi nummerdatud koopiaid ei looda.

Uue testi jaoks kopeeri uued sama nimega failid kausta `excel_files/input`. Windows küsib, kas olemasolevad failid asendada. Kinnita asendamine ja käivita `MRO_Test.bat`.

Sisendkaustas peab olema täpselt kolm faili: `STAT_INITIAL - Copy.xlsx`, `AMOS_uus_süsteem_info.xlsx` ja `AMOS_UOM.xlsx`.

`MRO_Puhasta.bat` eemaldab pärast kinnitust kõik `config.json` failis määratud sisend- ja väljund-Exceli failid. Kasuta seda ainult enne uute sisendfailide laadimist.

Failinimed ja STAT-i veerunumbrid on määratud failis `config.json`. Hinnavaliku, ajaloolise hinnalisa, kinnitamislimiidi ning mõõtühikute aliaseid saab muuta jaotises `business_rules`. `approval_limit_eur` võib olla 500, 1000 või mõni muu positiivne summa. Limiit märgib hinna kontrollimiseks, kuid ei vähenda arvutatud hinda.

Kinnitatud arvutusloogika kasutab eurodes väljendatud AMOS-i `avg_price` väärtust ja teisendab koguse ainult täpse suunatud UOM-reegli alusel:

`teisendatud kogus = STAT-i kogus × rate / exchange_base`

Ühikut `LT` käsitletakse ühiku `L` aliasena. Kui AMOS-i `avg_price` on tühi või null või vajalik UOM-reegel puudub, kasutatakse sama STAT-i materjalirea ajaloolist ühikuhinda eurodes (veerg K), lisades sellele 20%. Task Cardi vana `AVG`-väärtust ei kasutata materjali ühikuhinnana: seda võrreldakse alles uue Task Cardi AVG-ga ning kliendihinnaks valitakse neist suurem.

Kui detailinumber puudub AMOS-ist, vasteid on mitu, UOM-reegel on vastuoluline või vajalik ajalooline ühikuhind puudub, jäetakse rida käsitsi kontrollimiseks. Seda ei muudeta nulliks ega arvestata kinnitatud kogusumma hulka.

Raportis eristatakse välju `current_price_status` ja `fallback_status`. Võimalikud olekud on näiteks `CURRENT_PRICE_VALID`, `CURRENT_PRICE_ZERO`, `CURRENT_PRICE_MISSING`, `CURRENT_PRICE_INVALID`, `MATERIAL_NOT_FOUND`, `HISTORICAL_FALLBACK_APPLIED` ja `MISSING_HISTORICAL_PRICE`. Eduka fallback'i korral on hinnaallikas `HISTORICAL_PRICE_PLUS_20_PERCENT` ning arvutuse olek `FALLBACK_APPLIED`. Kui ei praegust ega ajaloolist hinda saa kasutada, kuvatakse valitud hinna väljal `ANDMED PUUDUVAD`. STAT-i lõpuread, näiteks `GRAND TOTAL`, ei ole materjaliread ning neid ei kaasata Reference'i arvutusse.

`approved_physical_price_conversions` on eraldi loend üksnes kirjalikult kinnitatud füüsiliste hinnaühikute teisenduste jaoks. Igal aktiivsel kirjel peavad olema `amos_unit`, `stat_unit`, positiivne `price_multiplier`, `enabled: true` ja `approval_reference`. Esmalt kasutatakse alati AMOS UOM-i detailinumbripõhist reeglit.

Füüsilist erandit kasutatakse ainult siis, kui ametlik reegel puudub. Muul juhul kehtib ajaloolise materjalihinna +20% varureegel. Näidisteisendus `ME → MM` on seadistuses tahtlikult välja lülitatud, kuni Magnetic MRO kinnitab, et AMOS-i `ME` tähendab meetrit, ja kiidab hinnateguri `0.001` heaks.

Arvutamise järjekord on järgmine: materjalirea hind → materjalide summa `Reference` kaupa → `Reference`-summade keskmine `Task Card` kaupa → `MAX(vana AVG, uus AVG)` → lõplike Task Cardi hindade summa.

EUR/USD vahetuskurss laaditakse automaatselt Euroopa Keskpanga päevakursina. Võrguvea korral kasutatakse viimast edukalt salvestatud ECB kurssi. Kurss näitab, mitu USA dollarit vastab ühele eurole:

`USD = EUR × EUR/USD`

`approval_limit_eur` on ühikuhinna kontrollimise lävend, mitte hinnalagi: süsteem ei vähenda selle tõttu hinda.

Projektis valmistatakse ette Magnetic MRO AMOS-i võrdlusandmed ning võrreldakse STAT-i materjaliridu. Exceli arendaja jaoks on peamine tulemus tööleht `STAT decisions`: üks kirje iga lehe `INITIAL` sisuka rea kohta, sealhulgas probleemsete võtmeväljadega read. Maksumuse arvutamise ja mõõtühikute teisendamise lubatavuse määrab meeskonna arvutusprotsess.

## Käivitamine Ubuntu keskkonnas

Käivita käsud projekti juurkaustas Python 3.12 virtuaalkeskkonnas:

```bash
cd ~/MRO_hakk
source .venv/bin/activate
python scripts/amos_prepare.py
python scripts/sorter.py --amos "excel_files/input/AMOS_uus_süsteem_info.xlsx"
python scripts/compare_stat_uom_full.py
```

Projekti uue koopia jaoks loo esmalt virtuaalkeskkond käsuga `python3 -m venv .venv` ning paigalda sõltuvused käsuga `.venv/bin/python -m pip install -r requirements.txt`.

Juurkaustast käivitatav käsk `python amos_prepare.py` valmistab samuti andmed ette. Käsk `python scripts/sorter.py` võrdleb ainult STAT-i ja UOM-i andmeid. Iga skripti parameetreid saab vaadata valikuga `--help`.

## Failid

Paiguta failid `STAT_INITIAL - Copy.xlsx`, `AMOS_UOM.xlsx` ja `AMOS_uus_süsteem_info.xlsx` kausta `excel_files/input/`. Vaikimisi toetavad skriptid ka sisendfailide varasemat asukohta projekti juurkaustas. Tulemused salvestatakse kausta `excel_files/output/`.

| Tulemuse Exceli fail | Otstarve |
| --- | --- |
| `amos_prepared.xlsx` | Lehel `Prepared` on üks rida iga AMOS-i kirje kohta, mõlemad hinnad, nende algväärtused ja olekud. `UOM_Map` sisaldab teisendusreegleid, algseid koefitsiente, kirjeldusi ja olemasolevaid täiendavaid välju. `Issues` ja `Summary` näitavad kontrollimist vajavaid probleeme ning kontrollarve. |
| `STAT_UOM_vordlusraport.xlsx` | Peamine lõpparuanne. `Executive Summary` näitab kontrollarve ja lõppsummat; `Material Calculations` iga rea arvutust; `Reference Summary` summasid Reference'i kaupa; `Task Card Summary` uut AVG-d, vana AVG-d, valitud suuremat väärtust ning allikat `OLD` või `NEW`; `Exceptions` varureegli rakendamise juhte ja käsitsi kontrollimist vajavaid ridu. Tehnilistel lehtedel säilitatakse diagnostilised seosed ja UOM-i probleemid. |
| `STAT_UOM_vordlusraport_friend.xlsx` | Üksikasjalik aruanne kõigi STAT × UOM paaride kohta. Võrdleb STAT-i mõõtühikut reegli sisemise mõõtühikuga. Unikaalseid detailinumbreid, STAT-i ridu ja paare loendatakse eraldi. |

Valemitega töövihikut tuleb hoida genereeritud tulemustest eraldi. Skripti uuesti käivitamisel asendatakse kogu sellele kuuluv aruanne, sealhulgas aruandesse käsitsi sisestatud väärtused. Salvestamine toimub eraldi ajutise faili kaudu, millele järgneb atomaarne asendamine.

Süsteem kaitseb juhusliku ülekirjutamise eest suvalist olemasolevat töövihikut, teise skripti tulemust, sisendfaili, väljundfailile osutavat sümbollinki ja olemasolevat Exceli valemitega töövihikut. Vanade aruannete esmakordsel uuendamisel tuvastatakse need tavapärase failinime ja lehtede koosseisu järgi. Sisendtöövihikud avatakse lugemisrežiimis. Kui sisend lahtrites on valemid, peab Excel olema salvestanud ka nende ajakohased arvutatud väärtused.

## Exceli andmevahetuse kokkulepe

- Arvutuse lähteandmeteks on lehe `INITIAL` iga materjalirida. Kirje võti on algse rea number. Vaikimisi kasutatakse järgmisi veerge: A — Reference; C — Task Card; D — NRC; E — Factor; F — detailinumber; G — kirjeldus; H — kogus; I — mõõtühik; K — ajalooline ühikuhind eurodes. Veerunumbreid saab muuta failis `config.json` või käsurea parameetritega.
- Väljad `partno`, `stat_unit`, `stat_quantity`, `task_card` ja `reference` kuuluvad algsele materjalireale. Kui sama detailinumber esineb kahel materjalireal, luuakse kaks eraldi kirjet. Lehel `Vastavused` võib ühe materjalirea kohta olla mitu rida. Need on mõeldud võimalike vastete otsimiseks ja neid ei tohi summeerida.
- Parameetri `--amos` kasutamisel eiratakse detailinumbri otsingus tähesuurust ja algus- või lõputühikuid. Üheainsa vaste korral (`amos_lookup_status=UNIQUE`) väljastatakse `amos_source_row`, `amos_unit`, `avg_price`, `purch_price` ja mõlema hinna olekud. Mitme vaste korral kasutatakse olekut `MULTIPLE`; suvalist kirjet ei valita. `NO_MATCH` ja `NOT_REQUESTED` eristavad puuduva vaste olukorda käivitusest, kus suurt AMOS-i faili ei kasutatud.
- `part_rule_rows` näitab detailinumbriga seotud reeglite arvu, `unit_match_rows` mõne poole mõõtühikuga kattuvate reeglite arvu ning `clean_unit_match_rows` selliste kirjete arvu, millel ei tuvastatud puudusi. `unit_decision` selgitab kontrolli tulemust. Selles versioonis on `conversion_approved` alati `FALSE`. Täpsem info asub kandidaatide lehe väljadel `match_side` ja `rule_issues`. Sama detailinumbri mitu erinevat välist mõõtühikut on lubatud. Täpseks korduseks loetakse kirje, milles kattuvad detailinumber, mõlemad mõõtühikud ja mõlemad arvulised koefitsiendid.
- Lehel `Prepared` kajastab väli `uom_status` tekstilist vastavust. Tulemuse tõlgendamisel tuleb vaadata ka välja `uom_conversion_status` ja mõlemat kandidaatide loendurit. Näiteks sisemise mõõtühiku kattumine nullkoefitsiendiga annab oleku `REVIEW_INVALID_RULE`. Hindade puhul eristatakse olekuid `MISSING`, `ZERO`, `NEGATIVE`, `INVALID` ja `POSITIVE`. Algseid hindu ja võtmevälju saab samuti kontrollida.
- AMOS-i `avg_price` kasutatakse ainult siis, kui detailinumbril on üks täpne vaste ja mõõtühikud on ühilduvad. Null- või tühja hinna või vajaliku suunatud UOM-reegli puudumise korral kasutatakse sama STAT-i rea ajaloolist ühikuhinda koos seadistatava juurdehindlusega. AMOS-ist puuduvale detailinumbrile ei määrata automaatselt hinda.
- `Reference Summary` saab summa ainult siis, kui kõik selle Reference'i read on arvutatud. `Task Card Summary` arvutab keskmise ainult täielikult arvutatud Reference'ite põhjal ning märgib Task Cardi olekuga `INCOMPLETE`, kui vähemalt üks Reference on lõpetamata. Seetõttu ei tohi puudulikku tulemust käsitleda lõpliku hinnana.
- Task Cardi lõplik hind on `MAX(vana Task Cardi AVG, uus Task Cardi AVG)`. Võrdsete väärtuste korral märgitakse allikaks järjekindlalt `NEW`. Kui uues arvutuses on lahendamata ridu, kuvatakse ärireegli kohaselt siiski maksimum, kuid Task Cardi ja kogutulemuse olek jääb `INCOMPLETE`: sellist hinda ei tohi esitada täielikult kontrollituna.

## Testimine

```bash
.venv/bin/python -B -m unittest discover -s tests -v
```

Testid loovad eraldi ajutises projektikoopias oma sisend- ja väljundtöövihikud. Kaasasoleva andmeväljavõtte kontrollnäitajad on 66 535 AMOS-i kirjet, 1 417 UOM-reeglit ning 42 373 kirjet, mille mõlemad hinnad on nullid. Detailinumbril `S15/90` on üks STAT-i rida, neli reeglit ja mõõtühik `LT`, mida nende reeglite hulgas ei ole. Need arvud kirjeldavad just praegu kaasasolevaid faile.
