"""
STAT ja Measure Unit Excel-failide võrdlusautomaatika
====================================================

MIDA SEE PROGRAMM TEEB?
- Loeb vana STAT Exceli faili töölehelt INITIAL.
- Loeb Measure Unit Exceli faili töölehelt Sheet1.
- Võrdleb STAT-faili partno väärtusi Measure Unit faili partno väärtustega.
- Võrdleb vastavate kirjete mõõtühikuid.
- Koostab eraldi Exceli raporti.
- EI muuda algfaile.
- EI lisa ega kustuta algandmeid.

PAIGALDAMINE:
    py -m pip install openpyxl

KÄIVITAMINE:
    py compare_stat_uom.py

FAILID:
    Pane see .py-fail ja mõlemad Exceli failid samasse kausta.

MÄRKUS:
STAT-failis ei ole päiserida. Praeguse näidisrea põhjal on:
    STAT veerg F = partno (nt S15/90)
    STAT veerg I = measure_unit (nt LT)

Kui sinu tegelikus STAT-failis on partno mõnes teises veerus,
muuda allpool STAT_PARTNO_COL väärtust.
Exceli veerunumbrid: A=1, B=2, C=3 jne.
"""

from pathlib import Path
from collections import defaultdict
from openpyxl import load_workbook, Workbook
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter


# ============================================================
# 1. SEADISTUSED
# ============================================================

KAUST = Path(__file__).resolve().parent

STAT_FAIL = KAUST / "excel_files" / "input" / "STAT_INITIAL - Copy.xlsx"
UOM_FAIL = KAUST / "AMOS_UOM.xlsx"
RAPORT_FAIL = KAUST / "excel_files" / "output" / "STAT_UOM_vordlusraport_friend.xlsx"

STAT_LEHT = "INITIAL"
UOM_LEHT = "Sheet1"

# STAT-failis pole päiseid, seega määrame veerud käsitsi.
STAT_PARTNO_COL = 6       # F-veerg
STAT_MEASURE_UNIT_COL = 9 # I-veerg

# Measure Unit faili päised.
UOM_PARTNO_HEADER = "partno"
UOM_MEASURE_UNIT_HEADER = "measure_unit"

# Võrdlemisel eemaldatakse ääretühikud ja muudetakse tekst suurtähtedeks.
# Algseid väärtusi raportis ei muudeta.
IGNORE_CASE = True


# ============================================================
# 2. ABIFUNKTSIOONID
# ============================================================

def normaliseeri(väärtus):
    """Teeb väärtused võrdlemiseks ühtlaseks."""
    if väärtus is None:
        return ""

    tekst = str(väärtus).strip()

    if IGNORE_CASE:
        tekst = tekst.upper()

    return tekst


def leia_veeru_number(tööleht, päise_nimi):
    """Leiаb Exceli veeru numbri esimese rea päise järgi."""
    otsitav = normaliseeri(päise_nimi)

    for lahter in tööleht[1]:
        if normaliseeri(lahter.value) == otsitav:
            return lahter.column

    raise ValueError(
        f"Veeru päist '{päise_nimi}' ei leitud töölehelt "
        f"'{tööleht.title}'. Kontrolli Exceli päiserida."
    )


def stiili_tööleht(tööleht):
    """Vormindab raporti töölehe loetavaks."""
    tööleht.freeze_panes = "A2"

    if tööleht.max_row >= 1 and tööleht.max_column >= 1:
        tööleht.auto_filter.ref = tööleht.dimensions

    for lahter in tööleht[1]:
        lahter.font = Font(bold=True, color="FFFFFF")
        lahter.fill = PatternFill(
            fill_type="solid",
            fgColor="1F4E78"
        )
        lahter.alignment = Alignment(
            wrap_text=True,
            vertical="center"
        )

    tööleht.row_dimensions[1].height = 30

    for veerg in tööleht.columns:
        veeru_täht = get_column_letter(veerg[0].column)
        max_pikkus = 0

        for lahter in veerg[:3000]:
            if lahter.value is not None:
                max_pikkus = max(
                    max_pikkus,
                    len(str(lahter.value))
                )

        tööleht.column_dimensions[veeru_täht].width = min(
            max(max_pikkus + 2, 12),
            45
        )


def lisa_raporti_leht(töövihik, nimi, päised, read):
    """Lisab raportisse ühe töölehe."""
    tööleht = töövihik.create_sheet(title=nimi)
    tööleht.append(päised)

    for rida in read:
        tööleht.append(rida)

    stiili_tööleht(tööleht)
    return tööleht


# ============================================================
# 3. STAT-FAILI LUGEMINE
# ============================================================

def loe_stat_fail(failitee):
    """Loeb STAT-faili INITIAL töölehe andmed."""
    töövihik = load_workbook(
        failitee,
        data_only=True,
        read_only=True
    )

    if STAT_LEHT not in töövihik.sheetnames:
        olemas = ", ".join(töövihik.sheetnames)
        töövihik.close()
        raise ValueError(
            f"STAT töölehte '{STAT_LEHT}' ei leitud. "
            f"Failis olevad töölehed: {olemas}"
        )

    tööleht = töövihik[STAT_LEHT]
    tulemused = []

    for rea_nr, rida in enumerate(
        tööleht.iter_rows(values_only=True),
        start=1
    ):
        väärtused = list(rida)

        # Jäta täiesti tühjad read vahele.
        if not any(v is not None for v in väärtused):
            continue

        partno = (
            väärtused[STAT_PARTNO_COL - 1]
            if len(väärtused) >= STAT_PARTNO_COL
            else None
        )

        mõõtühik = (
            väärtused[STAT_MEASURE_UNIT_COL - 1]
            if len(väärtused) >= STAT_MEASURE_UNIT_COL
            else None
        )

        tulemused.append({
            "rea_nr": rea_nr,
            "partno": partno,
            "partno_võti": normaliseeri(partno),
            "mõõtühik": mõõtühik,
            "mõõtühik_võti": normaliseeri(mõõtühik),
            "andmed": väärtused
        })

    töövihik.close()
    return tulemused


# ============================================================
# 4. MEASURE UNIT FAILI LUGEMINE
# ============================================================

def loe_uom_fail(failitee):
    """Loeb Measure Unit faili ja leiab vajalikud veerud päise järgi."""
    töövihik = load_workbook(
        failitee,
        data_only=True,
        read_only=True
    )

    if UOM_LEHT not in töövihik.sheetnames:
        olemas = ", ".join(töövihik.sheetnames)
        töövihik.close()
        raise ValueError(
            f"Measure Unit töölehte '{UOM_LEHT}' ei leitud. "
            f"Failis olevad töölehed: {olemas}"
        )

    tööleht = töövihik[UOM_LEHT]
    # AMOS Excel võib deklareerida valed töölehe mõõtmed.
    tööleht.reset_dimensions()

    partno_veerg = leia_veeru_number(
        tööleht,
        UOM_PARTNO_HEADER
    )

    mõõtühiku_veerg = leia_veeru_number(
        tööleht,
        UOM_MEASURE_UNIT_HEADER
    )

    tulemused = []

    for rea_nr, rida in enumerate(
        tööleht.iter_rows(min_row=2, values_only=True),
        start=2
    ):
        väärtused = list(rida)

        if not any(v is not None for v in väärtused):
            continue

        partno = (
            väärtused[partno_veerg - 1]
            if len(väärtused) >= partno_veerg
            else None
        )

        mõõtühik = (
            väärtused[mõõtühiku_veerg - 1]
            if len(väärtused) >= mõõtühiku_veerg
            else None
        )

        tulemused.append({
            "rea_nr": rea_nr,
            "partno": partno,
            "partno_võti": normaliseeri(partno),
            "mõõtühik": mõõtühik,
            "mõõtühik_võti": normaliseeri(mõõtühik),
            "andmed": väärtused
        })

    töövihik.close()
    return tulemused


# ============================================================
# 5. ANDMETE VÕRDLEMINE
# ============================================================

def võrdle_andmeid(stat_read, uom_read):
    """
    Võrdleb kirjeid partno alusel.

    Tulemused:
    - vastavused;
    - mõõtühikute erinevused;
    - ainult STAT-failis olevad kirjed;
    - ainult UOM-failis olevad kirjed;
    - võimalikud duplikaadid;
    - puuduva partno väärtusega read.
    """

    stat_indeks = defaultdict(list)
    uom_indeks = defaultdict(list)

    stat_partno_puudub = []
    uom_partno_puudub = []

    for kirje in stat_read:
        if kirje["partno_võti"]:
            stat_indeks[kirje["partno_võti"]].append(kirje)
        else:
            stat_partno_puudub.append([
                kirje["rea_nr"],
                kirje["mõõtühik"]
            ])

    for kirje in uom_read:
        if kirje["partno_võti"]:
            uom_indeks[kirje["partno_võti"]].append(kirje)
        else:
            uom_partno_puudub.append([
                kirje["rea_nr"],
                kirje["mõõtühik"]
            ])

    vastavused = []
    erinevused = []
    ainult_stat = []
    ainult_uom = []
    duplikaadid = []

    kõik_võtmed = set(stat_indeks) | set(uom_indeks)

    for võti in sorted(kõik_võtmed):
        stat_kirjed = stat_indeks.get(võti, [])
        uom_kirjed = uom_indeks.get(võti, [])

        if len(stat_kirjed) > 1 or len(uom_kirjed) > 1:
            duplikaadid.append([
                võti,
                len(stat_kirjed),
                len(uom_kirjed),
                ", ".join(str(x["rea_nr"]) for x in stat_kirjed),
                ", ".join(str(x["rea_nr"]) for x in uom_kirjed)
            ])

        if stat_kirjed and uom_kirjed:
            for stat in stat_kirjed:
                for uom in uom_kirjed:
                    rida = [
                        võti,
                        stat["rea_nr"],
                        stat["partno"],
                        stat["mõõtühik"],
                        uom["rea_nr"],
                        uom["partno"],
                        uom["mõõtühik"]
                    ]

                    vastavused.append(rida)

                    if (
                        stat["mõõtühik_võti"]
                        != uom["mõõtühik_võti"]
                    ):
                        erinevused.append(rida)

        elif stat_kirjed:
            for stat in stat_kirjed:
                ainult_stat.append([
                    võti,
                    stat["rea_nr"],
                    stat["partno"],
                    stat["mõõtühik"]
                ])

        else:
            for uom in uom_kirjed:
                ainult_uom.append([
                    võti,
                    uom["rea_nr"],
                    uom["partno"],
                    uom["mõõtühik"]
                ])

    return {
        "vastavused": vastavused,
        "erinevused": erinevused,
        "ainult_stat": ainult_stat,
        "ainult_uom": ainult_uom,
        "duplikaadid": duplikaadid,
        "stat_partno_puudub": stat_partno_puudub,
        "uom_partno_puudub": uom_partno_puudub
    }


# ============================================================
# 6. EXCELI RAPORTI KOOSTAMINE
# ============================================================

def koosta_raport(stat_read, uom_read, tulemused, failitee):
    """Salvestab võrdluse tulemused uude Exceli faili."""

    töövihik = Workbook()
    kokkuvõte = töövihik.active
    kokkuvõte.title = "Kokkuvote"

    kokkuvõtte_read = [
        ["Näitaja", "Tulemus"],
        ["STAT-faili sisuga read", len(stat_read)],
        ["Measure Unit faili sisuga read", len(uom_read)],
        ["Vastavused partno alusel", len(tulemused["vastavused"])],
        ["Mõõtühikute erinevused", len(tulemused["erinevused"])],
        ["Ainult STAT-failis", len(tulemused["ainult_stat"])],
        ["Ainult Measure Unit failis", len(tulemused["ainult_uom"])],
        ["Võimalikud duplikaadid", len(tulemused["duplikaadid"])],
        ["STAT read puuduva partno väärtusega",
         len(tulemused["stat_partno_puudub"])],
        ["UOM read puuduva partno väärtusega",
         len(tulemused["uom_partno_puudub"])],
        ["Algfaile muudeti", "EI"],
        ["Uusi andmeid lisati", "EI"],
        ["Võrdlusvõti", "STAT veerg C ↔ UOM partno"],
        ["Mõõtühiku võrdlus", "STAT veerg I ↔ UOM measure_unit"]
    ]

    for rida in kokkuvõtte_read:
        kokkuvõte.append(rida)

    päised_vastavus = [
        "partno võti",
        "STAT rea nr",
        "STAT partno",
        "STAT measure_unit",
        "UOM rea nr",
        "UOM partno",
        "UOM measure_unit"
    ]

    lisa_raporti_leht(
        töövihik,
        "Vastavused",
        päised_vastavus,
        tulemused["vastavused"]
    )

    lisa_raporti_leht(
        töövihik,
        "Unit erinevused",
        päised_vastavus,
        tulemused["erinevused"]
    )

    lisa_raporti_leht(
        töövihik,
        "Ainult STAT",
        ["partno võti", "STAT rea nr", "STAT partno", "STAT unit"],
        tulemused["ainult_stat"]
    )

    lisa_raporti_leht(
        töövihik,
        "Ainult UOM",
        ["partno võti", "UOM rea nr", "UOM partno", "UOM unit"],
        tulemused["ainult_uom"]
    )

    lisa_raporti_leht(
        töövihik,
        "Duplikaadid",
        [
            "partno võti",
            "STAT kirjete arv",
            "UOM kirjete arv",
            "STAT reanumbrid",
            "UOM reanumbrid"
        ],
        tulemused["duplikaadid"]
    )

    lisa_raporti_leht(
        töövihik,
        "STAT puuduv partno",
        ["STAT rea nr", "STAT measure_unit"],
        tulemused["stat_partno_puudub"]
    )

    lisa_raporti_leht(
        töövihik,
        "UOM puuduv partno",
        ["UOM rea nr", "UOM measure_unit"],
        tulemused["uom_partno_puudub"]
    )

    stiili_tööleht(kokkuvõte)
    kokkuvõte.column_dimensions["A"].width = 43
    kokkuvõte.column_dimensions["B"].width = 55

    failitee.parent.mkdir(parents=True, exist_ok=True)
    töövihik.save(failitee)
    töövihik.close()


# ============================================================
# 7. PROGRAMMI KÄIVITAMINE
# ============================================================

def main():
    print("=" * 60)
    print("STAT + MEASURE UNIT VÕRDLUSAUTOMAATIKA")
    print("=" * 60)

    # Kontrollime failide olemasolu.
    if not STAT_FAIL.exists():
        raise FileNotFoundError(
            f"STAT-faili ei leitud:\n{STAT_FAIL}"
        )

    if not UOM_FAIL.exists():
        raise FileNotFoundError(
            f"Measure Unit faili ei leitud:\n{UOM_FAIL}"
        )

    print("\n1. Loen STAT-faili...")
    stat_read = loe_stat_fail(STAT_FAIL)

    print("2. Loen Measure Unit faili...")
    uom_read = loe_uom_fail(UOM_FAIL)

    print("3. Võrdlen andmeid partno alusel...")
    tulemused = võrdle_andmeid(stat_read, uom_read)

    print("4. Koostan Exceli raporti...")
    koosta_raport(
        stat_read,
        uom_read,
        tulemused,
        RAPORT_FAIL
    )

    print("\n" + "=" * 60)
    print("VÕRDLUS VALMIS")
    print("=" * 60)

    print(f"STAT sisuga ridu: {len(stat_read)}")
    print(f"UOM sisuga ridu: {len(uom_read)}")
    print(f"Vastavusi: {len(tulemused['vastavused'])}")
    print(f"Mõõtühikute erinevusi: {len(tulemused['erinevused'])}")
    print(f"Ainult STAT-is: {len(tulemused['ainult_stat'])}")
    print(f"Ainult UOM-is: {len(tulemused['ainult_uom'])}")
    print(f"Võimalikke duplikaate: {len(tulemused['duplikaadid'])}")

    print("\nRaport salvestati:")
    print(RAPORT_FAIL)

    print("\nTURVALISUS:")
    print("- Algseid Exceli faile ei muudetud.")
    print("- Andmeid ei lisatud ega kustutatud.")
    print("- Tulemuseks loodi eraldi võrdlusraport.")


if __name__ == "__main__":
    try:
        main()
    except Exception as viga:
        print("\nPROGRAMMI TÖÖS TEKKIS VIGA:")
        print(viga)

    # Terminalis ei nõua täiendavat klahvivajutust.
