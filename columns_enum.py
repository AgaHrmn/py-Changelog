from enum import Enum

class Elektronarzedzia(Enum): 
    ID = 0
    NAZWA = 1
    OPIS = 2
    TYP_OSTRZA = 3
    MOC_SILNIKA = 4
    TYP_SILNIKA = 5
    TYP_ZASILANIA = 6
    CENA = 7


class Ostrza(Enum): 
    ID = 0
    NAZWA = 1
    TYP = 2
    OPIS = 3
    DLUGOSC = 4
    SREDNICA = 5
    GRUBOSC = 6
    MATERIAL = 7
    LICZBA_ZEBOW = 8
    ZASTOSOWANIE = 9
    CENA = 10 

class ElektronarzedziaReport(Enum):
    STATUS = 0
    ID = 1
    NAZWA_2023 = 2
    NAZWA_2024 = 3
    OPIS_2023 = 4
    OPIS_2024 = 5
    TYP_OSTRZA_23 = 6
    TYP_OSTRZA_24 = 7
    MOC_SILNIKA_23 = 8
    MOC_SILNIKA_24 = 9
    TYP_SILNIKA_23 = 10
    TYP_SILNIKA_24 = 11
    TYP_ZASILANIA_23 = 12
    TYP_ZASILANIA_24 = 13
    CENA_23 = 14
    CENA_24 = 15
