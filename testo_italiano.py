# -*- coding: utf-8 -*-
"""
Piccoli aiuti di lingua per i testi generati.

L'articolo determinativo italiano davanti a un numero dipende da come il numero si
legge, non da come si scrive: "lo 0,26%", "l'8,26%", "l'11,55%" ma "il 5,43%" e
"il 159,50%". Nei testi del report l'articolo era fisso ("pari al ..."), quindi
sbagliava ogni volta che il valore iniziava per 0, 8 o 11.
"""

# Forme articolate: (preposizione, articolo semplice) -> forma contratta
_ARTICOLATE = {
    ('a', 'il'): 'al',   ('a', 'lo'): 'allo',   ('a', "l'"): "all'",
    ('di', 'il'): 'del', ('di', 'lo'): 'dello', ('di', "l'"): "dell'",
}


def articolo_numero(valore_formattato):
    """
    Articolo determinativo maschile singolare corretto per un numero già formattato
    all'italiana (es. "0,26", "8,26", "1.234,50").

    Si guarda solo la prima cifra della parte intera, perché è quella che determina
    il suono iniziale: 0 → "zero" (lo), 1 e 11 → "uno"/"undici" (l'), 8/80/800/8.000
    → "otto"/"ottanta"/"ottocento"/"ottomila" (l'). Attenzione: 18 si legge
    "diciotto" e 110 "centodieci", entrambi con consonante, quindi il controllo su
    1 e 11 è di uguaglianza e non di prefisso.
    """
    intero = str(valore_formattato).split(',')[0].replace('.', '').lstrip('-+').strip()
    if not intero.isdigit():
        return 'il'
    numero = int(intero)
    if numero == 0:
        return 'lo'
    if numero in (1, 11) or intero[0] == '8':
        return "l'"
    return 'il'


def con_articolo(valore_formattato, preposizione=None):
    """
    Antepone al numero l'articolo giusto, eventualmente unito a una preposizione:

        con_articolo("0,26")            -> "lo 0,26"
        con_articolo("0,26", 'a')       -> "allo 0,26"
        con_articolo("8,26", 'a')       -> "all'8,26"
        con_articolo("159,50", 'di')    -> "del 159,50"
    """
    articolo = articolo_numero(valore_formattato)
    if preposizione:
        articolo = _ARTICOLATE[(preposizione, articolo)]
    separatore = '' if articolo.endswith("'") else ' '
    return f"{articolo}{separatore}{valore_formattato}"
