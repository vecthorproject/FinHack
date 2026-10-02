"""
Il periodo di riferimento del report, in un posto solo.

Quando l'estrazione ORBIS aggiunge un esercizio, qui si aggiunge l'anno a ANNI
e tutto il resto segue: nomi delle colonne, intestazioni delle tabelle, assi dei
grafici, testi ("nel quinquennio", "nel 2025") e fogli Excel.

ANNI deve elencare gli esercizi nell'ordine in cui compaiono nei nomi delle
colonne dell'estrazione ("Margine EBITDA (*) % 2025"), dal piu' vecchio al piu'
recente.
"""

ANNI = ('2021', '2022', '2023', '2024', '2025')

PRIMO = ANNI[0]
ULTIMO = ANNI[-1]                      # l'esercizio che determina le classi
PENULTIMO = ANNI[-2]                   # l'esercizio di confronto piu' recente
PERIODO = f'{PRIMO}-{ULTIMO}'
QUANTI_ANNI = len(ANNI)

# Come si chiama il periodo nei testi: "nel quinquennio", "del quadriennio"...
_NOMI_PERIODO = {2: 'biennio', 3: 'triennio', 4: 'quadriennio', 5: 'quinquennio',
                 6: 'sessennio', 7: 'settennio'}
NOME_PERIODO = _NOMI_PERIODO.get(QUANTI_ANNI, 'periodo')

_NUMERI = {2: 'due', 3: 'tre', 4: 'quattro', 5: 'cinque', 6: 'sei', 7: 'sette'}
QUANTI_ANNI_LETTERE = _NUMERI.get(QUANTI_ANNI, str(QUANTI_ANNI))


# I template Word e PowerPoint sono stati scritti quando l'ultimo esercizio era il
# 2024: alla lavorazione gli anni citati nel loro testo fisso vengono riallineati al
# periodo di cui sopra. Se un giorno i template verranno riscritti, qui si aggiorna
# l'anno con cui sono scritti e il riallineamento diventa un'operazione a vuoto.
ANNO_TEMPLATE = '2024'
PERIODO_TEMPLATE = '2021-2024'
NOME_PERIODO_TEMPLATE = 'quadriennio'


def col(base, anno=ULTIMO):
    """Il nome della colonna ORBIS: "Margine EBITDA (*) %" + anno."""
    return f'{base} {anno}'


def colonne(base, anni=ANNI):
    """I nomi delle colonne di una variabile per tutti gli esercizi."""
    return [col(base, anno) for anno in anni]
