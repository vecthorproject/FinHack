# -*- coding: utf-8 -*-
"""
Foglio di alert sulle oscillazioni, per chi deve commentare il trend.

Il report e la presentazione raccontano il percorso di ogni indicatore, ma chi
scrive il commento finale ha bisogno di sapere in fretta *dove guardare*: quali
indicatori si sono mossi davvero nel quadriennio, fra quali anni, e di quanto.
Questo foglio mette in fila i nove indicatori ordinati per quanto hanno oscillato,
con la domanda a cui il commento dovrebbe rispondere.

I numeri e la lettura del percorso arrivano dallo stesso motore della
presentazione, quindi foglio e slide non possono dire cose diverse.
"""

import io

import pandas as pd
import xlsxwriter

from identificazione_azienda import riga_target
from report_breve_corp import analizza_percorso

ANNI = ('2021', '2022', '2023', '2024')

# (chiave, colonna base, nome leggibile, nome con l'articolo, unita', indicatore inverso)
INDICATORI_ALERT = [
    ('ebitda',    'Margine EBITDA (*) %',                           'Margine EBITDA',                            'il Margine EBITDA',                            '%', False),
    ('ebit',      'Margine EBIT (*) %',                             'Margine EBIT',                              'il Margine EBIT',                              '%', False),
    ('profitto',  'Margine di Profitto (*) %',                      'Margine di Profitto',                       'il Margine di Profitto',                       '%', False),
    ('strut1',    'Indice di Struttura 1° livello (*)',             'Indice di Struttura di 1° livello',         "l'Indice di Struttura di 1° livello",          '',  False),
    ('strut2',    'Indice di Struttura 2° livello (*)',             'Indice di Struttura di 2° livello',         "l'Indice di Struttura di 2° livello",          '',  False),
    ('gearing',   'Gearing (*) %',                                  'Gearing',                                   'il Gearing',                                   '%', True),
    ('cr',        'Current Ratio (*)',                              'Current Ratio',                             'il Current Ratio',                             '',  False),
    ('qr',        'Quick Ratio (*)',                                'Quick Ratio',                               'il Quick Ratio',                               '',  False),
    ('rotazione', 'Indice di Rotazione del Capitale Investito (*)', 'Indice di Rotazione del Capitale Investito', "l'Indice di Rotazione del Capitale Investito", '',  False),
]

# Soglie di attenzione. Non hanno pretese statistiche: servono a mettere in cima
# alla lista quello che chi commenta non puo' permettersi di non aver guardato.
SOGLIA_ALTA_SALTO = 50.0
SOGLIA_ALTA_ESCURSIONE = 60.0
SOGLIA_MEDIA_SALTO = 20.0
SOGLIA_MEDIA_ESCURSIONE = 30.0

COLORI = {
    'ALTA':  {'sfondo': '#FBE9E7', 'testo': '#B71C1C'},
    'MEDIA': {'sfondo': '#FFF8E1', 'testo': '#B26A00'},
    'BASSA': {'sfondo': '#E8F5E9', 'testo': '#1B5E20'},
}


def _numero(valore):
    """Formattazione all'italiana, con la virgola decimale."""
    if valore is None or pd.isna(valore):
        return 'n.d.'
    return f"{valore:,.2f}".replace(',', 'X').replace('.', ',').replace('X', '.')


def _livello(salto_pct, escursione_pct):
    if abs(salto_pct) >= SOGLIA_ALTA_SALTO or escursione_pct >= SOGLIA_ALTA_ESCURSIONE:
        return 'ALTA'
    if abs(salto_pct) >= SOGLIA_MEDIA_SALTO or escursione_pct >= SOGLIA_MEDIA_ESCURSIONE:
        return 'MEDIA'
    return 'BASSA'


def _domanda(nome_con_articolo, dati, variazione_periodo):
    """La domanda a cui il commento dovrebbe rispondere."""
    salto = dati.get('salto')
    if not salto:
        return "Serie storica incompleta: verificare i dati prima di commentare."
    anno_da, anno_a, variazione = salto
    escursione = dati.get('escursione_pct', 0)
    if abs(variazione) < SOGLIA_MEDIA_SALTO and escursione < SOGLIA_MEDIA_ESCURSIONE:
        # Nessuno scalino, ma il quadriennio puo' comunque essersi spostato piano piano.
        if abs(variazione_periodo) >= 10:
            return ("Il valore si sposta in modo graduale, senza strappi fra un esercizio e "
                    "l'altro: al commento basta la direzione complessiva del quadriennio.")
        return "Nessun movimento rilevante: si pu\u00f2 confermare la stabilit\u00e0 del periodo."
    verso = 'salito' if variazione > 0 else 'sceso'
    domanda = (f"Che cosa spiega {nome_con_articolo} {verso} del {_numero(abs(variazione))}% "
               f"fra il {anno_da} e il {anno_a}?")
    if dati.get('trend') == 'altalenante':
        domanda += (" Attenzione: il valore torna vicino al punto di partenza, quindi il "
                    "confronto 2021-2024 da solo nasconde il movimento.")
    return domanda


def calcola_alert(df_orbis, azienda_target, chiave_target=None):
    """Una riga per indicatore, ordinate da quella che si e' mossa di piu'."""
    riga = riga_target(df_orbis, chiave_target, azienda_target)
    if riga.empty:
        raise ValueError("Azienda target non trovata nel campione.")
    riga = riga.iloc[0]

    righe = []
    for chiave, base, nome, nome_articolo, unita, inverso in INDICATORI_ALERT:
        serie = [(anno, pd.to_numeric(riga.get(f"{base} {anno}"), errors='coerce')) for anno in ANNI]
        dati = analizza_percorso(serie, inverso=inverso, unita=unita)
        salto = dati.get('salto')
        salto_pct = salto[2] if salto else 0.0
        escursione = dati.get('escursione_pct', 0.0)
        validi = [v for _, v in serie if v is not None and not pd.isna(v)]
        variazione_periodo = 0.0
        if len(validi) >= 2 and abs(validi[0]) > 0.01:
            variazione_periodo = (validi[-1] - validi[0]) / abs(validi[0]) * 100
        righe.append({
            'chiave': chiave,
            'nome': nome,
            'unita': unita,
            'valori': {anno: valore for anno, valore in serie},
            'trend': dati.get('trend', 'n.d.'),
            'percorso': dati.get('frase', ''),
            'escursione': escursione,
            'salto_anni': f"{salto[0]}→{salto[1]}" if salto else 'n.d.',
            'salto_pct': salto_pct,
            'livello': _livello(salto_pct, escursione),
            'variazione_periodo': variazione_periodo,
            'domanda': _domanda(nome_articolo, dati, variazione_periodo),
        })

    ordine = {'ALTA': 0, 'MEDIA': 1, 'BASSA': 2}
    righe.sort(key=lambda r: (ordine[r['livello']], -abs(r['salto_pct']), -r['escursione']))
    return righe


def genera_foglio_alert(df_orbis, azienda_target, settore_nace, chiave_target=None):
    """Costruisce il foglio di alert e lo restituisce in memoria."""
    righe = calcola_alert(df_orbis, azienda_target, chiave_target)

    uscita = io.BytesIO()
    libro = xlsxwriter.Workbook(uscita, {'in_memory': True})
    foglio = libro.add_worksheet('Alert Trend')
    foglio.hide_gridlines(2)

    f_titolo = libro.add_format({'bold': True, 'font_size': 16, 'font_color': '#1F3352'})
    f_sotto = libro.add_format({'font_size': 10, 'font_color': '#64748B'})
    f_intest = libro.add_format({
        'bold': True, 'font_size': 10, 'font_color': 'white', 'bg_color': '#1F3352',
        'align': 'center', 'valign': 'vcenter', 'text_wrap': True, 'border': 1, 'border_color': '#1F3352',
    })
    f_nome = libro.add_format({'bold': True, 'font_size': 10, 'valign': 'vcenter',
                               'text_wrap': True, 'border': 1, 'border_color': '#D9D9D9'})
    f_num = libro.add_format({'font_size': 10, 'align': 'right', 'valign': 'vcenter',
                              'border': 1, 'border_color': '#D9D9D9'})
    f_num_2024 = libro.add_format({'font_size': 10, 'bold': True, 'align': 'right',
                                   'valign': 'vcenter', 'border': 1, 'border_color': '#D9D9D9'})
    f_testo = libro.add_format({'font_size': 10, 'valign': 'top', 'text_wrap': True,
                                'border': 1, 'border_color': '#D9D9D9'})
    livelli = {
        nome: libro.add_format({
            'bold': True, 'font_size': 10, 'align': 'center', 'valign': 'vcenter',
            'bg_color': colori['sfondo'], 'font_color': colori['testo'],
            'border': 1, 'border_color': '#D9D9D9',
        })
        for nome, colori in COLORI.items()
    }

    foglio.write(0, 0, 'Alert sulle oscillazioni del quadriennio', f_titolo)
    foglio.write(1, 0, f"{azienda_target} · settore {settore_nace} · esercizi 2021-2024", f_sotto)
    foglio.write(2, 0, 'Gli indicatori sono ordinati da quello che si è mosso di più: '
                       'la colonna "Da chiarire nel commento" dice a cosa rispondere.', f_sotto)

    colonne = ['Indicatore', '2021', '2022', '2023', '2024', 'Andamento',
               'Escursione max-min', 'Scalino più forte', '%', 'Attenzione',
               'Il percorso', 'Da chiarire nel commento']
    larghezze = [30, 11, 11, 11, 11, 16, 17, 16, 10, 12, 52, 62]
    for i, (testo, larghezza) in enumerate(zip(colonne, larghezze)):
        foglio.set_column(i, i, larghezza)
        foglio.write(4, i, testo, f_intest)
    foglio.set_row(4, 32)
    foglio.freeze_panes(5, 1)

    for n, r in enumerate(righe):
        riga_xl = 5 + n
        foglio.set_row(riga_xl, 46)
        foglio.write(riga_xl, 0, r['nome'], f_nome)
        for j, anno in enumerate(ANNI, start=1):
            valore = r['valori'].get(anno)
            testo = 'n.d.' if valore is None or pd.isna(valore) else f"{_numero(valore)}{r['unita']}"
            foglio.write(riga_xl, j, testo, f_num_2024 if anno == '2024' else f_num)
        foglio.write(riga_xl, 5, r['trend'].capitalize(), f_num)
        foglio.write(riga_xl, 6, f"{_numero(r['escursione'])}%", f_num)
        foglio.write(riga_xl, 7, r['salto_anni'], f_num)
        foglio.write(riga_xl, 8, f"{_numero(r['salto_pct'])}%", f_num)
        foglio.write(riga_xl, 9, r['livello'], livelli[r['livello']])
        foglio.write(riga_xl, 10, r['percorso'], f_testo)
        foglio.write(riga_xl, 11, r['domanda'], f_testo)

    # --- come si leggono questi numeri ---------------------------------------
    metodo = libro.add_worksheet('Come si legge')
    metodo.hide_gridlines(2)
    metodo.set_column(0, 0, 110)
    f_par = libro.add_format({'font_size': 11, 'text_wrap': True, 'valign': 'top',
                              'font_color': '#1F3352'})
    f_tit = libro.add_format({'bold': True, 'font_size': 13, 'font_color': '#1F3352'})
    testi = [
        ('Escursione max-min',
         "Differenza fra il valore piu' alto e il piu' basso del quadriennio, rapportata al "
         "valore di partenza. Dice quanto l'indicatore si e' mosso, a prescindere da dove ha "
         "chiuso: un indicatore che parte e finisce allo stesso livello puo' avere "
         "un'escursione molto ampia."),
        ("Scalino piu' forte",
         "La variazione piu' marcata fra due esercizi consecutivi, con gli anni in cui e' "
         "avvenuta. E' il punto che un commento serio deve saper spiegare."),
        ('Andamento',
         "\"In crescita\" e \"in contrazione\" guardano il primo e l'ultimo anno. "
         "\"Altalenante\" segnala i casi in cui gli estremi si equivalgono ma nel mezzo il "
         "valore si e' mosso: li' dire \"stabile\" sarebbe falso. Sul Gearing la lettura si "
         "inverte, perche' un valore piu' basso e' un risultato migliore."),
        ('Attenzione',
         f"ALTA quando lo scalino supera il {int(SOGLIA_ALTA_SALTO)}% o l'escursione il "
         f"{int(SOGLIA_ALTA_ESCURSIONE)}%; MEDIA sopra il {int(SOGLIA_MEDIA_SALTO)}% e il "
         f"{int(SOGLIA_MEDIA_ESCURSIONE)}%; BASSA sotto. Sono soglie di lettura, non test "
         "statistici: servono a ordinare la lista."),
        ('Coerenza con gli altri documenti',
         "I valori, il percorso e le etichette di andamento sono calcolati dallo stesso "
         "motore che alimenta il report Word e la presentazione: i tre documenti dicono le "
         "stesse cose."),
    ]
    metodo.write(0, 0, 'Come si legge il foglio di alert', f_tit)
    riga_m = 2
    for titolo, corpo in testi:
        metodo.write(riga_m, 0, titolo, f_tit)
        metodo.write(riga_m + 1, 0, corpo, f_par)
        metodo.set_row(riga_m + 1, 58)
        riga_m += 3

    libro.close()
    uscita.seek(0)
    return uscita
