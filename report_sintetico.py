# -*- coding: utf-8 -*-
"""
Versione breve del report di posizionamento.

Non usa il template Word: costruisce il documento da zero con python-docx, in
quattro-cinque pagine. Serve a chi vuole il quadro in due minuti — sintesi,
tabella dei nove indicatori, priorità — e non l'analisi distesa.

Riusa il motore narrativo e gli aiuti di formattazione di report_corp, così i
numeri e il modo di raccontarli restano gli stessi del report esteso.
"""

import io

import numpy as np
import pandas as pd
from docx import Document
from docx.enum.section import WD_ORIENT
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Cm, Pt, RGBColor

from identificazione_azienda import riga_target
from report_corp import (
    ANNI_SERIE,
    con_articolo,
    format_euro,
    inserisci_in_ordine,
    nome_regione_breve,
    posizione_ordinale,
    preposizione_regione,
    serie_indicatore,
)

BLU = RGBColor(0x1F, 0x35, 0x5E)
GRIGIO = RGBColor(0x59, 0x59, 0x59)
VERDE = RGBColor(0x1B, 0x7F, 0x3B)
ROSSO = RGBColor(0xB3, 0x26, 0x26)

# (chiave, colonna base, nome esteso, unità, è inverso)
INDICATORI = [
    ('ebitda',    'Margine EBITDA (*) %',                           'Margine EBITDA',              '%', False),
    ('ebit',      'Margine EBIT (*) %',                             'Margine EBIT',                '%', False),
    ('profitto',  'Margine di Profitto (*) %',                      'Margine di Profitto',         '%', False),
    ('strut1',    'Indice di Struttura 1° livello (*)',             'Indice di Struttura 1° liv.', '',  False),
    ('strut2',    'Indice di Struttura 2° livello (*)',             'Indice di Struttura 2° liv.', '',  False),
    ('gearing',   'Gearing (*) %',                                  'Gearing',                     '%', True),
    ('cr',        'Current Ratio (*)',                              'Current Ratio',               '',  False),
    ('qr',        'Quick Ratio (*)',                                'Quick Ratio',                 '',  False),
    ('rotazione', 'Indice di Rotazione del Capitale Investito (*)', 'Rotazione Cap. Investito',    '',  False),
]

AREE = {
    'Economico':    ['ebitda', 'ebit', 'profitto'],
    'Patrimoniale': ['strut1', 'strut2', 'gearing'],
    'Finanziario':  ['cr', 'qr', 'rotazione'],
}


# ---------------------------------------------------------------- calcolo ---

def _terzili(serie):
    s = pd.to_numeric(serie, errors='coerce').dropna()
    return (s.quantile(1 / 3), s.quantile(2 / 3)) if not s.empty else (np.nan, np.nan)


def _classe(valore, t1, t2, inverso):
    """1 = terzile migliore. Stessa convenzione del report esteso."""
    if valore is None or pd.isna(valore) or pd.isna(t1):
        return 3
    if inverso:
        return 1 if valore <= t1 else (2 if valore <= t2 else 3)
    return 1 if valore >= t2 else (2 if valore >= t1 else 3)


def _lettera_area(somma):
    return 'A' if somma <= 4 else ('B' if somma <= 7 else 'C')


def calcola_quadro(df, chiave_target, azienda_target):
    """Raccoglie tutto ciò che serve al documento in un unico dizionario."""
    riga = riga_target(df, chiave_target, azienda_target)
    if riga.empty:
        raise ValueError("Azienda target non trovata nel campione.")
    riga = riga.iloc[0]

    quadro = {'indicatori': {}, 'punti': {}}
    for chiave, base, nome, unita, inverso in INDICATORI:
        col = f"{base} 2024"
        az, sett = serie_indicatore(df, riga, base)
        t1, t2 = _terzili(df[col]) if col in df.columns else (np.nan, np.nan)
        punti = _classe(az.get('2024'), t1, t2, inverso)
        quadro['punti'][chiave] = punti
        quadro['indicatori'][chiave] = {
            'nome': nome, 'unita': unita, 'inverso': inverso,
            'serie_az': az, 'serie_set': sett,
            'valore': az.get('2024'), 'mediana': sett.get('2024'),
            'classe': 'ABC'[punti - 1],
        }

    for area, chiavi in AREE.items():
        quadro[f'rating_{area.lower()}'] = _lettera_area(sum(quadro['punti'][k] for k in chiavi))
    quadro['rating_combinato'] = (quadro['rating_economico'] + quadro['rating_patrimoniale']
                                  + quadro['rating_finanziario'])

    col_reg = next((c for c in df.columns if 'nuts2' in str(c).lower()), None)
    reg = str(riga.get(col_reg, '')).split(' - ')[-1] if col_reg else ''
    quadro['regione'] = nome_regione_breve(reg)
    quadro['panel'] = len(df)
    quadro['vdp'] = pd.to_numeric(riga.get('Totale valore della produzione migl EUR 2024'), errors='coerce')
    quadro['attivo'] = pd.to_numeric(riga.get('Totale Attivo migl EUR 2024'), errors='coerce')
    quadro['dipendenti'] = pd.to_numeric(riga.get('Numero dipendenti 2024'), errors='coerce')

    col_vdp = 'Totale valore della produzione migl EUR 2024'
    if col_vdp in df.columns and pd.notna(quadro['vdp']):
        rango = df[col_vdp].rank(ascending=False, method='min')
        quadro['rango_vdp'] = int(rango.loc[riga.name]) if riga.name in rango.index else None
    else:
        quadro['rango_vdp'] = None
    return quadro


# ------------------------------------------------------------- documento ---

def _stile(doc):
    normale = doc.styles['Normal']
    normale.font.name = 'Calibri'
    normale.font.size = Pt(10)
    normale.paragraph_format.space_after = Pt(6)
    normale.paragraph_format.line_spacing = 1.15


def _titolo(doc, testo, dimensione=16, spazio_prima=14):
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(spazio_prima)
    p.paragraph_format.space_after = Pt(4)
    r = p.add_run(testo)
    r.bold = True
    r.font.size = Pt(dimensione)
    r.font.color.rgb = BLU
    return p


def _riga_tabella(cella, testo, grassetto=False, colore=None, dim=9, centro=False):
    par = cella.paragraphs[0]
    if centro:
        par.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = par.add_run(testo)
    run.bold = grassetto
    run.font.size = Pt(dim)
    if colore is not None:
        run.font.color.rgb = colore


def _bordi(tabella):
    tblPr = tabella._tbl.tblPr
    bordi = OxmlElement('w:tblBorders')
    for lato in ('top', 'left', 'bottom', 'right', 'insideH', 'insideV'):
        el = OxmlElement(f'w:{lato}')
        el.set(qn('w:val'), 'single')
        el.set(qn('w:sz'), '4')
        el.set(qn('w:color'), 'BFBFBF')
        bordi.append(el)
    inserisci_in_ordine(tblPr, bordi)


def _sfondo(cella, colore_hex):
    shd = OxmlElement('w:shd')
    shd.set(qn('w:val'), 'clear')
    shd.set(qn('w:fill'), colore_hex)
    inserisci_in_ordine(cella._tc.get_or_add_tcPr(), shd)


def _freccia(serie):
    """Verso dell'indicatore fra il primo e l'ultimo anno disponibile."""
    punti = [serie.get(a) for a in ANNI_SERIE if serie.get(a) is not None]
    if len(punti) < 2:
        return '–', GRIGIO
    if punti[-1] > punti[0]:
        return '▲', GRIGIO
    if punti[-1] < punti[0]:
        return '▼', GRIGIO
    return '=', GRIGIO


def genera_report_sintetico(df_orbis, azienda_target, settore_nace, chiave_target=None):
    """Costruisce il report breve e restituisce il .docx in memoria."""
    q = calcola_quadro(df_orbis, chiave_target, azienda_target)
    doc = Document()
    _stile(doc)
    sez = doc.sections[0]
    sez.top_margin = sez.bottom_margin = Cm(1.8)
    sez.left_margin = sez.right_margin = Cm(2.0)

    # ---- intestazione ----
    p = doc.add_paragraph()
    r = p.add_run(str(azienda_target))
    r.bold = True
    r.font.size = Pt(20)
    r.font.color.rgb = BLU
    p.paragraph_format.space_after = Pt(0)

    p = doc.add_paragraph()
    r = p.add_run(f"Report di posizionamento in forma breve · Settore {settore_nace} · Esercizio 2024")
    r.font.size = Pt(10)
    r.font.color.rgb = GRIGIO

    # ---- sintesi ----
    _titolo(doc, "In sintesi", 14, spazio_prima=10)
    p = doc.add_paragraph()
    r = p.add_run(f"Rating Combinato {q['rating_combinato']}")
    r.bold = True
    r.font.size = Pt(13)
    r.font.color.rgb = BLU
    r2 = p.add_run(f"    (Economico {q['rating_economico']} · Patrimoniale "
                   f"{q['rating_patrimoniale']} · Finanziario {q['rating_finanziario']})")
    r2.font.size = Pt(10)
    r2.font.color.rgb = GRIGIO

    ind = q['indicatori']
    sotto_mediana = [i['nome'] for i in ind.values()
                     if i['valore'] is not None and i['mediana'] is not None
                     and ((i['valore'] > i['mediana']) if i['inverso'] else (i['valore'] < i['mediana']))]
    doc.add_paragraph(
        f"Su nove indicatori, {len(ind) - len(sotto_mediana)} si collocano al di sopra del "
        f"riferimento di settore e {len(sotto_mediana)} al di sotto. "
        f"Il confronto è fatto sulle {format_euro(q['panel'], 0)} imprese del panel."
    )

    # ---- profilo ----
    _titolo(doc, "Profilo dell'impresa", 14)
    voci = []
    if pd.notna(q['vdp']):
        voci.append(f"un Valore della Produzione 2024 di € {format_euro(q['vdp'] / 1000)} mln")
    if pd.notna(q['attivo']):
        voci.append(f"un attivo di € {format_euro(q['attivo'] / 1000)} mln")
    if pd.notna(q['dipendenti']):
        voci.append(f"{format_euro(q['dipendenti'], 0)} dipendenti")
    testo = ""
    if voci:
        elenco = voci[0] if len(voci) == 1 else ", ".join(voci[:-1]) + f" e {voci[-1]}"
        testo = f"L'impresa presenta {elenco}."
    if q['regione']:
        testo += f" Ha sede {preposizione_regione(q['regione'])}."
    if q['rango_vdp']:
        testo += (f" Per dimensione occupa il {posizione_ordinale(q['rango_vdp'])} posto "
                  f"del panel per Valore della Produzione.")
    doc.add_paragraph(testo.strip())

    # ---- tabella dei nove indicatori ----
    _titolo(doc, "I nove indicatori a confronto", 14)
    t = doc.add_table(rows=1, cols=7)
    t.alignment = WD_TABLE_ALIGNMENT.CENTER
    _bordi(t)
    intest = ['Indicatore', '2021', '2022', '2023', '2024', 'Settore 2024', 'Classe']
    for i, testo_col in enumerate(intest):
        _riga_tabella(t.rows[0].cells[i], testo_col, grassetto=True, dim=9,
                      colore=RGBColor(0xFF, 0xFF, 0xFF), centro=(i > 0))
        _sfondo(t.rows[0].cells[i], '1F355E')

    for chiave, _, nome, unita, inverso in INDICATORI:
        dati = ind[chiave]
        riga = t.add_row()
        _riga_tabella(riga.cells[0], nome, grassetto=True, dim=9)
        for j, anno in enumerate(ANNI_SERIE, start=1):
            v = dati['serie_az'].get(anno)
            _riga_tabella(riga.cells[j], "n.d." if v is None else f"{format_euro(v)}{unita}",
                          dim=9, centro=True, grassetto=(anno == '2024'))
        med = dati['mediana']
        _riga_tabella(riga.cells[5], "n.d." if med is None else f"{format_euro(med)}{unita}",
                      dim=9, centro=True, colore=GRIGIO)
        colore = {'A': VERDE, 'B': GRIGIO, 'C': ROSSO}[dati['classe']]
        _riga_tabella(riga.cells[6], dati['classe'], grassetto=True, dim=10, centro=True, colore=colore)

    larghezze = [Cm(4.6), Cm(1.9), Cm(1.9), Cm(1.9), Cm(2.1), Cm(2.6), Cm(1.7)]
    for riga in t.rows:
        for cella, larghezza in zip(riga.cells, larghezze):
            cella.width = larghezza

    p = doc.add_paragraph()
    r = p.add_run("Classe A: primo terzile del settore · B: secondo · C: terzo. "
                  "Per il Gearing l'ordine è invertito, perché un valore più basso indica meno debito.")
    r.font.size = Pt(8)
    r.font.color.rgb = GRIGIO

    # ---- lettura per area ----
    _titolo(doc, "Come leggerli", 14)
    for area, chiavi in AREE.items():
        p = doc.add_paragraph()
        r = p.add_run(f"{area} (Classe {q['rating_' + area.lower()]}) — ")
        r.bold = True
        r.font.size = Pt(10)
        r.font.color.rgb = BLU
        # "meglio/peggio" invece di "sopra/sotto": per il Gearing un valore più alto
        # è un risultato peggiore, e "sopra il settore" suonerebbe come un pregio
        meglio = [ind[k]['nome'] for k in chiavi
                  if ind[k]['valore'] is not None and ind[k]['mediana'] is not None
                  and ((ind[k]['valore'] < ind[k]['mediana']) if ind[k]['inverso']
                       else (ind[k]['valore'] > ind[k]['mediana']))]
        peggio = [ind[k]['nome'] for k in chiavi if ind[k]['nome'] not in meglio]
        parti = []
        for etichetta, gruppo in (("meglio del settore", meglio), ("sotto il riferimento di settore", peggio)):
            if not gruppo:
                continue
            dettaglio = "su tutti e tre gli indicatori" if len(gruppo) == 3 else "su " + ", ".join(gruppo)
            parti.append(f"{etichetta} {dettaglio}")
        p.add_run("; ".join(parti) + ".").font.size = Pt(10)

    # ---- priorità ----
    _titolo(doc, "Dove guardare", 14)
    priorita = []
    if ind['ebitda']['valore'] is not None and ind['ebitda']['mediana'] is not None \
            and ind['ebitda']['valore'] < ind['ebitda']['mediana']:
        priorita.append("il recupero dei margini operativi")
    if ind['gearing']['valore'] is not None and ind['gearing']['mediana'] is not None \
            and ind['gearing']['valore'] > ind['gearing']['mediana']:
        priorita.append("il contenimento dell'indebitamento")
    if (ind['cr']['valore'] or 0) < 1 or (ind['qr']['valore'] or 0) < 1:
        priorita.append("il presidio della liquidità di breve periodo")
    if priorita:
        for voce in priorita:
            p = doc.add_paragraph(style='List Bullet')
            p.add_run(voce[0].upper() + voce[1:] + ".").font.size = Pt(10)
    else:
        doc.add_paragraph("Nessuna delle tre aree presenta scostamenti tali da richiedere un intervento "
                          "prioritario: il presidio ordinario è sufficiente.")

    # ---- nota ----
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(16)
    r = p.add_run(
        f"Nota — Il panel comprende {format_euro(q['panel'], 0)} imprese del settore {settore_nace} "
        f"estratte da ORBIS. Il settore è descritto dalla mediana, che a differenza della media non "
        f"viene spostata dai pochi valori estremi del campione. Le classi sono calcolate sui terzili "
        f"della distribuzione 2024."
    )
    r.font.size = Pt(8)
    r.font.color.rgb = GRIGIO

    uscita = io.BytesIO()
    doc.save(uscita)
    uscita.seek(0)
    return uscita
