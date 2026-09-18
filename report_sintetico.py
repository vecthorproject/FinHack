# -*- coding: utf-8 -*-
"""
Report di posizionamento in forma breve, impaginato nello stile "Industry Report".

Struttura ripresa dal modello di riferimento: copertina a fondo pieno, banda di
testata su ogni pagina, indice, "Big Picture" con i semafori delle tre aree,
tabella dei dati di sintesi con le variazioni, una sezione per area e il glossario
delle formule in chiusura.

Non usa il template Word: costruisce tutto con python-docx. Numeri, terzili e
classi arrivano dallo stesso motore del report esteso, quindi i due documenti
raccontano gli stessi dati.
"""

import io

import numpy as np
import pandas as pd
from docx import Document
from docx.enum.section import WD_SECTION
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_TAB_ALIGNMENT
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Cm, Pt, RGBColor

from identificazione_azienda import riga_target
from report_corp import (
    ANNI_SERIE,
    format_euro,
    inserisci_in_ordine,
    nome_regione_breve,
    posizione_ordinale,
    preposizione_regione,
    serie_indicatore,
)

# Palette del modello di riferimento
NAVY = '1F3352'
NAVY_CHIARO = '35507A'
CORALLO = 'E8613C'
VERDE = '3DB39E'
ROSSO = 'C0392B'
GRIGIO = '7F7F7F'
RIGA_ALT = 'F4F5F7'

C_NAVY = RGBColor(0x1F, 0x33, 0x52)
C_CORALLO = RGBColor(0xE8, 0x61, 0x3C)
C_VERDE = RGBColor(0x3D, 0xB3, 0x9E)
C_ROSSO = RGBColor(0xC0, 0x39, 0x2B)
C_GRIGIO = RGBColor(0x7F, 0x7F, 0x7F)
C_BIANCO = RGBColor(0xFF, 0xFF, 0xFF)

# (chiave, colonna base, nome, unità, inverso, area)
INDICATORI = [
    ('ebitda',    'Margine EBITDA (*) %',                           'Margine EBITDA',              '%', False, 'Economica'),
    ('ebit',      'Margine EBIT (*) %',                             'Margine EBIT',                '%', False, 'Economica'),
    ('profitto',  'Margine di Profitto (*) %',                      'Margine di Profitto',         '%', False, 'Economica'),
    ('strut1',    'Indice di Struttura 1° livello (*)',             'Indice di Struttura 1° liv.', '',  False, 'Patrimoniale'),
    ('strut2',    'Indice di Struttura 2° livello (*)',             'Indice di Struttura 2° liv.', '',  False, 'Patrimoniale'),
    ('gearing',   'Gearing (*) %',                                  'Gearing',                     '%', True,  'Patrimoniale'),
    ('cr',        'Current Ratio (*)',                              'Current Ratio',               '',  False, 'Finanziaria'),
    ('qr',        'Quick Ratio (*)',                                'Quick Ratio',                 '',  False, 'Finanziaria'),
    ('rotazione', 'Indice di Rotazione del Capitale Investito (*)', 'Rotazione Cap. Investito',    '',  False, 'Finanziaria'),
]

AREE = ['Economica', 'Patrimoniale', 'Finanziaria']
CHIAVI_AREA = {a: [k for k, _, _, _, _, ar in INDICATORI if ar == a] for a in AREE}

GLOSSARIO = [
    ('Margine EBITDA', 'Risultato operativo lordo / Valore della Produzione'),
    ('Margine EBIT', 'Risultato operativo / Valore della Produzione'),
    ('Margine di Profitto', 'Utile ante imposte / Valore della Produzione'),
    ('Indice di Struttura 1° livello', 'Patrimonio netto / Immobilizzazioni'),
    ('Indice di Struttura 2° livello', '(Patrimonio netto + Passività non correnti) / Immobilizzazioni'),
    ('Gearing', 'Debiti finanziari / Patrimonio netto'),
    ('Current Ratio', 'Attività correnti / Passività correnti'),
    ('Quick Ratio', '(Attività correnti − Rimanenze) / Passività correnti'),
    ('Rotazione Cap. Investito', 'Valore della Produzione / Capitale investito'),
]


# ------------------------------------------------------------ primitive ----

def _sfondo(elemento, colore):
    shd = OxmlElement('w:shd')
    shd.set(qn('w:val'), 'clear')
    shd.set(qn('w:fill'), colore)
    inserisci_in_ordine(elemento, shd)


def _sfondo_cella(cella, colore):
    _sfondo(cella._tc.get_or_add_tcPr(), colore)


def _senza_bordi(tabella):
    bordi = OxmlElement('w:tblBorders')
    for lato in ('top', 'left', 'bottom', 'right', 'insideH', 'insideV'):
        el = OxmlElement(f'w:{lato}')
        el.set(qn('w:val'), 'none')
        el.set(qn('w:sz'), '0')
        bordi.append(el)
    inserisci_in_ordine(tabella._tbl.tblPr, bordi)


def _linea_sotto(cella, colore=NAVY, spessore='4'):
    bordi = OxmlElement('w:tcBorders')
    el = OxmlElement('w:bottom')
    el.set(qn('w:val'), 'single')
    el.set(qn('w:sz'), spessore)
    el.set(qn('w:color'), colore)
    bordi.append(el)
    inserisci_in_ordine(cella._tc.get_or_add_tcPr(), bordi)


def _testo(contenitore, testo, dim=9, grassetto=False, colore=None,
           allineamento=None, spazio_dopo=2, primo=False):
    par = contenitore.paragraphs[0] if primo and contenitore.paragraphs else contenitore.add_paragraph()
    if primo and contenitore.paragraphs:
        par.text = ''
    par.paragraph_format.space_after = Pt(spazio_dopo)
    par.paragraph_format.space_before = Pt(0)
    if allineamento is not None:
        par.alignment = allineamento
    run = par.add_run(testo)
    run.font.size = Pt(dim)
    run.bold = grassetto
    run.font.color.rgb = colore if colore is not None else C_NAVY
    return par


def _altezza_riga(riga, cm):
    trPr = riga._tr.get_or_add_trPr()
    h = OxmlElement('w:trHeight')
    h.set(qn('w:val'), str(int(cm * 567)))
    h.set(qn('w:hRule'), 'atLeast')
    trPr.append(h)


# -------------------------------------------------------------- calcolo ----

def _classe(valore, t1, t2, inverso):
    if valore is None or pd.isna(valore) or pd.isna(t1):
        return 3
    if inverso:
        return 1 if valore <= t1 else (2 if valore <= t2 else 3)
    return 1 if valore >= t2 else (2 if valore >= t1 else 3)


def calcola_quadro(df, chiave_target, azienda_target):
    riga = riga_target(df, chiave_target, azienda_target)
    if riga.empty:
        raise ValueError("Azienda target non trovata nel campione.")
    riga = riga.iloc[0]

    q = {'ind': {}, 'punti': {}}
    for chiave, base, nome, unita, inverso, area in INDICATORI:
        col = f"{base} 2024"
        az, sett = serie_indicatore(df, riga, base)
        if col in df.columns:
            s = pd.to_numeric(df[col], errors='coerce').dropna()
            t1, t2 = (s.quantile(1 / 3), s.quantile(2 / 3)) if not s.empty else (np.nan, np.nan)
        else:
            t1 = t2 = np.nan
        punti = _classe(az.get('2024'), t1, t2, inverso)
        q['punti'][chiave] = punti
        v21, v24 = az.get('2021'), az.get('2024')
        variazione = None
        if v21 not in (None, 0) and v24 is not None and abs(v21) > 1e-9:
            variazione = (v24 - v21) / abs(v21) * 100
        q['ind'][chiave] = {
            'nome': nome, 'unita': unita, 'inverso': inverso, 'area': area,
            'az': az, 'set': sett, 'valore': v24, 'mediana': sett.get('2024'),
            'classe': 'ABC'[punti - 1], 'var': variazione,
        }

    lettera = lambda somma: 'A' if somma <= 4 else ('B' if somma <= 7 else 'C')
    for area in AREE:
        q[f'rating_{area}'] = lettera(sum(q['punti'][k] for k in CHIAVI_AREA[area]))
    q['rating'] = q['rating_Economica'] + q['rating_Patrimoniale'] + q['rating_Finanziaria']

    col_reg = next((c for c in df.columns if 'nuts2' in str(c).lower()), None)
    reg = str(riga.get(col_reg, '')).split(' - ')[-1] if col_reg else ''
    q['regione'] = nome_regione_breve(reg)
    q['panel'] = len(df)
    q['vdp'] = pd.to_numeric(riga.get('Totale valore della produzione migl EUR 2024'), errors='coerce')
    q['attivo'] = pd.to_numeric(riga.get('Totale Attivo migl EUR 2024'), errors='coerce')
    q['dipendenti'] = pd.to_numeric(riga.get('Numero dipendenti 2024'), errors='coerce')

    col_vdp = 'Totale valore della produzione migl EUR 2024'
    if col_vdp in df.columns and pd.notna(q['vdp']):
        rango = df[col_vdp].rank(ascending=False, method='min')
        q['rango'] = int(rango.loc[riga.name]) if riga.name in rango.index else None
    else:
        q['rango'] = None
    return q


# ------------------------------------------------------------ impaginato ---

def _copertina(doc, azienda, settore):
    t = doc.add_table(rows=1, cols=1)
    _senza_bordi(t)
    cella = t.rows[0].cells[0]
    _sfondo_cella(cella, NAVY)
    _altezza_riga(t.rows[0], 19.5)
    cella.width = Cm(17)

    _testo(cella, '', dim=28, primo=True)
    _testo(cella, 'Report di Posizionamento', dim=30, grassetto=True, colore=C_BIANCO, spazio_dopo=2)
    p = cella.add_paragraph()
    p.paragraph_format.space_after = Pt(18)
    r = p.add_run('BASE')
    r.font.size = Pt(26)
    r.bold = True
    r.font.color.rgb = C_CORALLO

    _testo(cella, azienda, dim=17, grassetto=True, colore=C_BIANCO, spazio_dopo=4)
    _testo(cella, f'Settore {settore}', dim=11, colore=C_CORALLO, spazio_dopo=20)
    _testo(cella, 'Bilanci', dim=15, colore=C_BIANCO, spazio_dopo=2)
    _testo(cella, '2024 · 2023 · 2022 · 2021', dim=13, colore=C_BIANCO, spazio_dopo=0)


def _banda_testata(sezione, azienda, settore):
    """Banda navy ripetuta in testa a ogni pagina, col numero di pagina a destra."""
    intestazione = sezione.header
    for p in intestazione.paragraphs:
        p.text = ''
    t = intestazione.add_table(rows=1, cols=2, width=Cm(17))
    _senza_bordi(t)
    sx, dx = t.rows[0].cells
    sx.width, dx.width = Cm(15), Cm(2)
    _sfondo_cella(sx, NAVY)
    _sfondo_cella(dx, CORALLO)
    _altezza_riga(t.rows[0], 1.1)
    _testo(sx, 'REPORT DI POSIZIONAMENTO', dim=9, grassetto=True, colore=C_BIANCO, spazio_dopo=0, primo=True)
    _testo(sx, f'{azienda} · {settore}', dim=7, colore=RGBColor(0xC8, 0xD0, 0xDC), spazio_dopo=0)

    par = dx.paragraphs[0]
    par.text = ''
    par.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = par.add_run()
    run.font.size = Pt(12)
    run.bold = True
    run.font.color.rgb = C_BIANCO
    fld = OxmlElement('w:fldSimple')
    fld.set(qn('w:instr'), 'PAGE')
    run._r.addnext(fld)


def _titolo_sezione(doc, testo):
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(14)
    p.paragraph_format.space_after = Pt(6)
    r = p.add_run('▎ ')
    r.font.size = Pt(14)
    r.font.color.rgb = C_CORALLO
    r = p.add_run(testo)
    r.font.size = Pt(14)
    r.bold = True
    r.font.color.rgb = C_NAVY
    return p


def _indice(doc):
    _titolo_sezione(doc, 'Indice')
    voci = [
        ('Big Picture', 'Semafori delle tre aree e dati di sintesi'),
        ('Area Economica', 'Margini a confronto con il settore'),
        ('Area Patrimoniale', 'Copertura degli investimenti e leva'),
        ('Area Finanziaria', 'Liquidità di breve e rotazione del capitale'),
        ('Posizionamento di sintesi', 'Rating Combinato e priorità'),
        ('Note metodologiche e glossario', 'Campione, terzili, formule'),
    ]
    t = doc.add_table(rows=0, cols=1)
    _senza_bordi(t)
    for titolo, sotto in voci:
        riga = t.add_row()
        riga.cells[0].width = Cm(17)
        _testo(riga.cells[0], titolo, dim=11, grassetto=True, spazio_dopo=0, primo=True)
        _testo(riga.cells[0], sotto, dim=8, colore=C_GRIGIO, spazio_dopo=6)
        _linea_sotto(riga.cells[0], 'D9D9D9', '2')


def _big_picture(doc, q):
    _titolo_sezione(doc, 'Big Picture')
    t = doc.add_table(rows=2, cols=3)
    _senza_bordi(t)
    t.alignment = WD_TABLE_ALIGNMENT.CENTER
    etichette = {'Economica': 'Area Economica', 'Patrimoniale': 'Area Patrimoniale',
                 'Finanziaria': 'Area Finanziaria'}
    for i, area in enumerate(AREE):
        cella = t.rows[0].cells[i]
        cella.width = Cm(5.6)
        _sfondo_cella(cella, RIGA_ALT)
        _testo(cella, etichette[area], dim=11, grassetto=True, colore=C_CORALLO,
               allineamento=WD_ALIGN_PARAGRAPH.CENTER, spazio_dopo=4, primo=True)
        classe = q[f'rating_{area}']
        colore = {'A': C_VERDE, 'B': C_NAVY, 'C': C_ROSSO}[classe]
        _testo(cella, '●', dim=30, colore=colore, allineamento=WD_ALIGN_PARAGRAPH.CENTER, spazio_dopo=2)
        _testo(cella, f'Classe {classe}', dim=12, grassetto=True, colore=colore,
               allineamento=WD_ALIGN_PARAGRAPH.CENTER, spazio_dopo=6)

    for i, area in enumerate(AREE):
        cella = t.rows[1].cells[i]
        sopra = [q['ind'][k]['nome'] for k in CHIAVI_AREA[area] if _meglio(q['ind'][k])]
        n = len(sopra)
        if n == 3:
            frase = 'Tutti e tre gli indicatori fanno meglio del settore.'
        elif n == 0:
            frase = 'Nessuno dei tre indicatori raggiunge il riferimento di settore.'
        else:
            frase = f"{n} indicatori su 3 fanno meglio del settore: {', '.join(sopra)}."
        _testo(cella, frase, dim=8, colore=C_GRIGIO,
               allineamento=WD_ALIGN_PARAGRAPH.CENTER, spazio_dopo=8, primo=True)


def _meglio(dati):
    v, m = dati['valore'], dati['mediana']
    if v is None or m is None:
        return False
    return (v < m) if dati['inverso'] else (v > m)


def _dati_sintesi(doc, q):
    _titolo_sezione(doc, 'Dati di sintesi')
    t = doc.add_table(rows=1, cols=7)
    _senza_bordi(t)
    intest = ['', '2021', '2022', '2023', '2024', 'Settore 2024', 'Var % 21/24']
    larghezze = [Cm(5.0), Cm(1.8), Cm(1.8), Cm(1.8), Cm(2.0), Cm(2.4), Cm(2.2)]
    for i, testo in enumerate(intest):
        cella = t.rows[0].cells[i]
        cella.width = larghezze[i]
        _testo(cella, testo, dim=8, grassetto=True, colore=C_NAVY,
               allineamento=None if i == 0 else WD_ALIGN_PARAGRAPH.RIGHT, spazio_dopo=3, primo=True)
        _linea_sotto(cella, NAVY, '8')

    for n, (chiave, _, nome, unita, inverso, _area) in enumerate(INDICATORI):
        d = q['ind'][chiave]
        riga = t.add_row()
        for i, larghezza in enumerate(larghezze):
            riga.cells[i].width = larghezza
            if n % 2 == 1:
                _sfondo_cella(riga.cells[i], RIGA_ALT)
        _testo(riga.cells[0], nome, dim=8.5, grassetto=True, spazio_dopo=3, primo=True)
        for j, anno in enumerate(ANNI_SERIE, start=1):
            v = d['az'].get(anno)
            _testo(riga.cells[j], 'n.d.' if v is None else f"{format_euro(v)}{unita}",
                   dim=8.5, grassetto=(anno == '2024'),
                   allineamento=WD_ALIGN_PARAGRAPH.RIGHT, spazio_dopo=3, primo=True)
        med = d['mediana']
        _testo(riga.cells[5], 'n.d.' if med is None else f"{format_euro(med)}{unita}",
               dim=8.5, colore=C_GRIGIO, allineamento=WD_ALIGN_PARAGRAPH.RIGHT, spazio_dopo=3, primo=True)
        if d['var'] is None:
            _testo(riga.cells[6], 'n.d.', dim=8.5, colore=C_GRIGIO,
                   allineamento=WD_ALIGN_PARAGRAPH.RIGHT, spazio_dopo=3, primo=True)
        else:
            buono = (d['var'] < 0) if inverso else (d['var'] > 0)
            _testo(riga.cells[6], f"{format_euro(d['var'])}%", dim=8.5, grassetto=True,
                   colore=C_VERDE if buono else C_ROSSO,
                   allineamento=WD_ALIGN_PARAGRAPH.RIGHT, spazio_dopo=3, primo=True)

    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(3)
    r = p.add_run('Verde e rosso indicano il verso favorevole o sfavorevole della variazione; '
                  'per il Gearing un calo è un miglioramento.')
    r.font.size = Pt(7.5)
    r.font.color.rgb = C_GRIGIO


def _sezione_area(doc, q, area):
    _titolo_sezione(doc, f'Area {area}')
    classe = q[f'rating_{area}']
    p = doc.add_paragraph()
    p.paragraph_format.space_after = Pt(8)
    r = p.add_run(f'Classe {classe}')
    r.font.size = Pt(12)
    r.bold = True
    r.font.color.rgb = {'A': C_VERDE, 'B': C_NAVY, 'C': C_ROSSO}[classe]

    for chiave in CHIAVI_AREA[area]:
        d = q['ind'][chiave]
        v, m, u = d['valore'], d['mediana'], d['unita']
        if v is None:
            continue
        # "meglio/peggio" e non "sopra/sotto": sul Gearing un valore più alto è
        # un risultato peggiore, e "sopra la mediana" si leggerebbe come un pregio
        verso = 'meglio' if _meglio(d) else 'peggio'
        frase = (f"{d['nome']}: {format_euro(v)}{u} nel 2024, {verso} della mediana di settore "
                 f"({format_euro(m)}{u}).") if m is not None else \
                f"{d['nome']}: {format_euro(v)}{u} nel 2024."
        v21 = d['az'].get('2021')
        if v21 is not None:
            direzione = 'in crescita' if v > v21 else ('in calo' if v < v21 else 'stabile')
            frase += f" Dal 2021 il valore è {direzione} ({format_euro(v21)}{u} → {format_euro(v)}{u})."
        par = doc.add_paragraph(style='List Bullet')
        par.paragraph_format.space_after = Pt(4)
        run = par.add_run(frase)
        run.font.size = Pt(9)
        run.font.color.rgb = C_NAVY


def _sintesi(doc, q):
    _titolo_sezione(doc, 'Posizionamento di sintesi')
    t = doc.add_table(rows=1, cols=2)
    _senza_bordi(t)
    sx, dx = t.rows[0].cells
    sx.width, dx.width = Cm(5.2), Cm(11.8)
    _sfondo_cella(sx, NAVY)
    _testo(sx, 'Rating Combinato', dim=9, colore=RGBColor(0xC8, 0xD0, 0xDC),
           allineamento=WD_ALIGN_PARAGRAPH.CENTER, spazio_dopo=2, primo=True)
    _testo(sx, q['rating'], dim=30, grassetto=True, colore=C_BIANCO,
           allineamento=WD_ALIGN_PARAGRAPH.CENTER, spazio_dopo=2)
    _testo(sx, 'Economica · Patrimoniale · Finanziaria', dim=7,
           colore=RGBColor(0xC8, 0xD0, 0xDC), allineamento=WD_ALIGN_PARAGRAPH.CENTER, spazio_dopo=8)

    voci = []
    if pd.notna(q['vdp']):
        voci.append(f"Valore della Produzione 2024: € {format_euro(q['vdp'] / 1000)} mln")
    if pd.notna(q['attivo']):
        voci.append(f"Attivo: € {format_euro(q['attivo'] / 1000)} mln")
    if pd.notna(q['dipendenti']):
        voci.append(f"Organico: {format_euro(q['dipendenti'], 0)} dipendenti")
    if q['regione']:
        voci.append(f"Sede {preposizione_regione(q['regione'])}")
    if q['rango']:
        voci.append(f"{posizione_ordinale(q['rango'])} posto del panel per Valore della Produzione")
    primo = True
    for voce in voci:
        _testo(dx, f'›  {voce}', dim=9, spazio_dopo=3, primo=primo)
        primo = False

    priorita = []
    if not _meglio(q['ind']['ebitda']):
        priorita.append('recupero dei margini operativi')
    if not _meglio(q['ind']['gearing']):
        priorita.append("contenimento dell'indebitamento")
    if (q['ind']['cr']['valore'] or 0) < 1 or (q['ind']['qr']['valore'] or 0) < 1:
        priorita.append('presidio della liquidità di breve periodo')
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(10)
    r = p.add_run('Dove guardare  ')
    r.bold = True
    r.font.size = Pt(10)
    r.font.color.rgb = C_CORALLO
    testo = ('; '.join(priorita) + '.') if priorita else \
        'Nessuna delle tre aree presenta scostamenti tali da richiedere un intervento prioritario.'
    r = p.add_run(testo[0].upper() + testo[1:])
    r.font.size = Pt(9.5)
    r.font.color.rgb = C_NAVY


def _note(doc, q, settore):
    _titolo_sezione(doc, 'Note metodologiche e glossario')
    p = doc.add_paragraph()
    r = p.add_run(
        f"Il panel comprende {format_euro(q['panel'], 0)} imprese del settore {settore}, estratte da ORBIS. "
        f"Il settore è descritto dalla mediana e non dalla media, perché la mediana non viene spostata dai "
        f"pochi valori estremi presenti nel campione. Le classi nascono dai terzili della distribuzione 2024: "
        f"A per la fascia superiore, B per quella intermedia, C per quella inferiore. Per il Gearing l'ordine "
        f"si inverte, perché un valore più contenuto segnala meno debito."
    )
    r.font.size = Pt(8.5)
    r.font.color.rgb = C_NAVY

    t = doc.add_table(rows=0, cols=2)
    _senza_bordi(t)
    for n, (nome, formula) in enumerate(GLOSSARIO):
        riga = t.add_row()
        riga.cells[0].width, riga.cells[1].width = Cm(5.4), Cm(11.6)
        if n % 2 == 1:
            _sfondo_cella(riga.cells[0], RIGA_ALT)
            _sfondo_cella(riga.cells[1], RIGA_ALT)
        _testo(riga.cells[0], nome, dim=8, grassetto=True, colore=C_CORALLO, spazio_dopo=2, primo=True)
        _testo(riga.cells[1], formula, dim=8, colore=C_NAVY, spazio_dopo=2, primo=True)


def genera_report_sintetico(df_orbis, azienda_target, settore_nace, chiave_target=None):
    """Costruisce il report breve impaginato e restituisce il .docx in memoria."""
    q = calcola_quadro(df_orbis, chiave_target, azienda_target)
    doc = Document()
    normale = doc.styles['Normal']
    normale.font.name = 'Calibri'
    normale.font.size = Pt(9)
    normale.font.color.rgb = C_NAVY
    normale.paragraph_format.space_after = Pt(4)

    prima = doc.sections[0]
    prima.top_margin = prima.bottom_margin = Cm(0)
    prima.left_margin = prima.right_margin = Cm(2.0)
    _copertina(doc, str(azienda_target), settore_nace)

    corpo = doc.add_section(WD_SECTION.NEW_PAGE)
    corpo.top_margin = Cm(2.4)
    corpo.bottom_margin = Cm(1.6)
    corpo.left_margin = corpo.right_margin = Cm(2.0)
    corpo.header_distance = Cm(0.8)
    corpo.header.is_linked_to_previous = False
    _banda_testata(corpo, str(azienda_target), settore_nace)

    _indice(doc)
    _big_picture(doc, q)
    _dati_sintesi(doc, q)
    for area in AREE:
        _sezione_area(doc, q, area)
    _sintesi(doc, q)
    _note(doc, q, settore_nace)

    # Word non accetta una tabella come ultimo elemento del corpo
    doc.add_paragraph()

    uscita = io.BytesIO()
    doc.save(uscita)
    uscita.seek(0)
    return uscita
