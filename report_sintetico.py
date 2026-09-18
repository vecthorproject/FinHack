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
    commento_indicatore,
    con_articolo,
    format_euro,
    inserisci_in_ordine,
    nome_regione_breve,
    posizione_ordinale,
    pulisci_nome_orbis,
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

# (chiave, colonna base, nome, unità, soglia unitaria, inverso, area)
INDICATORI = [
    ('ebitda',    'Margine EBITDA (*) %',                           'Margine EBITDA',              '%', False, False, 'Economica'),
    ('ebit',      'Margine EBIT (*) %',                             'Margine EBIT',                '%', False, False, 'Economica'),
    ('profitto',  'Margine di Profitto (*) %',                      'Margine di Profitto',         '%', False, False, 'Economica'),
    ('strut1',    'Indice di Struttura 1° livello (*)',             'Indice di Struttura 1° liv.', '',  True,  False, 'Patrimoniale'),
    ('strut2',    'Indice di Struttura 2° livello (*)',             'Indice di Struttura 2° liv.', '',  True,  False, 'Patrimoniale'),
    ('gearing',   'Gearing (*) %',                                  'Gearing',                     '%', False, True,  'Patrimoniale'),
    ('cr',        'Current Ratio (*)',                              'Current Ratio',               '',  True,  False, 'Finanziaria'),
    ('qr',        'Quick Ratio (*)',                                'Quick Ratio',                 '',  True,  False, 'Finanziaria'),
    ('rotazione', 'Indice di Rotazione del Capitale Investito (*)', 'Rotazione Cap. Investito',    '',  False, False, 'Finanziaria'),
]

# In tabella servono nomi corti, ma il commento discorsivo antepone l'articolo:
# "Il Rotazione Cap. Investito" non si legge.
NOMI_NARRATIVI = {
    'strut1':    'Indice di Struttura di 1° livello',
    'strut2':    'Indice di Struttura di 2° livello',
    'rotazione': 'Indice di Rotazione del Capitale Investito',
}

# Riga di lettura aggiunta in coda al commento quando la soglia è rispettata:
# sono le stesse del report esteso.
LETTURE = {
    'strut1': "Un valore superiore all'unità indica che il patrimonio netto copre integralmente "
              "le immobilizzazioni.",
    'strut2': "Il capitale permanente, dato da patrimonio netto e passività non correnti, copre "
              "quindi le immobilizzazioni.",
    'cr':     "Le attività correnti coprono le passività di pari scadenza.",
    'qr':     "La copertura delle passività correnti regge anche al netto delle rimanenze.",
}

# Grandezze dimensionali e voci di conto economico presenti nell'estrazione,
# lette per tutti e quattro gli esercizi.
VOCI_DIMENSIONE = [
    ('Totale valore della produzione migl EUR', 'Valore della Produzione'),
    ('Totale Attivo migl EUR',                  'Totale Attivo'),
    ('Utile/Perdita al netto delle imposte migl EUR', 'Risultato netto'),
]
VOCI_CONTO_ECONOMICO = [
    ('Totale valore della produzione migl EUR',       'Valore della Produzione',  False),
    ('Costo del venduto migl EUR',                    'Costo del venduto',        True),
    ('Oneri diversi di gestione migl EUR',            'Oneri diversi di gestione', True),
    ('Proventi/oneri finanziari migl EUR',            'Proventi/oneri finanziari', False),
    ('Totale imposte migl EUR',                       'Imposte',                  True),
    ('Utile/Perdita al netto delle imposte migl EUR', 'Risultato netto',          False),
]

# ORBIS restituisce le macroaree NUTS1 in inglese.
MACROAREE = {
    'ITC': 'Nord-Ovest', 'ITH': 'Nord-Est', 'ITI': 'Centro',
    'ITF': 'Sud', 'ITG': 'Isole',
}

A_COSA_SERVE = {
    'ebitda':    'Quanta parte del valore della produzione resta dopo i costi operativi, prima di '
                 'ammortamenti e svalutazioni.',
    'ebit':      'Il margine che resta a valle anche degli ammortamenti: è il reddito della sola '
                 'gestione caratteristica.',
    'profitto':  'Quanto del valore della produzione arriva fino al risultato ante imposte, dopo '
                 'oneri finanziari e gestione extra-operativa.',
    'strut1':    'Quanta parte delle immobilizzazioni è coperta dal solo patrimonio netto, che non '
                 'ha scadenza.',
    'strut2':    'La stessa copertura, contando anche i debiti a medio-lungo termine.',
    'gearing':   'Quanto pesa il debito finanziario rispetto al patrimonio netto: qui un valore '
                 'più basso è un risultato migliore.',
    'cr':        'Se le attività a breve bastano a coprire i debiti che scadono entro l\'anno.',
    'qr':        'La stessa verifica, togliendo le rimanenze: misura la copertura senza contare il '
                 'magazzino.',
    'rotazione': 'Quante volte il capitale investito si è tradotto in valore della produzione nel '
                 'corso dell\'anno.',
}

AREE = ['Economica', 'Patrimoniale', 'Finanziaria']
CHIAVI_AREA = {a: [k for k, _, _, _, _, _, ar in INDICATORI if ar == a] for a in AREE}

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
    run = par.add_run(pulisci_nome_orbis(testo))
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


def _lettera(somma):
    return 'A' if somma <= 4 else ('B' if somma <= 7 else 'C')


def _punti_panel(serie, t1, t2, inverso):
    """Punteggio 1/2/3 di tutto il panel su un indicatore, con la stessa regola
    usata per l'azienda target."""
    if pd.isna(t1) or pd.isna(t2):
        return pd.Series(3, index=serie.index)
    if inverso:
        punti = np.where(serie <= t1, 1, np.where(serie <= t2, 2, 3))
    else:
        punti = np.where(serie >= t2, 1, np.where(serie >= t1, 2, 3))
    return pd.Series(np.where(serie.isna(), 3, punti), index=serie.index)


def _elenco(voci):
    """"A, B e C" invece di "A e B e C"."""
    if not voci:
        return ''
    if len(voci) == 1:
        return voci[0]
    return ', '.join(voci[:-1]) + ' e ' + voci[-1]


def _perc(valore, preposizione=None, decimali=2):
    """Percentuale con l'articolo concordato: "dell'11,10%", "lo 0,05%"."""
    return f"{con_articolo(format_euro(valore, decimali), preposizione)}%"


def _macroarea(valore):
    sigla = str(valore).split(' - ')[0].strip()[:3]
    return MACROAREE.get(sigla, str(valore).split(' - ')[-1].strip() or 'N.D.')


def _forma_breve(valore):
    v = str(valore).upper()
    if 'S.P.A' in v or 'AZIONI' in v:
        return 'S.p.A.'
    if 'SEMPLIFICATA' in v:
        return 'S.r.l. semplificata'
    if 'S.R.L' in v or 'RESPONSABILITÀ LIMITATA' in v:
        return 'S.r.l.'
    if 'COOPERATIVA' in v or 'COOP' in v:
        return 'Cooperativa'
    return 'Altre forme'


def calcola_quadro(df, chiave_target, azienda_target):
    riga = riga_target(df, chiave_target, azienda_target)
    if riga.empty:
        raise ValueError("Azienda target non trovata nel campione.")
    riga = riga.iloc[0]

    col_reg = next((c for c in df.columns if 'nuts2' in str(c).lower()), None)
    col_macro = next((c for c in df.columns if 'nuts1' in str(c).lower()), None)
    regione_raw = str(riga.get(col_reg, '')) if col_reg else ''
    regione = regione_raw.split(' - ')[-1].strip()
    stessa_regione = df[df[col_reg] == riga[col_reg]] if col_reg and pd.notna(riga.get(col_reg)) else df.iloc[0:0]

    q = {'ind': {}, 'punti': {}, 'terzili': {}, 'rank': {}, 'rank_reg': {}}
    punti_panel = {}

    for giro, (chiave, base, nome, unita, soglia, inverso, area) in enumerate(INDICATORI):
        col = f"{base} 2024"
        az, sett = serie_indicatore(df, riga, base)
        serie = pd.to_numeric(df[col], errors='coerce') if col in df.columns else pd.Series(dtype=float)
        validi = serie.dropna()
        t1, t2 = (validi.quantile(1 / 3), validi.quantile(2 / 3)) if not validi.empty else (np.nan, np.nan)
        q['terzili'][chiave] = (t1, t2)
        punti = _classe(az.get('2024'), t1, t2, inverso)
        q['punti'][chiave] = punti
        punti_panel[chiave] = _punti_panel(serie, t1, t2, inverso)

        v21, v24 = az.get('2021'), az.get('2024')
        variazione = None
        if v21 not in (None, 0) and v24 is not None and abs(v21) > 1e-9:
            variazione = (v24 - v21) / abs(v21) * 100

        # Posizione in classifica: sugli indicatori inversi il primo posto va al
        # valore più basso.
        def posizione(sottoinsieme):
            s = pd.to_numeric(sottoinsieme[col], errors='coerce').dropna() if col in sottoinsieme.columns else pd.Series(dtype=float)
            if s.empty or v24 is None:
                return None, len(s)
            migliori = (s < v24).sum() if inverso else (s > v24).sum()
            return int(migliori) + 1, int(len(s))

        q['rank'][chiave] = posizione(df)
        q['rank_reg'][chiave] = posizione(stessa_regione)

        lettura = LETTURE.get(chiave)
        if chiave in ('cr', 'qr', 'strut1', 'strut2') and (v24 is None or v24 < 1):
            lettura = None
        commento = commento_indicatore(NOMI_NARRATIVI.get(chiave, nome), az, sett, unita=unita,
                                       soglia_unitaria=soglia, inverso=inverso,
                                       lettura=lettura, giro=giro)

        q['ind'][chiave] = {
            'nome': nome, 'unita': unita, 'inverso': inverso, 'area': area,
            'az': az, 'set': sett, 'valore': v24, 'mediana': sett.get('2024'),
            'classe': 'ABC'[punti - 1], 'var': variazione,
            'commento': commento.lstrip('• ').strip(),
        }

    for area in AREE:
        q[f'rating_{area}'] = _lettera(sum(q['punti'][k] for k in CHIAVI_AREA[area]))
    q['rating'] = q['rating_Economica'] + q['rating_Patrimoniale'] + q['rating_Finanziaria']

    # --- come si distribuisce il panel sulle stesse regole -------------------
    lettere_panel = {}
    for area in AREE:
        somma = sum(punti_panel[k] for k in CHIAVI_AREA[area])
        lettere_panel[area] = somma.map(_lettera)
    q['distribuzione'] = {area: lettere_panel[area].value_counts() for area in AREE}
    combinato = lettere_panel['Economica'] + lettere_panel['Patrimoniale'] + lettere_panel['Finanziaria']
    q['comb_conteggi'] = combinato.value_counts()
    q['comb_target'] = int(q['comb_conteggi'].get(q['rating'], 0))
    q['comb_piu_diffuso'] = (q['comb_conteggi'].index[0], int(q['comb_conteggi'].iloc[0])) \
        if not q['comb_conteggi'].empty else (None, 0)

    # --- anagrafica e peso sul settore ---------------------------------------
    q['regione'] = nome_regione_breve(regione)
    q['macroarea'] = _macroarea(riga.get(col_macro, '')) if col_macro else ''
    q['forma'] = _forma_breve(riga.get('Forma giuridica nazionale', ''))
    q['piva'] = str(riga.get('Codice fiscale/Partita IVA', '') or '').strip()
    q['panel'] = len(df)
    q['imprese_regione'] = len(stessa_regione)
    q['vdp'] = pd.to_numeric(riga.get('Totale valore della produzione migl EUR 2024'), errors='coerce')
    q['attivo'] = pd.to_numeric(riga.get('Totale Attivo migl EUR 2024'), errors='coerce')
    q['dipendenti'] = pd.to_numeric(riga.get('Numero dipendenti 2024'), errors='coerce')
    q['utile'] = pd.to_numeric(riga.get('Utile/Perdita al netto delle imposte migl EUR 2024'), errors='coerce')
    ricavi_mln = (q['vdp'] / 1000) if pd.notna(q['vdp']) else 0
    q['classe_dim'] = 'Grande impresa' if ricavi_mln > 50 else ('Media impresa' if ricavi_mln > 10 else 'Piccola impresa')

    col_vdp = 'Totale valore della produzione migl EUR 2024'
    vdp_panel = pd.to_numeric(df[col_vdp], errors='coerce') if col_vdp in df.columns else pd.Series(dtype=float)
    q['vdp_panel'] = vdp_panel.sum()
    q['peso_vdp'] = (q['vdp'] / q['vdp_panel'] * 100) if q['vdp_panel'] else None
    if not vdp_panel.empty and pd.notna(q['vdp']):
        rango = vdp_panel.rank(ascending=False, method='min')
        q['rango'] = int(rango.loc[riga.name]) if riga.name in rango.index else None
    else:
        q['rango'] = None

    # --- voci di bilancio nel quadriennio ------------------------------------
    def serie_voce(base):
        valori = {}
        for anno in ANNI_SERIE:
            col = f"{base} {anno}"
            v = pd.to_numeric(riga.get(col), errors='coerce') if col in df.columns else np.nan
            valori[anno] = None if pd.isna(v) else float(v)
        return valori

    q['dimensione'] = [(nome, serie_voce(base)) for base, nome in VOCI_DIMENSIONE]
    q['conto_economico'] = [(nome, serie_voce(base), costo)
                            for base, nome, costo in VOCI_CONTO_ECONOMICO]

    # --- il settore in numeri -------------------------------------------------
    forme = df['Forma giuridica nazionale'].map(_forma_breve) if 'Forma giuridica nazionale' in df.columns else pd.Series(dtype=object)
    q['forme'] = forme.value_counts()
    if col_macro:
        macro = df[col_macro].map(_macroarea)
        raggr = pd.DataFrame({
            'macro': macro,
            'vdp': vdp_panel,
            'dip': pd.to_numeric(df.get('Numero dipendenti 2024'), errors='coerce'),
        })
        gruppi = raggr.groupby('macro', dropna=False)
        q['macro_tabella'] = pd.DataFrame({
            'imprese': gruppi.size(),
            'vdp': gruppi['vdp'].sum(),
            'dipendenti': gruppi['dip'].sum(),
        }).sort_values('imprese', ascending=False)
    else:
        q['macro_tabella'] = pd.DataFrame()
    return q


# ------------------------------------------------------------ impaginato ---

def _copertina(doc, azienda, settore, q):
    t = doc.add_table(rows=1, cols=1)
    _senza_bordi(t)
    cella = t.rows[0].cells[0]
    _sfondo_cella(cella, NAVY)
    _altezza_riga(t.rows[0], 19.5)
    cella.width = Cm(17)

    _testo(cella, '', dim=24, primo=True)
    _testo(cella, 'Report di Posizionamento', dim=30, grassetto=True, colore=C_BIANCO, spazio_dopo=2)
    p = cella.add_paragraph()
    p.paragraph_format.space_after = Pt(18)
    r = p.add_run('BASE')
    r.font.size = Pt(26)
    r.bold = True
    r.font.color.rgb = C_CORALLO

    _testo(cella, azienda, dim=17, grassetto=True, colore=C_BIANCO, spazio_dopo=4)
    _testo(cella, f'Settore {settore}', dim=11, colore=C_CORALLO, spazio_dopo=14)
    _testo(cella, f"Rating Combinato  {q['rating']}", dim=13, grassetto=True,
           colore=C_BIANCO, spazio_dopo=14)
    _testo(cella, 'Bilanci', dim=15, colore=C_BIANCO, spazio_dopo=2)
    _testo(cella, '2024 · 2023 · 2022 · 2021', dim=13, colore=C_BIANCO, spazio_dopo=2)
    _testo(cella, f"Confronto su {format_euro(q['panel'], 0)} imprese del settore", dim=9,
           colore=RGBColor(0xC8, 0xD0, 0xDC), spazio_dopo=0)


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


def _titolo_sezione(doc, testo, nuova_pagina=False):
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(14)
    p.paragraph_format.space_after = Pt(6)
    p.paragraph_format.keep_with_next = True
    if nuova_pagina:
        inserisci_in_ordine(p._p.get_or_add_pPr(), OxmlElement('w:pageBreakBefore'))
    r = p.add_run('▎ ')
    r.font.size = Pt(14)
    r.font.color.rgb = C_CORALLO
    r = p.add_run(testo)
    r.font.size = Pt(14)
    r.bold = True
    r.font.color.rgb = C_NAVY
    return p


def _sottotitolo(doc, testo):
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(10)
    p.paragraph_format.space_after = Pt(4)
    p.paragraph_format.keep_with_next = True
    r = p.add_run(testo.upper())
    r.font.size = Pt(8.5)
    r.bold = True
    r.font.color.rgb = C_CORALLO
    return p


def _capoverso(doc, testo, dim=9, colore=None, spazio_dopo=6):
    p = doc.add_paragraph()
    p.paragraph_format.space_after = Pt(spazio_dopo)
    p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    r = p.add_run(pulisci_nome_orbis(testo))
    r.font.size = Pt(dim)
    r.font.color.rgb = colore if colore is not None else C_NAVY
    return p


def _indice(doc):
    _titolo_sezione(doc, 'Indice')
    voci = [
        ("Profilo dell'impresa", 'Anagrafica, dimensione e peso sul settore'),
        ("L'impresa nel quadriennio", 'Produzione, attivo e risultato dal 2021 al 2024'),
        ('Big Picture', 'I semafori delle tre aree'),
        ('Dati di sintesi', 'I nove indicatori dal 2021 al 2024'),
        ('Struttura del conto economico', 'Dalle voci di bilancio ai margini'),
        ('Il settore in numeri', 'Composizione del panel e distribuzione territoriale'),
        ('Come nascono le classi', 'Terzili, soglie e regola di assegnazione'),
        ('Come si distribuisce il panel', 'Quante imprese in ciascuna classe'),
        ('Area Economica', 'Margini a confronto con il settore'),
        ('Area Patrimoniale', 'Copertura degli investimenti e leva finanziaria'),
        ('Area Finanziaria', 'Liquidità di breve periodo e rotazione del capitale'),
        ('Classifiche', 'Posizione nazionale e regionale sui nove indicatori'),
        ('Posizionamento di sintesi', 'Rating Combinato e dove guardare'),
        ('Note metodologiche e glossario', 'Campione, metodo, formule'),
    ]
    t = doc.add_table(rows=0, cols=1)
    _senza_bordi(t)
    for titolo, sotto in voci:
        riga = t.add_row()
        riga.cells[0].width = Cm(17)
        _testo(riga.cells[0], titolo, dim=10.5, grassetto=True, spazio_dopo=0, primo=True)
        _testo(riga.cells[0], sotto, dim=8, colore=C_GRIGIO, spazio_dopo=5)
        _linea_sotto(riga.cells[0], 'D9D9D9', '2')


def _profilo(doc, q, azienda, settore):
    _titolo_sezione(doc, "Profilo dell'impresa", nuova_pagina=True)

    kpi = [
        ('Valore della Produzione 2024', f"€ {format_euro(q['vdp'] / 1000)} mln" if pd.notna(q['vdp']) else 'n.d.'),
        ('Totale Attivo 2024', f"€ {format_euro(q['attivo'] / 1000)} mln" if pd.notna(q['attivo']) else 'n.d.'),
        ('Dipendenti 2024', format_euro(q['dipendenti'], 0) if pd.notna(q['dipendenti']) else 'n.d.'),
        ('Rating Combinato', q['rating']),
    ]
    t = doc.add_table(rows=2, cols=4)
    _senza_bordi(t)
    for i, (etichetta, valore) in enumerate(kpi):
        alto, basso = t.rows[0].cells[i], t.rows[1].cells[i]
        alto.width = basso.width = Cm(4.25)
        _sfondo_cella(alto, NAVY if i == 3 else RIGA_ALT)
        _sfondo_cella(basso, NAVY if i == 3 else RIGA_ALT)
        _testo(alto, etichetta.upper(), dim=7, grassetto=True,
               colore=RGBColor(0xC8, 0xD0, 0xDC) if i == 3 else C_GRIGIO,
               allineamento=WD_ALIGN_PARAGRAPH.CENTER, spazio_dopo=2, primo=True)
        _testo(basso, valore, dim=15, grassetto=True, colore=C_BIANCO if i == 3 else C_NAVY,
               allineamento=WD_ALIGN_PARAGRAPH.CENTER, spazio_dopo=6, primo=True)

    _sottotitolo(doc, 'Anagrafica')
    voci = [
        ('Ragione sociale', azienda),
        ('Partita IVA', q['piva'] or 'n.d.'),
        ('Settore NACE Rev.2', settore),
        ('Forma giuridica', q['forma']),
        ('Sede', f"{q['regione']} ({q['macroarea']})" if q['regione'] else 'n.d.'),
        ('Classe dimensionale', q['classe_dim']),
    ]
    if pd.notna(q['utile']):
        voci.append(('Risultato netto 2024', f"€ {format_euro(q['utile'] / 1000)} mln"))
    t = doc.add_table(rows=0, cols=2)
    _senza_bordi(t)
    for n, (etichetta, valore) in enumerate(voci):
        riga = t.add_row()
        riga.cells[0].width, riga.cells[1].width = Cm(5.4), Cm(11.6)
        if n % 2 == 1:
            _sfondo_cella(riga.cells[0], RIGA_ALT)
            _sfondo_cella(riga.cells[1], RIGA_ALT)
        _testo(riga.cells[0], etichetta, dim=8.5, grassetto=True, colore=C_GRIGIO, spazio_dopo=3, primo=True)
        _testo(riga.cells[1], str(valore), dim=8.5, spazio_dopo=3, primo=True)

    _sottotitolo(doc, 'Peso sul settore')
    frasi = []
    if q['peso_vdp'] is not None:
        frase = (f"Il Valore della Produzione dell'impresa vale {_perc(q['peso_vdp'])} del totale "
                 f"generato dalle {format_euro(q['panel'], 0)} imprese del panel")
        if q['rango']:
            frase += f", che la colloca al {posizione_ordinale(q['rango'])} posto per dimensione"
        frasi.append(frase + '.')
    if q['imprese_regione'] and q['regione']:
        sede = preposizione_regione(q['regione'])
        frasi.append(f"{sede[0].upper()}{sede[1:]} operano "
                     f"{format_euro(q['imprese_regione'], 0)} imprese del settore, pari "
                     f"{_perc(q['imprese_regione'] / q['panel'] * 100, 'a')} del campione.")
    if frasi:
        _capoverso(doc, ' '.join(frasi))


def _mln(valore):
    return 'n.d.' if valore is None else f"€ {format_euro(valore / 1000)}"


def _dimensione(doc, q):
    _titolo_sezione(doc, "L'impresa nel quadriennio", nuova_pagina=True)
    _capoverso(doc, "Prima degli indicatori, le grandezze da cui sono costruiti: quanto ha prodotto "
                    "l'impresa, quanto capitale ha impiegato e che risultato ne ha ricavato, "
                    "esercizio per esercizio. Importi in milioni di euro.")

    colonne = ['Voce'] + list(ANNI_SERIE) + ['Var % 21/24']
    larghezze = [Cm(5.4), Cm(2.3), Cm(2.3), Cm(2.3), Cm(2.4), Cm(2.3)]
    t = doc.add_table(rows=1, cols=len(colonne))
    _senza_bordi(t)
    for i, testo in enumerate(colonne):
        cella = t.rows[0].cells[i]
        cella.width = larghezze[i]
        _testo(cella, testo, dim=8, grassetto=True,
               allineamento=None if i == 0 else WD_ALIGN_PARAGRAPH.RIGHT, spazio_dopo=3, primo=True)
        _linea_sotto(cella, NAVY, '8')
    for n, (nome, serie) in enumerate(q['dimensione']):
        riga = t.add_row()
        for i, larghezza in enumerate(larghezze):
            riga.cells[i].width = larghezza
            if n % 2 == 1:
                _sfondo_cella(riga.cells[i], RIGA_ALT)
        _testo(riga.cells[0], nome, dim=8.5, grassetto=True, spazio_dopo=3, primo=True)
        for j, anno in enumerate(ANNI_SERIE, start=1):
            _testo(riga.cells[j], _mln(serie.get(anno)), dim=8.5, grassetto=(anno == '2024'),
                   allineamento=WD_ALIGN_PARAGRAPH.RIGHT, spazio_dopo=3, primo=True)
        v0, vN = serie.get('2021'), serie.get('2024')
        if v0 in (None, 0) or vN is None or abs(v0) < 1e-9:
            _testo(riga.cells[5], 'n.d.', dim=8.5, colore=C_GRIGIO,
                   allineamento=WD_ALIGN_PARAGRAPH.RIGHT, spazio_dopo=3, primo=True)
        else:
            var = (vN - v0) / abs(v0) * 100
            _testo(riga.cells[5], f"{format_euro(var)}%", dim=8.5, grassetto=True,
                   colore=C_VERDE if var >= 0 else C_ROSSO,
                   allineamento=WD_ALIGN_PARAGRAPH.RIGHT, spazio_dopo=3, primo=True)

    _sottotitolo(doc, 'Che cosa se ne ricava')
    vdp = dict(q['dimensione'])['Valore della Produzione']
    attivo = dict(q['dimensione'])['Totale Attivo']
    frasi = []
    if vdp.get('2021') and vdp.get('2024'):
        direzione = 'cresce' if vdp['2024'] > vdp['2021'] else 'arretra'
        var = abs((vdp['2024'] - vdp['2021']) / abs(vdp['2021']) * 100)
        frasi.append(f"Nel quadriennio il Valore della Produzione {direzione} "
                     f"{_perc(var, 'di')}, da {_mln(vdp['2021'])} mln a {_mln(vdp['2024'])} mln.")
    if attivo.get('2024') and vdp.get('2024'):
        frasi.append(f"A fronte di un attivo di {_mln(attivo['2024'])} mln, il rapporto fra "
                     f"produzione e capitale impiegato è quanto misura l'Indice di Rotazione "
                     f"del Capitale Investito.")
    if pd.notna(q['dipendenti']):
        frasi.append(f"L'organico 2024 conta {format_euro(q['dipendenti'], 0)} dipendenti.")
    if frasi:
        _capoverso(doc, ' '.join(frasi), dim=8.5)


def _conto_economico(doc, q):
    _titolo_sezione(doc, 'Struttura del conto economico', nuova_pagina=True)
    _capoverso(doc, "Le voci che spiegano il percorso dal Valore della Produzione al risultato "
                    "netto. L'ultima colonna mostra quanto pesa ogni voce sul Valore della "
                    "Produzione 2024: è la lettura che sta dietro ai margini EBITDA, EBIT e di "
                    "Profitto. Importi in milioni di euro.")

    colonne = ['Voce'] + list(ANNI_SERIE) + ['% su VdP 2024']
    larghezze = [Cm(5.0), Cm(2.2), Cm(2.2), Cm(2.2), Cm(2.4), Cm(3.0)]
    t = doc.add_table(rows=1, cols=len(colonne))
    _senza_bordi(t)
    for i, testo in enumerate(colonne):
        cella = t.rows[0].cells[i]
        cella.width = larghezze[i]
        _testo(cella, testo, dim=8, grassetto=True,
               allineamento=None if i == 0 else WD_ALIGN_PARAGRAPH.RIGHT, spazio_dopo=3, primo=True)
        _linea_sotto(cella, NAVY, '8')

    vdp24 = dict((n, s) for n, s, _ in q['conto_economico'])['Valore della Produzione'].get('2024')
    for n, (nome, serie, costo) in enumerate(q['conto_economico']):
        riga = t.add_row()
        for i, larghezza in enumerate(larghezze):
            riga.cells[i].width = larghezza
            if n % 2 == 1:
                _sfondo_cella(riga.cells[i], RIGA_ALT)
        _testo(riga.cells[0], nome, dim=8.5, grassetto=True, spazio_dopo=3, primo=True)
        for j, anno in enumerate(ANNI_SERIE, start=1):
            _testo(riga.cells[j], _mln(serie.get(anno)), dim=8.5, grassetto=(anno == '2024'),
                   allineamento=WD_ALIGN_PARAGRAPH.RIGHT, spazio_dopo=3, primo=True)
        v24 = serie.get('2024')
        if vdp24 and v24 is not None and abs(vdp24) > 1e-9:
            quota = v24 / abs(vdp24) * 100
            _testo(riga.cells[5], f"{format_euro(quota)}%", dim=8.5, colore=C_GRIGIO,
                   allineamento=WD_ALIGN_PARAGRAPH.RIGHT, spazio_dopo=3, primo=True)
        else:
            _testo(riga.cells[5], 'n.d.', dim=8.5, colore=C_GRIGIO,
                   allineamento=WD_ALIGN_PARAGRAPH.RIGHT, spazio_dopo=3, primo=True)

    _testo(doc, "I costi sono esposti con il segno dell'estrazione: la percentuale indica quanta "
                "parte del Valore della Produzione assorbono.", dim=7.5, colore=C_GRIGIO, spazio_dopo=0)

    _sottotitolo(doc, 'Il legame con i margini')
    _capoverso(doc,
               f"Il Margine EBITDA ({format_euro(q['ind']['ebitda']['valore'])}%) misura quanto "
               f"resta del Valore della Produzione dopo i costi operativi; il Margine EBIT "
               f"({format_euro(q['ind']['ebit']['valore'])}%) toglie anche gli ammortamenti; il "
               f"Margine di Profitto ({format_euro(q['ind']['profitto']['valore'])}%) arriva al "
               f"risultato ante imposte, dopo gli oneri finanziari e la gestione extra-operativa. "
               f"La distanza fra i tre valori dice dove si consuma la redditività.", dim=8.5)


def _big_picture(doc, q):
    _titolo_sezione(doc, 'Big Picture', nuova_pagina=True)
    _capoverso(doc, 'Ogni area riceve una classe A, B o C a partire dai suoi tre indicatori: '
                    'il semaforo riassume il risultato del 2024 rispetto al settore.', dim=8.5,
               colore=C_GRIGIO, spazio_dopo=8)
    t = doc.add_table(rows=2, cols=3)
    _senza_bordi(t)
    t.alignment = WD_TABLE_ALIGNMENT.CENTER
    for i, area in enumerate(AREE):
        cella = t.rows[0].cells[i]
        cella.width = Cm(5.6)
        _sfondo_cella(cella, RIGA_ALT)
        _testo(cella, f'Area {area}', dim=11, grassetto=True, colore=C_CORALLO,
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

    _sottotitolo(doc, 'Che cosa dice il Rating Combinato')
    _capoverso(doc, _frase_rating(q))


def _frase_rating(q):
    letture = {
        'Economica': ('la redditività si colloca nella fascia migliore del settore',
                      'la redditività resta in linea con il grosso del comparto',
                      'la redditività si colloca sotto i riferimenti di settore'),
        'Patrimoniale': ('la struttura delle fonti è fra le più solide del panel',
                         'la struttura delle fonti è in linea con il comparto',
                         'la struttura delle fonti mostra una dipendenza dal debito superiore alla media'),
        'Finanziaria': ('la liquidità di breve periodo è fra le più ampie del settore',
                        'la liquidità di breve periodo è in linea con il comparto',
                        'la liquidità di breve periodo resta sotto i riferimenti di settore'),
    }
    pezzi = [letture[a]['ABC'.index(q[f'rating_{a}'])] for a in AREE]
    testo = (f"Il Rating Combinato “{q['rating']}” va letto una lettera alla volta, "
             f"nell'ordine Economica, Patrimoniale e Finanziaria: {pezzi[0]}, {pezzi[1]} e "
             f"{pezzi[2]}.")
    if q['comb_target']:
        testo += (f" La stessa combinazione ricorre in {format_euro(q['comb_target'], 0)} imprese "
                  f"del panel, pari {_perc(q['comb_target'] / q['panel'] * 100, 'a')} del campione.")
    return testo


def _meglio(dati):
    v, m = dati['valore'], dati['mediana']
    if v is None or m is None:
        return False
    return (v < m) if dati['inverso'] else (v > m)


def _dati_sintesi(doc, q):
    _titolo_sezione(doc, 'Dati di sintesi', nuova_pagina=True)
    t = doc.add_table(rows=1, cols=8)
    _senza_bordi(t)
    intest = ['', '2021', '2022', '2023', '2024', 'Settore 2024', 'Var % 21/24', 'Classe']
    larghezze = [Cm(4.6), Cm(1.7), Cm(1.7), Cm(1.7), Cm(1.9), Cm(2.3), Cm(2.0), Cm(1.1)]
    for i, testo in enumerate(intest):
        cella = t.rows[0].cells[i]
        cella.width = larghezze[i]
        _testo(cella, testo, dim=8, grassetto=True, colore=C_NAVY,
               allineamento=None if i == 0 else WD_ALIGN_PARAGRAPH.RIGHT, spazio_dopo=3, primo=True)
        _linea_sotto(cella, NAVY, '8')

    for n, (chiave, _base, nome, unita, _soglia, inverso, _area) in enumerate(INDICATORI):
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
        _testo(riga.cells[7], d['classe'], dim=9, grassetto=True,
               colore={'A': C_VERDE, 'B': C_NAVY, 'C': C_ROSSO}[d['classe']],
               allineamento=WD_ALIGN_PARAGRAPH.CENTER, spazio_dopo=3, primo=True)

    _testo(doc, 'Verde e rosso indicano il verso favorevole o sfavorevole della variazione; '
                'per il Gearing un calo è un miglioramento.', dim=7.5, colore=C_GRIGIO, spazio_dopo=0)


def _settore_in_numeri(doc, q, settore):
    _titolo_sezione(doc, 'Il settore in numeri', nuova_pagina=True)
    _capoverso(doc, f"Il confronto si basa su {format_euro(q['panel'], 0)} imprese italiane del "
                    f"settore {settore} con bilanci disponibili per tutti e quattro gli esercizi. "
                    f"Nel 2024 il panel esprime un Valore della Produzione aggregato di "
                    f"€ {format_euro(q['vdp_panel'] / 1000)} mln.")

    _sottotitolo(doc, 'Forma giuridica')
    t = doc.add_table(rows=1, cols=3)
    _senza_bordi(t)
    for i, testo in enumerate(['Forma', 'Imprese', 'Quota']):
        cella = t.rows[0].cells[i]
        cella.width = [Cm(9.0), Cm(4.0), Cm(4.0)][i]
        _testo(cella, testo, dim=8, grassetto=True,
               allineamento=None if i == 0 else WD_ALIGN_PARAGRAPH.RIGHT, spazio_dopo=3, primo=True)
        _linea_sotto(cella, NAVY, '8')
    for n, (forma, conteggio) in enumerate(q['forme'].items()):
        riga = t.add_row()
        for i, larghezza in enumerate([Cm(9.0), Cm(4.0), Cm(4.0)]):
            riga.cells[i].width = larghezza
            if n % 2 == 1:
                _sfondo_cella(riga.cells[i], RIGA_ALT)
        _testo(riga.cells[0], forma, dim=8.5, spazio_dopo=3, primo=True)
        _testo(riga.cells[1], format_euro(conteggio, 0), dim=8.5,
               allineamento=WD_ALIGN_PARAGRAPH.RIGHT, spazio_dopo=3, primo=True)
        _testo(riga.cells[2], f"{format_euro(conteggio / q['panel'] * 100)}%", dim=8.5,
               colore=C_GRIGIO, allineamento=WD_ALIGN_PARAGRAPH.RIGHT, spazio_dopo=3, primo=True)

    if not q['macro_tabella'].empty:
        _sottotitolo(doc, 'Distribuzione territoriale')
        colonne = ['Macroarea', 'Imprese', 'Quota', 'Valore Produzione', 'Dipendenti']
        larghezze = [Cm(4.4), Cm(2.6), Cm(2.4), Cm(4.4), Cm(3.2)]
        t = doc.add_table(rows=1, cols=len(colonne))
        _senza_bordi(t)
        for i, testo in enumerate(colonne):
            cella = t.rows[0].cells[i]
            cella.width = larghezze[i]
            _testo(cella, testo, dim=8, grassetto=True,
                   allineamento=None if i == 0 else WD_ALIGN_PARAGRAPH.RIGHT, spazio_dopo=3, primo=True)
            _linea_sotto(cella, NAVY, '8')
        for n, (macro, dati) in enumerate(q['macro_tabella'].iterrows()):
            riga = t.add_row()
            evidenzia = (macro == q['macroarea'])
            for i, larghezza in enumerate(larghezze):
                riga.cells[i].width = larghezza
                if evidenzia:
                    _sfondo_cella(riga.cells[i], 'E4ECF5')
                elif n % 2 == 1:
                    _sfondo_cella(riga.cells[i], RIGA_ALT)
            _testo(riga.cells[0], str(macro), dim=8.5, grassetto=evidenzia, spazio_dopo=3, primo=True)
            _testo(riga.cells[1], format_euro(dati['imprese'], 0), dim=8.5,
                   allineamento=WD_ALIGN_PARAGRAPH.RIGHT, spazio_dopo=3, primo=True)
            _testo(riga.cells[2], f"{format_euro(dati['imprese'] / q['panel'] * 100)}%", dim=8.5,
                   colore=C_GRIGIO, allineamento=WD_ALIGN_PARAGRAPH.RIGHT, spazio_dopo=3, primo=True)
            _testo(riga.cells[3], f"€ {format_euro(dati['vdp'] / 1000)} mln", dim=8.5,
                   allineamento=WD_ALIGN_PARAGRAPH.RIGHT, spazio_dopo=3, primo=True)
            _testo(riga.cells[4], format_euro(dati['dipendenti'], 0), dim=8.5,
                   colore=C_GRIGIO, allineamento=WD_ALIGN_PARAGRAPH.RIGHT, spazio_dopo=3, primo=True)
        _testo(doc, 'La riga evidenziata è la macroarea di sede dell\'impresa.', dim=7.5,
               colore=C_GRIGIO, spazio_dopo=0)


def _come_nascono_le_classi(doc, q):
    _titolo_sezione(doc, 'Come nascono le classi', nuova_pagina=True)
    _capoverso(doc, "Per ogni indicatore la distribuzione 2024 del panel viene divisa in tre parti "
                    "uguali. Chi sta nel terzo migliore prende 1 punto, chi sta in quello di mezzo "
                    "2, chi sta nel terzo peggiore 3. I tre punteggi di un'area si sommano: fino a "
                    "4 punti la classe è A, fino a 7 è B, oltre è C. Per il Gearing l'ordine si "
                    "inverte, perché un debito più contenuto è un risultato migliore.")
    _sottotitolo(doc, 'Le soglie del settore e la posizione dell\'impresa')
    colonne = ['Indicatore', '1° terzile', '2° terzile', 'Valore 2024', 'Punti', 'Classe']
    larghezze = [Cm(5.4), Cm(2.5), Cm(2.5), Cm(2.6), Cm(1.8), Cm(2.2)]
    t = doc.add_table(rows=1, cols=len(colonne))
    _senza_bordi(t)
    for i, testo in enumerate(colonne):
        cella = t.rows[0].cells[i]
        cella.width = larghezze[i]
        _testo(cella, testo, dim=8, grassetto=True,
               allineamento=None if i == 0 else WD_ALIGN_PARAGRAPH.RIGHT, spazio_dopo=3, primo=True)
        _linea_sotto(cella, NAVY, '8')
    for n, (chiave, _base, nome, unita, _soglia, _inverso, _area) in enumerate(INDICATORI):
        d = q['ind'][chiave]
        t1, t2 = q['terzili'][chiave]
        riga = t.add_row()
        for i, larghezza in enumerate(larghezze):
            riga.cells[i].width = larghezza
            if n % 2 == 1:
                _sfondo_cella(riga.cells[i], RIGA_ALT)
        _testo(riga.cells[0], nome, dim=8.5, grassetto=True, spazio_dopo=3, primo=True)
        for i, valore in ((1, t1), (2, t2)):
            _testo(riga.cells[i], 'n.d.' if pd.isna(valore) else f"{format_euro(valore)}{unita}",
                   dim=8.5, colore=C_GRIGIO, allineamento=WD_ALIGN_PARAGRAPH.RIGHT,
                   spazio_dopo=3, primo=True)
        _testo(riga.cells[3], 'n.d.' if d['valore'] is None else f"{format_euro(d['valore'])}{unita}",
               dim=8.5, grassetto=True, allineamento=WD_ALIGN_PARAGRAPH.RIGHT, spazio_dopo=3, primo=True)
        _testo(riga.cells[4], str(q['punti'][chiave]), dim=8.5,
               allineamento=WD_ALIGN_PARAGRAPH.RIGHT, spazio_dopo=3, primo=True)
        _testo(riga.cells[5], d['classe'], dim=9, grassetto=True,
               colore={'A': C_VERDE, 'B': C_NAVY, 'C': C_ROSSO}[d['classe']],
               allineamento=WD_ALIGN_PARAGRAPH.CENTER, spazio_dopo=3, primo=True)

    _sottotitolo(doc, 'Dai punti alla lettera')
    t = doc.add_table(rows=1, cols=3)
    _senza_bordi(t)
    for i, area in enumerate(AREE):
        cella = t.rows[0].cells[i]
        cella.width = Cm(5.6)
        _sfondo_cella(cella, RIGA_ALT)
        somma = sum(q['punti'][k] for k in CHIAVI_AREA[area])
        conti = ' + '.join(str(q['punti'][k]) for k in CHIAVI_AREA[area])
        _testo(cella, f'Area {area}', dim=9, grassetto=True, colore=C_CORALLO,
               allineamento=WD_ALIGN_PARAGRAPH.CENTER, spazio_dopo=2, primo=True)
        _testo(cella, f"{conti} = {somma} punti", dim=8.5, colore=C_GRIGIO,
               allineamento=WD_ALIGN_PARAGRAPH.CENTER, spazio_dopo=2)
        _testo(cella, f"Classe {q[f'rating_{area}']}", dim=11, grassetto=True,
               colore={'A': C_VERDE, 'B': C_NAVY, 'C': C_ROSSO}[q[f'rating_{area}']],
               allineamento=WD_ALIGN_PARAGRAPH.CENTER, spazio_dopo=6)


def _distribuzione_panel(doc, q):
    _titolo_sezione(doc, 'Come si distribuisce il panel', nuova_pagina=True)
    _capoverso(doc, 'Le stesse regole applicate a tutte le imprese del campione mostrano quanto è '
                    'affollata ciascuna classe, e quindi quanto è comune il risultato ottenuto '
                    "dall'impresa.")
    colonne = ['Benchmark', 'Classe A', 'Classe B', 'Classe C', 'Classe dell\'impresa']
    larghezze = [Cm(4.2), Cm(3.2), Cm(3.2), Cm(3.2), Cm(3.2)]
    t = doc.add_table(rows=1, cols=len(colonne))
    _senza_bordi(t)
    for i, testo in enumerate(colonne):
        cella = t.rows[0].cells[i]
        cella.width = larghezze[i]
        _testo(cella, testo, dim=8, grassetto=True,
               allineamento=None if i == 0 else WD_ALIGN_PARAGRAPH.CENTER, spazio_dopo=3, primo=True)
        _linea_sotto(cella, NAVY, '8')
    for n, area in enumerate(AREE):
        conteggi = q['distribuzione'][area]
        riga = t.add_row()
        for i, larghezza in enumerate(larghezze):
            riga.cells[i].width = larghezza
            if n % 2 == 1:
                _sfondo_cella(riga.cells[i], RIGA_ALT)
        _testo(riga.cells[0], area, dim=8.5, grassetto=True, spazio_dopo=3, primo=True)
        for i, lettera in enumerate('ABC', start=1):
            n_imprese = int(conteggi.get(lettera, 0))
            _testo(riga.cells[i], f"{format_euro(n_imprese, 0)}  ({format_euro(n_imprese / q['panel'] * 100)}%)",
                   dim=8.5, colore=C_GRIGIO, allineamento=WD_ALIGN_PARAGRAPH.CENTER,
                   spazio_dopo=3, primo=True)
        classe = q[f'rating_{area}']
        _testo(riga.cells[4], classe, dim=11, grassetto=True,
               colore={'A': C_VERDE, 'B': C_NAVY, 'C': C_ROSSO}[classe],
               allineamento=WD_ALIGN_PARAGRAPH.CENTER, spazio_dopo=3, primo=True)

    _sottotitolo(doc, 'Le combinazioni più frequenti')
    t = doc.add_table(rows=1, cols=3)
    _senza_bordi(t)
    for i, testo in enumerate(['Rating Combinato', 'Imprese', 'Quota']):
        cella = t.rows[0].cells[i]
        cella.width = [Cm(6.0), Cm(5.5), Cm(5.5)][i]
        _testo(cella, testo, dim=8, grassetto=True,
               allineamento=None if i == 0 else WD_ALIGN_PARAGRAPH.RIGHT, spazio_dopo=3, primo=True)
        _linea_sotto(cella, NAVY, '8')
    prime = q['comb_conteggi'].head(6)
    if q['rating'] not in prime.index and q['comb_target']:
        prime = pd.concat([prime, pd.Series({q['rating']: q['comb_target']})])
    for n, (combinazione, conteggio) in enumerate(prime.items()):
        riga = t.add_row()
        evidenzia = (combinazione == q['rating'])
        for i, larghezza in enumerate([Cm(6.0), Cm(5.5), Cm(5.5)]):
            riga.cells[i].width = larghezza
            if evidenzia:
                _sfondo_cella(riga.cells[i], 'E4ECF5')
            elif n % 2 == 1:
                _sfondo_cella(riga.cells[i], RIGA_ALT)
        _testo(riga.cells[0], str(combinazione), dim=9, grassetto=True,
               colore=C_CORALLO if evidenzia else C_NAVY, spazio_dopo=3, primo=True)
        _testo(riga.cells[1], format_euro(conteggio, 0), dim=8.5,
               allineamento=WD_ALIGN_PARAGRAPH.RIGHT, spazio_dopo=3, primo=True)
        _testo(riga.cells[2], f"{format_euro(conteggio / q['panel'] * 100)}%", dim=8.5,
               colore=C_GRIGIO, allineamento=WD_ALIGN_PARAGRAPH.RIGHT, spazio_dopo=3, primo=True)
    _testo(doc, "La riga evidenziata è la combinazione dell'impresa.", dim=7.5,
           colore=C_GRIGIO, spazio_dopo=0)


APERTURE_AREA = {
    'Economica': "L'area economica misura quanto del valore prodotto resta all'impresa lungo il "
                 "conto economico: dal margine operativo lordo al risultato ante imposte.",
    'Patrimoniale': "L'area patrimoniale guarda alla correlazione fra investimenti durevoli e "
                    "fonti che li finanziano, e al peso del debito rispetto al capitale proprio.",
    'Finanziaria': "L'area finanziaria verifica la capacità di far fronte agli impegni entro "
                   "l'anno e l'efficienza con cui il capitale investito si traduce in produzione.",
}


def _scheda_indicatore(doc, q, chiave):
    """Blocco di un singolo indicatore: numeri a sinistra, serie a destra, commento sotto."""
    d = q['ind'][chiave]
    unita = d['unita']
    colore_classe = {'A': C_VERDE, 'B': C_NAVY, 'C': C_ROSSO}[d['classe']]

    t = doc.add_table(rows=1, cols=2)
    _senza_bordi(t)
    sx, dx = t.rows[0].cells
    sx.width, dx.width = Cm(6.4), Cm(10.6)
    _sfondo_cella(sx, RIGA_ALT)

    _testo(sx, d['nome'].upper(), dim=8, grassetto=True, colore=C_GRIGIO, spazio_dopo=2, primo=True)
    _testo(sx, 'n.d.' if d['valore'] is None else f"{format_euro(d['valore'])}{unita}",
           dim=20, grassetto=True, colore=colore_classe, spazio_dopo=2)
    med = d['mediana']
    _testo(sx, f"mediana di settore {format_euro(med)}{unita}" if med is not None
           else 'mediana di settore n.d.', dim=8, colore=C_GRIGIO, spazio_dopo=2)
    _testo(sx, f"Classe {d['classe']}", dim=10, grassetto=True, colore=colore_classe, spazio_dopo=6)

    _testo(dx, A_COSA_SERVE[chiave], dim=8, colore=C_GRIGIO, spazio_dopo=4, primo=True)
    interna = dx.add_table(rows=2, cols=5)
    _senza_bordi(interna)
    for j, anno in enumerate(['Anno'] + list(ANNI_SERIE)):
        cella = interna.rows[0].cells[j]
        cella.width = Cm(2.0)
        _testo(cella, anno, dim=7.5, grassetto=True, colore=C_GRIGIO,
               allineamento=WD_ALIGN_PARAGRAPH.CENTER, spazio_dopo=1, primo=True)
        _linea_sotto(cella, 'D9D9D9', '4')
    _testo(interna.rows[1].cells[0], 'Impresa', dim=7.5, grassetto=True, colore=C_GRIGIO,
           allineamento=WD_ALIGN_PARAGRAPH.CENTER, spazio_dopo=2, primo=True)
    for j, anno in enumerate(ANNI_SERIE, start=1):
        v = d['az'].get(anno)
        cella = interna.rows[1].cells[j]
        cella.width = Cm(2.0)
        _testo(cella, 'n.d.' if v is None else f"{format_euro(v)}{unita}", dim=8.5,
               grassetto=(anno == '2024'), allineamento=WD_ALIGN_PARAGRAPH.CENTER,
               spazio_dopo=2, primo=True)
    riga_set = interna.add_row()
    _testo(riga_set.cells[0], 'Settore', dim=7.5, grassetto=True, colore=C_GRIGIO,
           allineamento=WD_ALIGN_PARAGRAPH.CENTER, spazio_dopo=2, primo=True)
    for j, anno in enumerate(ANNI_SERIE, start=1):
        v = d['set'].get(anno)
        cella = riga_set.cells[j]
        cella.width = Cm(2.0)
        _testo(cella, 'n.d.' if v is None else f"{format_euro(v)}{unita}", dim=8.5,
               colore=C_GRIGIO, allineamento=WD_ALIGN_PARAGRAPH.CENTER, spazio_dopo=2, primo=True)
    # Word vuole un paragrafo dopo una tabella annidata
    _testo(dx, '', dim=4, spazio_dopo=0)

    _capoverso(doc, d['commento'], dim=8.5, spazio_dopo=3)

    coda = []
    pos, totale = q['rank'][chiave]
    if pos:
        coda.append(f"{posizione_ordinale(pos)} posto su {format_euro(totale, 0)} imprese del panel")
    pos_r, totale_r = q['rank_reg'][chiave]
    if pos_r and q['regione']:
        coda.append(f"{posizione_ordinale(pos_r)} su {format_euro(totale_r, 0)} "
                    f"{preposizione_regione(q['regione'])}")
    if d['var'] is not None:
        verso = 'migliora' if ((d['var'] < 0) if d['inverso'] else (d['var'] > 0)) else 'peggiora'
        coda.append(f"dal 2021 {verso} {_perc(abs(d['var']), 'di')}")
    if coda:
        _testo(doc, 'Classifica:  ' + '  ·  '.join(coda), dim=7.5, colore=C_GRIGIO, spazio_dopo=10)


def _sezione_area(doc, q, area):
    _titolo_sezione(doc, f'Area {area}', nuova_pagina=True)
    classe = q[f'rating_{area}']
    p = doc.add_paragraph()
    p.paragraph_format.space_after = Pt(6)
    r = p.add_run(f'Classe {classe}')
    r.font.size = Pt(12)
    r.bold = True
    r.font.color.rgb = {'A': C_VERDE, 'B': C_NAVY, 'C': C_ROSSO}[classe]

    _capoverso(doc, APERTURE_AREA[area], dim=8.5, colore=C_GRIGIO, spazio_dopo=8)
    for chiave in CHIAVI_AREA[area]:
        if q['ind'][chiave]['valore'] is None:
            continue
        _scheda_indicatore(doc, q, chiave)

    _sottotitolo(doc, 'In sintesi')
    _capoverso(doc, _sintesi_area(q, area), dim=8.5, spazio_dopo=2)
    # La lettera nasce dai terzili: un indicatore puo' stare sotto la mediana e
    # restare comunque nel terzo centrale della distribuzione.
    _testo(doc, 'La classe dipende dai terzili della distribuzione, non dal solo confronto con '
                'la mediana.', dim=7.5, colore=C_GRIGIO, spazio_dopo=0)


def _sintesi_area(q, area):
    nomi_meglio = [q['ind'][k]['nome'] for k in CHIAVI_AREA[area] if _meglio(q['ind'][k])]
    nomi_peggio = [q['ind'][k]['nome'] for k in CHIAVI_AREA[area] if not _meglio(q['ind'][k])]
    classe = q[f'rating_{area}']
    pezzi = []
    if nomi_meglio:
        verbo = 'fa' if len(nomi_meglio) == 1 else 'fanno'
        pezzi.append(f"{_elenco(nomi_meglio)} {verbo} meglio della mediana di settore")
    if nomi_peggio:
        verbo = 'resta' if len(nomi_peggio) == 1 else 'restano'
        pezzi.append(f"{_elenco(nomi_peggio)} {verbo} al di sotto del riferimento")
    testo = f"Nel 2024 l'area {area.lower()} appartiene alla classe {classe}"
    if pezzi:
        testo += ': ' + ', mentre '.join(pezzi) if len(pezzi) == 2 else ': ' + pezzi[0]
    testo += '.'
    migliorati = [q['ind'][k]['nome'] for k in CHIAVI_AREA[area]
                  if q['ind'][k]['var'] is not None
                  and ((q['ind'][k]['var'] < 0) if q['ind'][k]['inverso'] else (q['ind'][k]['var'] > 0))]
    if migliorati and len(migliorati) < 3:
        testo += f" Nel quadriennio la direzione è favorevole su {_elenco(migliorati)}."
    elif len(migliorati) == 3:
        testo += " Nel quadriennio tutti e tre gli indicatori si muovono nella direzione giusta."
    elif not migliorati:
        testo += " Nel quadriennio nessuno dei tre indicatori si muove nella direzione favorevole."
    return testo


def _classifiche(doc, q):
    _titolo_sezione(doc, 'Classifiche', nuova_pagina=True)
    _capoverso(doc, "Per ogni indicatore l'impresa viene ordinata insieme a tutte le altre del "
                    "panel: il primo posto va al risultato migliore, che per il Gearing è il "
                    "valore più basso. Accanto, la stessa classifica ristretta alle imprese della "
                    "regione di sede.")
    colonne = ['Indicatore', 'Valore 2024', 'Italia', 'Meglio di', q['regione'] or 'Regione']
    larghezze = [Cm(5.0), Cm(2.6), Cm(3.4), Cm(2.6), Cm(3.4)]
    t = doc.add_table(rows=1, cols=len(colonne))
    _senza_bordi(t)
    for i, testo in enumerate(colonne):
        cella = t.rows[0].cells[i]
        cella.width = larghezze[i]
        _testo(cella, testo, dim=8, grassetto=True,
               allineamento=None if i == 0 else WD_ALIGN_PARAGRAPH.RIGHT, spazio_dopo=3, primo=True)
        _linea_sotto(cella, NAVY, '8')
    for n, (chiave, _base, nome, unita, _soglia, _inverso, _area) in enumerate(INDICATORI):
        d = q['ind'][chiave]
        pos, totale = q['rank'][chiave]
        pos_r, totale_r = q['rank_reg'][chiave]
        riga = t.add_row()
        for i, larghezza in enumerate(larghezze):
            riga.cells[i].width = larghezza
            if n % 2 == 1:
                _sfondo_cella(riga.cells[i], RIGA_ALT)
        _testo(riga.cells[0], nome, dim=8.5, grassetto=True, spazio_dopo=3, primo=True)
        _testo(riga.cells[1], 'n.d.' if d['valore'] is None else f"{format_euro(d['valore'])}{unita}",
               dim=8.5, allineamento=WD_ALIGN_PARAGRAPH.RIGHT, spazio_dopo=3, primo=True)
        _testo(riga.cells[2], 'n.d.' if not pos else f"{posizione_ordinale(pos)} su {format_euro(totale, 0)}",
               dim=8.5, allineamento=WD_ALIGN_PARAGRAPH.RIGHT, spazio_dopo=3, primo=True)
        if pos and totale:
            percentile = (1 - pos / totale) * 100
            _testo(riga.cells[3], f"{format_euro(percentile, 0)}%", dim=8.5, colore=C_GRIGIO,
                   allineamento=WD_ALIGN_PARAGRAPH.RIGHT, spazio_dopo=3, primo=True)
        else:
            _testo(riga.cells[3], 'n.d.', dim=8.5, colore=C_GRIGIO,
                   allineamento=WD_ALIGN_PARAGRAPH.RIGHT, spazio_dopo=3, primo=True)
        _testo(riga.cells[4], 'n.d.' if not pos_r else f"{posizione_ordinale(pos_r)} su {format_euro(totale_r, 0)}",
               dim=8.5, colore=C_GRIGIO, allineamento=WD_ALIGN_PARAGRAPH.RIGHT, spazio_dopo=3, primo=True)
    _testo(doc, '"Meglio di" è la quota di imprese del panel che restano dietro: il primo posto '
                'corrisponde al 100%.', dim=7.5, colore=C_GRIGIO, spazio_dopo=0)


def _sintesi(doc, q):
    _titolo_sezione(doc, 'Posizionamento di sintesi', nuova_pagina=True)
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
    if q['comb_target']:
        voci.append(f"{format_euro(q['comb_target'], 0)} imprese del panel condividono lo stesso "
                    f"Rating Combinato")
    primo = True
    for voce in voci:
        _testo(dx, f'›  {voce}', dim=9, spazio_dopo=3, primo=primo)
        primo = False

    _sottotitolo(doc, 'Lettura complessiva')
    _capoverso(doc, _frase_rating(q), dim=9)

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

    _sottotitolo(doc, 'Punti di forza e di attenzione')
    t = doc.add_table(rows=1, cols=2)
    _senza_bordi(t)
    forza, attenzione = t.rows[0].cells
    forza.width = attenzione.width = Cm(8.5)
    _sfondo_cella(forza, 'EAF6F4')
    _sfondo_cella(attenzione, 'FBEDEA')
    _testo(forza, 'FANNO MEGLIO DEL SETTORE', dim=8, grassetto=True, colore=C_VERDE, spazio_dopo=3, primo=True)
    _testo(attenzione, 'SOTTO IL RIFERIMENTO', dim=8, grassetto=True, colore=C_ROSSO, spazio_dopo=3, primo=True)
    sopra = [i[0] for i in INDICATORI if _meglio(q['ind'][i[0]])]
    sotto = [i[0] for i in INDICATORI if not _meglio(q['ind'][i[0]])]
    for chiave in sopra:
        d = q['ind'][chiave]
        _testo(forza, f"·  {d['nome']}  {format_euro(d['valore'])}{d['unita']} "
                      f"(settore {format_euro(d['mediana'])}{d['unita']})", dim=8, spazio_dopo=2)
    if not sopra:
        _testo(forza, '·  Nessun indicatore sopra il riferimento di settore.', dim=8, spazio_dopo=2)
    for chiave in sotto:
        d = q['ind'][chiave]
        if d['valore'] is None or d['mediana'] is None:
            continue
        _testo(attenzione, f"·  {d['nome']}  {format_euro(d['valore'])}{d['unita']} "
                           f"(settore {format_euro(d['mediana'])}{d['unita']})", dim=8, spazio_dopo=2)
    if not sotto:
        _testo(attenzione, '·  Tutti gli indicatori raggiungono il riferimento.', dim=8, spazio_dopo=2)
    _testo(forza, '', dim=4, spazio_dopo=0)
    _testo(attenzione, '', dim=4, spazio_dopo=0)


def _note(doc, q, settore):
    _titolo_sezione(doc, 'Note metodologiche e glossario', nuova_pagina=True)
    _capoverso(doc,
               f"Il panel comprende {format_euro(q['panel'], 0)} imprese del settore {settore}, "
               f"estratte da ORBIS e selezionate in modo da avere i dati disponibili su tutti e "
               f"quattro gli esercizi 2021-2024.")
    _capoverso(doc,
               "Il settore è descritto dalla mediana e non dalla media, perché la mediana non "
               "viene spostata dai pochi valori estremi presenti nel campione. Sulle mediane è "
               "calcolata la variazione anno su anno, che mostra la direzione del comparto nel "
               "quadriennio.")
    _capoverso(doc,
               "Le classi nascono dai terzili della distribuzione 2024: A per la fascia superiore, "
               "B per quella intermedia, C per quella inferiore. Per il Gearing l'ordine si "
               "inverte, perché un valore più contenuto segnala meno debito. Il 2024 determina il "
               "posizionamento, mentre il quadriennio ne mostra la direzione.")
    _capoverso(doc,
               "I numeri, i terzili e le classi sono gli stessi del report esteso: questo "
               "documento ne riprende i risultati in forma più breve, senza i grafici e le tabelle "
               "di dettaglio.")

    _sottotitolo(doc, 'Glossario delle formule')
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
    _copertina(doc, str(azienda_target), settore_nace, q)

    corpo = doc.add_section(WD_SECTION.NEW_PAGE)
    corpo.top_margin = Cm(2.4)
    corpo.bottom_margin = Cm(1.6)
    corpo.left_margin = corpo.right_margin = Cm(2.0)
    corpo.header_distance = Cm(0.8)
    corpo.header.is_linked_to_previous = False
    _banda_testata(corpo, str(azienda_target), settore_nace)

    _indice(doc)
    _profilo(doc, q, str(azienda_target), settore_nace)
    _dimensione(doc, q)
    _big_picture(doc, q)
    _dati_sintesi(doc, q)
    _conto_economico(doc, q)
    _settore_in_numeri(doc, q, settore_nace)
    _come_nascono_le_classi(doc, q)
    _distribuzione_panel(doc, q)
    for area in AREE:
        _sezione_area(doc, q, area)
    _classifiche(doc, q)
    _sintesi(doc, q)
    _note(doc, q, settore_nace)

    # Word non accetta una tabella come ultimo elemento del corpo
    doc.add_paragraph()

    uscita = io.BytesIO()
    doc.save(uscita)
    uscita.seek(0)
    return uscita
