# -*- coding: utf-8 -*-
"""
Identificazione univoca dell'azienda target.

Il campione ORBIS contiene sistematicamente decine di ragioni sociali che si
contengono a vicenda (nel NACE 41.20 ci sono 4 "COSTRUZIONI GENERALI S.R.L."
piu' una "SISTEM COSTRUZIONI GENERALI S.R.L." piu' oltre 180 nomi che contengono
quella stessa frase). Cercare l'azienda per nome — anche solo parziale — significa
agganciare la prima riga che capita, che quasi mai e' quella richiesta.

Qui l'azienda viene identificata SEMPRE tramite una chiave univoca (Numero BvD ID,
in subordine P.IVA/Codice Fiscale) e mai tramite la ragione sociale, che resta solo
un'etichetta da stampare.
"""

import pandas as pd


def trova_colonne_identita(df):
    """Restituisce (col_ragione_sociale, col_piva, col_bvd) del dataframe ORBIS."""
    col_ragione = next((c for c in df.columns if 'ragione' in str(c).lower()), None)
    col_piva = next(
        (c for c in df.columns
         if 'partita iva' in str(c).lower() or 'codice fiscale' in str(c).lower()),
        None
    )
    col_bvd = next((c for c in df.columns if 'bvd' in str(c).lower()), None)
    return col_ragione, col_piva, col_bvd


def normalizza_bvd(valore):
    """Il BvD ID e' gia' una stringa (es. 'IT02359130602'): serve solo normalizzarla."""
    if valore is None or (isinstance(valore, float) and pd.isna(valore)):
        return None
    testo = str(valore).strip().upper()
    if testo in ('', 'N.D.', 'NAN', 'NONE'):
        return None
    return testo


def normalizza_piva(valore):
    """
    Excel/Pandas legge la P.IVA come numero e ne perde gli zeri iniziali
    (00380570166 diventa 380570166.0): la ricostruiamo a 11 cifre.
    I Codici Fiscali alfanumerici (16 caratteri) restano invariati.
    """
    if valore is None or (isinstance(valore, float) and pd.isna(valore)):
        return None
    testo = str(valore).strip()
    if testo.lower() in ('', 'n.d.', 'nan', 'none'):
        return None
    try:
        return str(int(float(testo))).zfill(11)
    except ValueError:
        return testo.upper()


def chiave_da_riga(riga, df=None):
    """
    Costruisce la chiave univoca dell'azienda a partire da una riga del dataframe.
    Ritorna un dict con ragione_sociale / piva / bvd (i campi assenti restano None).
    """
    if df is None:
        col_ragione, col_piva, col_bvd = trova_colonne_identita(pd.DataFrame([riga]))
    else:
        col_ragione, col_piva, col_bvd = trova_colonne_identita(df)
    return {
        'ragione_sociale': str(riga[col_ragione]).strip() if col_ragione else None,
        'piva': normalizza_piva(riga[col_piva]) if col_piva else None,
        'bvd': normalizza_bvd(riga[col_bvd]) if col_bvd else None,
    }


def normalizza_chiave(chiave, azienda_target=None):
    """
    Accetta una chiave in qualunque forma (dict, stringa di P.IVA/BvD, None) e la
    riporta al dict canonico. Se manca del tutto, si ripiega sulla sola ragione
    sociale: un confronto ESATTO, mai un 'contains'.
    """
    if isinstance(chiave, dict):
        out = {
            'ragione_sociale': chiave.get('ragione_sociale'),
            'piva': normalizza_piva(chiave.get('piva')),
            'bvd': normalizza_bvd(chiave.get('bvd')),
        }
    elif chiave:
        testo = str(chiave).strip()
        out = {'ragione_sociale': None, 'piva': None, 'bvd': None}
        if testo.upper().startswith('IT') or any(c.isalpha() for c in testo):
            out['bvd'] = normalizza_bvd(testo)
        else:
            out['piva'] = normalizza_piva(testo)
    else:
        out = {'ragione_sociale': None, 'piva': None, 'bvd': None}

    if not out['ragione_sociale'] and azienda_target:
        out['ragione_sociale'] = str(azienda_target).strip()
    return out


def maschera_target(df, chiave, azienda_target=None):
    """
    Maschera booleana che seleziona la riga (o le righe) dell'azienda target.
    Ordine di precedenza: BvD ID esatto -> P.IVA esatta -> ragione sociale esatta.
    Non viene mai usato un confronto parziale sul nome.
    """
    chiave = normalizza_chiave(chiave, azienda_target)
    col_ragione, col_piva, col_bvd = trova_colonne_identita(df)

    if chiave['bvd'] and col_bvd is not None:
        m = df[col_bvd].map(normalizza_bvd) == chiave['bvd']
        if m.any():
            return m

    if chiave['piva'] and col_piva is not None:
        m = df[col_piva].map(normalizza_piva) == chiave['piva']
        if m.any():
            return m

    if chiave['ragione_sociale'] and col_ragione is not None:
        nome = chiave['ragione_sociale'].strip().lower()
        return df[col_ragione].astype(str).str.strip().str.lower() == nome

    return pd.Series(False, index=df.index)


def riga_target(df, chiave, azienda_target=None):
    """Sotto-dataframe con la sola riga dell'azienda target (vuoto se non trovata)."""
    return df[maschera_target(df, chiave, azienda_target)]


def risolvi_ricerca(df, testo_ricerca):
    """
    Ricerca manuale dell'utente (ragione sociale, P.IVA o BvD ID).

    Ritorna (indice_riga, chiave, n_omonime) oppure (None, None, 0) se non trova nulla.
    La priorita' e' sulle chiavi univoche: solo se la stringa non corrisponde a nessun
    BvD ID / P.IVA si passa al nome, dove il match esatto viene comunque preferito a
    quello parziale, per non agganciare l'omonima sbagliata.
    """
    testo = str(testo_ricerca or '').strip()
    if not testo:
        return None, None, 0

    col_ragione, col_piva, col_bvd = trova_colonne_identita(df)

    candidati = None
    if col_bvd is not None:
        bvd = normalizza_bvd(testo)
        m = df[col_bvd].map(normalizza_bvd) == bvd
        if m.any():
            candidati = df[m]

    if candidati is None and col_piva is not None:
        piva = normalizza_piva(testo)
        if piva:
            m = df[col_piva].map(normalizza_piva) == piva
            if m.any():
                candidati = df[m]

    if candidati is None and col_ragione is not None:
        nomi = df[col_ragione].astype(str).str.strip().str.lower()
        m_esatto = nomi == testo.lower()
        if m_esatto.any():
            candidati = df[m_esatto]
        else:
            m_parziale = nomi.str.contains(testo.lower(), regex=False, na=False)
            if m_parziale.any():
                candidati = df[m_parziale]

    if candidati is None or candidati.empty:
        return None, None, 0

    idx = candidati.index[0]
    return idx, chiave_da_riga(df.loc[idx], df), len(candidati)


def etichetta_azienda(riga, df=None):
    """Etichetta leggibile 'RAGIONE SOCIALE (P.IVA/CF ..., BvD ID ...)' per i menu."""
    chiave = chiave_da_riga(riga, df)
    dettagli = []
    if chiave['piva']:
        dettagli.append(f"P.IVA/CF {chiave['piva']}")
    if chiave['bvd']:
        dettagli.append(f"BvD ID {chiave['bvd']}")
    nome = chiave['ragione_sociale'] or 'N.D.'
    return f"{nome} ({', '.join(dettagli)})" if dettagli else nome
